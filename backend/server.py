from fastapi import FastAPI, APIRouter, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from dotenv import load_dotenv
from starlette.middleware.cors import CORSMiddleware
from motor.motor_asyncio import AsyncIOMotorClient
import os
import io
import logging
import uuid
from pathlib import Path
from pydantic import BaseModel, Field, ConfigDict
from typing import List, Optional
from datetime import datetime, timezone
from difflib import SequenceMatcher
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from college_data import COLLEGES

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

mongo_url = os.environ['MONGO_URL']
client = AsyncIOMotorClient(mongo_url)
db = client[os.environ['DB_NAME']]

app = FastAPI()
api_router = APIRouter(prefix="/api")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Models ───────────────────────────────────────────────────────────────────
class StatusCheck(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    client_name: str
    timestamp: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class StatusCheckCreate(BaseModel):
    client_name: str

# ── Helpers ──────────────────────────────────────────────────────────────────
def normalize(s: str) -> str:
    return s.lower().strip() if s else ""

def match_college(name: str, colleges: list) -> Optional[dict]:
    """Match college name (full or short) against DB. Returns college doc or None."""
    if not name or not name.strip():
        return None
    n = normalize(name)
    # 1. Exact match on full name
    for c in colleges:
        if normalize(c["college_name"]) == n:
            return c
    # 2. Exact match on any short name
    for c in colleges:
        for s in c.get("short_names", []):
            if normalize(s) == n:
                return c
    # 3. Fuzzy match (threshold 0.82)
    best, best_score = None, 0
    for c in colleges:
        candidates = [c["college_name"]] + c.get("short_names", [])
        for cand in candidates:
            score = SequenceMatcher(None, n, normalize(cand)).ratio()
            if score > best_score:
                best_score = score
                best = c
    if best_score >= 0.82:
        return best
    return None

def rank_display(c: Optional[dict]) -> str:
    if c is None:
        return ""
    return str(c["rank"]) if c.get("rank") is not None else "null"

def excel_thin_border():
    t = Side(style="thin", color="D1D5DB")
    return Border(left=t, right=t, top=t, bottom=t)

# ── Basic Routes ─────────────────────────────────────────────────────────────
@api_router.get("/")
async def root():
    return {"message": "Hello World"}

@api_router.post("/status", response_model=StatusCheck)
async def create_status_check(input: StatusCheckCreate):
    status_dict = input.model_dump()
    status_obj = StatusCheck(**status_dict)
    doc = status_obj.model_dump()
    doc['timestamp'] = doc['timestamp'].isoformat()
    await db.status_checks.insert_one(doc)
    return status_obj

@api_router.get("/status", response_model=List[StatusCheck])
async def get_status_checks():
    status_checks = await db.status_checks.find({}, {"_id": 0}).to_list(1000)
    for check in status_checks:
        if isinstance(check['timestamp'], str):
            check['timestamp'] = datetime.fromisoformat(check['timestamp'])
    return status_checks

# ── College DB Routes ─────────────────────────────────────────────────────────
@api_router.post("/colleges/seed")
async def seed_colleges():
    """Seed all 300 colleges into MongoDB."""
    await db.colleges.drop()
    await db.colleges.create_index("college_name")
    result = await db.colleges.insert_many(COLLEGES)
    return {"seeded": len(result.inserted_ids), "message": "College database ready"}

@api_router.get("/colleges")
async def get_colleges(skip: int = 0, limit: int = 300):
    colleges = await db.colleges.find({}, {"_id": 0}).skip(skip).limit(limit).to_list(limit)
    total = await db.colleges.count_documents({})
    return {"colleges": colleges, "total": total}

@api_router.get("/colleges/stats")
async def college_stats():
    total = await db.colleges.count_documents({})
    ranked = await db.colleges.count_documents({"rank": {"$ne": None}})
    return {"total": total, "ranked": ranked, "unranked": total - ranked}

# ── Excel Processing Route ────────────────────────────────────────────────────
@api_router.post("/process-excel")
async def process_excel(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "Please upload an Excel file (.xlsx or .xls)")

    # Load colleges from DB
    colleges = await db.colleges.find({}, {"_id": 0}).to_list(500)
    if not colleges:
        colleges = COLLEGES  # fallback to in-memory

    # Read uploaded Excel
    content = await file.read()
    wb_in = openpyxl.load_workbook(io.BytesIO(content))
    ws_in = wb_in.active

    # Find column indices (1-based)
    headers = [ws_in.cell(row=1, column=c).value for c in range(1, ws_in.max_column + 1)]
    headers_lower = [str(h).lower().strip() if h else "" for h in headers]

    def find_col(keywords):
        for kw in keywords:
            for i, h in enumerate(headers_lower):
                if kw in h:
                    return i + 1  # 1-based
        return None

    ug_col  = find_col(["ug university", "ug uni"])
    ug_rank_col = find_col(["rank ug", "ug rank"])
    pg_col  = find_col(["pg university", "pg uni"])
    pg_rank_col = find_col(["rank pg", "pg rank"])

    # ── Build output workbook ──────────────────────────────────────────────
    wb_out = openpyxl.Workbook()

    # ─ Sheet 1: Original data + ranks filled ──────────────────────────────
    ws1 = wb_out.active
    ws1.title = "Applicant Data with Ranks"

    hdr_fill  = PatternFill("solid", fgColor="1A73E8")
    hdr_font  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    data_font = Font(name="Calibri", size=10)
    rank_font = Font(name="Calibri", bold=True, color="1557B0", size=10)
    null_font = Font(name="Calibri", italic=True, color="9CA3AF", size=10)
    alt_fill  = PatternFill("solid", fgColor="F0F7FF")

    # Copy headers
    for c, h in enumerate(headers, 1):
        cell = ws1.cell(row=1, column=c, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = excel_thin_border()
    ws1.row_dimensions[1].height = 28

    # Track matched colleges for Tab 2
    matched_pairs = {}  # {college_name: {"rank": ..., "short_names": [...]}}

    # Copy rows + fill ranks
    for row_idx in range(2, ws_in.max_row + 1):
        is_alt = (row_idx % 2 == 0)
        row_fill = alt_fill if is_alt else None

        for col_idx in range(1, ws_in.max_column + 1):
            val = ws_in.cell(row=row_idx, column=col_idx).value
            cell = ws1.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            if row_fill:
                cell.fill = row_fill
            cell.border = excel_thin_border()
            cell.alignment = Alignment(vertical="center", wrap_text=False)

        ws1.row_dimensions[row_idx].height = 16

        # Fill UG rank
        if ug_col:
            ug_name = ws_in.cell(row=row_idx, column=ug_col).value
            if ug_name and str(ug_name).strip():
                matched = match_college(str(ug_name), colleges)
                rank_val = rank_display(matched)
                if rank_val:
                    r_cell = ws1.cell(row=row_idx, column=ug_rank_col) if ug_rank_col else None
                    if r_cell:
                        r_cell.value = rank_val
                        r_cell.font = null_font if rank_val == "null" else rank_font
                        r_cell.alignment = Alignment(horizontal="center", vertical="center")
                    if matched:
                        matched_pairs[matched["college_name"]] = matched

        # Fill PG rank
        if pg_col:
            pg_name = ws_in.cell(row=row_idx, column=pg_col).value
            if pg_name and str(pg_name).strip():
                matched = match_college(str(pg_name), colleges)
                rank_val = rank_display(matched)
                if rank_val:
                    r_cell = ws1.cell(row=row_idx, column=pg_rank_col) if pg_rank_col else None
                    if r_cell:
                        r_cell.value = rank_val
                        r_cell.font = null_font if rank_val == "null" else rank_font
                        r_cell.alignment = Alignment(horizontal="center", vertical="center")
                    if matched:
                        matched_pairs[matched["college_name"]] = matched

    # Auto-width for first few key columns
    for col in ws1.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws1.column_dimensions[col[0].column_letter].width = min(max_len + 3, 40)

    ws1.freeze_panes = "A2"

    # ─ Sheet 2: Ranking Reference ──────────────────────────────────────────
    ws2 = wb_out.create_sheet("Ranking Reference")

    # Title
    ws2.merge_cells("A1:B1")
    t = ws2["A1"]
    t.value = "College Ranking Reference — Matched from Job Post Excel"
    t.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor="0F4C81")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 32

    # Headers
    for c, h in enumerate(["UG University/institute Name", "Rank"], 1):
        cell = ws2.cell(row=2, column=c, value=h)
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill("solid", fgColor="1A73E8")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = excel_thin_border()
    ws2.row_dimensions[2].height = 24
    ws2.column_dimensions["A"].width = 65
    ws2.column_dimensions["B"].width = 12

    # Data rows: full name + short name(s) per college
    row_num = 3
    green_fill  = PatternFill("solid", fgColor="E8F5E9")
    green_fill2 = PatternFill("solid", fgColor="F1FAF1")

    sorted_colleges = sorted(matched_pairs.values(),
                              key=lambda x: (x["rank"] is None, x["rank"] or 9999))

    for i, college in enumerate(sorted_colleges):
        grp_fill = green_fill if i % 2 == 0 else green_fill2
        rank_str = str(college["rank"]) if college.get("rank") is not None else "null"

        # Full name row
        cell_name = ws2.cell(row=row_num, column=1, value=college["college_name"])
        cell_name.font = Font(name="Calibri", bold=True, size=10, color="111827")
        cell_name.fill = grp_fill
        cell_name.border = excel_thin_border()
        cell_name.alignment = Alignment(horizontal="center", vertical="center")

        cell_rank = ws2.cell(row=row_num, column=2, value=rank_str)
        cell_rank.font = Font(name="Calibri", bold=True, size=10,
                              color="1A73E8" if rank_str != "null" else "9CA3AF")
        cell_rank.fill = grp_fill
        cell_rank.border = excel_thin_border()
        cell_rank.alignment = Alignment(horizontal="center", vertical="center")
        ws2.row_dimensions[row_num].height = 18
        row_num += 1

        # Short name rows (one per short name)
        for sname in college.get("short_names", []):
            cell_sn = ws2.cell(row=row_num, column=1, value=sname)
            cell_sn.font = Font(name="Calibri", italic=True, size=10, color="374151")
            cell_sn.fill = grp_fill
            cell_sn.border = excel_thin_border()
            cell_sn.alignment = Alignment(horizontal="center", vertical="center")

            cell_sr = ws2.cell(row=row_num, column=2, value=rank_str)
            cell_sr.font = Font(name="Calibri", size=10,
                                color="1A73E8" if rank_str != "null" else "9CA3AF")
            cell_sr.fill = grp_fill
            cell_sr.border = excel_thin_border()
            cell_sr.alignment = Alignment(horizontal="center", vertical="center")
            ws2.row_dimensions[row_num].height = 18
            row_num += 1

    ws2.freeze_panes = "A3"

    # ── Stream back as download ──────────────────────────────────────────────
    out_buffer = io.BytesIO()
    wb_out.save(out_buffer)
    out_buffer.seek(0)

    return StreamingResponse(
        out_buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="BluBridge_Processed_{file.filename}"'}
    )

# ── Static Excel Downloads ────────────────────────────────────────────────────
@api_router.get("/download-college-list")
async def download_college_list():
    file_path = Path("/app/college_rankings.xlsx")
    if not file_path.exists():
        raise HTTPException(404, "File not found")
    return FileResponse(str(file_path),
        filename="BluBridge_College_Rankings_NIRF_2025.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@api_router.get("/download-college-list-shortnames")
async def download_college_list_shortnames():
    file_path = Path("/app/college_rankings_with_shortnames.xlsx")
    if not file_path.exists():
        raise HTTPException(404, "File not found")
    return FileResponse(str(file_path),
        filename="BluBridge_College_Rankings_ShortNames_NIRF_2025.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── App Setup ─────────────────────────────────────────────────────────────────
app.include_router(api_router)
app.add_middleware(
    CORSMiddleware,
    allow_credentials=True,
    allow_origins=os.environ.get('CORS_ORIGINS', '*').split(','),
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.on_event("shutdown")
async def shutdown_db_client():
    client.close()
