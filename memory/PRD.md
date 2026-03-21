# BluBridge - College Ranking Processor

## Original Problem Statement
Build a college ranking processor application with:
1. A premium registration page
2. A homepage with Upload and Download options for Excel processing
3. Store ~300 colleges (rank, full name, short names, city, state) in a database
4. Process uploaded "Job Post appraise" Excel sheets by matching college names against the database and adding NIRF ranks
5. Center the logo in the global header and remove Register/Login/Logout buttons
6. Replace null ranks with rank bands (RB:101-150, RB:151-200, RB:201-300); show "NL" for not-listed colleges

## Tech Stack
- **Frontend**: React, Tailwind CSS, Shadcn/UI
- **Backend**: FastAPI, Pandas, Openpyxl
- **Database**: MongoDB

## Architecture
```
/app/
├── backend/
│   ├── .env
│   ├── college_data.py       # 300 colleges with ranks + rank bands
│   ├── requirements.txt
│   ├── server.py
│   └── tests/test_rank_bands.py
├── frontend/
│   ├── .env
│   └── src/
│       ├── App.js
│       └── components/
│           ├── BluBridgeHome.jsx
│           └── RegisterPremium.jsx
├── college_rankings.xlsx
└── college_rankings_with_shortnames.xlsx
```

## Key API Endpoints
- `POST /api/process-excel` - Upload and process Excel with college rank matching
- `GET /api/colleges/stats` - Stats: 300 total, 100 ranked, 200 rank band
- `GET /api/download-college-list` - Download NIRF rankings (with rank bands)
- `GET /api/download-college-list-shortnames` - Download NIRF rankings with short names (with rank bands)
- `POST /api/colleges/seed` - Seed college database

## DB Schema
- **colleges**: `{ name: str, short_names: List[str], city: str, state: str, rank: Union[int, str] }`
  - rank: `int` (1-100) for top colleges, `"RB:101-150"` / `"RB:151-200"` / `"RB:201-300"` for band colleges

## Rank System
- **1-100**: Specific NIRF rank (integer)
- **101-150**: Rank band 101-150 (50 colleges)
- **151-200**: Rank band 151-200 (52 colleges)
- **201-300**: Rank band 201-300 (99 colleges)
- **NL**: Not Listed — shown in processed Excel for colleges not in the database

## Completed Features
- [x] Premium Registration Page (`/register-premium`)
- [x] College data seeding (300 colleges with rank bands)
- [x] Homepage with upload/download functionality
- [x] Backend Excel processing with rank matching + NL for unmatched
- [x] Header: Logo centered globally, no auth buttons
- [x] Rank bands replacing null values in DB, download sheets, and processing
- [x] Static download Excel files regenerated with rank bands
- [x] All tests passing (iteration_1 + iteration_2)

- [x] Generated comprehensive Indian colleges Excel (42,927 colleges, alphabetically sorted)
  - File: `/frontend/public/indian_colleges_sorted.xlsx`
  - Source: AISHE government dataset (data.gov.in) via GitHub mirror

- [x] Generated comprehensive Indian colleges Excel (37,244 colleges, alphabetically sorted)
  - File: `/frontend/public/indian_colleges_sorted.xlsx`
  - Source: AISHE government dataset (data.gov.in) via GitHub mirror
- [x] All India Colleges seeded to MongoDB (`all_india_colleges` collection) with NIRF rank matching
  - 240 colleges matched to NIRF ranks, 37,004 marked as "NL"
  - API: `GET /api/download-all-india-colleges`, `GET /api/all-india-colleges/stats`, `POST /api/all-india-colleges/seed`
- [x] Download card added to homepage for "All India Colleges" Excel

- [x] Excel Matcher page (`/excel-matcher`) — Upload 2 Excel sheets, match by phone number, tag Shortlist/Reject
  - Backend: `POST /api/excel-matcher/preview`, `POST /api/excel-matcher/preview-sheet`, `POST /api/excel-matcher/process`
  - Multi-sheet support: User selects sheet tab for left file, right file searches ALL sheets automatically
  - Auto-detects header row (skips title rows), normalizes phone numbers (last 10 digits)
  - Downloads result with color-coded Shortlist (green) / Reject (red) / Not Found (gray)

## No Pending Tasks
All user-requested features have been implemented and verified.
