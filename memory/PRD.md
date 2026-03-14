# PRD - BluBridge College Ranking Processor

## Problem Statement
BluBridge platform needs:
1. Premium Registration Form page (/register-premium)
2. Home page (/) with College Ranking Excel Processor
3. 300 NIRF 2025 colleges saved in MongoDB
4. Upload Job Post Excel → auto-fill UG/PG ranks → download 2-tab processed Excel
5. Download NIRF reference Excel files

## Architecture
- Frontend: React + Tailwind CSS + shadcn UI
- Backend: FastAPI
- Database: MongoDB
- Fonts: Manrope + Inter (Google Fonts)

## What's Been Implemented

### /register-premium (RegisterPremium.jsx)
- Premium card with blue gradient accent bar, Manrope heading font
- Frosted-glass sticky header with BluBridge logo
- All original fields: Full Name, Email, Phone (+WhatsApp note), Age, State dropdown, City, Grad Year, College, Degree, Course
- Section dividers, confirmation checkbox, PROCEED button (activates only when checked)
- Dark footer, mobile responsive, all data-testids

### Home Page / (BluBridgeHome.jsx)
- Premium BluBridge-branded homepage (beige bg, white cards, blue accents)
- Live stats (Total Colleges / NIRF Ranked / Rank Band) from /api/colleges/stats
- Upload Card: drag-and-drop Job Post Excel with 3-step visual guide
- Download Card: 3 download options (Processed Excel + 2 NIRF reference files)
- Seed DB button: one-click seed 300 colleges to MongoDB

### Backend APIs
- POST /api/colleges/seed — Seeds all 300 colleges into MongoDB
- GET /api/colleges — Paginated college list
- GET /api/colleges/stats — Returns total/ranked/unranked counts
- POST /api/process-excel — Main feature:
  - Accepts Job Post Excel upload (.xlsx/.xls)
  - Finds UG University/institute Name column (col 22), Rank column (col 23)
  - Finds PG university/institute name column (col 27), Rank column (col 28)
  - Matches each college name via: exact → short name → fuzzy (threshold 0.82)
  - Top-100 colleges → specific rank (1–100)
  - Band colleges (101–300) → "null"
  - Unrecognized colleges → blank
  - Returns 2-tab Excel:
    * Tab 1: All original applicant data with Rank UG + Rank PG filled in
    * Tab 2 "Ranking Reference": Exact original input names (as typed by applicants) + matched rank, same order as input
- GET /api/download-college-list — Original rankings Excel
- GET /api/download-college-list-shortnames — Rankings with short names Excel

### MongoDB College Data (300 colleges)
- Fields: rank (int or null), college_name, short_names[], city, state
- 100 with NIRF ranks 1–100
- 200 in rank bands 101–150, 151–200, 201–300 (rank = null)
- Source: NIRF 2025 Engineering Rankings

## Excel Files Available
- /app/college_rankings.xlsx — Rank | Name | City | State
- /app/college_rankings_with_shortnames.xlsx — Rank | Name | Short Name | City | State

## Tab 2 Format (Ranking Reference) — Exact screenshot match
| UG University/institute Name | Rank |
|------------------------------|------|
| Vellore Institute of Technology | 16 |
| VIT | 16 |
| Jadavpur University | 18 |
| JU | 18 |

## Next Action Items (Backlog)
- P1: College dropdown in Registration Form → search from 300 DB colleges
- P1: Form submission with backend API + save to DB
- P2: Form validation (required fields, email/phone format)
- P2: Success confirmation screen/modal after PROCEED
- P2: Auth/login system
