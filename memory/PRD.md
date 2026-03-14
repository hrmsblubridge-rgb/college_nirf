# BluBridge - College Ranking Processor

## Original Problem Statement
Build a college ranking processor application with:
1. A premium registration page
2. A homepage with Upload and Download options for Excel processing
3. Store ~300 colleges (rank, full name, short names, city, state) in a database
4. Process uploaded "Job Post appraise" Excel sheets by matching college names against the database and adding NIRF ranks
5. Center the logo in the global header and remove Register/Login/Logout buttons

## Core Requirements
- **Homepage**: Upload/Download functionality for Excel processing
- **College Database**: 300 colleges with rank, name, short names, city, state
- **Excel Processing**: Match college names (full + short) to NIRF ranks, insert Rank column, generate second tab with original names + ranks
- **UI**: Centered logo header, no auth buttons globally

## Tech Stack
- **Frontend**: React, Tailwind CSS, Shadcn/UI, Axios
- **Backend**: FastAPI, Pandas, Openpyxl
- **Database**: MongoDB

## Architecture
```
/app/
├── backend/
│   ├── .env
│   ├── models/college.py
│   ├── requirements.txt
│   └── server.py
├── frontend/
│   ├── .env
│   ├── package.json
│   └── src/
│       ├── App.js
│       └── components/
│           ├── BluBridgeHome.jsx
│           └── RegisterPremium.jsx
```

## Key API Endpoints
- `POST /api/process-excel` - Upload and process Excel with college rank matching
- `GET /api/colleges/stats` - Get college database statistics
- `GET /api/download-college-list` - Download NIRF rankings (original)
- `GET /api/download-college-list-shortnames` - Download NIRF rankings with short names
- `POST /api/colleges/seed` - Seed college database

## DB Schema
- **colleges**: `{ name: str, short_names: List[str], city: str, state: str, rank: Union[int, None] }`

## Completed Features (as of 2026-03-14)
- [x] Premium Registration Page (`/register-premium`)
- [x] College data seeding (300 colleges into MongoDB)
- [x] Homepage with upload/download functionality
- [x] Backend Excel processing endpoint with rank matching
- [x] Bug fix: Data overwrite issue (Graduation Year column preserved)
- [x] Bug fix: Correct name matching in second tab
- [x] Header: Logo centered globally
- [x] Header: Register/Login/Logout buttons removed globally
- [x] E2E testing passed

## No Pending Tasks
All user-requested features have been implemented and verified.
