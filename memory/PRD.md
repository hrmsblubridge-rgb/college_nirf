# BluBridge Registration Form - Premium Design

## Problem Statement
Create a premium redesigned version of the BluBridge Registration Form as a new page at `/register-premium`. Same fields, same labels, same color palette — just upgraded to a premium design.

## Architecture
- Frontend: React + Tailwind CSS + shadcn UI components
- Backend: FastAPI (not modified for this task)
- Database: MongoDB (not modified for this task)
- Fonts: Manrope (headings) + Inter (body) via Google Fonts

## What's Been Implemented (Feb 2026)

## What's Been Implemented (Feb 2026)

## What's Been Implemented (Mar 2026)

### Home Page (`/`) — BluBridgeHome.jsx
- Premium BluBridge-branded homepage with beige background, white cards, blue accents
- **Stats section**: Total Colleges (dynamic from API), NIRF Ranked, Rank Band
- **Upload Card**: Drag-and-drop or browse to upload Job Post Excel (.xlsx/.xls)
  - 3-step visual guide (Upload → Auto Match → Download)
  - Processing state with spinner
  - Success/error messages with download link
- **Download Card**: 3 download options
  - Processed Excel (appears after processing)
  - NIRF Rankings with Short Names
  - NIRF Rankings (Original)
- **Seed DB Card**: One-click seed/refresh 300 colleges to MongoDB

### Backend APIs
- `POST /api/colleges/seed` - Seeds all 300 colleges into MongoDB
- `GET /api/colleges` - Paginated college list
- `GET /api/colleges/stats` - Returns total, ranked, unranked counts
- `POST /api/process-excel` - Upload Job Post Excel → returns processed Excel:
  - Tab 1: All original data with Rank UG + Rank PG auto-filled
  - Tab 2: Ranking Reference (full name + short names + rank pairs)
  - Matching: exact → short name → fuzzy (threshold 0.82)
  - Unranked band colleges → "null"

### College Database (MongoDB)
- 300 colleges: rank, college_name, short_names[], city, state
- 100 with specific NIRF ranks (1-100)
- 200 in rank bands (rank = null)

### Excel Files
- `/api/download-college-list` → Original rankings (Rank | Name | City | State)
- `/api/download-college-list-shortnames` → With short names (Rank | Name | Short Name | City | State)

## Colors (Same as Original)
- Background: #F5F2E9 (beige/cream)
- Card: #FFFFFF
- Primary CTA: #1A73E8 (blue)
- Footer: #222222
- Labels: #374151 (gray-700)

## Routes
- `/` - Original home page
- `/register-premium` - New premium registration form page

## Backlog / Next Action Items
- P1: Add form submission with backend API
- P1: Add form validation (required fields, email format, phone format)
- P2: Add success confirmation screen/modal after PROCEED
- P2: Toast notification on submission
- P2: Add smooth page entrance animations
