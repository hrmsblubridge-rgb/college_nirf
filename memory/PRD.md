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

### /register-premium page (RegisterPremium.jsx)
- Premium card with rounded-3xl corners, large box-shadow
- Blue gradient accent bar at top of card
- Manrope font for the Registration Form heading
- Sticky header with frosted glass effect, BluBridge logo, Logout button
- All original fields: Full Name, Email, Phone (+WhatsApp note), Age, State dropdown, City, Grad Year, College, Degree, Course
- Section dividers, confirmation checkbox, PROCEED button
- Dark footer, mobile responsive, data-testid on all elements

### College Rankings Excel (BluBridge_College_Rankings_NIRF_2025.xlsx)
- Source: NIRF 2025 Engineering Rankings
- Total colleges: 301 (100 ranked + 201 null-rank)
- Columns: Rank | College Name | City | State
- Top 100: specific NIRF ranks 1-100
- Band 101-150: 50 colleges (null rank)
- Band 151-200: 52 colleges (null rank)
- Band 201-300: 99 colleges (null rank)
- Color-coded Excel: Blue header, ranked rows in light blue, null rows in light gray
- Band separator rows for easy navigation
- Download endpoint: GET /api/download-college-list

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
