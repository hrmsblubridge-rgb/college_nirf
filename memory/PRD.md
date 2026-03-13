# BluBridge Registration Form - Premium Design

## Problem Statement
Create a premium redesigned version of the BluBridge Registration Form as a new page at `/register-premium`. Same fields, same labels, same color palette — just upgraded to a premium design.

## Architecture
- Frontend: React + Tailwind CSS + shadcn UI components
- Backend: FastAPI (not modified for this task)
- Database: MongoDB (not modified for this task)
- Fonts: Manrope (headings) + Inter (body) via Google Fonts

## What's Been Implemented (Feb 2026)

### /register-premium page (RegisterPremium.jsx)
- Premium card with rounded-3xl corners, large box-shadow
- Blue gradient accent bar at top of card
- Manrope font for the Registration Form heading
- Sticky header with frosted glass effect (backdrop-blur), BluBridge logo with blue "B" box, Logout button
- All original fields preserved:
  - Full Name (with User icon)
  - Email Address (with Mail icon)
  - Phone Number (with Phone icon + WhatsApp note)
  - Age (with Hash icon)
  - Current Location/State (shadcn Select dropdown)
  - Preferred Location/City (text input with MapPin icon)
  - Year of Graduation (shadcn Select dropdown)
  - College (text input with Building2 icon)
  - Degree (shadcn Select dropdown)
  - Course (shadcn Select dropdown)
- Section dividers: Personal Information / Location Details / Academic Information
- Uppercase tracked small labels (same naming as original)
- Confirmation checkbox with blue highlight when checked
- PROCEED button: blue when checkbox checked, grayed when unchecked, hover lift effect
- Dark footer (#222222) with copyright
- Full mobile responsiveness (1-col on mobile, 2-col on desktop)
- data-testid on all interactive elements

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
