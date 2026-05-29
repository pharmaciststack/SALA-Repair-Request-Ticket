# Sala-Repair — Repair Ticket Management System
### บริษัท ศาลาโอสถรีเทล จำกัด

---

## What It Is

Sala-Repair is a lightweight, web-based maintenance request and tracking system built specifically for Salaosot Retail — a pharmacy chain operating **34 branches** across Eastern Thailand. It replaces paper forms and LINE/phone-based reporting with a structured, auditable digital workflow.

**Live URL:** https://goodyearzph.github.io/salaosot-repair-system/
**Current Version:** 1.1.0

---

## The Problem It Solves

Before this system, maintenance requests at branch level were handled ad-hoc — phone calls, LINE messages, or handwritten notes. This meant:

- No central record of what was broken or who was working on it
- No visibility on overdue repairs
- No automatic notifications to technicians or requesters
- No audit trail or completion documentation

---

## Key Features

### For Branch Staff (Requester View)
- Submit repair tickets in under 2 minutes via Google Sign-In (no password, no account creation)
- Attach up to 8 photos per ticket (auto-uploaded to Google Drive)
- Set urgency level with built-in SLA guidance:
  - Low — 7–30 days
  - Medium — 2–7 days
  - High — 1–2 days
- Track the status of their own tickets in real time
- Receive automatic email confirmation on submission and on completion

### For Admins & Technicians (Dashboard View)
- Live dashboard with counters: total, pending, in-progress, done, overdue
- Filter tickets by status, urgency, category, branch, or overdue flag
- Full-text search across requester name, branch, equipment, and description
- Update ticket status and add technician notes in one click
- Generate a printable PDF completion report (with signature blocks) for any resolved ticket
- Auto-email notification to technicians on new submissions; auto-email to requesters on completion

### For Super Admins (Settings View)
- Manage the admin access list dynamically — no code changes needed
- Manage the technician email notification list dynamically
- All changes persist in Google Sheets in real time

---

## Technical Architecture

| Layer | Technology |
|---|---|
| Frontend | Vanilla JS + Tailwind CSS, hosted on GitHub Pages |
| Auth | Google Identity Services (OAuth 2.0 ID tokens) |
| Backend | Google Apps Script (Web App) |
| Database | Google Sheets (3 tabs: Tickets, AdminEmails, TechEmails) |
| File Storage | Google Drive (auto-created `RepairTicketImages` folder) |
| Notifications | Gmail via Apps Script `MailApp` |

**Zero infrastructure cost.** The entire system runs on Google's free tier — no server, no database subscription, no hosting fees.

---

## Security

All backend endpoints are protected by Google ID token verification. Every API request from the frontend includes the signed JWT issued at login; the backend validates it before processing any read or write. Specific fixes applied in the latest release:

- Backend auth required on every request (previously unprotected)
- JWT expiry checked on session restore
- XSS mitigated in PDF generation and image handling
- Field whitelist on update endpoint (only `status` and `note` accepted)
- No developer bypass or hardcoded credentials in production code

---

## Coverage

- **34 branches** pre-configured, with branch search autocomplete
- **6 equipment categories:** Electrical, Plumbing, Air Conditioning, IT/Computer, Furniture, Other
- Fully bilingual UI (Thai primary, English labels in backend/reports)
- Overdue badge auto-calculates based on urgency SLA from submission date

---

## Development Timeline

| Date | Milestone |
|---|---|
| Early 2026 | Initial release — ticket submission + local storage |
| v1.0.1 | PDF repair reports + dynamic admin management |
| v1.1.0 | Email notifications (technicians + requester auto-notify), overdue tracking |
| 2026-05-15 | Security hardening — all Critical & High issues resolved |

---

## Built By

Developed in-house by the IT team. Contact: ฝ่าย IT / ภก.ปาล์ม

---

*© 2026 Salaosot. All rights reserved.*
