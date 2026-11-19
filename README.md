# Kovie Student Actor Signup

Public signup page for UARK theater students performing Kovie's live roommate stories on Union Mall, Tuesday, December 1, 2026, 11:00 a.m. to 2:00 p.m.

**Live page:** https://kea1e.github.io/kovie-actor-signup/

## How it works

- `index.html` is a static page on GitHub Pages. It shows six half-hour slots with two actor seats each.
- Tapping an open seat places a 5 minute hold on it, so nobody else can take it while the actor fills in their name and UARK email.
- Booking writes a row to the **Kovie Actor Signup** Google Sheet, sends a confirmation email, and sends a Google Calendar invite from the **Kovie Actor Signups** calendar.
- The confirmation email has a cancel link (`?cancel=TOKEN`) that frees the seat and withdraws the invite.
- The page shows names of signed-up actors. Emails are never returned by the API; they live only in the sheet.
- `@uark.edu` addresses only. One person can take several slots, but not two seats at the same time. Signups close at 11:00 a.m. on December 1.

## Backend

`apps-script/` is a Google Apps Script web app bound to the sheet, deployed with [clasp](https://github.com/google/clasp).

```bash
cd apps-script/..
clasp push --force
clasp redeploy <deploymentId> --description "update"   # keeps the same /exec URL
```

The sheet has two tabs: **Roster** (same layout as the original spreadsheet, rebuilt on every change) and **Signups** (the full log, including cancellations).

Event details, slot times and the contact line are in `CONFIG` at the top of `apps-script/Code.gs` and in the text of `index.html`.
