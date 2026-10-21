# Drishti

SLA and KPI monitor for Purview Healthcare clients.

## Files in this repository

- `index.html` is the app.
- `config.js` holds two settings: the web address of Drishti's Google Sheet backend, and the
  Application (client) ID of the "Purview Drishti" Microsoft app, which switches on OneDrive.
  Both are public identifiers, not passwords.

## Where the data lives

Accounts, managers, clients and KPI numbers are kept in one Google Sheet, through a
Google Apps Script web app. Everyone sees the same data from any computer.

Report files are read inside each person's browser and are never uploaded anywhere.
Only the figures worked out from them are saved, so no patient names or claim lines
are stored. Do not upload real patient reports to this repository.

If `DRISHTI_API` in `config.js` is empty, Drishti falls back to saving in each person's own browser.

## OneDrive

Report files can also be fetched from OneDrive. They sit in the OneDrive of
Purviewautomate@purview-healthcare.com: Drishti, then one folder per supervisor, one per client,
one per month (for example October 2026, holding the month's files) and inside it one per week
(Week 1, Week 2 and so on). A week runs Monday to Friday and goes in the month its Monday is in.

Fetch reads a month's or week's files in the person's browser, with their own Microsoft sign-in,
and shows the usual checklist. Update writes Drishti's own file, "Drishti summary - client.xlsx",
in the client's folder. Drishti never changes the team's report files. Nothing runs by itself.

If `DRISHTI_MS_CLIENT_ID` in `config.js` is empty, OneDrive is switched off and files are
dropped by hand as before.

## History

Every month, week and day that is uploaded is kept. Each client has a History tab that puts
them side by side (for example October, November and December), with a trend line per metric
and a Download Excel button. The Google Sheet keeps the same numbers in its KPI History tab.

## Who sees what

- Admin (one account, set in the Apps Script settings): every client.
- Manager: their own clients and every client maintained by the people who report to them.
- Supervisor, Team Lead, SME: only the clients they maintain.

## Updating

Replace `index.html` with the new version. Leave `config.js` as it is. Data is not affected.
