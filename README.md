# Drishti

SLA and KPI monitor for Purview Healthcare clients.

## Files in this repository

- `index.html` is the app.
- `config.js` holds the settings: the web address of Drishti's Google Sheet backend, the
  Application (client) ID of the "Purview Drishti" Microsoft app (it switches on OneDrive), the
  Microsoft account everyone signs in with for OneDrive, and an optional OneDrive folder link that
  is left empty. None of them is a password.

## Where the data lives

Accounts, managers, clients and KPI numbers are kept in one Google Sheet, through a
Google Apps Script web app. Everyone sees the same data from any computer.

Report files are read inside each person's browser, straight from OneDrive, and are never
uploaded anywhere else. Only the figures worked out from them are saved, so no patient names
or claim lines are stored. Do not upload real patient reports to this repository.

If `DRISHTI_API` in `config.js` is empty, Drishti falls back to saving in each person's own browser.

## OneDrive

Every report file comes from OneDrive. Nothing is uploaded from a computer. Everyone signs in to
Microsoft as the team's shared account, Purviewautomate@purview-healthcare.com, and
"Upload from OneDrive" opens My files in its OneDrive. The button is on the big box, for all of a
period's files at once (each matched to its report by its name), and on each report's own box, for
one file. The person opens the folders, ticks the files, and the usual checklist shows. The files go
into the month or week on screen. Drishti only accepts the shared account.

Suggested layout in My files: Drishti, then one folder per supervisor, one per client, one per month
(for example October 2026, holding the month's files) and inside it one per week (Week 1, Week 2
and so on). A week runs Monday to Friday and goes in the month its Monday is in.

"Update Drishti summary" writes Drishti's own file, "Drishti summary - client.xlsx", in the client's
folder: the one above the month and week folders the files came from. Drishti never changes the
team's report files. Nothing runs by itself.

If `DRISHTI_MS_CLIENT_ID` in `config.js` is empty, OneDrive is switched off and files are dropped
from the computer as before.

## History

Every month, week and day that is uploaded is kept. Each client has a History tab that puts
them side by side (for example October, November and December), with a trend line per metric
and a Download Excel button. The Google Sheet keeps the same numbers in its KPI History tab.

## Who sees what

- Admin (one account, set in the Apps Script settings): every client.
- Manager: their own clients and every client maintained by the people who report to them.
- Supervisor, Team Lead, SME: only the clients they maintain.

## Updating

Replace `index.html` with the new version. `config.js` changes only when a setting changes.
Data is not affected.
