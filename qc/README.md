# Ermine QC — فحص الجودة اليومي

Bilingual (Arabic / English) daily quality checklist. Employees fill it on a phone and
submit; every report lands in a Google Sheet that works as the history log.

- Page: `qc/index.html` (served by GitHub Pages, same address as the dashboard + `/qc/`)
- Backend: `qc/backend/Code.gs` (Google Apps Script, bound to a Google Sheet)
- Address of the backend: `qc/config.js`

## Setup (about 15 minutes, once)

1. In Google Drive create a new **Google Sheet** named `Ermine QC`.
2. In the sheet: **Extensions → Apps Script**. Delete the starter code and paste all of
   `qc/backend/Code.gs`.
3. At the top change the two codes (8 characters or more, different from each other):
   - `TEAM_CODE`: the code employees type once on their phone.
   - `MANAGER_CODE`: the code that opens History and export.
4. Pick `setup` in the function menu and press **Run**. Allow the permissions
   (*Advanced → Go to project → Allow*). The tabs `Reports`, `Items`, `Config` appear.
5. **Deploy → New deployment → Web app**. *Execute as*: **Me**. *Who has access*: **Anyone**.
   Press **Deploy** and copy the **Web app URL**.
6. In this repository open `qc/config.js`, paste the URL between the quotes, commit.
7. Wait two or three minutes, then open the page, enter the team code, and send a test report.
8. Add your people in the `Config` tab: a new row with `kind = employee`, the name in `ar`
   (and `en` if you like). The same tab holds fridge names and limits, areas and hygiene items.
   Tick `hide` to remove a row without deleting it.

If you change `Code.gs` later: **Deploy → Manage deployments → Edit → New version**.

## What is stored where

| Sheet tab | One row per | Use |
|---|---|---|
| `Reports` | report | the history log, filter by date or employee |
| `Items` | checked item | temperatures and statuses, ready for a pivot table or chart |
| `Config` | employee / fridge / area / hygiene item | edit names and limits here |

## About access

The page is public like any GitHub Pages site, but it holds no secrets and no data. Reports
are only accepted with the team code, and History is only returned with the manager code.
Both codes live inside the Apps Script, never in this repository. Choose codes that are not
guessable, and change them if someone leaves.
