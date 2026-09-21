# Parfum Compliance Dashboard — Developer Reference / Entwicklerreferenz

Status: 2026-09-21. Repo: `digitakt-consulting/ParfumeDashboard` (moved from `meikerensen/DraftParfuemKunde`).
Live: https://digitakt-consulting.github.io/ParfumeDashboard/ (the old `meikerensen.github.io` URL no longer works).

## Data flow / Datenfluss

1. Sophie and Zoltan maintain one Excel file on SharePoint: `Recent_ExcelListSummarized.xlsx` (10 tabs). / Sophie und Zoltan pflegen eine Excel-Datei auf SharePoint (10 Tabs).
2. GitHub workflow `.github/workflows/weekly-dashboard.yml` runs **every Monday 07:00 Berlin time** (two UTC crons + a Berlin-hour guard, so it is correct in summer and winter). It can also be started manually (Actions → "Weekly dashboard refresh").
3. `generate.js` fetches the tabs via Microsoft Graph (`lib/ms-auth.js` client-credentials token, `lib/msgraph-lite.js` worksheet `usedRange`), builds `dist/index.html`.
4. The workflow copies it to the repo-root `index.html` and pushes only if it changed; GitHub Pages serves it.

`.github/workflows/test-generate-msgraph.yml` = manual dry run, no push, uploads `dist/index.html` as artifact.

## Secrets / Zugangsdaten

GitHub repo secrets: `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`. Never paste values into chat or code. The client secret **expires** — check the expiry in the Azure app registration; an expired secret makes the Monday run fail.

## What the dashboard shows / Inhalt

GPSR status + weekly trend (tab "Numbers per week": weekly block + "Anzahl offener Fälle" Start/Aktuell block), processed listings ("3. New Listings"), Account Violations, prohibited ingredients ("Lilial"), brand approvals PD ("PD-brand approvals"). The Blocked ASINs section appears only if "4. Blocked ASINs" has data.

Deliberately removed (not in the SharePoint file): revenue, GPSR case log by market, priority status. The daily block in "Numbers per week" (Week/Day/PD/PS) is not read — unit unclear, values partly converted to dates.

## Gotchas / Stolperfallen

- **Graph `usedRange` starts at the first used row/column.** `msgraph-lite.js` restores absolute positions via `rowIndex` and `columnIndex`; without it, a sheet with an empty column A (e.g. "3. New Listings") shifts every column. `generate.js` fails loudly if New Listings has >50 rows but no dates.
- Graph returns native types (numbers, booleans, date serials); `msgraph-lite.js` normalises all cells to strings.
- Tab names and column order are read by fixed name/position — do not rename tabs or insert columns in the Excel file. Missing tab → empty section plus a note under "Daten-Hinweise".
- Verify numbers, not just the exit code: after changes, run the test workflow and compare key figures (e.g. 19 listings in 30 days, 23 weeks in the trend, 681 prohibited-ingredient entries, GPSR open now PD 151 / PS 22 at 2026-09-21).
- Local test without Microsoft access: stub `fetch` with a JSON fixture of the `usedRange` responses (must include `rowIndex`/`columnIndex` like the real API).
- Windows PowerShell 5.1 does not support `&&` — run git commands on separate lines.

## Open items / Offene Punkte

- Meaning/unit of the daily block in "Numbers per week".
- Zoltan to update Parfum Store numbers.
- SharePoint access for editors without a Digitakt email (not yet checked).
- Replace any shared `meikerensen.github.io` links.
- Note the client-secret expiry date and set a reminder.
- Concept document (bilingual): https://claude.ai/artifact/5FxxtPN9uDzncACw6VSeFg
