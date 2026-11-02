// Reads the shared Google Sheet (published as "anyone with the link can view")
// through its public CSV export - one request per tab, no credentials needed.
// Exposes the same { lastModified, listSheets(), getSheet(name), getRows(name) }
// interface as lib/msgraph-lite.js, so generate.js's row logic is unchanged.
//
// Cell values arrive as the sheet's DISPLAY text (CSV export), not as Excel
// serial numbers: dates are text like "2.11" / "14.12.2024" (handled by
// parseFlexibleDate in generate.js), numbers are plain digits.
// rowNum is the 1-based sheet row, like the Graph loader.
//
// Google gives no last-edit time via the export, so lastModified is null and
// generate.js omits the "Excel zuletzt bearbeitet" chip.
const { parseCSV } = require("./csv-lite");

const BASE = "https://docs.google.com/spreadsheets/d";

async function getText(url, what) {
  const res = await fetch(url, { redirect: "follow" });
  if (!res.ok) throw new Error(`Google Sheets GET ${what} failed: ${res.status} ${res.statusText}`);
  return res.text();
}

// Tab name -> gid, read from the sheet's own htmlview so a re-created tab with
// a new gid keeps working. Throws if the sheet is not publicly viewable.
async function listTabGids(sheetId) {
  const html = await getText(`${BASE}/${sheetId}/htmlview`, "htmlview");
  const map = {};
  const re = /items[.]push[(][{]name: "([^"]*)"[^}]*?gid: "([0-9]+)"/g;
  let m;
  while ((m = re.exec(html))) map[m[1]] = m[2];
  if (!Object.keys(map).length) {
    throw new Error("Keine Tabs im Google Sheet gefunden - ist es noch für 'Jeder mit dem Link' freigegeben?");
  }
  return map;
}

async function loadWorkbookFromGoogleSheet(sheetId, tabNames) {
  const gids = await listTabGids(sheetId);
  const missing = tabNames.filter((n) => !(n in gids));
  if (missing.length) throw new Error(`Tab(s) fehlen im Google Sheet: ${missing.join(", ")}`);

  const cache = {};
  for (const name of tabNames) {
    const csv = await getText(`${BASE}/${sheetId}/export?format=csv&gid=${gids[name]}`, `tab "${name}"`);
    if (/^\s*<(!doctype|html)/i.test(csv)) throw new Error(`Tab "${name}": Google lieferte HTML statt CSV (Freigabe/Login?)`);
    cache[name] = parseCSV(csv).map((row, idx) => {
      const arr = row.map((v) => (v == null ? "" : String(v)));
      arr.rowNum = idx + 1;
      return arr;
    });
  }

  return {
    lastModified: null,
    listSheets: () => tabNames,
    getSheet: (name) => cache[name],
    getRows: (name) => {
      if (!(name in cache)) throw new Error(`Tab "${name}" was not fetched - add it to the tabNames list.`);
      return cache[name];
    },
  };
}

module.exports = { loadWorkbookFromGoogleSheet, listTabGids };
