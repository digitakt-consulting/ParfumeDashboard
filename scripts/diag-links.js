// One-off diagnostic: what do the "AMZ Case Link" cells really contain?
// Compares usedRange.values (display values) with usedRange.formulas, because
// a =HYPERLINK("url","text") cell or a cell hyperlink shows only its text in
// values. Prints counts + samples only, no secrets.
const { getAccessToken } = require("../lib/ms-auth");
const { resolveShareLink, graphGet } = require("../lib/msgraph-lite");
const SHARE_URL = "https://digitaktco-my.sharepoint.com/:x:/g/personal/zh_digitakt_co/IQBKluJ8RMXdRKmybCnHIALtAVOSdcbdI_Kq3nZ7But7-Rs";
const TABS = [["3. New Listings", 10, 8, 9], ["2. Account Violations", 11, 9, 10]]; // name, linkCol, statusCol, commentCol
(async () => {
  const token = await getAccessToken(process.env.MS_TENANT_ID, process.env.MS_CLIENT_ID, process.env.MS_CLIENT_SECRET);
  const { driveId, itemId } = await resolveShareLink(SHARE_URL, token);
  for (const [name, lc, sc, cc] of TABS) {
    const r = await graphGet(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${encodeURIComponent(name)}/usedRange?$select=values,formulas,text,valueTypes,rowIndex,columnIndex,address`, token);
    const c0 = r.columnIndex || 0, li = lc - c0, si = sc - c0, ci = cc - c0;
    const cls = { empty: 0, urlValue: 0, hyperlinkFormula: 0, otherText: 0 };
    const samples = { otherText: [], hyperlinkFormula: [] };
    const byStatus = {};
    r.values.forEach((row, i) => {
      if (i < 3) return;
      const v = String(row[li] ?? "").trim(), f = String((r.formulas[i] || [])[li] ?? "").trim();
      const st = String(row[si] ?? "").trim() || "(leer)";
      if (!row.some((x) => String(x ?? "").trim())) return;
      let k;
      if (/^=\s*HYPERLINK/i.test(f)) k = "hyperlinkFormula";
      else if (/^https?:\/\/\S+$/i.test(v)) k = "urlValue";
      else if (v === "") k = "empty";
      else k = "otherText";
      cls[k]++;
      byStatus[st] = byStatus[st] || { total: 0, withLink: 0 };
      byStatus[st].total++;
      if (k === "urlValue" || k === "hyperlinkFormula") byStatus[st].withLink++;
      if (samples[k] && samples[k].length < 12) samples[k].push({ row: (r.rowIndex || 0) + i + 1, value: v.slice(0, 80), formula: f.slice(0, 160) });
    });
    console.log(`\n=== ${name}  address=${r.address} ===`);
    console.log("classes:", JSON.stringify(cls));
    console.log("by status:", JSON.stringify(byStatus));
    console.log("samples:", JSON.stringify(samples, null, 1));
  }
})().catch((e) => { console.error(e.message || e); process.exit(1); });
