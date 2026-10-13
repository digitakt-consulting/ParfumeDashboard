// One-off diagnostic: why do most blocked "New Listings" rows have no ASIN?
const { getAccessToken } = require("../lib/ms-auth");
const { resolveShareLink, graphGet } = require("../lib/msgraph-lite");
const SHARE_URL = "https://digitaktco-my.sharepoint.com/:x:/g/personal/zh_digitakt_co/IQBKluJ8RMXdRKmybCnHIALtAVOSdcbdI_Kq3nZ7But7-Rs";
(async () => {
  const token = await getAccessToken(process.env.MS_TENANT_ID, process.env.MS_CLIENT_ID, process.env.MS_CLIENT_SECRET);
  const { driveId, itemId } = await resolveShareLink(SHARE_URL, token);
  const r = await graphGet(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${encodeURIComponent("3. New Listings")}/usedRange?$select=values,rowIndex,columnIndex`, token);
  const c0 = r.columnIndex || 0;
  const col = (row, i) => String(row[i - c0] ?? "").trim();
  const isAsin = (s) => /^B0[A-Z0-9]{8}$/i.test(s);
  let blocked = 0, noAsin = 0, asinOtherForm = 0, eanOnly = 0, productOnly = 0;
  const shapes = {}, samples = [];
  r.values.forEach((row, i) => {
    if (i < 3 || col(row, 8).toLowerCase() !== "blocked") return;
    blocked++;
    const a6 = col(row, 6), a14 = col(row, 14), ean = col(row, 5), prod = col(row, 7);
    if (isAsin(a6) || isAsin(a14)) return;
    noAsin++;
    const shape = (v) => (v === "" ? "empty" : /^\d{8,14}$/.test(v) ? "digits8-14" : /^[A-Z0-9]{10}$/i.test(v) ? "alnum10" : "text");
    const k = `c6=${shape(a6)} c14=${shape(a14)} ean=${shape(ean)} product=${prod ? "yes" : "empty"}`;
    shapes[k] = (shapes[k] || 0) + 1;
    if (samples.length < 15) samples.push({ row: (r.rowIndex || 0) + i + 1, c6: a6.slice(0, 20), c14: a14.slice(0, 20), ean: ean.slice(0, 20), product: prod.slice(0, 50), country: col(row, 3), comment: col(row, 9).slice(0, 60) });
  });
  console.log(JSON.stringify({ blocked, noAsin }));
  console.log("shapes:", JSON.stringify(shapes, null, 1));
  console.log("samples:", JSON.stringify(samples, null, 1));
})().catch((e) => { console.error(e.message || e); process.exit(1); });
