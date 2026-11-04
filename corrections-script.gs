// Name corrections for the Invigilator Duty Finder.
// Paste into a Google Sheet (Extensions > Apps Script), then Deploy > New deployment >
// Web app, Execute as: Me, Who has access: Anyone. Copy the /exec URL into FIX_URL in index.html.
// Each correction is one row; the newest row per person wins. Delete a row to undo it.

const SHEET = 'Corrections';

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['Time', 'Slug', 'Original name', 'Corrected name']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function dist_(a, b) {
  let p = [...Array(b.length + 1).keys()];
  for (let i = 1; i <= a.length; i++) {
    const c = [i];
    for (let j = 1; j <= b.length; j++) c[j] = Math.min(p[j] + 1, c[j - 1] + 1, p[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    p = c;
  }
  return p[b.length];
}

// Same rule as the page: allowed characters only, and close to the original spelling.
function ok_(from, to) {
  return /^[A-Za-z][A-Za-z .'()\-]{2,59}$/.test(to) &&
    dist_(from.toLowerCase(), to.toLowerCase()) <= Math.max(3, Math.floor(0.4 * Math.max(from.length, to.length)));
}

function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const out = {};
  for (const [, slug, from, to] of rows) if (slug && from && to) out[slug] = { from: String(from), to: String(to) };
  return json_(out);
}

function doPost(e) {
  try {
    const { slug, from, to } = JSON.parse(e.postData.contents);
    const name = String(to || '').trim().replace(/\s+/g, ' ');
    if (!/^[a-z0-9-]{1,80}$/.test(slug) || !from || String(from).length > 60 || !ok_(String(from), name)) return json_({ ok: false, error: 'invalid' });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try { sheet_().appendRow([new Date(), slug, String(from), name]); } finally { lock.releaseLock(); }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'invalid' });
  }
}
