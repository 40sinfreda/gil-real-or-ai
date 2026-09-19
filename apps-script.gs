function doGet(e) {
  const p = e.parameter || {};
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('scores');
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName('scores');
    if (sh.getLastRow() === 0) sh.appendRow(['name', 'score', 'wrong', 'ms', 'when']);
  }
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  if (p.action === 'add' && p.name) {
    sh.appendRow([
      String(p.name).slice(0, 80),
      Number(p.score) || 0,
      Number(p.wrong) || 0,
      Number(p.ms) || 0,
      new Date()
    ]);
  }
  const rows = sh.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    items.push({ name: String(r[0]), score: Number(r[1]) || 0, wrong: Number(r[2]) || 0, ms: Number(r[3]) || 0 });
  }
  items.sort(function (a, b) { return (b.score - a.score) || (a.ms - b.ms); });
  return out.setContent(JSON.stringify({ ok: true, items: items.slice(0, 50) }));
}
