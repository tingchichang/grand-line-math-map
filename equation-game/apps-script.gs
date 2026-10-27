/**
 * 一元一次方程式闖關島｜作答紀錄接收程式（Google Apps Script）
 * 使用方式：在試算表中「擴充功能 → Apps Script」，把這個檔案的內容全部貼上，
 * 再「部署 → 新增部署作業 → 網頁應用程式」，執行身分選「我」，存取權選「所有人」。
 *
 * - 遊戲以 POST 送出作答紀錄，寫入「作答紀錄」工作表。
 * - 教師後台以 GET ?action=all&key=密碼 讀取全部紀錄；密碼放在「設定」工作表 B2。
 * - 「設定」B4 改成「關」可暫停接收作答。
 */
const SHEET = '作答紀錄';
const FIELDS = ['t', 'cls', 'seat', 'name', 'sid', 'dev', 'round', 'isl', 'islName', 'phase', 'qid', 'q', 'ans', 'correct', 'ok', 'myth', 'hint', 'sec'];
const MAX_ROWS_PER_POST = 200;

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (String(ss.getSheetByName('設定').getRange('B4').getValue()).trim() === '關') return json({ ok: false, error: 'closed' });
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'bad_json' }); }
  const rows = (body.rows || []).slice(0, MAX_ROWS_PER_POST).map(r => FIELDS.map(f => clean(r[f])));
  if (!rows.length) return json({ ok: true, n: 0 });
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = ss.getSheetByName(SHEET);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, FIELDS.length).setValues(rows);
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true, n: rows.length });
}

function doGet(e) {
  const p = e.parameter || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (p.action === 'ping') return json({ ok: true });
  if (p.action === 'all') {
    const key = String(ss.getSheetByName('設定').getRange('B2').getValue()).trim();
    if (!key || String(p.key || '').trim() !== key) return json({ ok: false, error: 'key' });
    const sh = ss.getSheetByName(SHEET);
    const n = sh.getLastRow() - 1;
    const rows = n > 0 ? sh.getRange(2, 1, n, FIELDS.length).getValues() : [];
    return json({ ok: true, fields: FIELDS, rows: rows.map(r => r.map(v => v instanceof Date ? v.toISOString() : v)) });
  }
  return json({ ok: false, error: 'action' });
}

// 避免以 = + - @ 開頭的內容被試算表當成公式執行
function clean(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  let s = String(v).slice(0, 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
