/**
 * Expense Tracker — Google Sheets backend
 *
 * Paste this into Extensions → Apps Script of your private Google Sheet,
 * set a secret key (see README), then Deploy → New deployment → Web app.
 *
 * Every request must carry the secret key; without it the script returns
 * nothing. The sheet itself stays private in your Drive.
 */

const TX_SHEET = 'Transactions';
const BUDGET_SHEET = 'Budgets';
const TX_HEADERS = ['id', 'date', 'type', 'category', 'counterparty', 'amount', 'note', 'source', 'updated'];
const BUDGET_HEADERS = ['category', 'monthly_budget'];

/** Run once from the editor: creates the tabs and asks nothing else. */
function setup() {
  sheet_(TX_SHEET, TX_HEADERS);
  sheet_(BUDGET_SHEET, BUDGET_HEADERS);
  if (!PropertiesService.getScriptProperties().getProperty('API_KEY')) {
    throw new Error('Now set API_KEY: Project Settings → Script properties → Add property API_KEY');
  }
}

function doGet(e) {
  return handle_(e.parameter || {});
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return out_({ ok: false, error: 'Bad JSON' }); }
  return handle_(body);
}

function handle_(req) {
  const key = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!key || req.key !== key) return out_({ ok: false, error: 'Unauthorized' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (req.action) {
      case 'list':      return out_({ ok: true, transactions: readTx_(), budgets: readBudgets_() });
      case 'add':       return out_({ ok: true, tx: upsert_([req.tx])[0] });
      case 'update':    return out_({ ok: true, tx: upsert_([req.tx])[0] });
      case 'delete':    return out_({ ok: true, deleted: delete_(req.id) });
      case 'import':    return out_({ ok: true, added: importRows_(req.rows || []) });
      case 'budgets':   return out_({ ok: true, budgets: saveBudgets_(req.budgets || {}) });
      case 'ping':      return out_({ ok: true });
      default:          return out_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readTx_() {
  const sh = sheet_(TX_SHEET, TX_HEADERS);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const vals = sh.getRange(2, 1, n, TX_HEADERS.length).getValues();
  return vals.filter(r => r[0] !== '').map(rowToTx_);
}

function rowToTx_(r) {
  const d = r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[1]);
  return { id: String(r[0]), date: d, type: r[2], category: r[3], counterparty: r[4], amount: Number(r[5]), note: r[6], source: r[7] };
}

function txToRow_(t) {
  return [t.id, t.date, t.type === 'in' ? 'in' : 'out', t.category || 'Uncategorised', t.counterparty || '', Number(t.amount) || 0, t.note || '', t.source || 'manual', new Date()];
}

function idIndex_(sh) {
  const n = sh.getLastRow() - 1;
  const map = {};
  if (n < 1) return map;
  sh.getRange(2, 1, n, 1).getValues().forEach((r, i) => { if (r[0] !== '') map[String(r[0])] = i + 2; });
  return map;
}

function upsert_(list) {
  const sh = sheet_(TX_SHEET, TX_HEADERS);
  const idx = idIndex_(sh);
  return list.map(t => {
    if (!t || !t.date || !(Number(t.amount) > 0)) throw new Error('date and a positive amount are required');
    if (!t.id) t.id = Utilities.getUuid();
    const row = txToRow_(t);
    if (idx[t.id]) sh.getRange(idx[t.id], 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return t;
  });
}

function delete_(id) {
  const sh = sheet_(TX_SHEET, TX_HEADERS);
  const row = idIndex_(sh)[String(id)];
  if (!row) return false;
  sh.deleteRow(row);
  return true;
}

/** Bulk import; rows whose id already exists are skipped, so re-importing is safe. */
function importRows_(rows) {
  const sh = sheet_(TX_SHEET, TX_HEADERS);
  const idx = idIndex_(sh);
  const fresh = rows.filter(t => t && t.id && !idx[t.id] && Number(t.amount) > 0).map(txToRow_);
  if (fresh.length) sh.getRange(sh.getLastRow() + 1, 1, fresh.length, TX_HEADERS.length).setValues(fresh);
  return fresh.length;
}

function readBudgets_() {
  const sh = sheet_(BUDGET_SHEET, BUDGET_HEADERS);
  const n = sh.getLastRow() - 1;
  const b = {};
  if (n < 1) return b;
  sh.getRange(2, 1, n, 2).getValues().forEach(r => { if (r[0] !== '' && Number(r[1]) > 0) b[r[0]] = Number(r[1]); });
  return b;
}

function saveBudgets_(budgets) {
  const sh = sheet_(BUDGET_SHEET, BUDGET_HEADERS);
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, 2).clearContent();
  const rows = Object.keys(budgets).filter(k => Number(budgets[k]) > 0).map(k => [k, Number(budgets[k])]);
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
  return readBudgets_();
}
