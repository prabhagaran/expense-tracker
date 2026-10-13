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
const SETTINGS_SHEET = 'Settings';
const SETTINGS_HEADERS = ['key', 'value'];
const TRASH_SHEET = 'Trash';
const TRASH_HEADERS = TX_HEADERS.concat(['deleted_at']);
const TRASH_DAYS = 30;

/** Run once from the editor: creates the tabs and asks nothing else. */
function setup() {
  sheet_(TX_SHEET, TX_HEADERS);
  sheet_(BUDGET_SHEET, BUDGET_HEADERS);
  sheet_(SETTINGS_SHEET, SETTINGS_HEADERS);
  sheet_(TRASH_SHEET, TRASH_HEADERS);
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
  // Reads don't need the lock and are served from cache when nothing has changed.
  if (req.action === 'list') return out_(listCached_());
  if (req.action === 'ping') return out_({ ok: true });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (req.action) {
      case 'list':      return out_({ ok: true, transactions: readTx_(), budgets: readBudgets_(), settings: readSettings_() });
      case 'add':       return out_({ ok: true, tx: upsert_([req.tx])[0] });
      case 'update':    return out_({ ok: true, tx: upsert_([req.tx])[0] });
      case 'delete':    return out_({ ok: true, deleted: delete_(req.id) });
      case 'updateMany':return out_({ ok: true, count: upsert_(req.txs || []).length });
      case 'deleteMany':return out_({ ok: true, count: deleteMany_(req.ids || []) });
      case 'import':    return out_({ ok: true, added: importRows_(req.rows || []) });
      case 'budgets':   return out_({ ok: true, budgets: saveBudgets_(req.budgets || {}) });
      case 'settings':  return out_({ ok: true, settings: saveSettings_(req.settings || {}) });
      case 'restore':   return out_({ ok: true, count: restore_(req.ids || []) });
      case 'ping':      return out_({ ok: true });
      default:          return out_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    clearListCache_();
    lock.releaseLock();
  }
}

/* ---------- list cache: the full list is kept in CacheService (in chunks) until the next write ---------- */
const CACHE_CHUNK = 90000;
function listCached_() {
  const cache = CacheService.getScriptCache();
  const meta = cache.get('list:n');
  if (meta) {
    const keys = []; for (let i = 0; i < Number(meta); i++) keys.push('list:' + i);
    const parts = cache.getAll(keys);
    if (keys.every(k => parts[k] != null)) { const o = JSON.parse(keys.map(k => parts[k]).join('')); o.cached = true; return o; }
  }
  const o = { ok: true, transactions: readTx_(), budgets: readBudgets_(), settings: readSettings_() };
  try {
    const s = JSON.stringify(o), put = {};
    const n = Math.ceil(s.length / CACHE_CHUNK);
    for (let i = 0; i < n; i++) put['list:' + i] = s.slice(i * CACHE_CHUNK, (i + 1) * CACHE_CHUNK);
    put['list:n'] = String(n);
    cache.putAll(put, 21600);
  } catch (e) { /* cache is an optimisation only */ }
  return o;
}
function clearListCache_() { try { CacheService.getScriptCache().remove('list:n'); } catch (e) {} }

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
  const d = normDate_(r[1]);
  return { id: String(r[0]), date: d, type: r[2], category: r[3], counterparty: r[4], amount: Number(r[5]), note: r[6], source: r[7] };
}

/** Accepts Date objects, yyyy-mm-dd, dd-mm-yyyy and dd/mm/yyyy; returns yyyy-MM-dd. */
let TZ_ = null;
function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ_ || (TZ_ = Session.getScriptTimeZone()), 'yyyy-MM-dd');
  const s = String(v || '').trim().replace(/^'/, '');
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})$/);
  if (m) return (m[3].length === 2 ? '20' + m[3] : m[3]) + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return s;
}

function txToRow_(t) {
  return [t.id, "'" + normDate_(t.date), t.type === 'in' ? 'in' : 'out', t.category || 'Uncategorised', t.counterparty || '', Number(t.amount) || 0, t.note || '', t.source || 'manual', new Date()];
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
  return deleteMany_([id]) > 0;
}

/** Deleted rows are moved to the Trash tab and kept for 30 days, so a delete can be undone. */
function deleteMany_(ids) {
  const sh = sheet_(TX_SHEET, TX_HEADERS);
  const tr = sheet_(TRASH_SHEET, TRASH_HEADERS);
  const idx = idIndex_(sh);
  const rows = ids.map(id => idx[String(id)]).filter(Boolean).sort((a, b) => b - a);
  if (!rows.length) return 0;
  const now = new Date();
  const moved = rows.map(r => sh.getRange(r, 1, 1, TX_HEADERS.length).getValues()[0].concat([now]));
  tr.getRange(tr.getLastRow() + 1, 1, moved.length, TRASH_HEADERS.length).setValues(moved);
  rows.forEach(r => sh.deleteRow(r));
  purgeTrash_();
  return rows.length;
}

function restore_(ids) {
  const sh = sheet_(TX_SHEET, TX_HEADERS);
  const tr = sheet_(TRASH_SHEET, TRASH_HEADERS);
  const n = tr.getLastRow() - 1;
  if (n < 1) return 0;
  const want = {}; ids.forEach(id => want[String(id)] = true);
  const vals = tr.getRange(2, 1, n, TRASH_HEADERS.length).getValues();
  const back = [], drop = [];
  vals.forEach((r, i) => { if (want[String(r[0])]) { back.push(r.slice(0, TX_HEADERS.length)); drop.push(i + 2); } });
  if (back.length) sh.getRange(sh.getLastRow() + 1, 1, back.length, TX_HEADERS.length).setValues(back);
  drop.sort((a, b) => b - a).forEach(r => tr.deleteRow(r));
  return back.length;
}

function purgeTrash_() {
  const tr = sheet_(TRASH_SHEET, TRASH_HEADERS);
  const n = tr.getLastRow() - 1;
  if (n < 1) return;
  const cutoff = Date.now() - TRASH_DAYS * 864e5;
  const col = tr.getRange(2, TRASH_HEADERS.length, n, 1).getValues();
  const old = [];
  col.forEach((r, i) => { const t = r[0] instanceof Date ? r[0].getTime() : Date.parse(r[0]); if (t && t < cutoff) old.push(i + 2); });
  old.sort((a, b) => b - a).forEach(r => tr.deleteRow(r));
}

/** Settings are stored as key → JSON value (transfer categories, bills, opening balance). */
function readSettings_() {
  const sh = sheet_(SETTINGS_SHEET, SETTINGS_HEADERS);
  const n = sh.getLastRow() - 1;
  const s = {};
  if (n < 1) return s;
  sh.getRange(2, 1, n, 2).getValues().forEach(r => { if (r[0] !== '') { try { s[r[0]] = JSON.parse(r[1]); } catch (e) { s[r[0]] = r[1]; } } });
  return s;
}

function saveSettings_(obj) {
  const sh = sheet_(SETTINGS_SHEET, SETTINGS_HEADERS);
  const cur = readSettings_();
  Object.keys(obj).forEach(k => cur[k] = obj[k]);
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, 2).clearContent();
  const rows = Object.keys(cur).map(k => [k, JSON.stringify(cur[k])]);
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
  return cur;
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
