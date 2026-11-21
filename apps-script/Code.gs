// Workout Buddy: connection between the app and the Workout Buddy Google Sheet.
// Deployed as a web app (Execute as: me, Who has access: Anyone).
// The key lives only in the deployed copy (FALLBACK_KEY or Script Property KEY); this public copy holds no secrets.
//
// People: every per-person tab (Sets, Events, Progress, Plan, Feedback, Exercise Overrides) has a Person column.
// A blank Person means Josh (rows from before Michelle joined). Settings rows with a blank Person apply to everyone;
// a row with a Person overrides that setting for that person. Exercises is the shared library; per-person changes
// (tempo, pace, on/off, form tips) live in Exercise Overrides.

const FALLBACK_KEY = '';   // filled in only in the deployed copy, never in this repo
const SHEET_ID = '1hmU_XKOhSxIM_eculF_hw0agBoI6TKY7EnNb10x5go0';
const ID_COL = { Sets: 18, Events: 8, Progress: 5 };   // column that holds each row's unique id
const WIDTH = { Sets: 18, Events: 8, Progress: 5 };    // columns before Person
const TZ = 'America/New_York';
const DEFAULT_PERSON = 'Josh';

function doGet(e) { return handle_(e.parameter || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad json' }); }
  return handle_(p);
}

function handle_(p) {
  const key = PropertiesService.getScriptProperties().getProperty('KEY') || FALLBACK_KEY;
  if (!key || p.key !== key) return out_({ ok: false, error: 'bad key' });
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const person = p.person || DEFAULT_PERSON;
  try {
    if (p.action === 'ping') return out_({ ok: true, now: new Date().toISOString() });
    if (p.action === 'state') return out_(Object.assign({ ok: true, person: person }, state_(ss, person)));
    // generic actions (rows as objects keyed by header): new features use these, so the script rarely needs redeploying
    if (p.action === 'read') return out_(gRead_(ss, p));
    if (p.action === 'append' && p.tab) return out_(gAppend_(ss, p));
    if (p.action === 'update') return out_(gUpdate_(ss, p));
    if (p.action === 'delete') return out_(gDelete_(ss, p));
    // positional appends for Sets / Events / Progress (the app's offline outbox)
    if (p.action === 'append') return out_(Object.assign({ ok: true }, append_(ss, p.sheet, p.rows || [], person)));
    if (p.action === 'updateExercise') return out_(Object.assign({ ok: true }, updateExercise_(ss, p.exercise, p.field, p.value, person)));
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
const who_ = v => String(v || '').trim() || DEFAULT_PERSON;

// Read a tab as an array of objects keyed by header, using displayed text.
function table_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const cols = sh.getLastColumn();
  const head = sh.getRange(1, 1, 1, cols).getDisplayValues()[0];
  const vals = sh.getRange(2, 1, last - 1, cols).getDisplayValues();
  return vals.filter(r => r.some(c => c !== '')).map(r => {
    const o = {};
    head.forEach((h, i) => { if (h) o[h] = r[i]; });
    return o;
  });
}
function mine_(rows, person, max) {
  const m = rows.filter(r => who_(r['Person']) === person);
  return max ? m.slice(-max) : m;
}

function state_(ss, person) {
  // settings: shared rows first, then this person's rows override
  const settings = {};
  const srows = table_(ss, 'Settings');
  srows.filter(r => !String(r['Person'] || '').trim()).forEach(r => { settings[r['Setting']] = r['Value']; });
  srows.filter(r => String(r['Person'] || '').trim() === person).forEach(r => { settings[r['Setting']] = r['Value']; });
  // exercises: shared library with this person's overrides applied (latest wins)
  const exercises = table_(ss, 'Exercises');
  const byName = {};
  exercises.forEach(e => { byName[e['Exercise']] = e; });
  mine_(table_(ss, 'Exercise Overrides'), person).forEach(o => {
    const e = byName[o['Exercise']];
    if (e && o['Field']) e[o['Field']] = o['Value'];
  });
  return {
    exercises: exercises,
    settings: settings,
    plan: mine_(table_(ss, 'Plan'), person),
    sets: mine_(table_(ss, 'Sets'), person, 3000),
    events: mine_(table_(ss, 'Events'), person, 500),
    progress: mine_(table_(ss, 'Progress'), person, 500),
    serverTime: new Date().toISOString()
  };
}

// Append rows (arrays in the tab's column order, without Person), adding the person. Skips ids already present, so retries are safe.
function append_(ss, name, rows, person) {
  if (!ID_COL[name]) throw new Error('cannot append to ' + name);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let fresh = [], result;
  try {
    const sh = ss.getSheetByName(name);
    const idc = ID_COL[name], w = WIDTH[name];
    const last = sh.getLastRow();
    const have = new Set();
    if (last >= 2) {
      const from = Math.max(2, last - 3000);
      sh.getRange(from, idc, last - from + 1, 1).getValues().forEach(r => { if (r[0]) have.add(String(r[0])); });
    }
    fresh = rows.filter(r => r.length === w && r[w - 1] && !have.has(String(r[w - 1])));
    if (fresh.length) sh.getRange(last + 1, 1, fresh.length, w + 1).setValues(fresh.map(r => r.concat([person])));
    result = { added: fresh.length, skipped: rows.length - fresh.length };
  } finally {
    lock.releaseLock();
  }
  // New watch calories: start the instant calorie sync right away (hourly task is the backup)
  if (name === 'Events') {
    const cal = fresh.filter(r => r[2] === 'watch_calories' && String(r[5]).trim() !== '');
    if (cal.length) result.fired = fireCalorieSync_(ss, cal, person);
  }
  return result;
}

function fireCalorieSync_(ss, rows, person) {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('FIRE_URL'), token = props.getProperty('FIRE_TOKEN');
  if (!url || !token) return 'not set up';
  let code, note;
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'Authorization': 'Bearer ' + token, 'anthropic-version': '2023-06-01', 'anthropic-beta': 'experimental-cc-routine-2026-04-01' },
      payload: JSON.stringify({ text: 'New watch calories saved for ' + person + ': ' + rows.map(r => r[5] + ' cal (Row ID ' + r[7] + ')').join(', ') })
    });
    code = res.getResponseCode(); note = code < 300 ? 'started' : String(res.getContentText()).slice(0, 200);
  } catch (err) { code = 'error'; note = String(err).slice(0, 200); }
  // record the attempt so a failed start is visible in the sheet
  ss.getSheetByName('Events').appendRow([now_(), rows[0][1], 'loseit_fire', '', '', String(code), note, 'lf-' + Date.now(), person]);
  return String(code);
}

// Change one exercise setting for one person (for example Tempo after a tempo check). Saved in Exercise Overrides.
const EDITABLE = ['Tempo', 'Sec per Rep', 'Active', 'Pause at Turn (s)', 'Form Tips'];
function updateExercise_(ss, exercise, field, value, person) {
  if (EDITABLE.indexOf(field) < 0) throw new Error('field not editable: ' + field);
  if (field === 'Active' && value !== 'Yes' && value !== 'No') throw new Error('Active must be Yes or No');
  const sh = gSheet_(ss, 'Exercise Overrides', ['Person', 'Exercise', 'Field', 'Value', 'Updated At', 'Row ID']);
  const vals = sh.getDataRange().getValues();
  for (let r = 1; r < vals.length; r++) {
    if (who_(vals[r][0]) === person && vals[r][1] === exercise && vals[r][2] === field) {
      sh.getRange(r + 1, 4, 1, 2).setValues([[value, now_()]]);
      return { updated: true };
    }
  }
  sh.appendRow([person, exercise, field, value, now_(), 'xo-' + Utilities.getUuid().slice(0, 8)]);
  return { updated: true };
}

// ---------- generic actions (same shape as the Food Log script) ----------
function gSheet_(ss, tab, headers) {
  let s = ss.getSheetByName(tab);
  if (!s) {
    if (!headers || !headers.length) throw new Error('no such tab: ' + tab);
    s = ss.insertSheet(tab);
    s.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}
function gHeader_(s) { return s.getRange(1, 1, 1, Math.max(s.getLastColumn(), 1)).getValues()[0].map(String); }
function gIdCol_(h) {
  let i = h.indexOf('EntryID'); if (i < 0) i = h.indexOf('Row ID'); if (i < 0) i = h.indexOf('ID');
  return i; // 0-based, -1 if none
}
function gAddCols_(s, h, keys) {
  keys.filter(k => h.indexOf(k) < 0).forEach(k => {
    if (h.length + 1 > s.getMaxColumns()) s.insertColumnsAfter(s.getMaxColumns(), 1);
    s.getRange(1, h.length + 1).setValue(k).setFontWeight('bold');
    h.push(k);
  });
  return h;
}
function gCell_(v, r) {
  if (typeof v === 'string' && v.charAt(0) === '=') return v.replace(/\{r\}/g, r);
  return v === null || v === undefined ? '' : v;
}

/** read: {tab, person?, limit?} -> rows as objects (display text). A blank Person counts as Josh. */
function gRead_(ss, p) {
  let rows = table_(ss, p.tab);
  if (p.person && rows.length && 'Person' in rows[0]) rows = mine_(rows, p.person);
  if (p.limit) rows = rows.slice(-p.limit);
  return { ok: true, rows: rows };
}

/** append: {tab, rows:[{Header: value}], headers?: [...] (to create the tab)}; skips ids already present */
function gAppend_(ss, p) {
  const rows = p.rows || []; if (!rows.length) return { ok: true, added: 0 };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = gSheet_(ss, p.tab, p.headers || Object.keys(rows[0]));
    const h = gAddCols_(s, gHeader_(s), [...new Set(rows.flatMap(Object.keys))]);
    const ic = gIdCol_(h);
    const have = ic >= 0 && s.getLastRow() > 1 ? new Set(s.getRange(2, ic + 1, s.getLastRow() - 1, 1).getValues().flat().map(String)) : new Set();
    let r = s.getLastRow() + 1, added = 0;
    rows.forEach(o => {
      if (ic >= 0 && o[h[ic]] && have.has(String(o[h[ic]]))) return;
      s.getRange(r, 1, 1, h.length).setValues([h.map(k => gCell_(o[k], r))]);
      r++; added++;
    });
    return { ok: true, added: added };
  } finally { lock.releaseLock(); }
}

/** update: {tab, id, values:{Header: value}} */
function gUpdate_(ss, p) {
  const s = gSheet_(ss, p.tab); let h = gHeader_(s); const ic = gIdCol_(h);
  if (ic < 0) throw new Error('tab has no ID column');
  const hit = s.getRange(2, ic + 1, Math.max(s.getLastRow() - 1, 1), 1).createTextFinder(String(p.id)).matchEntireCell(true).findNext();
  if (!hit) return { ok: false, error: 'not found' };
  const r = hit.getRow();
  h = gAddCols_(s, h, Object.keys(p.values || {}));
  Object.keys(p.values || {}).forEach(k => s.getRange(r, h.indexOf(k) + 1).setValue(gCell_(p.values[k], r)));
  return { ok: true };
}

/** delete: {tab, id} */
function gDelete_(ss, p) {
  const s = gSheet_(ss, p.tab); const ic = gIdCol_(gHeader_(s));
  if (ic < 0) throw new Error('tab has no ID column');
  const hit = s.getRange(2, ic + 1, Math.max(s.getLastRow() - 1, 1), 1).createTextFinder(String(p.id)).matchEntireCell(true).findNext();
  if (hit) s.deleteRow(hit.getRow());
  return { ok: true };
}

// Run once from the editor so Google asks for the 'connect to an external service' permission.
function authorizeOutsideCalls() {
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
}
