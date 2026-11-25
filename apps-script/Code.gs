// Workout Buddy: connection between the app and the Workout Buddy Google Sheet.
// Deployed as a web app (Execute as: me, Who has access: Anyone).
// Keys never live in this repo. Script Properties, made by running newKeys() from the editor:
//   KEYS  {"<key>": {"p": "Josh", "t": "<made>"}}: the person comes from the key, never from the request.
//   KEY / FALLBACK_KEY: the old shared key, still accepted until dropOldKeys() runs (sets LEGACY_OFF).
// Rotate: newKeys() -> open the new setup links on each device -> dropOldKeys(). No key configured = every call refused.
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
const PEOPLE = ['Josh', 'Michelle'];
const APP_URL = 'https://joshdouglasfl-byte.github.io/workout-buddy/';
// Generic actions only reach these tabs (per-person tabs: only the caller's rows)
const GENERIC_TABS = { read: ['Coach Notes'], append: ['Feedback'], update: ['Coach Notes'], delete: [] };

function doGet(e) { return handle_(e.parameter || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad json' }); }
  return handle_(p);
}

// ---------- keys ----------
function keyMap_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('KEYS') || '{}'); } catch (e) { return {}; }
}
/** -> person, or '' when the key is wrong. Fails closed. */
function auth_(p) {
  const k = String(p.key || ''); if (!k) return '';
  const o = keyMap_()[k];
  if (o && PEOPLE.indexOf(o.p) >= 0) return o.p;
  const props = PropertiesService.getScriptProperties();
  const old = props.getProperty('KEY') || FALLBACK_KEY;
  if (old && !props.getProperty('LEGACY_OFF') && k === old) {
    const want = p.person || DEFAULT_PERSON;
    return PEOPLE.indexOf(want) >= 0 ? want : '';
  }
  return '';
}
function newKey_() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); }   // 64 hex chars
/** Run from the editor: adds a fresh key for each person (old keys keep working) and logs each person's setup link. */
function newKeys() {
  const props = PropertiesService.getScriptProperties(), t = new Date().toISOString();
  // The deployed web app URL (Deploy > Manage deployments > Web app URL), saved once as Script Property WEB_APP_URL.
  // Not ScriptApp.getService().getUrl(): run from the editor it gives the /dev test URL, which the app can't reach.
  const url = String(props.getProperty('WEB_APP_URL') || '').trim();
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(url)) {
    Logger.log('No keys made. Add Script Property WEB_APP_URL = the Web app URL from Deploy > Manage deployments (ends in /exec), then run newKeys again.');
    return;
  }
  const keys = keyMap_();
  PEOPLE.forEach(person => {
    const k = newKey_(); keys[k] = { p: person, t: t };
    const blob = Utilities.base64EncodeWebSafe(JSON.stringify({ url: url, key: k })).replace(/=+$/, '');
    Logger.log(`${person} setup link (open on each of ${person}'s devices, then Add to Home Screen):\n${APP_URL}${person.toLowerCase()}/#setup=${blob}`);
  });
  props.setProperty('KEYS', JSON.stringify(keys));
}
/** Run from the editor once every device uses the new links: keeps each person's newest key and turns the old shared key off. */
function dropOldKeys() {
  const props = PropertiesService.getScriptProperties(), m = keyMap_(), newest = {};
  Object.keys(m).forEach(k => { const o = m[k]; if (!newest[o.p] || o.t > m[newest[o.p]].t) newest[o.p] = k; });
  const keep = {}; Object.keys(newest).forEach(person => { keep[newest[person]] = m[newest[person]]; });
  props.setProperty('KEYS', JSON.stringify(keep));
  props.deleteProperty('KEY'); props.setProperty('LEGACY_OFF', '1');
  Logger.log('Old keys off. Kept: ' + Object.values(keep).map(o => o.p + ' ' + o.t).join(', ') + '. You can also blank FALLBACK_KEY in the code.');
}

function handle_(p) {
  const person = auth_(p);
  if (!person) return out_({ ok: false, error: 'bad key' });
  p.person = person;
  if (p.tab && GENERIC_TABS[p.action] && GENERIC_TABS[p.action].indexOf(p.tab) < 0) return out_({ ok: false, error: 'tab not allowed' });
  if (p.action === 'delete') return out_({ ok: false, error: 'tab not allowed' });
  const ss = SpreadsheetApp.openById(SHEET_ID);
  try {
    if (p.action === 'ping') return out_({ ok: true, now: new Date().toISOString() });
    if (p.action === 'state') return out_(Object.assign({ ok: true, person: person }, state_(ss, person)));
    // generic actions (rows as objects keyed by header): new features use these, so the script rarely needs redeploying
    if ((p.action === 'read' || p.action === 'update') && GENERIC_TABS[p.action].indexOf(p.tab) < 0) return out_({ ok: false, error: 'tab not allowed' });
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
    if (fresh.length) sh.getRange(last + 1, 1, fresh.length, w + 1).setValues(fresh.map(r => safeRow_(r).concat([person])));
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
      sh.getRange(r + 1, 4, 1, 2).setValues([[safe_(value), now_()]]);
      return { updated: true };
    }
  }
  sh.appendRow(safeRow_([person, exercise, field, value, now_(), 'xo-' + Utilities.getUuid().slice(0, 8)]));
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
/** Typed text is stored as text: a value starting with = + @ (or - not followed by a number) would otherwise run as a formula. */
function safe_(v) { return typeof v === 'string' && /^(=|\+|@|-(?![\d.]))/.test(v) ? "'" + v : v; }
const safeRow_ = row => row.map(safe_);
/** Row r (1-based) belongs to person, or the tab has no Person column. Blank Person = Josh. */
function ownsRow_(s, r, person) {
  const pi = gHeader_(s).indexOf('Person'); if (pi < 0) return true;
  return who_(s.getRange(r, pi + 1).getValue()) === person;
}
// A formula must be sent as {f: '=...'} ("{r}" becomes the row); formulas that reach the web are refused.
function gCell_(v, r) {
  if (v && typeof v === 'object' && typeof v.f === 'string' && v.f.charAt(0) === '=') {
    if (/IMPORT|IMAGE|HYPERLINK|WEBSERVICE/i.test(v.f)) throw new Error('formula not allowed');
    return v.f.replace(/\{r\}/g, r);
  }
  return v === null || v === undefined ? '' : safe_(v);
}

/** read: {tab, person?, limit?} -> rows as objects (display text). A blank Person counts as Josh. */
function gRead_(ss, p) {
  let rows = table_(ss, p.tab);
  if (rows.length && 'Person' in rows[0]) rows = mine_(rows, p.person);
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
    if (h.indexOf('Person') >= 0) rows.forEach(o => { o.Person = p.person; });
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
  if (!hit || !ownsRow_(s, hit.getRow(), p.person)) return { ok: false, error: 'not found' };
  const r = hit.getRow();
  if ('Person' in (p.values || {})) return { ok: false, error: 'Person cannot be changed' };
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
