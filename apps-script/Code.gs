// Workout Buddy: connection between the iPad app and the Workout Buddy Google Sheet.
// Deployed as a web app (Execute as: me, Who has access: Anyone).
// The key lives only in the deployed copy (FALLBACK_KEY or Script Property KEY); this public copy holds no secrets.

const FALLBACK_KEY = '';   // filled in only in the deployed copy, never in this repo
const SHEET_ID = '1hmU_XKOhSxIM_eculF_hw0agBoI6TKY7EnNb10x5go0';
const ID_COL = { Sets: 18, Events: 8, Progress: 5 };   // column that holds each row's unique id
const WIDTH = { Sets: 18, Events: 8, Progress: 5 };

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
  try {
    if (p.action === 'ping') return out_({ ok: true, now: new Date().toISOString() });
    if (p.action === 'state') return out_(Object.assign({ ok: true }, state_(ss)));
    if (p.action === 'append') return out_(Object.assign({ ok: true }, append_(ss, p.sheet, p.rows || [])));
    if (p.action === 'updateExercise') return out_(Object.assign({ ok: true }, updateExercise_(ss, p.exercise, p.field, p.value)));
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Read a tab as an array of objects keyed by header, using displayed text.
function table_(ss, name, maxRows) {
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const cols = sh.getLastColumn();
  const start = maxRows ? Math.max(2, last - maxRows + 1) : 2;
  const head = sh.getRange(1, 1, 1, cols).getDisplayValues()[0];
  const vals = sh.getRange(start, 1, last - start + 1, cols).getDisplayValues();
  return vals.filter(r => r.some(c => c !== '')).map(r => {
    const o = {};
    head.forEach((h, i) => { if (h) o[h] = r[i]; });
    return o;
  });
}

function state_(ss) {
  const settings = {};
  table_(ss, 'Settings').forEach(r => { settings[r['Setting']] = r['Value']; });
  return {
    exercises: table_(ss, 'Exercises'),
    settings: settings,
    plan: table_(ss, 'Plan'),
    sets: table_(ss, 'Sets', 3000),
    events: table_(ss, 'Events', 500),
    progress: table_(ss, 'Progress', 500),
    serverTime: new Date().toISOString()
  };
}

// Append rows, skipping any whose id (last column) is already in the tab. Safe to retry.
function append_(ss, name, rows) {
  if (!ID_COL[name]) throw new Error('cannot append to ' + name);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName(name);
    const idc = ID_COL[name], w = WIDTH[name];
    const last = sh.getLastRow();
    const have = new Set();
    if (last >= 2) {
      const from = Math.max(2, last - 2000);
      sh.getRange(from, idc, last - from + 1, 1).getValues().forEach(r => { if (r[0]) have.add(String(r[0])); });
    }
    const fresh = rows.filter(r => r.length === w && r[w - 1] && !have.has(String(r[w - 1])));
    if (fresh.length) {
      const range = sh.getRange(last + 1, 1, fresh.length, w);
      range.setValues(fresh);
    }
    var result = { added: fresh.length, skipped: rows.length - fresh.length };
  } finally {
    lock.releaseLock();
  }
  // New watch calories: start the instant Lose It! sync right away (hourly task is the backup)
  if (name === 'Events') {
    const cal = fresh.filter(r => r[2] === 'watch_calories' && String(r[5]).trim() !== '');
    if (cal.length) result.fired = fireLoseItSync_(ss, cal);
  }
  return result;
}

function fireLoseItSync_(ss, rows) {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('FIRE_URL'), token = props.getProperty('FIRE_TOKEN');
  if (!url || !token) return 'not set up';
  let code, note;
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'Authorization': 'Bearer ' + token, 'anthropic-version': '2023-06-01', 'anthropic-beta': 'experimental-cc-routine-2026-04-01' },
      payload: JSON.stringify({ text: 'New watch calories saved: ' + rows.map(r => r[5] + ' cal (Row ID ' + r[7] + ')').join(', ') })
    });
    code = res.getResponseCode(); note = code < 300 ? 'started' : String(res.getContentText()).slice(0, 200);
  } catch (err) { code = 'error'; note = String(err).slice(0, 200); }
  // record the attempt so a failed start is visible in the sheet
  ss.getSheetByName('Events').appendRow([Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd HH:mm:ss'), rows[0][1], 'loseit_fire', '', '', String(code), note, 'lf-' + Date.now()]);
  return String(code);
}

// Change one field of one exercise (for example Tempo after a tempo check).
const EDITABLE = ['Tempo', 'Sec per Rep'];
function updateExercise_(ss, exercise, field, value) {
  if (EDITABLE.indexOf(field) < 0) throw new Error('field not editable: ' + field);
  const sh = ss.getSheetByName('Exercises');
  const vals = sh.getDataRange().getValues();
  const col = vals[0].indexOf(field);
  for (let r = 1; r < vals.length; r++) {
    if (vals[r][0] === exercise) { sh.getRange(r + 1, col + 1).setValue(value); return { updated: true }; }
  }
  return { updated: false };
}

