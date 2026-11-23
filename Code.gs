/**
 * Heart & Lifestyle Health Check – backend (Google Apps Script)
 * Dr. Smita Karpe – Madhavbaug Clinic, Airoli
 *
 * Stores every patient result in the Google Sheet identified by SHEET_ID below.
 * The dashboard login (ADMIN_USER / ADMIN_PASS) is checked HERE, on Google's server.
 * Keep this file private: do not paste it anywhere public or into the website.
 * After ANY edit: Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
 */
var SHEET_ID   = '1AtbbZZ0KjLMmU4oQ4t1q1dzph2_x78cnU_yP46TxU1s';   // from your sheet link: /spreadsheets/d/<THIS PART>/edit
var ADMIN_USER = 'admin';
var ADMIN_PASS = 'clinic123';
var SHEET_NAME = 'Records';
var HEADERS = ['id','ts','date','name','age','sex','phone','bmi','waist','score','level','chest_flag','lang',
               'body','activity','diet','habits','mind','medical','level_label'];
var TEXT_COLS = [1, 3, 4, 7];            // id, date, name, phone: stored as plain text (blocks formula injection)
var CATS = ['body','activity','diet','habits','mind','medical'];
var LEVELS = ['Low risk','Moderate risk','High risk','Very high risk'];
var MAX_ROWS = 5000;
var TOKEN_TTL = 6 * 60 * 60;             // seconds (CacheService maximum)
var MAX_FAILS = 8;                       // wrong logins allowed per 10 minutes

/** Run this once from the editor to authorise the script and create the sheet. */
function setup() { sheet_(); }

function doGet() { return out_({ ok: true, service: 'health-check' }); }

function doPost(e) {
  try {
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    switch (d.action) {
      case 'save':   return out_(save_(d.rec));
      case 'login':  return out_(login_(d.user, d.pass));
      case 'logout': CacheService.getScriptCache().remove('tok_' + String(d.token || '')); return out_({ ok: true });
      case 'list':
      case 'delete':
      case 'clear':
        if (!valid_(d.token)) return out_({ ok: false, error: 'auth' });
        if (d.action === 'list') return out_(list_());
        if (d.action === 'delete') return out_(delete_(d.id));
        return out_(clear_());
    }
    return out_({ ok: false, error: 'bad_request' });
  } catch (err) {
    return out_({ ok: false, error: 'server' });
  }
}

/* ---------- helpers ---------- */
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_() {
  var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    TEXT_COLS.forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}
function same_(a, b) {                   // constant-time string compare
  a = String(a || ''); b = String(b || '');
  var r = a.length ^ b.length;
  for (var i = 0; i < Math.min(a.length, b.length); i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}
function str_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, max); }
function num_(v, min, max) { v = Number(v); return (isFinite(v) && v >= min && v <= max) ? v : null; }

/* ---------- auth ---------- */
function login_(user, pass) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) return { ok: false, error: 'locked' };
  if (same_(user, ADMIN_USER) && same_(pass, ADMIN_PASS)) {
    var t = Utilities.getUuid();
    cache.put('tok_' + t, '1', TOKEN_TTL);
    cache.remove('fails');
    return { ok: true, token: t };
  }
  cache.put('fails', String(fails + 1), 600);
  return { ok: false, error: 'login' };
}
function valid_(token) {
  token = String(token || '');
  return token.length >= 30 && !!CacheService.getScriptCache().get('tok_' + token);
}

/* ---------- save (public) ---------- */
function clean_(r) {
  if (!r || typeof r !== 'object') return null;
  var id = str_(r.id, 40);
  if (!/^[A-Za-z0-9]{4,40}$/.test(id)) return null;
  var name = str_(r.name, 80);
  if (!name) return null;
  var age = num_(r.age, 10, 100), bmi = num_(r.bmi, 0, 100), waist = num_(r.waist, 40, 200);
  var score = num_(r.score, 0, 100), level = num_(r.level, 0, 3);
  if ([age, bmi, waist, score, level].indexOf(null) > -1) return null;
  var sex = r.sex === 'm' ? 'm' : r.sex === 'f' ? 'f' : null;
  if (!sex) return null;
  var cats = {};
  for (var i = 0; i < CATS.length; i++) {
    var v = num_(r.cats && r.cats[CATS[i]], 0, 100);
    if (v === null) return null;
    cats[CATS[i]] = v;
  }
  return { id: id, name: name, age: age, sex: sex, phone: str_(r.phone, 20).replace(/[^0-9+ ]/g, ''),
           bmi: bmi, waist: waist, score: score, level: Math.round(level), flag: !!r.flag,
           lang: r.lang === 'mr' ? 'mr' : 'en', cats: cats };
}
function save_(r) {
  var rec = clean_(r);
  if (!rec) return { ok: false, error: 'invalid' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(), n = sh.getLastRow() - 1;
    if (n >= MAX_ROWS) return { ok: false, error: 'full' };
    if (n > 0) {                          // ignore duplicates (client retries)
      var ids = sh.getRange(2, 1, n, 1).getValues();
      for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === rec.id) return { ok: true, dup: true };
    }
    var now = new Date();
    var row = [rec.id, now.getTime(), Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd HH:mm'),
               rec.name, rec.age, rec.sex, rec.phone, rec.bmi, rec.waist, rec.score, rec.level,
               rec.flag ? 'yes' : 'no', rec.lang];
    CATS.forEach(function (c) { row.push(rec.cats[c]); });
    row.push(LEVELS[rec.level]);
    sh.appendRow(row);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------- admin (token required) ---------- */
function list_() {
  var sh = sheet_(), n = sh.getLastRow() - 1, recs = [];
  if (n > 0) {
    sh.getRange(2, 1, n, HEADERS.length).getValues().forEach(function (v) {
      if (!v[0]) return;
      var cats = {};
      CATS.forEach(function (c, i) { cats[c] = Number(v[13 + i]) || 0; });
      recs.push({ id: String(v[0]), ts: Number(v[1]) || 0, name: String(v[3]), age: Number(v[4]) || 0,
                  sex: String(v[5]), phone: String(v[6] || ''), bmi: Number(v[7]) || 0, waist: Number(v[8]) || 0,
                  score: Number(v[9]) || 0, level: Number(v[10]) || 0, flag: v[11] === 'yes',
                  lang: String(v[12] || 'en'), cats: cats });
    });
  }
  recs.sort(function (a, b) { return b.ts - a.ts; });
  return { ok: true, records: recs };
}
function delete_(id) {
  id = String(id || '');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(), n = sh.getLastRow() - 1;
    if (n > 0) {
      var ids = sh.getRange(2, 1, n, 1).getValues();
      for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) { sh.deleteRow(i + 2); break; }
    }
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function clear_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(), n = sh.getLastRow() - 1;
    if (n > 0) sh.getRange(2, 1, n, HEADERS.length).clearContent();
    return { ok: true };
  } finally { lock.releaseLock(); }
}
