/**
 * מעקב הזמנות מלונות — Backend (Google Apps Script)
 * גיליון: "מעקב הזמנות מלונות" (חשבון zachi.daniel2@gmail.com)
 *
 * התקנה (פעם אחת):
 *  1. בגיליון: תוספים → Apps Script → להדביק את הקובץ הזה → שמירה.
 *  2. להריץ את setup() (יוצר את הטאבים "הזמנות" + "היסטוריה" ומייצר API_TOKEN).
 *     ה-token מודפס ב-Logger (תצוגה → יומן ביצוע) וגם נשמר ב-Project Settings → Script properties.
 *  3. פריסה → פריסה חדשה → אפליקציית אינטרנט: Execute as Me, Who has access: Anyone → להעתיק את ה-URL.
 *  4. תזכורות טלגרם לפני סיום ביטול חינם: Script properties TELEGRAM_BOT_TOKEN + TELEGRAM_CHAT_ID (+ אופציונלי REMINDER_DAYS, ברירת מחדל 14,7,3,1,0)
 *     → להריץ installTriggers() פעם אחת (09:00 יומי) → sendTestReminder() לבדיקה. האפליקציה מראה 🔔 כשזה פעיל.
 *  ⚠ אחרי כל שינוי בקוד: פריסה → ניהול פריסות → עריכה → גרסה חדשה. שמירה בלבד לא מעדכנת את ה-URL.
 */

var SHEET_BOOKINGS = 'הזמנות';
var SHEET_HISTORY = 'היסטוריה';
var SHEET_ARCHIVE = 'ארכיון';

// [key, header, type]  type: text | date | num | bool
var COLS = [
  ['id', 'מזהה', 'text'],
  ['hotel', 'מלון', 'text'],
  ['destination', 'יעד', 'text'],
  ['platform', 'פלטפורמה', 'text'],
  ['account', 'יוזר', 'text'],
  ['checkIn', 'צ׳ק-אין', 'date'],
  ['checkOut', 'צ׳ק-אאוט', 'date'],
  ['nights', 'לילות', 'num'],
  ['roomType', 'סוג החדר', 'text'],
  ['breakfast', 'ארוחת בוקר', 'bool'],
  ['price', 'מחיר במטבע ששולם', 'num'],
  ['currency', 'מטבע', 'text'],
  ['priceIls', 'המרה לשקל', 'num'],
  ['freeCancelUntil', 'ביטול חינם עד', 'date'],
  ['paid', 'שולם', 'bool'],
  ['confirmation', 'מספר הזמנה', 'text'],
  ['link', 'קישור', 'text'],
  ['notes', 'הערות', 'text'],
  ['source', 'מקור', 'text'],
  ['createdAt', 'נוצר', 'text'],
  ['updatedAt', 'עודכן', 'text']
];
var HISTORY_COLS = [['ts', 'זמן', 'text'], ['action', 'פעולה', 'text'], ['by', 'מי', 'text'], ['note', 'פירוט', 'text']].concat(COLS);
// ארכיון: הזמנות שנמחקו/הוחלפו. אפשר לשחזר או למחוק לצמיתות. (ההיסטוריה היא יומן בלבד.)
var ARCHIVE_COLS = COLS.concat([['archivedAt', 'הועבר לארכיון', 'text'], ['archiveAction', 'סיבה', 'text'], ['archivedBy', 'ע"י', 'text'], ['archiveNote', 'הערה', 'text']]);

var PLATFORMS = ['Booking', 'Agoda', 'Expedia', 'Hotels.com', 'Trip.com', 'Airbnb', 'ישירות מול המלון', 'אחר'];

// ---------- setup ----------
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var bookings = ss.getSheetByName(SHEET_BOOKINGS);
  if (!bookings) {
    // אם יש רק גיליון אחד ריק (רק כותרות) — משתמשים בו במקום ליצור חדש
    var first = sheets[0];
    if (sheets.length === 1 && first.getLastRow() <= 1) { bookings = first; bookings.setName(SHEET_BOOKINGS); bookings.clear(); }
    else bookings = ss.insertSheet(SHEET_BOOKINGS, 0);
  }
  writeHeader_(bookings, COLS);
  var history = ss.getSheetByName(SHEET_HISTORY) || ss.insertSheet(SHEET_HISTORY);
  writeHeader_(history, HISTORY_COLS);
  var archive = ss.getSheetByName(SHEET_ARCHIVE) || ss.insertSheet(SHEET_ARCHIVE);
  writeHeader_(archive, ARCHIVE_COLS);

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('API_TOKEN')) props.setProperty('API_TOKEN', Utilities.getUuid().replace(/-/g, ''));
  Logger.log('API_TOKEN: ' + props.getProperty('API_TOKEN'));
  Logger.log('הטאבים מוכנים. עכשיו: פריסה → פריסה חדשה → אפליקציית אינטרנט.');
}

function writeHeader_(sh, cols) {
  var headers = cols.map(function (c) { return c[1]; });
  var existing = sh.getLastRow() >= 1 ? sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0] : [];
  var same = existing.length >= headers.length && headers.every(function (h, i) { return existing[i] === h; });
  if (!same && sh.getLastRow() > 1) throw new Error('הטאב "' + sh.getName() + '" מכיל נתונים עם כותרות שונות. גבה אותו/מחק לפני setup().');
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8f0fe');
  sh.setFrozenRows(1);
  sh.setRightToLeft(true);
  // עמודות טקסט לתאריכים ומזהים כדי שהגיליון לא "יתקן" אותם
  cols.forEach(function (c, i) {
    if (c[2] === 'date' || c[2] === 'text') sh.getRange(2, i + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  });
  sh.autoResizeColumns(1, headers.length);
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyReminders') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyReminders').timeBased().everyDays(1).atHour(9).create();
  Logger.log('טריגר יומי (09:00) לתזכורות ביטול הותקן.');
}

// ---------- helpers ----------
function props_() { return PropertiesService.getScriptProperties(); }
function checkToken_(t) { var s = props_().getProperty('API_TOKEN'); return !!s && t === s; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function now_() { return Utilities.formatDate(new Date(), 'Asia/Jerusalem', "yyyy-MM-dd'T'HH:mm:ss"); }
function newId_() { return Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyMMddHHmmss') + Math.random().toString(36).slice(2, 6); }

function normDate_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Jerusalem', 'yyyy-MM-dd');
  var s = String(v).trim();
  var m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return m[1] + '-' + pad_(m[2]) + '-' + pad_(m[3]);
  if ((m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})$/))) { var y = m[3].length === 2 ? '20' + m[3] : m[3]; return y + '-' + pad_(m[2]) + '-' + pad_(m[1]); }
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'Asia/Jerusalem', 'yyyy-MM-dd');
  return s;
}
function pad_(n) { n = String(n); return n.length < 2 ? '0' + n : n; }
function normBool_(v) {
  if (typeof v === 'boolean') return v;
  var s = String(v === undefined || v === null ? '' : v).trim().toLowerCase();
  return ['true', 'yes', 'y', '1', 'כן', 'v', 'כולל', 'included'].indexOf(s) >= 0;
}
function normNum_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v).replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? '' : n;
}
function nightsBetween_(a, b) {
  if (!a || !b) return '';
  var d1 = new Date(a + 'T00:00:00Z'), d2 = new Date(b + 'T00:00:00Z');
  var n = Math.round((d2 - d1) / 86400000);
  return n > 0 ? n : '';
}
function normPlatform_(p) {
  var s = String(p || '').trim().toLowerCase();
  if (!s) return '';
  if (s.indexOf('booking') >= 0 || s.indexOf('בוקינג') >= 0) return 'Booking';
  if (s.indexOf('agoda') >= 0 || s.indexOf('אגודה') >= 0) return 'Agoda';
  if (s.indexOf('expedia') >= 0 || s.indexOf('אקספדיה') >= 0) return 'Expedia';
  if (s.indexOf('hotels.com') >= 0 || s === 'hotels') return 'Hotels.com';
  if (s.indexOf('trip.com') >= 0 || s === 'trip') return 'Trip.com';
  if (s.indexOf('airbnb') >= 0) return 'Airbnb';
  if (s.indexOf('ישיר') >= 0 || s.indexOf('direct') >= 0) return 'ישירות מול המלון';
  return String(p).trim();
}

/** מקבל אובייקט חופשי ומחזיר רשומה מנורמלת ומלאה */
function normalize_(b, existing) {
  var r = {};
  COLS.forEach(function (c) { r[c[0]] = existing ? existing[c[0]] : ''; });
  COLS.forEach(function (c) {
    var k = c[0];
    if (b[k] === undefined) return;
    var v = b[k];
    if (c[2] === 'date') v = normDate_(v);
    else if (c[2] === 'bool') v = normBool_(v);
    else if (c[2] === 'num') v = normNum_(v);
    else v = v === null ? '' : String(v).trim();
    r[k] = v;
  });
  r.platform = normPlatform_(r.platform);
  r.currency = String(r.currency || 'ILS').toUpperCase().replace('₪', 'ILS').replace('NIS', 'ILS').replace('$', 'USD').replace('€', 'EUR').replace('£', 'GBP').replace('฿', 'THB').trim();
  if (r.checkIn && r.checkOut) r.nights = nightsBetween_(r.checkIn, r.checkOut);
  else if (r.checkIn && r.nights && !r.checkOut) { var d = new Date(r.checkIn + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + Number(r.nights)); r.checkOut = d.toISOString().slice(0, 10); }
  // המרה לשקל: ידני גובר; אחרת מחשבים כשאין ערך או כשהמחיר/המטבע השתנו
  var manualIls = b.priceIls !== undefined && b.priceIls !== null && String(b.priceIls) !== '';
  var priceChanged = !existing || String(existing.price) !== String(r.price) || String(existing.currency) !== String(r.currency);
  if (!manualIls && r.price !== '' && (r.priceIls === '' || priceChanged)) {
    if (r.currency === 'ILS') r.priceIls = r.price;
    else { var rate = fxRate_(r.currency); r.priceIls = rate ? Math.round(r.price * rate) : ''; }
  }
  return r;
}

var FX_CACHE_ = {};
function fxRate_(cur) {
  if (!cur || cur === 'ILS') return 1;
  if (FX_CACHE_[cur]) return FX_CACHE_[cur];
  try {
    var cache = CacheService.getScriptCache();
    var hit = cache.get('fx_' + cur);
    if (hit) return FX_CACHE_[cur] = parseFloat(hit);
    var res = UrlFetchApp.fetch('https://api.frankfurter.app/latest?from=' + encodeURIComponent(cur) + '&to=ILS', { muteHttpExceptions: true });
    var data = JSON.parse(res.getContentText());
    var rate = data && data.rates && data.rates.ILS;
    if (rate) { cache.put('fx_' + cur, String(rate), 21600); return FX_CACHE_[cur] = rate; }
  } catch (e) { }
  return null;
}

function bookingsSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BOOKINGS); }
function historySheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HISTORY); }
function archiveSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_ARCHIVE);
  if (!sh) { sh = ss.insertSheet(SHEET_ARCHIVE); writeHeader_(sh, ARCHIVE_COLS); }
  return sh;
}
function archive_(rec, action, by, note) {
  var row = {}; ARCHIVE_COLS.forEach(function (c) { row[c[0]] = rec[c[0]] === undefined ? '' : rec[c[0]]; });
  row.archivedAt = now_(); row.archiveAction = action; row.archivedBy = by || ''; row.archiveNote = note || '';
  archiveSheet_().appendRow(toRow_(row, ARCHIVE_COLS));
}
function listArchive_() { return readAll_(archiveSheet_(), ARCHIVE_COLS); }
function findArchived_(id) {
  var all = listArchive_();
  for (var i = all.length - 1; i >= 0; i--) if (String(all[i].id) === String(id)) return all[i];
  return null;
}

function readAll_(sh, cols) {
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, cols.length).getValues();
  var out = [];
  values.forEach(function (row, i) {
    var r = { _row: i + 2 }, empty = true;
    cols.forEach(function (c, j) {
      var v = row[j];
      if (c[2] === 'date') v = normDate_(v);
      else if (c[2] === 'bool') v = normBool_(v);
      else if (v instanceof Date) v = Utilities.formatDate(v, 'Asia/Jerusalem', ['ts', 'createdAt', 'updatedAt'].indexOf(c[0]) >= 0 ? "yyyy-MM-dd'T'HH:mm:ss" : 'yyyy-MM-dd');
      r[c[0]] = v;
      if (v !== '' && v !== null && v !== false) empty = false;
    });
    if (!empty) out.push(r);
  });
  return out;
}
function toRow_(r, cols) { return cols.map(function (c) { var v = r[c[0]]; if (v === undefined || v === null) return ''; if (c[2] === 'bool') return v ? 'כן' : 'לא'; return v; }); }

function listBookings_() { return readAll_(bookingsSheet_(), COLS).map(function (r) { delete r._row; return r; }); }

function findRow_(id) {
  var all = readAll_(bookingsSheet_(), COLS);
  for (var i = 0; i < all.length; i++) if (String(all[i].id) === String(id)) return all[i];
  return null;
}

function logHistory_(action, by, note, rec) {
  var h = historySheet_();
  var row = { ts: now_(), action: action, by: by || '', note: note || '' };
  COLS.forEach(function (c) { row[c[0]] = rec ? rec[c[0]] : ''; });
  h.appendRow(toRow_(row, HISTORY_COLS));
}

function appendBooking_(rec) { bookingsSheet_().appendRow(toRow_(rec, COLS)); }
function writeBooking_(rowNum, rec) { bookingsSheet_().getRange(rowNum, 1, 1, COLS.length).setValues([toRow_(rec, COLS)]); }
function deleteRow_(rowNum) { bookingsSheet_().deleteRow(rowNum); }

function remindersStatus_() {
  var token = props_().getProperty('TELEGRAM_BOT_TOKEN'), chat = props_().getProperty('TELEGRAM_CHAT_ID');
  var days = (props_().getProperty('REMINDER_DAYS') || '14,7,3,1,0').split(',').map(function (x) { return Number(x.trim()); }).filter(function (n) { return !isNaN(n); });
  var installed = false;
  try { installed = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'dailyReminders'; }); } catch (e) { }
  return { configured: !!(token && chat), triggerInstalled: installed, active: !!(token && chat) && installed, days: days, hour: 9 };
}

function summary_(list) {
  var nights = 0, ils = 0, missingIls = 0;
  list.forEach(function (r) { nights += Number(r.nights) || 0; if (r.priceIls !== '' && r.priceIls !== null) ils += Number(r.priceIls) || 0; else if (r.price !== '') missingIls++; });
  return { count: list.length, nights: nights, totalIls: Math.round(ils), missingIls: missingIls, targetNights: Number(props_().getProperty('TARGET_NIGHTS') || 21) };
}

function norm_(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9֐-׿]+/g, ' ').trim(); }
function matchCandidates_(all, b) {
  var hotel = norm_(b.hotel);
  if (!hotel) return [];
  var words = hotel.split(' ').filter(function (w) { return w.length > 2 && ['hotel', 'resort', 'the', 'and', 'villa', 'villas', 'boutique', 'spa', 'מלון'].indexOf(w) < 0; });
  return all.filter(function (r) {
    var h = norm_(r.hotel);
    if (h === hotel) return true;
    if (h.indexOf(hotel) >= 0 || hotel.indexOf(h) >= 0) return true;
    var hits = words.filter(function (w) { return h.indexOf(w) >= 0; }).length;
    return words.length > 0 && hits >= Math.max(1, Math.ceil(words.length * 0.6));
  });
}

// ---------- HTTP ----------
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'ping') return json_({ ok: true, ts: now_() });
  if (!checkToken_(p.token)) return json_({ ok: false, error: 'טוקן שגוי' });
  return handle_(p.action || 'list', p, 'web');
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'JSON לא תקין' }); }
  if (!checkToken_(body.token)) return json_({ ok: false, error: 'טוקן שגוי' });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return handle_(body.action, body, body.source || body.by || 'api'); }
  catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
  finally { lock.releaseLock(); }
}

function handle_(action, p, by) {
  by = p.by || p.source || by;
  switch (action) {
    case 'list': {
      var list = listBookings_();
      return json_({ ok: true, bookings: list, summary: summary_(list), platforms: PLATFORMS, reminders: remindersStatus_() });
    }
    case 'history': {
      var h = readAll_(historySheet_(), HISTORY_COLS).map(function (r) { delete r._row; return r; }).reverse();
      return json_({ ok: true, history: h.slice(0, Number(p.limit) || 300) });
    }
    case 'get': {
      var rec = findRow_(p.id); if (!rec) return json_({ ok: false, error: 'לא נמצא' });
      delete rec._row; return json_({ ok: true, booking: rec });
    }
    case 'add': {
      var b = p.booking || p; if (!b.hotel) return json_({ ok: false, error: 'חסר שם מלון' });
      var replaced = null;
      if (p.replaceId) { var old = findRow_(p.replaceId); if (!old) return json_({ ok: false, error: 'replaceId לא נמצא' }); deleteRow_(old._row); delete old._row; replaced = old; }
      var rec = normalize_(b); rec.id = newId_(); rec.source = by; rec.createdAt = now_(); rec.updatedAt = rec.createdAt;
      if (replaced) { var why = 'הוחלף בהזמנה ' + rec.id + ' (' + rec.platform + ' ' + rec.price + ' ' + rec.currency + ')'; archive_(replaced, 'הוחלף', by, why); logHistory_('הוחלף', by, why, replaced); }
      appendBooking_(rec);
      logHistory_(replaced ? 'נוסף (החלפה)' : 'נוסף', by, replaced ? 'החליף את ' + replaced.id : (p.note || ''), rec);
      return json_({ ok: true, booking: rec, replaced: replaced, summary: summary_(listBookings_()) });
    }
    case 'upsert': {
      // לרוברט: מחפש הזמנה קיימת לאותו מלון. 0 → מוסיף. 1 → מחליף. יותר → מבקש replaceId.
      var b = p.booking || p; if (!b.hotel) return json_({ ok: false, error: 'חסר שם מלון' });
      var all = readAll_(bookingsSheet_(), COLS);
      var cands = p.replaceId ? all.filter(function (r) { return String(r.id) === String(p.replaceId); }) : matchCandidates_(all, b);
      if (cands.length > 1) {
        return json_({ ok: false, needsChoice: true, error: 'נמצאו כמה הזמנות דומות — שלח replaceId', candidates: cands.map(function (r) { delete r._row; return r; }) });
      }
      if (cands.length === 1 && p.replaceId === undefined && p.confirmReplace === false) {
        var c0 = cands[0]; delete c0._row;
        return json_({ ok: false, needsConfirm: true, error: 'קיימת הזמנה דומה — שלח confirmReplace:true או replaceId', candidate: c0 });
      }
      var q = { token: p.token, booking: b, by: by, note: p.note };
      if (cands.length === 1) q.replaceId = cands[0].id;
      return handle_('add', q, by);
    }
    case 'update': {
      var old = findRow_(p.id); if (!old) return json_({ ok: false, error: 'לא נמצא' });
      var rowNum = old._row; delete old._row;
      var b = p.booking || p.fields || p;
      var rec = normalize_(b, old); rec.id = old.id; rec.createdAt = old.createdAt; rec.source = old.source || by; rec.updatedAt = now_();
      var changed = COLS.filter(function (c) { return String(old[c[0]]) !== String(rec[c[0]]) && c[0] !== 'updatedAt'; }).map(function (c) { return c[1] + ': ' + old[c[0]] + ' → ' + rec[c[0]]; });
      logHistory_('עודכן', by, changed.join(' | '), old);
      writeBooking_(rowNum, rec);
      return json_({ ok: true, booking: rec, changes: changed });
    }
    case 'delete': {
      var old = findRow_(p.id); if (!old) return json_({ ok: false, error: 'לא נמצא' });
      deleteRow_(old._row); delete old._row;
      archive_(old, 'נמחק', by, p.note || p.reason || '');
      logHistory_('נמחק (לארכיון)', by, p.note || p.reason || '', old);
      return json_({ ok: true, deleted: old, archived: true, summary: summary_(listBookings_()) });
    }
    case 'archive': {
      var a = listArchive_().map(function (r) { delete r._row; return r; }).reverse();
      return json_({ ok: true, archive: a });
    }
    case 'restore': {
      // מחזיר הזמנה מהארכיון לרשימה הפעילה
      var src = findArchived_(p.id); if (!src) return json_({ ok: false, error: 'לא נמצא בארכיון' });
      if (findRow_(p.id)) return json_({ ok: false, error: 'ההזמנה כבר פעילה' });
      var rec = {}; COLS.forEach(function (c) { rec[c[0]] = src[c[0]]; }); rec.updatedAt = now_();
      appendBooking_(rec);
      archiveSheet_().deleteRow(src._row);
      logHistory_('שוחזר מהארכיון', by, '', rec);
      return json_({ ok: true, booking: rec, summary: summary_(listBookings_()) });
    }
    case 'purge': {
      // מחיקה לצמיתות מהארכיון (נשאר רישום ביומן ההיסטוריה בלבד)
      var src = findArchived_(p.id); if (!src) return json_({ ok: false, error: 'לא נמצא בארכיון' });
      archiveSheet_().deleteRow(src._row); delete src._row;
      logHistory_('נמחק לצמיתות', by, p.note || '', src);
      return json_({ ok: true, purged: src });
    }
    case 'setTarget': { props_().setProperty('TARGET_NIGHTS', String(Number(p.nights) || 21)); return json_({ ok: true }); }
    default: return json_({ ok: false, error: 'פעולה לא מוכרת: ' + action });
  }
}

// ---------- תזכורות טלגרם (אופציונלי) ----------
function dailyReminders() {
  var token = props_().getProperty('TELEGRAM_BOT_TOKEN'), chat = props_().getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chat) return;
  var today = new Date(Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd') + 'T00:00:00Z');
  var thresholds = remindersStatus_().days;
  var lines = [];
  listBookings_().forEach(function (r) {
    if (!r.freeCancelUntil) return;
    var d = new Date(r.freeCancelUntil + 'T00:00:00Z');
    var days = Math.round((d - today) / 86400000);
    if (thresholds.indexOf(days) >= 0) {
      lines.push((days === 0 ? '🔴 היום' : days === 1 ? '🟠 מחר' : '⏰ בעוד ' + days + ' ימים') + ' נגמר הביטול החינמי: *' + r.hotel + '* (' + r.platform + (r.account ? ', יוזר ' + r.account : '') + ', ' + r.destination + ')\n' +
        r.checkIn + ' → ' + r.checkOut + ' · ' + r.nights + ' לילות · ' + (r.priceIls ? r.priceIls + ' ₪' : r.price + ' ' + r.currency) + (r.confirmation ? ' · #' + r.confirmation : ''));
    }
  });
  if (!lines.length) return;
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ chat_id: chat, text: lines.join('\n\n'), parse_mode: 'Markdown' })
  });
}

/** בדיקה: שולח הודעת טלגרם מיידית (דורש TELEGRAM_BOT_TOKEN + TELEGRAM_CHAT_ID ב-Script properties) */
function sendTestReminder() {
  var token = props_().getProperty('TELEGRAM_BOT_TOKEN'), chat = props_().getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chat) throw new Error('חסר TELEGRAM_BOT_TOKEN או TELEGRAM_CHAT_ID ב-Script properties');
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ chat_id: chat, text: '🏨 תזכורות ביטול חינם פעילות. ימים לפני: ' + remindersStatus_().days.join(', ') + ' (בשעה 09:00).' })
  });
  Logger.log(res.getContentText());
}

// בדיקה מהירה בעורך
function testAdd() {
  var res = handle_('add', { booking: { hotel: 'Test Hotel', destination: 'Bangkok', platform: 'booking', checkIn: '2027-08-10', checkOut: '2027-08-13', price: 300, currency: 'USD', breakfast: 'כן', freeCancelUntil: '2027-08-01' } }, 'test');
  Logger.log(res.getContent());
}
