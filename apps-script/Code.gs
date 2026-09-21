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
  ['paymentDate', 'מועד חיוב', 'date'],
  ['confirmation', 'מספר הזמנה', 'text'],
  ['link', 'קישור', 'text'],
  ['notes', 'הערות', 'text'],
  ['group', 'חדרים במקביל', 'text'],
  ['extraPrice', 'תוספת במלון', 'num'],
  ['extraNote', 'פירוט התוספת', 'text'],
  ['extraIls', 'תוספת בשקלים', 'num'],
  ['source', 'מקור', 'text'],
  ['createdAt', 'נוצר', 'text'],
  ['updatedAt', 'עודכן', 'text']
];
var SHEET_EXPENSES = 'הוצאות';
var SHEET_SETTINGS = 'הגדרות';
var SHEET_SUMMARY = 'סיכום';
var SETTINGS_COLS = [['key', 'מפתח', 'text'], ['value', 'ערך', 'text'], ['desc', 'הסבר', 'text']];
var EXP_CATEGORIES = ['טיסה', 'מזומן', 'תחבורה', 'אטרקציות', 'ביטוח', 'אחר'];
var EXP_COLS = [
  ['id', 'מזהה', 'text'],
  ['category', 'קטגוריה', 'text'],
  ['title', 'תיאור', 'text'],
  ['date', 'תאריך', 'date'],
  ['date2', 'תאריך חזור', 'date'],
  ['route', 'מסלול', 'text'],
  ['price', 'מחיר', 'num'],
  ['currency', 'מטבע', 'text'],
  ['priceIls', 'בשקלים', 'num'],
  ['account', 'יוזר', 'text'],
  ['confirmation', 'מספר הזמנה', 'text'],
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
  var expenses = ss.getSheetByName(SHEET_EXPENSES) || ss.insertSheet(SHEET_EXPENSES);
  writeHeader_(expenses, EXP_COLS);
  settingsSheet_(); syncSummary_();

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('API_TOKEN')) props.setProperty('API_TOKEN', Utilities.getUuid().replace(/-/g, ''));
  Logger.log('API_TOKEN: ' + props.getProperty('API_TOKEN'));
  Logger.log('הטאבים מוכנים. עכשיו: פריסה → פריסה חדשה → אפליקציית אינטרנט.');
}

/** מוודא שכל הכותרות קיימות; עמודות חסרות מתווספות בסוף (סדר העמודות לא משנה — הגישה לפי שם כותרת). */
function writeHeader_(sh, cols) {
  var headers = cols.map(function (c) { return c[1]; });
  var lastCol = sh.getLastColumn();
  var existing = (sh.getLastRow() >= 1 && lastCol > 0) ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  while (existing.length && existing[existing.length - 1] === '') existing.pop();
  var missing = headers.filter(function (h) { return existing.indexOf(h) < 0; });
  if (existing.length === 0) { sh.getRange(1, 1, 1, headers.length).setValues([headers]); existing = headers.slice(); }
  else if (missing.length) { sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]); existing = existing.concat(missing); }
  sh.getRange(1, 1, 1, existing.length).setFontWeight('bold').setBackground('#e8f0fe');
  sh.setFrozenRows(1);
  sh.setRightToLeft(true);
  // עמודות טקסט לתאריכים ומזהים כדי שהגיליון לא "יתקן" אותם
  cols.forEach(function (c) {
    var i = existing.indexOf(c[1]);
    if (i >= 0 && (c[2] === 'date' || c[2] === 'text')) sh.getRange(2, i + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  });
  return existing;
}
var HEADER_CACHE_ = {};
function headerOf_(sh, cols) {
  var key = sh.getSheetId();
  if (!HEADER_CACHE_[key]) HEADER_CACHE_[key] = writeHeader_(sh, cols);
  return HEADER_CACHE_[key];
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
  r.group = existing ? existing.group : '';
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
  // תוספת שתשולם במלון (למשל מיטה נוספת) — באותו מטבע של ההזמנה
  if (r.extraPrice === '' || r.extraPrice === null) r.extraIls = '';
  else if (b.extraPrice !== undefined || b.currency !== undefined || r.extraIls === '') {
    if (r.currency === 'ILS') r.extraIls = r.extraPrice;
    else { var rate2 = fxRate_(r.currency); r.extraIls = rate2 ? Math.round(r.extraPrice * rate2) : ''; }
  }
  return r;
}

function normalizeExpense_(b, existing) {
  var r = {};
  EXP_COLS.forEach(function (c) { r[c[0]] = existing ? existing[c[0]] : ''; });
  EXP_COLS.forEach(function (c) {
    var k = c[0];
    if (b[k] === undefined) return;
    var v = b[k];
    if (c[2] === 'date') v = normDate_(v);
    else if (c[2] === 'num') v = normNum_(v);
    else v = v === null ? '' : String(v).trim();
    r[k] = v;
  });
  var cat = String(r.category || '').trim().toLowerCase();
  if (cat.indexOf('flight') >= 0 || cat.indexOf('טיס') >= 0 || cat === 'air') r.category = 'טיסה';
  else if (cat.indexOf('cash') >= 0 || cat.indexOf('atm') >= 0 || cat.indexOf('מזומן') >= 0 || cat.indexOf('כספומט') >= 0) r.category = 'מזומן';
  else if (cat.indexOf('taxi') >= 0 || cat.indexOf('train') >= 0 || cat.indexOf('ferry') >= 0 || cat.indexOf('תחבור') >= 0 || cat.indexOf('מונית') >= 0 || cat.indexOf('רכבת') >= 0 || cat.indexOf('מעבורת') >= 0) r.category = 'תחבורה';
  else if (cat.indexOf('insur') >= 0 || cat.indexOf('ביטוח') >= 0) r.category = 'ביטוח';
  else if (cat.indexOf('attr') >= 0 || cat.indexOf('אטרק') >= 0 || cat.indexOf('טיול') >= 0 || cat.indexOf('tour') >= 0) r.category = 'אטרקציות';
  else if (EXP_CATEGORIES.indexOf(r.category) < 0) r.category = r.category ? 'אחר' : 'אחר';
  r.currency = String(r.currency || 'ILS').toUpperCase().replace('₪', 'ILS').replace('NIS', 'ILS').replace('$', 'USD').replace('€', 'EUR').replace('£', 'GBP').replace('฿', 'THB').trim();
  var manualIls = b.priceIls !== undefined && b.priceIls !== null && String(b.priceIls) !== '';
  var priceChanged = !existing || String(existing.price) !== String(r.price) || String(existing.currency) !== String(r.currency);
  if (!manualIls && r.price !== '' && (r.priceIls === '' || priceChanged)) {
    if (r.currency === 'ILS') r.priceIls = r.price;
    else { var rate3 = fxRate_(r.currency); r.priceIls = rate3 ? Math.round(r.price * rate3) : ''; }
  }
  if (!r.title) r.title = r.category === 'טיסה' ? ('טיסה ' + (r.route || '')).trim() : r.category === 'מזומן' ? 'משיכת מזומן' : r.category;
  return r;
}
function settingsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SETTINGS); writeHeader_(sh, SETTINGS_COLS);
    sh.getRange(2, 1, 2, 3).setValues([['TARGET_NIGHTS', props_().getProperty('TARGET_NIGHTS') || 21, 'יעד לילות לטיול'], ['TARGET_BUDGET', props_().getProperty('TARGET_BUDGET') || 0, 'תקציב כולל בשקלים (0 = לא הוגדר)']]);
  }
  return sh;
}
function getSetting_(key, def) {
  var rows = readAll_(settingsSheet_(), SETTINGS_COLS);
  for (var i = 0; i < rows.length; i++) if (String(rows[i].key) === key) return rows[i].value === '' ? def : rows[i].value;
  return def;
}
function setSetting_(key, value, desc) {
  var sh = settingsSheet_();
  var rows = readAll_(sh, SETTINGS_COLS);
  for (var i = 0; i < rows.length; i++) if (String(rows[i].key) === key) { sh.getRange(rows[i]._row, 2).setValue(value); return; }
  sh.appendRow([key, value, desc || '']);
}
/** טאב "סיכום" — תמונת מצב של התקציב, נכתב מחדש אחרי כל שינוי (גיבוי קריא בגיליון) */
function syncSummary_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_SUMMARY) || ss.insertSheet(SHEET_SUMMARY);
  var list = listBookings_(); var sum = summary_(list); var exps = listExpenses_(); var B = budget_(sum, exps);
  var rows = [['סעיף', 'ערך'], ['לילות סגורים', sum.nights + ' מתוך ' + sum.targetNights], ['מלונות (הזמנות)', sum.count], ['מלונות ₪', sum.totalIls], ['תוספות במלון ₪', sum.extrasIls]];
  EXP_CATEGORIES.forEach(function (c) { rows.push([c + ' ₪', Math.round(B.byCategory[c] || 0)]); });
  var paid = 0, left = 0;
  list.forEach(function (r) { var v = Number(r.priceIls) || 0; if (r.paid) paid += v; else left += v; left += Number(r.extraIls) || 0; });
  exps.forEach(function (e) { paid += Number(e.priceIls) || 0; });
  rows.push(['סה"כ הטיול ₪', B.totalIls]);
  rows.push(['שולם ₪', Math.round(paid)]);
  rows.push(['נשאר לשלם ₪ (מלונות שלא שולמו + תוספות)', Math.round(left)]);
  rows.push(['תקציב יעד ₪', B.targetBudget || '']);
  rows.push(['נשאר ₪', B.targetBudget ? B.targetBudget - B.totalIls : '']);
  rows.push(['ממוצע ללילה ₪ (כולל תוספות)', sum.nights ? Math.round(sum.hotelsIls / sum.nights) : '']);
  rows.push(['עודכן', now_()]);
  sh.clearContents();
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e8f0fe');
  sh.setRightToLeft(true); sh.setFrozenRows(1); sh.autoResizeColumns(1, 2);
}
function expensesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_EXPENSES);
  if (!sh) { sh = ss.insertSheet(SHEET_EXPENSES); writeHeader_(sh, EXP_COLS); }
  return sh;
}
function listExpenses_() { return readAll_(expensesSheet_(), EXP_COLS); }
function findExpense_(id) { var all = listExpenses_(); for (var i = 0; i < all.length; i++) if (String(all[i].id) === String(id)) return all[i]; return null; }
function logExpenseHistory_(action, by, note, e) {
  var rec = { hotel: '[' + e.category + '] ' + e.title, destination: e.route || '', checkIn: e.date, checkOut: e.date2, price: e.price, currency: e.currency, priceIls: e.priceIls, account: e.account, confirmation: e.confirmation, notes: e.notes, id: e.id, platform: 'הוצאה', source: e.source, createdAt: e.createdAt, updatedAt: e.updatedAt };
  logHistory_(action, by, note, rec);
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
  var ash = archiveSheet_(); ash.appendRow(toRow_(row, ARCHIVE_COLS, ash));
}
function listArchive_() { return readAll_(archiveSheet_(), ARCHIVE_COLS); }
function findArchived_(id) {
  var all = listArchive_();
  for (var i = all.length - 1; i >= 0; i--) if (String(all[i].id) === String(id)) return all[i];
  return null;
}

function readAll_(sh, cols) {
  if (!sh || sh.getLastRow() < 2) return [];
  var header = headerOf_(sh, cols);
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, header.length).getValues();
  var idx = {}; cols.forEach(function (c) { idx[c[0]] = header.indexOf(c[1]); });
  var out = [];
  values.forEach(function (row, i) {
    var r = { _row: i + 2 }, empty = true;
    cols.forEach(function (c) {
      var v = idx[c[0]] >= 0 ? row[idx[c[0]]] : '';
      if (c[2] === 'date') v = normDate_(v);
      else if (c[2] === 'bool') v = normBool_(v);
      else if (v instanceof Date) v = Utilities.formatDate(v, 'Asia/Jerusalem', ['ts', 'createdAt', 'updatedAt', 'archivedAt'].indexOf(c[0]) >= 0 ? "yyyy-MM-dd'T'HH:mm:ss" : 'yyyy-MM-dd');
      r[c[0]] = v;
      if (v !== '' && v !== null && v !== false) empty = false;
    });
    if (!empty) out.push(r);
  });
  return out;
}
/** שורה בסדר העמודות של הגיליון (לפי כותרות) */
function toRow_(r, cols, sh) {
  var header = headerOf_(sh, cols);
  var byHeader = {}; cols.forEach(function (c) { byHeader[c[1]] = c; });
  return header.map(function (h) {
    var c = byHeader[h]; if (!c) return '';
    var v = r[c[0]]; if (v === undefined || v === null) return '';
    if (c[2] === 'bool') return v ? 'כן' : 'לא';
    return v;
  });
}

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
  h.appendRow(toRow_(row, HISTORY_COLS, h));
}

function appendBooking_(rec) { var sh = bookingsSheet_(); sh.appendRow(toRow_(rec, COLS, sh)); }
function writeBooking_(rowNum, rec) { var sh = bookingsSheet_(); var row = toRow_(rec, COLS, sh); sh.getRange(rowNum, 1, 1, row.length).setValues([row]); }
function deleteRow_(rowNum) { bookingsSheet_().deleteRow(rowNum); }

function remindersStatus_() {
  var token = props_().getProperty('TELEGRAM_BOT_TOKEN'), chat = props_().getProperty('TELEGRAM_CHAT_ID');
  var days = (props_().getProperty('REMINDER_DAYS') || '14,7,3,1,0').split(',').map(function (x) { return Number(x.trim()); }).filter(function (n) { return !isNaN(n); });
  var installed = false;
  try { installed = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'dailyReminders'; }); } catch (e) { }
  return { configured: !!(token && chat), triggerInstalled: installed, active: !!(token && chat) && installed, days: days, paymentDays: [7, 1, 0], hour: 9 };
}

// ---------- חדרים במקביל (אותו מלון, תאריכים חופפים) ----------
var GROUP_COLORS = ['#fff3b0', '#cfe4ff', '#c9f2d0', '#ffd6e7', '#e3d9ff', '#ffe0c2', '#c8f4f0', '#f0e0c0'];
function overlaps_(a, b) { return a.checkIn && a.checkOut && b.checkIn && b.checkOut && a.checkIn < b.checkOut && b.checkIn < a.checkOut; }
/** מחזיר map id → {n, size, color} רק לקבוצות של 2+ הזמנות */
function computeGroups_(list) {
  var parent = list.map(function (_, i) { return i; });
  function find(i) { while (parent[i] !== i) { parent[i] = parent[parent[i]]; i = parent[i]; } return i; }
  for (var i = 0; i < list.length; i++) for (var j = i + 1; j < list.length; j++) {
    if (norm_(list[i].hotel) === norm_(list[j].hotel) && overlaps_(list[i], list[j])) parent[find(i)] = find(j);
  }
  var members = {};
  list.forEach(function (r, i) { var root = find(i); (members[root] = members[root] || []).push(r); });
  var groups = Object.keys(members).map(function (k) { return members[k]; }).filter(function (g) { return g.length > 1; });
  groups.sort(function (a, b) { return String(a[0].checkIn).localeCompare(String(b[0].checkIn)); });
  var out = {};
  groups.forEach(function (g, gi) { g.forEach(function (r, ri) { out[r.id] = { n: gi + 1, size: g.length, idx: ri + 1, color: GROUP_COLORS[gi % GROUP_COLORS.length] }; }); });
  return out;
}
/** צובע בגיליון שורות של חדרים במקביל באותו צבע וכותב תווית בעמודה "חדרים במקביל" */
function syncGroups_() {
  var sh = bookingsSheet_();
  var all = readAll_(sh, COLS);
  if (!all.length) return;
  var groups = computeGroups_(all);
  var header = headerOf_(sh, COLS);
  var gi = header.indexOf('חדרים במקביל');
  var last = sh.getLastRow();
  if (last < 2) return;
  var bg = [], labels = [];
  for (var row = 2; row <= last; row++) { bg.push(header.map(function () { return null; })); labels.push(['']); }
  all.forEach(function (r) {
    var g = groups[r.id];
    if (g) { bg[r._row - 2] = header.map(function () { return g.color; }); labels[r._row - 2] = ['קבוצה ' + g.n + ' · חדר ' + g.idx + '/' + g.size]; }
  });
  sh.getRange(2, 1, last - 1, header.length).setBackgrounds(bg);
  if (gi >= 0) sh.getRange(2, gi + 1, last - 1, 1).setValues(labels);
}
/** לילות ייחודיים (איחוד טווחי תאריכים) — חדרים במקביל נספרים פעם אחת */
function distinctNights_(list) {
  var iv = list.filter(function (r) { return r.checkIn && r.checkOut && r.checkIn < r.checkOut; }).map(function (r) { return [r.checkIn, r.checkOut]; }).sort();
  var total = 0, cur = null;
  iv.forEach(function (x) {
    if (!cur || x[0] >= cur[1]) { if (cur) total += nightsBetween_(cur[0], cur[1]) || 0; cur = [x[0], x[1]]; }
    else if (x[1] > cur[1]) cur[1] = x[1];
  });
  if (cur) total += nightsBetween_(cur[0], cur[1]) || 0;
  return total;
}

function summary_(list) {
  var nights = 0, ils = 0, missingIls = 0;
  list.forEach(function (r) { nights += Number(r.nights) || 0; if (r.priceIls !== '' && r.priceIls !== null) ils += Number(r.priceIls) || 0; else if (r.price !== '') missingIls++; });
  var extras = 0; list.forEach(function (r) { extras += Number(r.extraIls) || 0; });
  return { count: list.length, nights: distinctNights_(list), roomNights: nights, totalIls: Math.round(ils), extrasIls: Math.round(extras), hotelsIls: Math.round(ils + extras), missingIls: missingIls, targetNights: Number(getSetting_('TARGET_NIGHTS', 21)) || 21 };
}
function budget_(bookingsSummary, expenses) {
  var byCat = {}; EXP_CATEGORIES.forEach(function (c) { byCat[c] = 0; });
  expenses.forEach(function (e) { var v = Number(e.priceIls) || 0; byCat[e.category] = (byCat[e.category] || 0) + v; });
  var expTotal = 0; Object.keys(byCat).forEach(function (k) { expTotal += byCat[k]; });
  var total = (bookingsSummary.hotelsIls || 0) + expTotal;
  return { hotelsIls: bookingsSummary.totalIls, extrasIls: bookingsSummary.extrasIls, byCategory: byCat, expensesIls: Math.round(expTotal), totalIls: Math.round(total), targetBudget: Number(getSetting_('TARGET_BUDGET', 0)) || 0, categories: EXP_CATEGORIES };
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
  var res = handleInner_(action, p, by);
  if (['add', 'upsert', 'update', 'delete', 'restore'].indexOf(action) >= 0) { try { syncGroups_(); } catch (e) { } }
  if (['add', 'upsert', 'update', 'delete', 'restore', 'addExpense', 'updateExpense', 'deleteExpense', 'setTarget', 'setBudget'].indexOf(action) >= 0) { try { syncSummary_(); } catch (e) { } }
  return res;
}
function handleInner_(action, p, by) {
  by = p.by || p.source || by;
  switch (action) {
    case 'list': {
      var list = listBookings_();
      var groups = computeGroups_(list);
      list.forEach(function (r) { r.groupInfo = groups[r.id] || null; });
      var sum = summary_(list);
      var exps = listExpenses_().map(function (r) { delete r._row; return r; });
      return json_({ ok: true, bookings: list, summary: sum, expenses: exps, budget: budget_(sum, exps), platforms: PLATFORMS, reminders: remindersStatus_() });
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
      if (p.parallel) return handleInner_('add', { token: p.token, booking: b, by: by, note: p.note || 'חדר נוסף במקביל' }, by);
      var all = readAll_(bookingsSheet_(), COLS);
      var cands = p.replaceId ? all.filter(function (r) { return String(r.id) === String(p.replaceId); }) : matchCandidates_(all, b);
      if (cands.length > 1) {
        return json_({ ok: false, needsChoice: true, error: 'נמצאו כמה הזמנות דומות — שלח replaceId', candidates: cands.map(function (r) { delete r._row; return r; }) });
      }
      if (cands.length === 1 && !p.replaceId && p.confirmReplace !== true && cands[0].confirmation && b.confirmation && String(cands[0].confirmation) !== String(b.confirmation)) {
        var c1 = cands[0]; delete c1._row;
        return json_({ ok: false, needsConfirm: true, error: 'קיימת הזמנה לאותו מלון עם מספר הזמנה אחר — חדר נוסף במקביל (parallel:true) או החלפה (confirmReplace:true)?', candidate: c1 });
      }
      if (cands.length === 1 && p.replaceId === undefined && p.confirmReplace === false) {
        var c0 = cands[0]; delete c0._row;
        return json_({ ok: false, needsConfirm: true, error: 'קיימת הזמנה דומה — שלח confirmReplace:true או replaceId', candidate: c0 });
      }
      var q = { token: p.token, booking: b, by: by, note: p.note };
      if (cands.length === 1) q.replaceId = cands[0].id;
      return handleInner_('add', q, by);
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
    case 'setTarget': { setSetting_('TARGET_NIGHTS', Number(p.nights) || 21, 'יעד לילות לטיול'); return json_({ ok: true }); }
    case 'setBudget': { setSetting_('TARGET_BUDGET', Number(p.ils) || 0, 'תקציב כולל בשקלים (0 = לא הוגדר)'); return json_({ ok: true }); }
    case 'expenses': {
      var ex = listExpenses_().map(function (r) { delete r._row; return r; });
      return json_({ ok: true, expenses: ex, budget: budget_(summary_(listBookings_()), ex) });
    }
    case 'addExpense': {
      var eb = p.expense || p; if (eb.price === undefined || eb.price === '') return json_({ ok: false, error: 'חסר מחיר' });
      var er = normalizeExpense_(eb); er.id = 'e' + newId_(); er.source = by; er.createdAt = now_(); er.updatedAt = er.createdAt;
      var esh = expensesSheet_(); esh.appendRow(toRow_(er, EXP_COLS, esh));
      logExpenseHistory_('נוספה הוצאה', by, p.note || '', er);
      var exs = listExpenses_();
      return json_({ ok: true, expense: er, budget: budget_(summary_(listBookings_()), exs) });
    }
    case 'updateExpense': {
      var eo = findExpense_(p.id); if (!eo) return json_({ ok: false, error: 'לא נמצא' });
      var erow = eo._row; delete eo._row;
      var er2 = normalizeExpense_(p.expense || p.fields || p, eo); er2.id = eo.id; er2.createdAt = eo.createdAt; er2.source = eo.source || by; er2.updatedAt = now_();
      var ch = EXP_COLS.filter(function (c) { return String(eo[c[0]]) !== String(er2[c[0]]) && c[0] !== 'updatedAt'; }).map(function (c) { return c[1] + ': ' + eo[c[0]] + ' → ' + er2[c[0]]; });
      var esh2 = expensesSheet_(); var row2 = toRow_(er2, EXP_COLS, esh2); esh2.getRange(erow, 1, 1, row2.length).setValues([row2]);
      logExpenseHistory_('עודכנה הוצאה', by, ch.join(' | '), eo);
      return json_({ ok: true, expense: er2, changes: ch });
    }
    case 'deleteExpense': {
      var ed = findExpense_(p.id); if (!ed) return json_({ ok: false, error: 'לא נמצא' });
      expensesSheet_().deleteRow(ed._row); delete ed._row;
      logExpenseHistory_('נמחקה הוצאה', by, p.note || '', ed);
      return json_({ ok: true, deleted: ed, budget: budget_(summary_(listBookings_()), listExpenses_()) });
    }
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
  listBookings_().forEach(function (r) {
    if (!r.paymentDate || r.paid) return;
    var d = new Date(r.paymentDate + 'T00:00:00Z');
    var days = Math.round((d - today) / 86400000);
    if ([7, 1, 0].indexOf(days) >= 0) {
      lines.push((days === 0 ? '💳 היום' : days === 1 ? '💳 מחר' : '💳 בעוד ' + days + ' ימים') + ' יורד התשלום: *' + r.hotel + '* (' + r.platform + (r.account ? ', יוזר ' + r.account : '') + ')\n' +
        (r.priceIls ? r.priceIls + ' ₪' : r.price + ' ' + r.currency) + (r.confirmation ? ' · #' + r.confirmation : ''));
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
