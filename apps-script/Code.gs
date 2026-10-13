/**
 * יום עיון "קירות נופלים בטיפול" – קליטת הרשמות לגיליון Google Sheets + מייל אישור לנרשם/ת.
 *
 * התקנה (בקצרה – פירוט מלא ב-README.md):
 *   1. בגיליון Google חדש: תוספים (Extensions) > Apps Script, ולהדביק את כל הקובץ הזה.
 *   2. להריץ פעם אחת את הפונקציה setup (ולאשר הרשאות).
 *   3. Deploy > New deployment > Web app | Execute as: Me | Who has access: Anyone.
 *   4. להעתיק את כתובת ה-Web app אל CONFIG.scriptUrl בקובץ index.html.
 */

const CONFIG = {
  SHEET_NAME: 'הרשמות',
  // מייל לקבלת עדכון על כל הרשמה חדשה. להשאיר ריק ('') כדי לא לשלוח.
  ORGANIZER_EMAIL: '',
  // כתובת שאליה יגיעו תשובות למייל האישור. ריק = הכתובת של בעל/ת הסקריפט.
  REPLY_TO: '',
  SENDER_NAME: 'יום עיון קרימינולוגים קליניים – שער מנשה',
  TIMEZONE: 'Asia/Jerusalem',
};

const EVENT = {
  title: 'קירות נופלים בטיפול',
  quote: '"מה קורה שם בחוץ?"',
  dedication: 'לזכרה של תמי קרוק',
  date: 'יום שני, כ"ז בכסלו התשפ"ז, 7 בדצמבר 2026',
  time: 'התכנסות וכיבוד בוקר ב-8:30, סיום משוער ב-13:30',
  place: 'אולם ע"ש בלה שניידר, המרכז לבריאות הנפש שער מנשה',
};

const HEADERS = [
  'תאריך הרשמה',
  'שם פרטי',
  'שם משפחה',
  'טלפון',
  'דוא"ל',
  'מקצוע',
  'מקום עבודה',
  'סטודנט/ית לקרימינולוגיה קלינית',
  'תשלום',
  'מייל אישור',
  'הגיע/ה',
];

const FIELDS = ['firstName', 'lastName', 'phone', 'email', 'profession', 'workplace', 'student'];

/** להריץ פעם אחת מתוך עורך Apps Script: יוצר את הגיליון, הכותרות והעיצוב. */
function setup() {
  const sheet = getSheet_();
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setFontWeight('bold')
    .setBackground('#F2E7DA')
    .setFontColor('#2B2420');
  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);
  sheet.setColumnWidths(1, HEADERS.length, 150);
  sheet.setColumnWidth(5, 220);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(8, 120);
  SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetTimeZone(CONFIG.TIMEZONE);
}

/** בדיקה מהירה בדפדפן שה-Web app פעיל. */
function doGet() {
  return ContentService.createTextOutput('OK – טופס ההרשמה פעיל');
}

function doPost(e) {
  const params = (e && e.parameter) || {};

  // מלכודת בוטים: שדה נסתר שאנשים לא ממלאים
  if (params.website) return json_({ ok: true });

  const data = {};
  FIELDS.forEach(k => { data[k] = clean_(params[k]); });

  const missing = FIELDS.filter(k => !data[k]);
  if (missing.length) return json_({ ok: false, error: 'missing', fields: missing });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(data.email)) return json_({ ok: false, error: 'email' });
  if (data.student !== 'כן' && data.student !== 'לא') return json_({ ok: false, error: 'student' });

  const payment = paymentFor_(data);
  const lock = LockService.getScriptLock();
  let sheet, row;
  try {
    lock.waitLock(20000);
    sheet = getSheet_();
    sheet.appendRow([
      Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      safe_(data.firstName),
      safe_(data.lastName),
      "'" + data.phone, // שומר את ה-0 בתחילת המספר
      safe_(data.email),
      safe_(data.profession),
      safe_(data.workplace),
      data.student,
      payment.label,
      'בשליחה…',
      false,
    ]);
    row = sheet.getLastRow();
    sheet.getRange(row, HEADERS.indexOf('הגיע/ה') + 1).insertCheckboxes();
  } catch (err) {
    return json_({ ok: false, error: 'sheet: ' + err.message });
  } finally {
    lock.releaseLock();
  }

  // ההרשמה כבר נשמרה – תקלה במייל לא מפילה אותה
  let mailStatus = 'נשלח';
  try {
    sendConfirmation_(data, payment);
  } catch (err) {
    mailStatus = 'שגיאה: ' + err.message;
  }
  sheet.getRange(row, HEADERS.indexOf('מייל אישור') + 1).setValue(mailStatus);

  if (CONFIG.ORGANIZER_EMAIL) {
    try { notifyOrganizer_(data, payment, row - 1); } catch (err) { /* לא קריטי */ }
  }

  return json_({ ok: true });
}

/* ---------------- עזרים ---------------- */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);
}

function clean_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, 200);
}

/** מונע הזרקת נוסחאות לגיליון (ערך שמתחיל ב- = + - @). */
function safe_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function paymentFor_(data) {
  if (data.student === 'כן') {
    return { label: 'ללא עלות – סטודנט/ית', text: 'השתתפותך ללא עלות (סטודנטים לקרימינולוגיה קלינית).' };
  }
  if (/שער[\s\-–]*מנשה/.test(data.workplace)) {
    return { label: 'ללא עלות – עובד/ת המרכז', text: 'השתתפותך ללא עלות (עובדי המרכז לבריאות הנפש שער מנשה).' };
  }
  return {
    label: '50 ₪ במזומן',
    text: 'דמי ההשתתפות: 50 ₪, בתשלום במזומן בלבד בעת ההגעה והרישום. לעובדי המרכז לבריאות הנפש שער מנשה – ללא עלות.',
  };
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function sendConfirmation_(data, payment) {
  const row = (label, value) =>
    '<tr><td style="padding:4px 0 4px 16px;color:#6E6259;white-space:nowrap;vertical-align:top">' + label +
    '</td><td style="padding:4px 0;color:#2B2420">' + esc_(value) + '</td></tr>';

  const html =
    '<div dir="rtl" style="font-family:Arial,Helvetica,sans-serif;font-size:16px;line-height:1.7;color:#2B2420;max-width:560px;margin:0 auto">' +
      '<p style="margin:0 0 16px">שלום ' + esc_(data.firstName) + ',</p>' +
      '<p style="margin:0 0 16px">תודה שנרשמת ליום העיון השנתי של סקטור קרימינולוגים קליניים. ההרשמה התקבלה, ונשמח לפגוש אותך.</p>' +
      '<div style="background:#F6F1EA;border-radius:12px;padding:18px 20px;margin:0 0 18px">' +
        '<div style="color:#564A42;font-size:15px">' + esc_(EVENT.quote) + '</div>' +
        '<div style="font-size:22px;font-weight:bold;margin:2px 0 4px">' + esc_(EVENT.title) + '</div>' +
        '<div style="color:#564A42;font-size:15px;margin-bottom:12px">' + esc_(EVENT.dedication) + '</div>' +
        '<div><b>מתי:</b> ' + esc_(EVENT.date) + '<br>' + esc_(EVENT.time) + '</div>' +
        '<div style="margin-top:6px"><b>איפה:</b> ' + esc_(EVENT.place) + '</div>' +
        '<div style="margin-top:6px"><b>השתתפות:</b> ' + esc_(payment.text) + '</div>' +
      '</div>' +
      '<p style="margin:0 0 6px;font-weight:bold">פרטי ההרשמה</p>' +
      '<table style="border-collapse:collapse;font-size:15px;margin:0 0 18px">' +
        row('שם', data.firstName + ' ' + data.lastName) +
        row('טלפון', data.phone) +
        row('מקצוע', data.profession) +
        row('מקום עבודה', data.workplace) +
        row('סטודנט/ית לקרימינולוגיה קלינית', data.student) +
      '</table>' +
      '<p style="margin:0">להתראות,<br>צוות סקטור קרימינולוגים קליניים<br>המרכז לבריאות הנפש שער מנשה</p>' +
    '</div>';

  const plain =
    'שלום ' + data.firstName + ',\n\n' +
    'תודה שנרשמת ליום העיון "' + EVENT.title + '". ההרשמה התקבלה.\n\n' +
    'מתי: ' + EVENT.date + ' – ' + EVENT.time + '\n' +
    'איפה: ' + EVENT.place + '\n' +
    'השתתפות: ' + payment.text + '\n\n' +
    'להתראות,\nצוות סקטור קרימינולוגים קליניים';

  const options = { name: CONFIG.SENDER_NAME, htmlBody: html };
  if (CONFIG.REPLY_TO) options.replyTo = CONFIG.REPLY_TO;

  MailApp.sendEmail(data.email, 'אישור הרשמה – יום העיון "' + EVENT.title + '" | 7.12.2026', plain, options);
}

function notifyOrganizer_(data, payment, total) {
  MailApp.sendEmail(
    CONFIG.ORGANIZER_EMAIL,
    'הרשמה חדשה ליום העיון: ' + data.firstName + ' ' + data.lastName + ' (סה"כ ' + total + ')',
    [
      'שם: ' + data.firstName + ' ' + data.lastName,
      'טלפון: ' + data.phone,
      'דוא"ל: ' + data.email,
      'מקצוע: ' + data.profession,
      'מקום עבודה: ' + data.workplace,
      'סטודנט/ית לקרימינולוגיה קלינית: ' + data.student,
      'תשלום: ' + payment.label,
      '',
      'סה"כ נרשמים עד כה: ' + total,
    ].join('\n'),
    { name: CONFIG.SENDER_NAME }
  );
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
