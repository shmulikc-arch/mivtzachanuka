/**
 * מבצע חנוכה עם אבא · פנסאים × לב חב"ד
 * קולט הרשמות מהאתר לגוגל שיטס, ושומר על מכסה של 10 אבות למשמרת.
 *
 * התקנה:
 * 1. בגוגל שיטס החדש: הרחבות → Apps Script → להדביק את כל הקובץ הזה במקום מה שיש.
 * 2. לשמור, ואז: פריסה (Deploy) → פריסה חדשה → סוג: אפליקציית אינטרנט (Web app)
 *    הפעלה בתור: אני | למי יש גישה: כולם (Anyone) → פריסה → לאשר הרשאות.
 * 3. להעתיק את הכתובת שמסתיימת ב־/exec ולשלוח לשמוליק (נכנסת ל־SCRIPT_URL באתר).
 *
 * ביטול הרשמה: לכתוב "בוטל" בעמודת "סטטוס" בגיליון "הרשמות". המקום מתפנה מיד.
 */

const CAPACITY = 10;

// חייב להתאים לרשימה ב־index.html
const HOSPITALS = { kaplan: 'בית החולים קפלן', wolfson: 'בית החולים וולפסון' };
const DATES = ['2026-12-06', '2026-12-07', '2026-12-08', '2026-12-09', '2026-12-10'];

const REG_SHEET = 'הרשמות';
const KIDS_SHEET = 'ילדים';
const REG_HEAD = ['נרשם בתאריך', 'מספר הרשמה', 'בית רפואה', 'תאריך עברי', 'תאריך', 'שעה',
  'שם פרטי (אב)', 'שם משפחה (אב)', 'שם האם (אב)', 'טלפון', 'עיר', 'מספר ילדים', 'שמות הילדים', 'סטטוס', 'קוד משמרת'];
const KIDS_HEAD = ['מספר הרשמה', 'בית רפואה', 'תאריך עברי', 'שם פרטי', 'שם משפחה', 'שם האם',
  'כיתה', 'סניף', 'שם האב', 'טלפון האב', 'סטטוס'];
const COL_STATUS = REG_HEAD.indexOf('סטטוס') + 1;
const COL_SHIFT = REG_HEAD.indexOf('קוד משמרת') + 1;

function validShift(id) {
  return Object.keys(HOSPITALS).some(h => DATES.some(d => id === h + '-' + d));
}

function sheet(name, head) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#FCE7BE');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

function counts() {
  const sh = sheet(REG_SHEET, REG_HEAD);
  const out = {};
  const n = sh.getLastRow() - 1;
  if (n < 1) return out;
  const rows = sh.getRange(2, 1, n, REG_HEAD.length).getValues();
  rows.forEach(r => {
    if (String(r[COL_STATUS - 1]).trim() === 'בוטל') return;
    const id = r[COL_SHIFT - 1];
    if (id) out[id] = (out[id] || 0) + 1;
  });
  return out;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return json({ ok: true, capacity: CAPACITY, counts: counts() });
}

function clean(v, max) { return String(v == null ? '' : v).trim().slice(0, max || 80); }

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (x) { return json({ ok: false, error: 'BAD_REQUEST' }); }

  const shift = clean(body.shift, 40);
  if (!validShift(shift)) return json({ ok: false, error: 'SHIFT_CLOSED' });

  const f = body.father || {};
  const kids = Array.isArray(body.children) ? body.children.slice(0, 15) : [];
  const father = { first: clean(f.first), last: clean(f.last), mother: clean(f.mother), phone: clean(f.phone, 20), city: clean(f.city) };
  if (!father.first || !father.last || !father.mother || !father.phone || !father.city || !kids.length)
    return json({ ok: false, error: 'MISSING_FIELDS' });
  const children = kids.map(k => ({ first: clean(k.first), last: clean(k.last), mother: clean(k.mother), grade: clean(k.grade, 20), branch: clean(k.branch) }));
  if (children.some(k => !k.first || !k.last || !k.grade || !k.branch)) return json({ ok: false, error: 'MISSING_FIELDS' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json({ ok: false, error: 'BUSY' });
  try {
    if ((counts()[shift] || 0) >= CAPACITY) return json({ ok: false, error: 'SHIFT_FULL' });

    const id = Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
    const hospital = HOSPITALS[shift.split('-')[0]];
    const now = new Date();
    sheet(REG_SHEET, REG_HEAD).appendRow([now, id, hospital, clean(body.heb_date, 20), clean(body.date, 12), clean(body.time, 20),
      father.first, father.last, father.mother, "'" + father.phone, father.city,
      children.length, children.map(k => k.first).join(', '), 'רשום', shift]);
    const ks = sheet(KIDS_SHEET, KIDS_HEAD);
    children.forEach(k => ks.appendRow([id, hospital, clean(body.heb_date, 20), k.first, k.last, k.mother, k.grade, k.branch,
      father.first + ' ' + father.last, "'" + father.phone, 'רשום']));
    SpreadsheetApp.flush();
    return json({ ok: true, id: id });
  } finally {
    lock.releaseLock();
  }
}

/** להרצה ידנית פעם אחת (לא חובה): יוצר את שני הגיליונות עם הכותרות. */
function setup() {
  sheet(REG_SHEET, REG_HEAD);
  sheet(KIDS_SHEET, KIDS_HEAD);
}
