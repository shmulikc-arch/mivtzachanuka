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
 * ביטול הרשמה: לכתוב "בוטל" בעמודת "סטטוס" בגיליון "הרשמות". המקום מתפנה מיד,
 * והסטטוס מועתק אוטומטית לכל הילדים של אותה הרשמה בגיליון "ילדים".
 *
 * לוח שיבוץ: גיליון "לוח שיבוץ" מתעדכן לבד אחרי כל הרשמה ואחרי כל שינוי בגיליון "הרשמות".
 * אפשר גם לרענן ידנית מהתפריט "מבצע חנוכה" שבראש השיטס.
 */

const CAPACITY = 10;

// חייב להתאים לרשימה ב־index.html
const HOSPITALS = { kaplan: 'בית החולים קפלן', wolfson: 'בית החולים וולפסון' };
const DATES = ['2026-12-06', '2026-12-07', '2026-12-08', '2026-12-09', '2026-12-10'];

// לתצוגת הלוח בלבד
const BOARD = [
  { id: 'kaplan',  title: 'בית החולים קפלן · רחובות · 17:00–19:00' },
  { id: 'wolfson', title: 'בית החולים וולפסון · חולון · 16:30–18:30' },
];
const DAY_LABELS = ['ראשון · כ"ו כסלו', 'שני · כ"ז כסלו', 'שלישי · כ"ח כסלו', 'רביעי · כ"ט כסלו', "חמישי · ל' כסלו"];
const BOARD_SHEET = 'לוח שיבוץ';

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
    const regRow = [now, id, hospital, clean(body.heb_date, 20), clean(body.date, 12), clean(body.time, 20),
      father.first, father.last, father.mother, "'" + father.phone, father.city,
      children.length, children.map(k => k.first).join(', '), 'רשום', shift];
    const reg = sheet(REG_SHEET, REG_HEAD);
    reg.getRange(reg.getLastRow() + 1, 1, 1, REG_HEAD.length).setValues([regRow]);
    const ks = sheet(KIDS_SHEET, KIDS_HEAD);
    const kidRows = children.map(k => [id, hospital, clean(body.heb_date, 20), k.first, k.last, k.mother, k.grade, k.branch,
      father.first + ' ' + father.last, "'" + father.phone, 'רשום']);
    ks.getRange(ks.getLastRow() + 1, 1, kidRows.length, KIDS_HEAD.length).setValues(kidRows);   // כתיבה אחת לכל הילדים
    try { fillBoard(); } catch (err) {}   // עדכון מהיר של הלוח; לא יעצור הרשמה
    return json({ ok: true, id: id });
  } finally {
    lock.releaseLock();
  }
}

/** להרצה ידנית פעם אחת: יוצר את הגיליונות ובונה את הלוח. */
function setup() {
  sheet(REG_SHEET, REG_HEAD);
  sheet(KIDS_SHEET, KIDS_HEAD);
  buildBoard();
}

/* ============================================================
   לוח שיבוץ: עמודה לכל יום, עד 10 משבצות, לוח לכל בית רפואה
   layoutBoard: עיצוב מלא (איטי, רץ רק ב־setup / רענון ידני)
   fillBoard:   רק ערכים וצבעים (מהיר, רץ אחרי כל הרשמה ועריכה)
   ============================================================ */
const BLOCK = CAPACITY + 4;   // כותרת, שורת ימים, 10 משבצות, סה"כ, רווח
const W = DATES.length + 1;

function boardSheet() {
  const ss = SpreadsheetApp.getActive();
  let b = ss.getSheetByName(BOARD_SHEET);
  if (!b) { b = ss.insertSheet(BOARD_SHEET, 0); b.setRightToLeft(true); layoutBoard(b); }
  return b;
}

function layoutBoard(b) {
  b = b || boardSheet();
  b.clear(); b.clearFormats();
  b.getRange(1, 1, b.getMaxRows(), b.getMaxColumns()).getMergedRanges().forEach(m => m.breakApart());
  BOARD.forEach((h, k) => {
    const r = 1 + k * BLOCK;
    b.getRange(r, 1, 1, W).merge().setValue(h.title)
      .setFontWeight('bold').setFontSize(14).setBackground('#3A2716').setFontColor('#FFD23A').setHorizontalAlignment('center');
    b.getRange(r + 1, 1, 1, W).setValues([['#'].concat(DAY_LABELS)])
      .setFontWeight('bold').setBackground('#FCE7BE').setHorizontalAlignment('center');
    const g = b.getRange(r + 2, 1, CAPACITY, W);
    g.setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, '#E3D3B5', SpreadsheetApp.BorderStyle.SOLID);
    b.getRange(r + 2, 1, CAPACITY, 1).setHorizontalAlignment('center').setFontColor('#7A6650')
      .setValues(Array.from({ length: CAPACITY }, (_, i) => [i + 1]));
    b.getRange(r + 2 + CAPACITY, 1, 1, W).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#F7EFE0');
  });
  b.setColumnWidth(1, 44);
  for (let c = 2; c <= W; c++) b.setColumnWidth(c, 230);
  b.getRange(1 + BOARD.length * BLOCK, 1).setValue('מתעדכן אוטומטית מגיליון "הרשמות".').setFontColor('#7A6650').setFontSize(9);
}

function fillBoard() {
  const b = boardSheet();
  const reg = sheet(REG_SHEET, REG_HEAD);
  const n = reg.getLastRow() - 1;
  const rows = n > 0 ? reg.getRange(2, 1, n, REG_HEAD.length).getValues() : [];
  const by = {};
  rows.forEach(r => {
    if (String(r[COL_STATUS - 1]).trim() === 'בוטל') return;
    const id = r[COL_SHIFT - 1]; if (!id) return;
    const kids = r[11];
    (by[id] = by[id] || []).push(r[6] + ' ' + r[7] + ' · ' + String(r[9]).replace(/^'/, '') + ' · ' + (kids == 1 ? 'ילד אחד' : kids + ' ילדים'));
  });
  BOARD.forEach((h, k) => {
    const r = 1 + k * BLOCK;
    const grid = [];
    for (let i = 0; i < CAPACITY; i++) grid.push(DATES.map(d => (by[h.id + '-' + d] || [])[i] || ''));
    const g = b.getRange(r + 2, 2, CAPACITY, DATES.length);
    g.setValues(grid);
    g.setBackgrounds(grid.map(row => row.map(v => v ? '#E1F3EA' : '#FFFFFF')));
    const totals = ['סה"כ'].concat(DATES.map(d => {
      const c = (by[h.id + '-' + d] || []).length;
      return c >= CAPACITY ? 'מלא ' + c + '/' + CAPACITY : c + '/' + CAPACITY;
    }));
    const t = b.getRange(r + 2 + CAPACITY, 1, 1, W);
    t.setValues([totals]);
    t.setFontColors([totals.map(v => String(v).indexOf('מלא') === 0 ? '#B4232F' : '#2B1E12')]);
  });
}

/** רענון מלא: עיצוב מחדש + מילוי */
function buildBoard() { layoutBoard(); fillBoard(); }

/** כשמשנים משהו בגיליון "הרשמות": הסטטוס עובר לילדים, והלוח מתעדכן */
function onEdit(e) {
  if (!e || e.range.getSheet().getName() !== REG_SHEET) return;
  try { syncKidsStatus(e.range); } catch (err) {}
  try { fillBoard(); } catch (err) {}
}

/** מעתיק את הסטטוס ("בוטל" / "רשום") מכל הרשמה שנערכה לכל הילדים שלה בגיליון "ילדים" */
function syncKidsStatus(range) {
  const first = range.getColumn(), last = first + range.getNumColumns() - 1;
  if (COL_STATUS < first || COL_STATUS > last) return;           // לא נגעו בעמודת הסטטוס
  const top = Math.max(2, range.getRow()), bottom = range.getRow() + range.getNumRows() - 1;
  if (bottom < top) return;
  const reg = range.getSheet();
  const vals = reg.getRange(top, 1, bottom - top + 1, REG_HEAD.length).getValues();
  const status = {};
  vals.forEach(r => { if (r[1]) status[String(r[1])] = String(r[COL_STATUS - 1]).trim() || 'רשום'; });
  if (!Object.keys(status).length) return;

  const ks = sheet(KIDS_SHEET, KIDS_HEAD);
  const n = ks.getLastRow() - 1;
  if (n < 1) return;
  const col = KIDS_HEAD.indexOf('סטטוס') + 1;
  const ids = ks.getRange(2, 1, n, 1).getValues();
  const cur = ks.getRange(2, col, n, 1).getValues();
  let changed = false;
  ids.forEach((r, i) => {
    const s = status[String(r[0])];
    if (s !== undefined && cur[i][0] !== s) { cur[i][0] = s; changed = true; }
  });
  if (changed) ks.getRange(2, col, n, 1).setValues(cur);
}

/** להרצה ידנית אם צריך: מסנכרן את כל הסטטוסים מ"הרשמות" ל"ילדים" */
function syncAllKids() {
  const reg = sheet(REG_SHEET, REG_HEAD);
  if (reg.getLastRow() < 2) return;
  syncKidsStatus(reg.getRange(2, 1, reg.getLastRow() - 1, REG_HEAD.length));
}

/** תפריט בראש השיטס */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('מבצע חנוכה').addItem('רענון לוח השיבוץ', 'buildBoard').addItem('סנכרון סטטוס לילדים', 'syncAllKids').addToUi();
}
