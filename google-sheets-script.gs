/**
 * חיבור השאלון ל-Google Sheets
 *
 * התקנה (5 דקות, פעם אחת):
 * 1. פותחים גיליון Google Sheets חדש.
 * 2. בתפריט: תוספים (Extensions) > Apps Script. מוחקים את מה שיש ומדביקים את כל הקובץ הזה.
 * 3. לוחצים Deploy > New deployment > סוג: Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. מאשרים את ההרשאות, ומעתיקים את הכתובת שנגמרת ב-/exec.
 * 5. מדביקים אותה ב-index.html בתוך CONFIG.SHEETS_ENDPOINT.
 *
 * הסקריפט גמיש: כל שאלה חדשה שתוסיף ב-index.html מקבלת עמודה לבד,
 * אין צורך לגעת כאן. כל לקוח הוא שורה בלשונית "תשובות", ובנוסף מקבל לשונית משלו.
 */

const SHEET_NAME = 'תשובות';
const DATE_HEADER = 'תאריך';
const NOTIFY_BY_EMAIL = true; // מייל אליך על כל תשובה חדשה
const TAB_PER_PERSON = true;  // בנוסף לטבלה המרכזית, לשונית נפרדת לכל עונה

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = JSON.parse(e.postData.contents);
    if (data.website) return json({ ok: true }); // מלכודת בוטים
    const answers = data.answers || {};
    const labels = Object.keys(answers).slice(0, 80);
    if (!labels.length) return json({ ok: false, error: 'empty' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    // כותרות: מה שכבר קיים בגיליון + עמודה חדשה לכל שאלה שעוד אין לה
    let headers = sheet.getLastColumn()
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String)
      : [];
    if (!headers.length || !headers[0]) {
      headers = [DATE_HEADER];
      sheet.setRightToLeft(true);
      sheet.setFrozenRows(1);
    }
    labels.forEach(function (label) {
      if (headers.indexOf(label) === -1) headers.push(label);
    });
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

    sheet.appendRow(headers.map(function (header) {
      return header === DATE_HEADER ? new Date() : clean(answers[header]);
    }));

    if (TAB_PER_PERSON) writePersonTab(ss, answers, labels);

    if (NOTIFY_BY_EMAIL) {
      const lines = labels
        .filter(function (label) { return answers[label]; })
        .map(function (label) { return label + ': ' + String(answers[label]).slice(0, 1500); });
      MailApp.sendEmail(
        Session.getEffectiveUser().getEmail(),
        'שאלון פתיחה חדש: ' + String(answers['שם'] || '').slice(0, 80),
        lines.join('\n\n')
      );
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}

// לשונית נפרדת לכל עונה: שאלה בעמודה אחת, תשובה בעמודה השנייה. נוח לקריאה ולהדפסה.
function writePersonTab(ss, answers, labels) {
  const base = (String(answers['שם'] || 'ללא שם').replace(/[\[\]\*\?\/\\:]/g, ' ').trim() || 'ללא שם').slice(0, 60);
  let name = base;
  for (let n = 2; ss.getSheetByName(name); n++) name = base + ' (' + n + ')';
  const tab = ss.insertSheet(name, ss.getNumSheets());
  const rows = [['שאלה', 'תשובה'], [DATE_HEADER, new Date()]].concat(
    labels.map(function (label) { return [label, clean(answers[label])]; })
  );
  tab.setRightToLeft(true);
  tab.getRange(1, 1, rows.length, 2).setValues(rows).setVerticalAlignment('top').setWrap(true);
  tab.getRange(1, 1, 1, 2).setFontWeight('bold');
  tab.getRange(1, 1, rows.length, 1).setFontWeight('bold');
  tab.setColumnWidth(1, 190);
  tab.setColumnWidth(2, 560);
  tab.setFrozenRows(1);
}

// בדיקה מהירה בדפדפן שה-Web App חי
function doGet() {
  return json({ ok: true, status: 'alive' });
}

// מונע מתשובה להפוך לנוסחה בגיליון, ושומר על אפס מוביל במספרים
function clean(value) {
  let s = String(value == null ? '' : value).slice(0, 2000);
  if (/^[=+\-@]/.test(s) || /^0\d+$/.test(s.replace(/[\s-]/g, ''))) s = "'" + s;
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
