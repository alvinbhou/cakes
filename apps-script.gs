// Alvin & Janice 喜餅寄送意願調查 — Google Apps Script backend
//
// Setup:
//   1. Create a new Google Sheet → Extensions → Apps Script → paste this file.
//   2. Run `setup` once (writes the header row, asks for permission).
//   3. Deploy → New deployment → Web app
//        Execute as: Me    Who has access: Anyone
//   4. Copy the /exec URL into SCRIPT_URL in index.html.
//
// After editing this script: Deploy → Manage deployments → Edit → Version: New version
// (keeps the same URL; "New deployment" would create a different URL).

const SHEET_NAME = '喜餅';
const TIMESTAMP_FORMAT = 'yyyy-mm-dd hh:mm:ss';

// Must match the keys the page appends to FormData, in column order.
const HEADERS = [
  'Timestamp',
  '收件人姓名',
  '聯絡電話',
  '與新人的關係',
  '寄送方式',
  '寄送地址',
  '對新人的祝福',
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function setup() {
  const sheet = getSheet_();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange('A2:A').setNumberFormat(TIMESTAMP_FORMAT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const sheet = getSheet_();
    if (sheet.getLastRow() === 0) setup();

    const p = e.parameter;
    const row = HEADERS.map(h => h === 'Timestamp' ? new Date() : (p[h] || '').toString().trim());
    const range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length);
    range.setValues([row]);
    range.getCell(1, 1).setNumberFormat(TIMESTAMP_FORMAT);

    return json_({ result: 'success' });
  } catch (err) {
    return json_({ result: 'error', error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
