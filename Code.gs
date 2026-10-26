// Paste this into Extensions > Apps Script of your Google Sheet.
// The sheet needs a tab named "Members" with headers in row 1: id | name | email | last
// The passphrase is NOT in this file: set it as a Script Property named SECRET.

const MAX_MEMBERS = 10;

function secretOk_(token) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  return !!secret && token === secret;
}

function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Stops text like "=1+1" from being run as a formula in the sheet.
function safe_(v) {
  v = String(v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function doGet(e) {
  if (!secretOk_(e.parameter.token)) return out_({ ok: false, error: 'unauthorized' });
  const rows = sheet_().getDataRange().getValues().slice(1)
    .filter(r => r[0] !== '')
    .map(r => ({
      id: String(r[0]),
      name: String(r[1]),
      email: String(r[2]),
      last: r[3] instanceof Date
        ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(r[3])
    }));
  return out_({ ok: true, members: rows });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if (!secretOk_(body.token)) return out_({ ok: false, error: 'unauthorized' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    sh.getRange('D:D').setNumberFormat('@'); // keep dates as plain text
    sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), 4).clearContent();
    const list = (body.members || []).slice(0, MAX_MEMBERS);
    if (list.length) {
      sh.getRange(2, 1, list.length, 4).setValues(
        list.map(m => [safe_(m.id), safe_(m.name), safe_(m.email), String(m.last)])
      );
    }
  } finally {
    lock.release();
  }
  return out_({ ok: true });
}
