/**
 * RSVP receiver for the "Srujana-Wedding-Visitors" Google Sheet.
 *
 * Setup (once):
 *  1. Open the sheet > Extensions > Apps Script.
 *  2. Replace everything in Code.gs with this file and click Save.
 *  3. Deploy > New deployment > type "Web app":
 *       Execute as: Me
 *       Who has access: Anyone
 *     Click Deploy and allow the permissions it asks for.
 *  4. Copy the Web app URL (ends in /exec) into RSVP_SHEET_URL on the website.
 *
 * Each RSVP becomes one row on a tab named "RSVPs" (created automatically).
 * The sheet itself can stay "Restricted"; the script writes to it as you.
 */

var SHEET_NAME = 'RSVPs';
var HEADERS = ['Submitted at', 'Name', 'Phone', 'Attending', 'Guests', 'Events', 'Message'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }
    var p = (e && e.parameter) || {};
    sheet.appendRow([
      new Date(),
      clean(p.name, 120),
      clean(p.phone, 30),
      clean(p.attending, 10),
      clean(p.guests, 4),
      clean(p.events, 200),
      clean(p.message, 2000)
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Visiting the URL in a browser shows this, which is a quick way to check it is live.
function doGet() {
  return ContentService.createTextOutput('RSVP receiver is running.');
}

// Trim, cap the length, and stop text that starts with = + - @ from being run as a formula.
function clean(value, max) {
  var s = String(value == null ? '' : value).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
