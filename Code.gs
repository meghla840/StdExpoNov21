/**
 * Student Spot Assessment: registration form -> Google Sheet
 *
 * Setup
 * 1. Create a Google Sheet. Open Extensions > Apps Script and paste this file.
 * 2. Click Deploy > New deployment > Web app.
 *    Execute as: Me. Who has access: Anyone.
 * 3. Copy the Web app URL and paste it into CONFIG.sheetUrl in index.html.
 * 4. If you change this code later, use Deploy > Manage deployments > Edit > New version.
 */

var SHEET_NAME = 'Registrations';
var HEADERS = [
  'Submitted at', 'Assessment ID', 'Name', 'Phone',
  'Email', 'Last education', 'Preferred countries'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var d = JSON.parse(raw);

    var sheet = getSheet_();

    // Skip repeated submissions of the same slip
    var id = clean_(d.id);
    if (id && sheet.getLastRow() > 1) {
      var ids = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === id) return json_({ ok: true, duplicate: true });
      }
    }

    var row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, HEADERS.length).setValues([[
      Utilities.formatDate(new Date(), 'Asia/Dhaka', 'dd MMM yyyy, hh:mm a'),
      id,
      clean_(d.name),
      clean_(d.phone),
      clean_(d.email),
      clean_(d.edu),
      clean_(d.country) || 'Not decided yet'
    ]]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// Open the web app URL in a browser to check that it is live
function doGet() {
  return json_({ ok: true, message: 'Student Spot Assessment form is live' });
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#0f1b3b').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.getRange('D:D').setNumberFormat('@'); // keep the leading 0 in phone numbers
    sheet.setColumnWidths(1, HEADERS.length, 170);
  }
  return sheet;
}

// Trim, limit length, and stop values starting with = + - @ from running as formulas
function clean_(v) {
  var t = String(v == null ? '' : v).trim().slice(0, 300);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
