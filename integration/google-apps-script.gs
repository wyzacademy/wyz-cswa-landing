/**
 * WYZ Academy — CSWA bookings → Google Sheet
 *
 * Setup (once):
 * 1. Create a new Google Sheet (name it e.g. "WYZ CSWA Bookings").
 * 2. Extensions > Apps Script. Delete the sample code, paste this whole file, Save.
 * 3. Deploy > New deployment > type "Web app"
 *      Execute as: Me
 *      Who has access: Anyone
 *    Click Deploy, approve the permissions, copy the Web app URL (ends with /exec).
 * 4. Send that URL so it goes into SHEET_ENDPOINT in index.html.
 *
 * Editing this code later: Deploy > Manage deployments > Edit > Version: New version.
 * (A brand-new deployment gives a NEW URL; editing the existing one keeps the same URL.)
 */

const SHEET_NAME = 'Bookings';
const HEADERS = ['Timestamp', 'Booking ID', 'Name', 'Phone', 'Email', 'Package', 'Price (EGP)',
                 'Language', 'Payment status', 'Notes', 'Page'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (!d.id || !d.phone) return json_({ ok: false, error: 'missing fields' });

    const sh = sheet_();
    sh.appendRow([
      new Date(),
      d.id,
      String(d.name || '').slice(0, 120),
      "'" + String(d.phone || '').slice(0, 20),        // leading ' keeps the 0 in 010…
      String(d.email || '').slice(0, 120),
      'Package ' + d.pkg + (d.pkgName ? ' · ' + d.pkgName : ''),
      Number(d.price) || '',
      d.lang || '',
      'Pending',
      '',
      String(d.page || '').slice(0, 300)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Opening the /exec URL in a browser shows this — handy to confirm the deployment works.
function doGet() {
  return ContentService.createTextOutput('WYZ bookings endpoint is live.');
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'Paid', 'Cancelled'], true).build();
    sh.getRange('I2:I').setDataValidation(rule);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
