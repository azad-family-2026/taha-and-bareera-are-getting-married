/*
  RSVP -> Google Sheet (no backend of your own; Google hosts it, free)
  =====================================================================

  ONE-TIME SETUP (~5 min):
   1. Create a new Google Sheet (this is where RSVPs will land).
   2. In that sheet: Extensions  ->  Apps Script.
   3. Delete any sample code, paste EVERYTHING below, click Save (disk icon).
   4. Click Deploy  ->  New deployment.
        - Gear icon -> select type "Web app".
        - Execute as: Me
        - Who has access: Anyone
        - Deploy  (approve/authorize the permissions when prompted).
   5. Copy the "Web app" URL — it ends in /exec.
   6. Open index.html, find the line:
          var RSVP_ENDPOINT='PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
      and paste your URL between the quotes. Save. Done.

  Responses appear as rows in the sheet. To get a CSV:
      File -> Download -> Comma-separated values (.csv)

  NOTE: if you ever change this script, do Deploy -> Manage deployments ->
  edit -> Version: "New version" so the change goes live (URL stays the same).
*/

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // avoid two submissions clobbering the same row
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('RSVPs') || ss.insertSheet('RSVPs');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Response', 'Phone']);
    }
    var p = (e && e.parameter) ? e.parameter : {};
    var row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setValue(new Date());
    sheet.getRange(row, 2).setValue(p.name || '');
    sheet.getRange(row, 3).setValue(p.response || '');
    // Phone: force PLAIN TEXT so a leading "+" is not treated as a formula (#ERROR!)
    var phoneCell = sheet.getRange(row, 4);
    phoneCell.setNumberFormat('@');
    phoneCell.setValue(p.phone || '');
    return ContentService.createTextOutput('OK');
  } finally {
    lock.releaseLock();
  }
}
