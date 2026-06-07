// ============================================================
// KEVIN & THEA — RSVP Google Apps Script
// ============================================================
// SETUP INSTRUCTIONS:
//
// 1. Go to https://sheets.google.com and create a new sheet.
//    Name the first tab: "RSVPs"
//    Add these headers in row 1:
//    A: Timestamp  B: Name  C: Guests  D: Attending  E: Message
//
// 2. In the sheet, click Extensions → Apps Script
//
// 3. Delete the default code and paste ALL of this file.
//
// 4. Click Deploy → New deployment
//    - Type: Web app
//    - Execute as: Me
//    - Who has access: Anyone
//    Click Deploy, copy the Web App URL.
//
// 5. In index.html find this line:
//    var APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE';
//    Replace with your copied URL.
//
// 6. Done! Test by submitting your own RSVP.
// ============================================================

var SHEET_NAME = 'RSVPs';

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    var data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(),           // Timestamp
      data.name || '',      // Full name
      data.guests || '1',   // Number of guests
      data.attending || 'yes', // Attending?
      data.message || ''    // Optional message
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var action = e.parameter.action;

    if (action === 'count') {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      // Subtract 1 for the header row
      var totalRows = sheet.getLastRow() - 1;
      var count = Math.max(0, totalRows);

      // Also count only those who are attending
      var attending = 0;
      var totalGuests = 0;
      if (count > 0) {
        var data = sheet.getRange(2, 1, count, 5).getValues();
        data.forEach(function(row) {
          if (row[3] === 'yes') {
            attending++;
            totalGuests += parseInt(row[2]) || 1;
          }
        });
      }

      return ContentService
        .createTextOutput(JSON.stringify({
          count: count,
          attending: attending,
          totalGuests: totalGuests
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
