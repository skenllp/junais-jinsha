/**
 * Google Apps Script for Muhammed Junais & Fathima Jinsha Wedding RSVP
 * Connects the website RSVP form directly to Google Sheets.
 *
 * HOW TO SET UP:
 * 1. Open your Google Sheet ("junais-jinsha").
 * 2. In the top menu, click Extensions > Apps Script.
 * 3. Delete any code in Code.gs and replace it with this file.
 * 4. Click the blue "Deploy" button (top right) > "New deployment".
 * 5. Click the gear icon next to "Select type" > select "Web app".
 * 6. Set:
 *    - Description: "RSVP Web App"
 *    - Execute as: "Me" (your email)
 *    - Who has access: "Anyone" (IMPORTANT!)
 * 7. Click "Deploy". Grant permissions if prompted (Advanced > Go to RSVP Web App (unsafe)).
 * 8. Copy the "Web app URL" (ends with /exec).
 * 9. Paste that URL into config.js as `rsvpScriptUrl: "https://script.google.com/macros/s/.../exec"`.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  // Wait up to 30 seconds for other concurrent submissions to complete
  try {
    lock.waitLock(30000);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: "Server busy, please try again" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // Try to get sheet named "junais-jinsha", otherwise get active sheet
    var sheet = ss.getSheetByName("junais-jinsha") || ss.getActiveSheet();

    // Auto-create header row if the sheet is completely blank
    if (sheet.getLastRow() === 0) {
      var headers = ["Timestamp", "Guest Name", "Attending", "Number of Guests", "Phone Number", "Note / Wishes"];
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#ab8644");
      headerRange.setFontColor("#ffffff");
      headerRange.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    // Extract parameters from form data or JSON body
    var params = {};
    if (e && e.parameter && Object.keys(e.parameter).length > 0) {
      params = e.parameter;
    } else if (e && e.postData && e.postData.contents) {
      try {
        params = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        params = {};
      }
    }

    var timestamp = params.timestamp || new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" });
    var name = params.name || "Anonymous";
    var attending = params.attending || "Joyfully yes";
    var guests = params.guests !== undefined ? params.guests : (attending === "Sadly no" ? "0" : "1");
    var phone = params.phone || "";
    var note = params.note || "";

    // Append the row
    sheet.appendRow([timestamp, name, attending, guests, phone, note]);

    // Optional: Format the newly added row
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, 6).setVerticalAlignment("middle");

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success", row: lastRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Fallback for browser testing (visiting the URL directly in browser)
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "active", message: "Muhammed Junais & Fathima Jinsha Wedding RSVP Endpoint is live!" }))
    .setMimeType(ContentService.MimeType.JSON);
}

