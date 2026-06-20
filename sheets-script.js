// Paste this entire file into Google Apps Script (instructions below).
// It receives contact form submissions and appends them as rows in your sheet.

const SHEET_NAME = "Submissions";

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    // Create the header row on first submission if the sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp",
        "Name",
        "Business",
        "Current Website",
        "Email",
        "Budget",
        "Message"
      ]);
      sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#1B1A16").setFontColor("#CE8E16");
      sheet.setFrozenRows(1);
    }

    var data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date().toLocaleString("en-US", { timeZone: "America/Los_Angeles" }),
      data.name        || "",
      data.business    || "",
      data.current_website || "",
      data.email       || "",
      data.budget      || "",
      data.message     || ""
    ]);

    // Auto-resize columns for readability
    sheet.autoResizeColumns(1, 7);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handles CORS preflight from the browser
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ result: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}
