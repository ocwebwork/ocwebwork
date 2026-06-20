// ============================================================
// OC Web Work — contact form handler for Google Sheets
// Paste this whole file into the Apps Script editor, Save, then
// Deploy > Manage deployments > edit (pencil) > New version > Deploy.
// IMPORTANT: "Who has access" MUST be set to "Anyone".
// ============================================================

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Add a header row the first time
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Name", "Business", "Current Website", "Email", "Budget", "Message"]);
      sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#1B1A16").setFontColor("#CE8E16");
      sheet.setFrozenRows(1);
    }

    // Accept BOTH form-encoded fields and raw JSON, so it works no matter how the site sends it
    var d = (e && e.parameter) ? e.parameter : {};
    if ((!d.name && !d.email) && e && e.postData && e.postData.contents) {
      try { d = JSON.parse(e.postData.contents); } catch (ignore) {}
    }

    sheet.appendRow([
      new Date().toLocaleString("en-US", { timeZone: "America/Los_Angeles" }),
      d.name            || "",
      d.business        || "",
      d.current_website || "",
      d.email           || "",
      d.budget          || "",
      d.message         || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Visiting the URL in a browser shows this — confirms the app is reachable.
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ result: "ok", message: "OC Web Work form endpoint is live." }))
    .setMimeType(ContentService.MimeType.JSON);
}
