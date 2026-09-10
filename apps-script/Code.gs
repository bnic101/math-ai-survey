/**
 * Google Apps Script backend for the "Mathematicians and AI" survey.
 *
 * Setup (one time):
 *  1. Create a new Google Sheet. Name the first tab "Responses" (or change SHEET_NAME).
 *  2. Extensions > Apps Script. Delete the default code, paste this file, save.
 *  3. Deploy > New deployment > type "Web app".
 *       Execute as:      Me
 *       Who has access:  Anyone
 *     Click Deploy, authorise, and copy the Web app URL (ends in /exec).
 *  4. Paste that URL into ENDPOINT at the bottom of index.html.
 *
 * Whenever you change this script you must Deploy > Manage deployments > edit >
 * "New version" for the change to go live at the same URL.
 */

var SHEET_NAME = "Responses";

var COLUMNS = [
  "submitted_at", "contact", "email",
  "q1", "q1_other", "q2", "q3", "q4", "q4_other", "q5",
  "q6", "q7", "q8", "q9", "q9_other", "q10", "q11", "q11_other",
  "q12", "q13", "q14", "q14_other", "q15", "q16", "q16_other",
  "q17", "q18", "user_agent"
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var body = (e && e.postData && e.postData.contents) || "{}";
    var data = JSON.parse(body);

    if (data.website) return respond({ ok: true }); // honeypot; pretend success

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    ensureHeader(sheet);

    var row = COLUMNS.map(function (c) {
      var v = data[c];
      if (v === undefined || v === null) return "";
      return String(v).slice(0, 5000);
    });
    sheet.appendRow(row);
    return respond({ ok: true });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return respond({ ok: true, message: "Survey endpoint is live. POST JSON to submit." });
}

function ensureHeader(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.appendRow(COLUMNS);
  sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
