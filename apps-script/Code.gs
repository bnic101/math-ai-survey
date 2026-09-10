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
  "submitted_at", "contact", "email", "email_results", "email_chat", "quote_ok",
  "role", "role_other", "area",
  "frequency", "hours", "tools", "tools_other", "uses",
  "real_idea", "trust", "disclose", "caught_error", "refereeing",
  "ns_reaction", "ns_reaction_other", "believe", "proof_no_human", "proof_no_human_other",
  "fraction_5y", "career_worry", "advise_phd", "advise_phd_other", "phd_for",
  "human_research_why", "keep_doing",
  "students_use", "course_policy", "dept_should", "dept_should_other", "pays",
  "should_have_asked", "anything_else"
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
