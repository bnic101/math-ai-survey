/**
 * Google Apps Script backend for the "Mathematicians and AI" survey.
 *
 * Setup (one time):
 *  1. Create a new Google Sheet. Name the first tab "Responses" (or change SHEET_NAME).
 *  2. Extensions > Apps Script. Delete the default code, paste this file, save.
 *  3. Project Settings > Script Properties: add RESULTS_KEY = <a long random string>.
 *     This key is never committed; it unlocks the raw export used for the private PDF.
 *  4. Deploy > New deployment > type "Web app".
 *       Execute as:      Me
 *       Who has access:  Anyone
 *     Click Deploy, authorise, and copy the Web app URL (ends in /exec).
 *  5. Paste that URL into ENDPOINT at the bottom of index.html.
 *
 * Whenever you change this script you must Deploy > Manage deployments > edit >
 * "New version" for the change to go live at the same URL.
 *
 * Endpoints:
 *   POST  /exec                  append a response (JSON body)
 *   GET   /exec?mode=summary     public aggregate counts + quotable free text
 *   GET   /exec?mode=raw&key=... private full export (all columns)
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

// Single-choice and multi-choice fields that the public summary counts.
var COUNT_FIELDS = [
  "role", "area", "frequency", "hours", "real_idea", "trust", "disclose",
  "caught_error", "refereeing", "believe", "proof_no_human", "fraction_5y",
  "career_worry", "advise_phd", "keep_doing", "students_use", "course_policy", "pays"
];
var MULTI_FIELDS = ["tools", "uses", "ns_reaction", "dept_should"];

// Free-text fields shown publicly only when the respondent ticked quote_ok.
var TEXT_FIELDS = [
  "phd_for", "human_research_why", "should_have_asked", "anything_else",
  "proof_no_human_other", "advise_phd_other", "ns_reaction_other",
  "role_other", "tools_other", "dept_should_other"
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var body = (e && e.postData && e.postData.contents) || "{}";
    var data = JSON.parse(body);

    if (data.website) return respond({ ok: true }); // honeypot; pretend success

    var sheet = getSheet();
    var row = COLUMNS.map(function (c) {
      var v = data[c];
      if (v === undefined || v === null) return "";
      return String(v).slice(0, 5000);
    });
    sheet.appendRow(row);
    CacheService.getScriptCache().remove("summary");
    return respond({ ok: true });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.mode === "summary") return respond(summary());
  if (p.mode === "raw") {
    var key = PropertiesService.getScriptProperties().getProperty("RESULTS_KEY");
    if (!key || p.key !== key) return respond({ ok: false, error: "forbidden" });
    return respond({ ok: true, generated_at: new Date().toISOString(), rows: readRows() });
  }
  return respond({ ok: true, message: "Survey endpoint is live. POST JSON to submit." });
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readRows() {
  var sheet = getSheet();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var values = sheet.getRange(2, 1, last - 1, COLUMNS.length).getValues();
  var rows = [];
  values.forEach(function (v) {
    var obj = {};
    COLUMNS.forEach(function (c, i) {
      var cell = v[i];
      obj[c] = cell instanceof Date ? cell.toISOString() : String(cell === null || cell === undefined ? "" : cell);
    });
    obj.hours = normalizeHours(obj.hours);
    if (isTestRow(obj)) return;
    rows.push(obj);
  });
  return rows;
}

// Early responses stored "1-5"/"5-15", which Sheets coerced into dates.
function normalizeHours(v) {
  if (v === "1-5" || v.indexOf("2026-01-05") === 0) return "1 to 5";
  if (v === "5-15" || v.indexOf("2026-05-15") === 0) return "5 to 15";
  return v;
}

function isTestRow(r) {
  var s = (r.submitted_at + " " + r.anything_else).toLowerCase();
  return s.indexOf("delete me") >= 0 || s.indexOf("test row") >= 0 || s.indexOf("live site test") >= 0;
}

function summary() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get("summary");
  if (hit) return JSON.parse(hit);

  var rows = readRows();

  function emptyCounts() {
    var c = {};
    COUNT_FIELDS.forEach(function (f) { c[f] = {}; });
    MULTI_FIELDS.forEach(function (f) { c[f] = {}; });
    return c;
  }
  function tally(c, r) {
    COUNT_FIELDS.forEach(function (f) {
      var v = (r[f] || "").trim();
      if (v) c[f][v] = (c[f][v] || 0) + 1;
    });
    MULTI_FIELDS.forEach(function (f) {
      (r[f] || "").split(";").forEach(function (part) {
        var v = part.trim();
        if (v) c[f][v] = (c[f][v] || 0) + 1;
      });
    });
  }

  var counts = emptyCounts();
  var byRole = {};          // role -> { n, counts }
  var quotes = {};
  TEXT_FIELDS.forEach(function (f) { quotes[f] = []; });

  rows.forEach(function (r) {
    tally(counts, r);
    var role = (r.role || "").trim() || "unspecified";
    if (!byRole[role]) byRole[role] = { n: 0, counts: emptyCounts() };
    byRole[role].n += 1;
    tally(byRole[role].counts, r);
    if ((r.quote_ok || "").toLowerCase() === "yes") {
      TEXT_FIELDS.forEach(function (f) {
        var v = (r[f] || "").trim();
        if (v) quotes[f].push({ text: v, role: r.role || "" });
      });
    }
  });

  var out = {
    ok: true,
    n: rows.length,
    generated_at: new Date().toISOString(),
    first_response: rows.length ? rows[0].submitted_at : null,
    last_response: rows.length ? rows[rows.length - 1].submitted_at : null,
    counts: counts,
    by_role: byRole,
    quotes: quotes
  };
  cache.put("summary", JSON.stringify(out), 120);
  return out;
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
