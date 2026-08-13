/**
 * GBI Meeting Answers backend
 * Google Apps Script web app that receives form submissions and appends
 * them to a Google Sheet as rows.
 *
 * DEPLOY (5 minutes):
 * 1. Go to https://sheets.new and create a spreadsheet. Name it e.g.
 *    "Client1 Meeting Answers". Copy the spreadsheet ID from the URL
 *    (the long string between /d/ and /edit).
 * 2. In that spreadsheet: Extensions > Apps Script. Delete the default
 *    content, paste this whole file, save (Ctrl+S). Name the project
 *    "GBI Meeting Answers Backend".
 * 3. Set SPREADSHEET_ID below to your sheet's ID, then save.
 * 4. Deploy > New deployment > type Web app:
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Click Deploy, authorize when prompted, copy the /exec URL.
 * 5. Open meeting-questions.html, expand "Form settings", paste the
 *    /exec URL, save. Done.
 *
 * Each submission becomes one row: Timestamp, Meta title, Client,
 * Respondent name, business, email, phone, then one column per question
 * id (b1q1...closeq40), then a JSON column with the full payload.
 */

const SPREADSHEET_ID = "REPLACE_WITH_YOUR_SPREADSHEET_ID";

// Respondent detail fields, in column order.
const RESPONDENT_FIELDS = ["name", "business", "email", "phone"];

// Question ids in display order, used to build stable columns.
const QUESTION_IDS = [
  "b1q1","b1q2","b1q3","b1q4","b1q5","b1q6",
  "b2q7","b2q8","b2q9","b2q10","b2q11","b2q12",
  "b3q13","b3q14","b3q15","b3q16","b3q17",
  "b4q18","b4q19","b4q20","b4q21",
  "b5q22","b5q23","b5q24","b5q25","b5q26",
  "b6q27","b6q28","b6q29","b6q30","b6q31","b6q32","b6q33","b6q34","b6q35","b6q36",
  "b7q37","b7q38","b7q39",
  "closeq40"
];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const answers = payload.answers || {};
    const respondent = payload.respondent || {};
    const sheet = getSheet_();

    // Ensure headers exist on first run.
    if (sheet.getLastRow() === 0) {
      const headers = ["timestamp", "meeting_title", "client",
        ...RESPONDENT_FIELDS.map(f => "respondent_" + f),
        ...QUESTION_IDS, "full_payload_json"];
      sheet.appendRow(headers);
    }

    const row = [
      new Date().toISOString(),
      (payload.meta && payload.meta.title) || "",
      (payload.meta && payload.meta.client) || "",
      ...RESPONDENT_FIELDS.map(f => (respondent[f] || "")),
      ...QUESTION_IDS.map(id => (answers[id] && answers[id].answer) || ""),
      JSON.stringify(payload)
    ];
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, row: sheet.getLastRow() }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput("GBI Meeting Answers backend is live. POST JSON to this URL.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Answers");
  if (!sheet) sheet = ss.insertSheet("Answers");
  return sheet;
}
