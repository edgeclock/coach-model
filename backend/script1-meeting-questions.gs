/**
 * SCRIPT 1: MEETING QUESTIONS BACKEND
 * Receives form submissions from meeting-questions.html and appends
 * each as a row in the Answers tab of the "Client Discovery" sheet.
 * No email, no PDF. Simple and stable.
 *
 * DEPLOY:
 * 1. Extensions > Apps Script (in the Client Discovery sheet).
 * 2. Paste this file, save.
 * 3. Deploy > New deployment > Web app:
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Deploy, authorize, copy the /exec URL.
 * 4. That URL is the FORM_ENDPOINT used by meeting-questions.html.
 */

const SPREADSHEET_ID = "1om_Q6oYbwZyY0IaaI1bMwztz6YZ-o6YFfLs7qO8WWQc";

const RESPONDENT_FIELDS = ["name", "business", "email", "phone"];

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
