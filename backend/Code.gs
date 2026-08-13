/**
 * GBI Meeting Answers + Signed Agreement backend
 * Google Apps Script web app that:
 *  1. Appends form submissions to a Google Sheet as rows.
 *  2. For signed agreements: builds a PDF of the agreement with the
 *     drawn signatures, emails it to the client and to Edge.
 *
 * DEPLOY:
 * 1. Create a Google Sheet: https://sheets.new (e.g. "Client Discovery").
 * 2. Extensions > Apps Script. Delete default content, paste this file, save.
 * 3. Set SPREADSHEET_ID below to your sheet's ID. Set EDGE_EMAIL to Edge's email.
 * 4. Deploy > New deployment > Web app:
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Deploy, authorize (this version also needs MailApp/DocumentApp/DriveApp
 *    scopes - accept them), copy the /exec URL.
 * 5. The form and agreement pages post to that URL (FORM_ENDPOINT baked in).
 *
 * IMPORTANT: after editing this file, re-deploy as a NEW VERSION and copy the
 * new /exec URL if it changed.
 */

const SPREADSHEET_ID = "1om_Q6oYbwZyY0IaaI1bMwztz6YZ-o6YFfLs7qO8WWQc";
const EDGE_EMAIL = "earlsaldajeno@gmail.com";

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

const AGREEMENT_CLAUSES = [
  ["1. Parties", "Network Provider: Earl Jones Saldajeno (\"Edge\"), Philippines. Client: Jhun Hadji Basher, accountant, bookkeeper and fractional CFO (\"Client\"), signing personally. Client owns his own business and client relationships. If Client later registers a business entity for the practice, Client shall cause that entity to adopt and be bound by this Agreement, and Edge's consent shall not be unreasonably withheld."],
  ["2. Purpose", "Edge assists Client to establish and grow a \"business solutions\" practice (accounting, bookkeeping, fractional CFO services) targeting clients in the US, UK, Canada, and Australia. Edge owns no equity in Client's business. Edge provides: business setup support, coaching, website, deck and marketing build, lead generation support, and a license to Edge's brand and method."],
  ["3. Network Fee", "Client pays Edge 10% of Adjusted Collected Revenue, monthly, during the term of this Agreement."],
  ["4. Adjusted Collected Revenue", "Cash actually collected from covered clients in a calendar month, minus: (a) VAT or percentage tax actually remitted to the BIR on that revenue, (b) documented pass-through costs reimbursed to Client at cost (no markup), and (c) refunds and chargebacks issued in the same month. No deduction for salaries, rent, marketing, software, owner draws, or any other expense. Revenue is measured on cash collected, not profit, billings, or accruals."],
  ["5. Covered Revenue", "All revenue of Client's practice, unless a legacy-client carve-out is agreed in writing before signing (attach list). Carve-out clients' revenue is excluded from Adjusted Collected Revenue; Edge still gets 10% of all new clients sourced after the start date."],
  ["6. Computation and Payment", "Client computes Adjusted Collected Revenue and reports it to Edge by the 10th of each month. Edge issues an invoice. Client pays by the 15th of the same month. Late payment: simple interest at 1% per month."],
  ["7. Verification", "Client grants Edge read-only access to the practice's books and records (accounting software and bank statements) sufficient to verify Adjusted Collected Revenue. Client keeps records for at least 3 years."],
  ["8. Term and Renewal", "12 months from the start date. Auto-renews for successive 12-month periods unless either party gives 30 days written notice before renewal."],
  ["9. Exit", "On termination, Client keeps his business, entity, and clients. Edge is entitled to a short collection tail: the Network Fee continues on revenue collected within 60 days after termination. Client may not copy, replicate, or sublicense Edge's method, brand, or materials for 36 months after termination. For 24 months after termination, Client may not solicit clients sourced through Edge's network or Edge's network personnel."],
  ["10. Relationship", "This Agreement creates no partnership, employment, agency, or joint venture. Client is an independent business owner. Neither party may bind the other."],
  ["11. Law and Venue", "Philippine law governs. Exclusive venue: the courts of the place where the Agreement is signed."],
  ["12. Severability", "If any provision is unenforceable, the rest remains in full force."],
  ["13. Signatures", "Both parties sign below. Client signs personally; if Client later registers a business entity for the practice, Client agrees to cause that entity to adopt this Agreement as well."]
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
        ...QUESTION_IDS, "signed_agreement_json", "full_payload_json"];
      sheet.appendRow(headers);
    }

    const row = [
      new Date().toISOString(),
      (payload.meta && payload.meta.title) || "",
      (payload.meta && payload.meta.client) || "",
      ...RESPONDENT_FIELDS.map(f => (respondent[f] || "")),
      ...QUESTION_IDS.map(id => (answers[id] && answers[id].answer) || ""),
      "",
      JSON.stringify(payload)
    ];
    sheet.appendRow(row);

    // If this is a signed agreement, build PDF and email it.
    const isAgreement = payload.meta && payload.meta.title &&
      String(payload.meta.title).toUpperCase().indexOf("AGREEMENT") !== -1;
    let emailResult = "not-agreement";
    if (isAgreement) {
      emailResult = sendAgreementEmails_(payload, answers, respondent);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, row: sheet.getLastRow(), email: emailResult }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput("Backend is live. POST JSON to this URL.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Answers");
  if (!sheet) sheet = ss.insertSheet("Answers");
  return sheet;
}

function sendAgreementEmails_(payload, answers, respondent) {
  try {
    const pdf = buildAgreementPdf_(payload, answers, respondent);
    const clientName = respondent.name || "Client";
    const clientEmail = (respondent.email || "").trim();

    const subject = "Your signed Network Fee Agreement";
    const body =
      "Dear " + clientName + ",\n\n" +
      "Please find attached your signed copy of the Network Fee Agreement " +
      "between you and Earl Jones Saldajeno (\"Edge\").\n\n" +
      "Keep this copy for your records. If you have any questions, reply to this email.\n\n" +
      "Best regards,\nEdge's Team";

    if (clientEmail) {
      MailApp.sendEmail({
        to: clientEmail,
        subject: subject,
        body: body,
        attachments: [pdf]
      });
    }

    const edgeBody =
      "Signed Network Fee Agreement received.\n\n" +
      "Client: " + clientName + "\n" +
      "Client email: " + (clientEmail || "(none provided)") + "\n" +
      "Signed at: " + new Date().toISOString() + "\n\n" +
      "PDF attached. The signed record is also in the Google Sheet.";

    MailApp.sendEmail({
      to: EDGE_EMAIL,
      subject: "Signed agreement: " + clientName,
      body: edgeBody,
      attachments: [pdf]
    });

    return clientEmail ? "sent-to-both" : "sent-to-edge-only";
  } catch (err) {
    return "email-failed: " + String(err);
  }
}

function buildAgreementPdf_(payload, answers, respondent) {
  const doc = DocumentApp.create("Network Fee Agreement - Signed " + new Date().toISOString());
  const body = doc.getBody();

  body.appendParagraph("NETWORK FEE AGREEMENT")
    .setHeading(DocumentApp.ParagraphHeading.TITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph("Executed electronically on " + new Date().toLocaleDateString() + "\n");
  body.appendParagraph("Date signed: " + ((answers.sig_client_date && answers.sig_client_date.answer) || ""));
  body.appendParagraph("");

  AGREEMENT_CLAUSES.forEach(function (pair) {
    body.appendParagraph(pair[0]).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(pair[1]);
  });

  body.appendParagraph("");
  body.appendParagraph("SIGNATURES").setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendParagraph("CLIENT");
  body.appendParagraph("Name: " + ((answers.sig_client_name && answers.sig_client_name.answer) || ""));
  if (answers.sig_client_image && answers.sig_client_image.answer) {
    body.appendParagraph("Signature:");
    body.appendImage(dataUrlToBlob_(answers.sig_client_image.answer, "client_signature.png"));
  }
  body.appendParagraph("Date: " + ((answers.sig_client_date && answers.sig_client_date.answer) || ""));

  body.appendParagraph("");
  body.appendParagraph("NETWORK PROVIDER (EDGE)");
  body.appendParagraph("Name: " + ((answers.sig_edge_name && answers.sig_edge_name.answer) || ""));
  if (answers.sig_edge_image && answers.sig_edge_image.answer) {
    body.appendParagraph("Signature:");
    body.appendImage(dataUrlToBlob_(answers.sig_edge_image.answer, "edge_signature.png"));
  }
  body.appendParagraph("Date: " + ((answers.sig_edge_date && answers.sig_edge_date.answer) || ""));

  doc.saveAndClose();

  const pdf = DriveApp.getFileById(doc.getId()).getAs("application/pdf");
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return pdf;
}

function dataUrlToBlob_(dataUrl, filename) {
  const parts = dataUrl.split(",");
  const mime = parts[0].match(/data:(.*?);/)[1];
  const bytes = Utilities.base64Decode(parts[1]);
  return Utilities.newBlob(bytes, mime, filename);
}
