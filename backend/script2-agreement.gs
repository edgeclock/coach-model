/**
 * SCRIPT 2: AGREEMENT BACKEND (PDF + EMAIL)
 * Receives signed agreement submissions from agreement.html.
 *  1. Appends the signed record to the Answers tab of the agreement sheet.
 *  2. Builds a PDF of the agreement with both drawn signatures embedded
 *     (via a temporary Google Doc, deleted after export).
 *  3. Emails the PDF to the client and to Edge.
 *
 * DEPLOY:
 * 1. Create a NEW Google Sheet for agreements: https://sheets.new
 *    (e.g. "Agreements"). Copy its ID from the URL (/d/.../edit).
 * 2. In that sheet: Extensions > Apps Script. Paste this file.
 * 3. Set SPREADSHEET_ID to the NEW sheet's ID. Verify EDGE_EMAIL.
 * 4. Deploy > New deployment > Web app:
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Deploy, authorize ALL scopes (Sheets, Mail, Docs, Drive).
 *    Copy the /exec URL and give it to Monica so she can point
 *    agreement.html at it.
 */

const SPREADSHEET_ID = "REPLACE_WITH_AGREEMENT_SHEET_ID";
const EDGE_EMAIL = "earlsaldajeno@gmail.com";

const RESPONDENT_FIELDS = ["name", "business", "email", "phone"];

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
        "client_name", "client_email", "client_date", "edge_name", "edge_date", "agreed",
        "signed_agreement_json", "full_payload_json"];
      sheet.appendRow(headers);
    }

    const row = [
      new Date().toISOString(),
      (payload.meta && payload.meta.title) || "",
      (payload.meta && payload.meta.client) || "",
      ...RESPONDENT_FIELDS.map(f => (respondent[f] || "")),
      (answers.sig_client_name && answers.sig_client_name.answer) || "",
      (answers.sig_client_email && answers.sig_client_email.answer) || (respondent.email || ""),
      (answers.sig_client_date && answers.sig_client_date.answer) || "",
      (answers.sig_edge_name && answers.sig_edge_name.answer) || "",
      (answers.sig_edge_date && answers.sig_edge_date.answer) || "",
      (answers.sig_agreed && answers.sig_agreed.answer) || "",
      JSON.stringify({ client_image: answers.sig_client_image || null, edge_image: answers.sig_edge_image || null }),
      JSON.stringify(payload)
    ];
    sheet.appendRow(row);

    const emailResult = sendAgreementEmails_(payload, answers, respondent);

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
    .createTextOutput("Agreement backend is live. POST JSON to this URL.")
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
    const clientName = (answers.sig_client_name && answers.sig_client_name.answer) || respondent.name || "Client";
    const clientEmail = ((answers.sig_client_email && answers.sig_client_email.answer) || (respondent.email || "")).trim();

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
      "PDF attached. The signed record is also in the agreement sheet.";

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
