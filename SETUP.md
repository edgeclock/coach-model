# Coach Model GBI Briefing Kit

Live site: https://edgeclock.github.io/coach-model/

## Pages

- **index.html** - landing page linking both tools
- **explainer.html** - "What is GBI and what it does to our business model" walkthrough (for Edge)
- **meeting-questions.html** - submittable 40-question meeting form (for the client meeting)

## How answer capture works

The form stores answers in browser localStorage (autosave) and can POST a JSON payload to a
Google Apps Script web app, which appends each submission as a row in a Google Sheet.
Monica reads the sheet directly via Google Sheets access.

### Deploy the backend (5 minutes)

1. Create a Google Sheet: https://sheets.new (name it, e.g. "Client1 Meeting Answers").
2. In the sheet: **Extensions > Apps Script**. Delete default content, paste
   `backend/Code.gs`, save. Set `SPREADSHEET_ID` and `EDGE_EMAIL` at the top of
   the file.
3. In `Code.gs`, replace `SPREADSHEET_ID` with your sheet's ID (the long string in the
   sheet URL between `/d/` and `/edit`). Save.
4. **Deploy > New deployment > Web app**:
   - Execute as: Me
   - Who has access: **Anyone** (this is critical, anonymous form posts get 401 without it)
   - Deploy, authorize, copy the `/exec` URL.
5. Tell Monica the `/exec` URL (or send it in chat). She bakes it into
   `meeting-questions.html` and `agreement.html` (the `FORM_ENDPOINT` constant)
   and re-pushes. Submissions now land in the sheet as rows.

## Signed agreement email flow

The agreement page posts the signed agreement (with drawn signature images).
The backend then:
1. Appends the signed record to the sheet (same Answers tab).
2. Builds a PDF of the agreement with both signatures embedded (via a temporary
   Google Doc, deleted after export).
3. Emails the PDF to the client (their email from the signature form) and to
   `EDGE_EMAIL`.

This version needs extra authorization scopes (MailApp, DocumentApp, DriveApp).
When you re-deploy and authorize, accept all scopes. If the `/exec` URL changes
after re-deploy, give Monica the new URL so she can update the forms.

### Reading the answers

The sheet gets one row per submission: timestamp, meeting title, client,
respondent name/business/email/phone, one column per question id, plus a
`full_payload_json` column. Monica queries this sheet to extract answers
after the meeting.

## Notes

- No backend reachable? The form still autosaves in the browser and can download
  answers as JSON (Download button).
- The form uses `mode: "no-cors"` for the POST, so no CORS setup is needed on the
  Apps Script side. Use `text/plain` content type (Apps Script reads `postData.contents`
  either way).
