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
   `backend/Code.gs`, save.
3. In `Code.gs`, replace `SPREADSHEET_ID` with your sheet's ID (the long string in the
   sheet URL between `/d/` and `/edit`). Save.
4. **Deploy > New deployment > Web app**:
   - Execute as: Me
   - Who has access: Anyone
   - Deploy, authorize, copy the `/exec` URL.
5. Open `meeting-questions.html`, expand **Form settings**, paste the `/exec` URL,
   save. Submissions now land in the sheet as rows.

### Reading the answers

The sheet gets one row per submission: timestamp, meeting title, client, one column
per question id, plus a `full_payload_json` column. Monica queries this sheet to
extract answers after the meeting.

## Notes

- No backend configured? The form still autosaves in the browser and can download
  answers as JSON.
- The form uses `mode: "no-cors"` for the POST, so no CORS setup is needed on the
  Apps Script side. Use `text/plain` content type (Apps Script reads `postData.contents`
  either way).
