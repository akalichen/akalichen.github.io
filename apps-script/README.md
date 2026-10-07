# CX43 usage log — backend setup (one-time, ~5 minutes)

The form on `microscope_log.html` is static; it needs somewhere to send entries.
This folder holds the Google Apps Script that receives them, appends them to a
Google Sheet, and emails you the cumulative log as an Excel attachment after every
submission.

## 1. Create the spreadsheet
1. In Google Drive (use the Google account you want the emails to come from — your
   OSU Google Workspace account is ideal so mail comes from `@oregonstate.edu`),
   create a new Google Sheet. Name it e.g. **CX43 Microscope Usage Log**.
2. Leave it empty; the script creates the `Log` tab and headers on first use.

## 2. Add the script
1. In the sheet: **Extensions → Apps Script**.
2. Delete the placeholder code and paste the whole contents of `Code.gs`.
3. Check the `CONFIG` block at the top (`NOTIFY_EMAIL`, confirmation on/off).
4. Save (Ctrl/Cmd+S); name the project anything.

## 3. Authorize and test
1. In the function dropdown pick **`testSubmission`** and click **Run**.
2. Approve the authorization prompt (Sheets, Gmail/Mail, external URL fetch —
   the last one is used to export the sheet as `.xlsx`). If Google shows
   "unverified app", click *Advanced → Go to … (unsafe)*; this is your own script.
3. You should receive an email with an `.xlsx` attachment and a test row in the sheet.
   Delete the test row afterwards (keep the header row).

## 4. Deploy as a web app
1. **Deploy → New deployment → ⚙ Select type → Web app**.
2. Description: anything. **Execute as: Me**. **Who has access: Anyone**
   (required so the public site can post without a Google login; the script only
   accepts POSTed log entries and never exposes the sheet).
3. **Deploy**, then copy the **Web app URL** (`https://script.google.com/macros/s/…/exec`).

## 5. Connect the form
1. Open `microscope_log.html` and replace `PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE`
   with the Web app URL.
2. Commit and push. Until this URL is set, the form falls back to opening a
   pre-filled email draft to `cheng.li@oregonstate.edu`, so nothing is lost.
3. Test from the live site; the status line under the button should read
   "your session has been logged (entry #n)".

## Updating the script later
After editing `Code.gs`, go to **Deploy → Manage deployments → ✎ Edit →
Version: New version → Deploy**. The URL stays the same; forgetting the new
version is the most common reason edits "don't take".

## Notes
- Quota: a consumer Gmail account can send ~100 emails/day from Apps Script,
  a Workspace (OSU) account ~1,500/day. Each submission sends 2 emails when
  `SEND_USER_CONFIRMATION` is on.
- The attached `.xlsx` is the **entire** spreadsheet at that moment, so every
  email carries the full history, as requested.
- To stop the per-user confirmation, set `SEND_USER_CONFIRMATION: false`.
- To add a form field: add the `<input>` in `microscope_log.html` and a matching
  `['field_name', 'Column title']` entry to `FIELDS` in `Code.gs`, then redeploy a
  new version. Existing rows keep their columns; add the new header manually.
