# Google Sheet + Drive setup

This connects the survey so **each participant becomes one row** in Google Sheets, and their **screen recording is saved in Google Drive** with the link stored in that row.

## 1. Create the Apps Script project

1. Open [script.google.com](https://script.google.com) (use Vaishali’s / research Gmail).
2. Click **New project**.
3. Delete any default code.
4. Paste everything from `Code.gs` in this folder.
5. Click **Save** and name it e.g. `Search Survey Backend`.

## 2. Create Sheet + Drive folder automatically

1. In the toolbar function dropdown, choose **`setupAll`**.
2. Click **Run**.
3. Click **Review permissions** → choose the Google account → **Advanced** → **Go to … (unsafe)** → **Allow**.
4. After it finishes: **Executions** (clock icon) or **View → Logs** and copy:
   - `spreadsheetUrl`
   - `driveFolderUrl`

You can re-check anytime by running **`getSetupInfo`**.

## 3. Deploy the web app

1. **Deploy → New deployment**
2. Type: **Web app**
3. Settings:
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Click **Deploy**
5. Copy the **Web app URL** (ends like `/exec`)

## 4. Connect the survey website

In `index.html`, set:

```js
const APPS_SCRIPT_URL = 'PASTE_THE_WEB_APP_URL_HERE';
```

Commit/push so GitHub Pages picks it up.

## 5. Test

1. Complete a short survey run (or jump to Submit after a tiny recording).
2. Confirm:
   - a new row appears in the Sheet
   - a `.webm` file appears in the Drive folder
   - **Drive Video URL** column has the link

## Notes

- Local `.webm` + metrics `.json` downloads still happen as backup.
- Very long recordings may fail Apps Script size limits; prefer sharing the survey browser tab / window rather than entire desktop if uploads fail.
- If you re-deploy the script after code changes, use **Deploy → Manage deployments → Edit → New version**.
