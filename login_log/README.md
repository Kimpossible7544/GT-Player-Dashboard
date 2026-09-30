# Login log

Every successful login is sent to a private Google Sheet: time, player name,
roster ID, access type (Player ID / Master ID / Master password / Alliance
code), page, and device. The hidden **Login Log** page (`gt_logins.html`,
linked by the faint symbol in the bottom-right corner of the landing page)
shows the log and lets you grant or remove all-players access for roster IDs.
It only opens with `ADMIN_PASSWORD` from `Code.gs`.

## Setup (one time)

1. Create a new Google Sheet (e.g. "GT Login Log").
2. In the sheet, open **Extensions -> Apps Script**, replace the editor
   contents with `Code.gs` from this folder, and save.
3. Click **Deploy -> New deployment**, choose type **Web app**, set
   **Execute as: Me** and **Who has access: Anyone**, then **Deploy** and
   authorize when prompted.
4. Copy the **Web app URL** (ends in `/exec`) and set it as `GT_LOG_URL` in
   `gt_log.js`.

The script creates a `Logins` tab and a `Master IDs` tab on first use. Master
IDs can be managed from the Login Log page or edited directly in the sheet.

If you edit `Code.gs` later, use **Deploy -> Manage deployments -> Edit ->
New version** so the same URL keeps working.
