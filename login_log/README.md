# Login log

Every successful login is sent to a Google Sheet: time, player name, roster
ID, access type (Player ID / Master ID / Master password / Alliance code),
page, and device. The site can only *read* the Excel workbook (Dropbox
download link), so the log needs somewhere writable; this Google Sheet is the
only piece outside Excel and you don't need to open it after setup.

## One-time setup

1. Create a new Google Sheet (e.g. "GT Login Log").
2. In the sheet, open **Extensions -> Apps Script**, replace the editor
   contents with `Code.gs` from this folder, and save.
3. Click **Deploy -> New deployment**, choose type **Web app**, set
   **Execute as: Me** and **Who has access: Anyone**, then **Deploy** and
   authorize when prompted.
4. Copy the **Web app URL** (ends in `/exec`) and set it as `GT_LOG_URL` in
   `gt_log.js`.

## Pull the log into Excel

1. In the workbook: **Data -> From Web**, and paste
   `<Web app URL>?action=csv&key=2026GT`.
2. In the Power Query window click **Load**. A new sheet with the logins is
   added (rename it "Login Log" if you like).
3. **Data -> Refresh All** pulls the latest logins any time.

## Viewing on the site

The hidden **Login Log** page (`gt_logins.html`, linked by the faint "π" in
the bottom-right corner of the landing page) shows the same log. It opens with
`ADMIN_PASSWORD` from `Code.gs`.

## All-players access (Master IDs)

Add a workbook sheet named **Master IDs** listing every player:

| A: ID | B: Player | C: Master |
|-------|-----------|-----------|
| 1030  | DeadPoolSurvivr | TRUE |
| 1007  | SomePlayer | FALSE |

Rows ticked in column C (a checkbox, or `TRUE` / `Yes` / `Y` / `X` / `1`) log
into the Dashboard with the full alliance view (player dropdown), the same as
the master password. Everyone else sees only their own data. A header row is
fine; rows without a numeric ID in column A are ignored.

If you edit `Code.gs` later, use **Deploy -> Manage deployments -> Edit ->
New version** so the same URL keeps working.
