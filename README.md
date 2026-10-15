# report-ads-spent-today

Export the configured reporting sheets as landscape PNG snapshots and send each image to Telegram with its caption.

## Capture behavior

- Existing export behavior is preserved: one PNG image is sent per configured sheet with the current caption text.
- Runs can be started manually from GitHub Actions or from the Google Sheets menu.
- Scheduled runs are dispatched by a Google Apps Script poller. The schedule is read live from the control tab; it is not hard-coded into GitHub Actions cron.
- Default daily times: `10:00`, `14:00`, `17:00`, and `00:30`. All four time slots are editable.
- Automatic capture can be turned on/off in cell `B3` of tab `CAU_HINH_CHUP_NO`. Manual captures remain available when automatic capture is off.
- Time zone is editable in cell `B9` using an IANA name, for example `Asia/Ho_Chi_Minh`, `America/Los_Angeles`, or `Europe/London`.
- Google Apps Script checks every minute and dispatches only at a configured slot. Apps Script/GitHub Actions are not hard real-time systems, so capture can start a little after the selected minute.

## Required GitHub repository secrets

Keep existing secrets configured for this exporter:

- `GOOGLE_SERVICE_ACCOUNT_JSON`
- `SPREADSHEET_ID`
- `SHEET_NAMES`
- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID`

The Google service account needs access to the reporting spreadsheet. Never commit tokens or service-account JSON.

## Set up the shared schedule controller (once only)

Install the same shared controller in **one** Google control spreadsheet, not separately in both repositories' spreadsheets; otherwise both controllers could dispatch duplicate jobs.

1. Open the control spreadsheet, then **Extensions → Apps Script**.
2. Copy the repository file `apps-script/Code.gs` into the Apps Script editor and save.
3. In **Project Settings → Script Properties**, add `GITHUB_TOKEN` with a fine-grained GitHub token that has **Contents: Read and write** access to both `sheet-export-telegram-bot` and `report-ads-spent-today`. Do not put the token in code.
4. Reload the spreadsheet and choose **CHỤP BÁO CÁO → Khởi tạo cấu hình (giữ thiết lập hiện có)** once. The first initialization sets the requested default times and time zone; later calls do not overwrite manually edited slots.
5. Choose **CHỤP BÁO CÁO → Cài kiểm tra lịch tự động** once and approve Google permissions.
6. In tab `CAU_HINH_CHUP_NO`, set `B3` to `BẬT` or `TẮT`, edit times in `B5:B8`, and edit the time zone in `B9`.
7. Use the **CHỤP BÁO CÁO** menu to start either report or both reports immediately, independent of the automatic switch.

The script writes the most recent **dispatch request** and its status to rows 11–13. A dispatch accepted by GitHub means the workflow was queued; it does not by itself confirm that Telegram delivery succeeded. Check **Actions** in the corresponding repository for the final job result.

## Run manually in GitHub

Open **Actions → Export Sheets to Telegram (scheduled and manual) → Run workflow**. Manual runs are independent of the automatic switch.

## Privacy and secrets

Never commit access tokens or service-account JSON. Keep report spreadsheet IDs and Telegram credentials in repository secrets.
