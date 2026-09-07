# Daily Sheet Change Summary

A Google Apps Script that watches a spreadsheet you only have **view**
access to, compares it against a snapshot from the last time it ran, and
emails you a daily summary of what changed (rows added, rows removed,
cells edited). It runs entirely inside your own Google account — no
external server or credentials needed.

Target sheet: `1VfQboGGRzDX2FbqWBtSes0oCVPBmW9mcNuGxLj37dkw`

## How it works

Because you don't have edit rights on the source spreadsheet, this is a
**standalone** script (created in your own Drive, not attached to the
shared sheet) rather than a bound one:

- It opens the source spreadsheet **read-only** by ID (`SpreadsheetApp.openById`
  works fine with just Viewer access — it only needs to call the source
  for edit-only operations, which this script never does).
- The first time it runs, it creates a new spreadsheet in **your own
  Drive** called "Moneytab Daily Diff Snapshots" to store yesterday's data
  — since you own that file, writing to it is never a permissions problem.
- Each day, it compares the source sheet's current contents against the
  matching tab in your snapshot spreadsheet, emails you the diff, then
  overwrites the snapshot with today's data.
- The very first run just establishes a baseline (nothing to compare
  against yet), so no email is sent that time.

By default rows are matched by **position** (row 2 vs row 2, etc.), which
works well for a ledger where new rows are appended at the bottom. If rows
get inserted in the middle or reordered, set `CONFIG.keyColumns` (see
below) so rows are matched by a stable key instead.

## Setup

1. Go to https://script.google.com and click **New project** (this creates
   a standalone script in your own account — do **not** open it via
   Extensions → Apps Script from inside the shared sheet, since that
   requires edit access you don't have).
2. Delete the boilerplate content and paste in the contents of
   [`Code.gs`](./Code.gs) from this repo.
3. At the top of `Code.gs`, review `CONFIG`:
   - `sourceSpreadsheetId`: already set to the sheet you shared.
   - `recipientEmail`: already set to `pt1010@gmail.com` — change if you
     want the summary sent elsewhere.
   - `sheetsToTrack`: leave as `[]` to watch every tab, or list specific
     tab names, e.g. `['Transactions', 'Budget']`.
   - `keyColumns`: leave as `[]` for position-based matching, or set to the
     header name(s) that uniquely identify a row, e.g. `['Date',
     'Description']`, if rows can be inserted/reordered in the source.
   - `triggerHour`: the hour (0–23) the daily email goes out.
4. Name the project (e.g. "Daily Diff") — it saves automatically.
5. In the function dropdown at the top, select `setupDailyTrigger` and
   click **Run**. Google will prompt you to authorize the script — it
   needs to: read the shared spreadsheet, create/edit a spreadsheet in
   your own Drive (for snapshots), and send email as you. Approve it.
6. Optionally, select `dailyDiffSummary` and click **Run** once by hand to
   establish the initial baseline immediately (instead of waiting for the
   trigger). It won't send an email on this first run.

After that, the script runs automatically once a day and emails you a
summary whenever something in the source sheet changed since the last run.

## Notes

- You never need edit access to the source sheet — the script only reads
  it. All writes go to the snapshot spreadsheet it creates in your Drive.
- Emails are sent via `MailApp`, which uses your own Gmail send quota
  (100/day on a personal account) — one email per day is well within that.
- To stop the daily emails, open the script project, click the clock
  ("Triggers") icon on the left, and delete the `dailyDiffSummary` trigger.
- If the source sheet's sharing is ever revoked, `dailyDiffSummary` will
  throw a permission error on its next run (visible in the script's
  **Executions** log) instead of silently failing.
