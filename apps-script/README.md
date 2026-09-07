# Daily Sheet Change Summary

A Google Apps Script that watches your spreadsheet, compares it against a
snapshot from the last time it ran, and emails you a daily summary of what
changed (rows added, rows removed, cells edited). It runs entirely inside
your Google account — no external server, credentials, or hosting needed.

Target sheet: `1VfQboGGRzDX2FbqWBtSes0oCVPBmW9mcNuGxLj37dkw`

## How it works

- On each run, every tracked sheet's current contents are compared against
  a hidden snapshot sheet (`__snapshot__<name>`) saved during the previous
  run.
- The diff is emailed as HTML to `CONFIG.recipientEmail`, then the snapshot
  is overwritten with today's data so tomorrow's run compares against today.
- The first run just establishes a baseline (nothing to compare against
  yet), so no email is sent that time.

By default rows are matched by position (row 2 vs row 2, etc.), which works
well for a ledger where new rows are appended at the bottom. If rows get
inserted in the middle or reordered, set `CONFIG.keyColumns` (see below) so
rows are matched by a stable key instead.

## Setup

1. Open the spreadsheet: https://docs.google.com/spreadsheets/d/1VfQboGGRzDX2FbqWBtSes0oCVPBmW9mcNuGxLj37dkw/edit
2. Go to **Extensions → Apps Script**.
3. Delete the boilerplate `Code.gs` content and paste in the contents of
   [`Code.gs`](./Code.gs) from this repo.
4. Click the **+** next to "Files", add a new script file, and also copy in
   [`appsscript.json`](./appsscript.json) via **Project Settings → Show
   "appsscript.json" manifest file in editor** (optional — only needed if
   you want to pin the timezone explicitly).
5. At the top of `Code.gs`, edit `CONFIG`:
   - `recipientEmail`: defaults to `pt1010@gmail.com` — change if you want
     the summary sent elsewhere.
   - `sheetsToTrack`: leave as `[]` to watch every tab, or list specific tab
     names, e.g. `['Transactions', 'Budget']`.
   - `keyColumns`: leave as `[]` for position-based matching, or set to the
     header name(s) that uniquely identify a row, e.g. `['Date',
     'Description']`, if rows can be inserted/reordered.
   - `triggerHour`: the hour (0–23) the daily email goes out.
6. Save the project (name it e.g. "Daily Diff").
7. In the function dropdown at the top, select `setupDailyTrigger` and click
   **Run**. Google will prompt you to authorize the script (it needs access
   to the spreadsheet and to send email as you) — approve it.
8. Reload the spreadsheet. You'll see a new **Daily Diff** menu with
   "Run diff now" (test it immediately) and "Enable daily email" (re-run
   setup if you ever change `triggerHour`).

That's it — after step 7, the script runs automatically once a day and
emails you a summary whenever something changed.

## Notes

- The hidden `__snapshot__*` sheets store yesterday's data for comparison.
  Don't delete or rename them (they'll just be recreated on the next run if
  you do, resetting the baseline).
- Emails are sent via `MailApp`, which uses your own Gmail send quota
  (100/day on a personal account) — one email per day is well within that.
- To stop the daily emails, go to the Apps Script project's **Triggers**
  page (clock icon) and delete the `dailyDiffSummary` trigger.
