/**
 * Daily Sheet Change Summary
 *
 * Compares every tracked sheet in this spreadsheet against a snapshot taken
 * the last time it ran, and emails a summary of rows added, removed, and
 * changed. Designed to run once per day via a time-driven trigger.
 *
 * Setup: see README.md in this folder.
 */

const CONFIG = {
  // Where the daily summary is sent.
  recipientEmail: 'pt1010@gmail.com',

  // Sheet (tab) names to watch. Leave empty to track every visible sheet.
  sheetsToTrack: [],

  // Column headers (from row 1) that uniquely identify a row, e.g.
  // ['Date', 'Description']. Rows are matched by this key across runs, so
  // insertions/reorders are handled correctly. Leave empty to match rows by
  // position instead (simplest, but a row inserted in the middle will show
  // up as changes to every row after it).
  keyColumns: [],

  // Hidden sheets used to store yesterday's data are prefixed with this.
  snapshotPrefix: '__snapshot__',

  // Hour of day (0-23, script timezone) the daily trigger fires.
  triggerHour: 7,

  // Send an email even on days with no changes.
  sendEvenIfNoChanges: false,
};

/** Adds a menu so you can run things by hand from the Sheets UI. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Daily Diff')
    .addItem('Run diff now', 'dailyDiffSummary')
    .addItem('Enable daily email (one-time setup)', 'setupDailyTrigger')
    .addToUi();
}

/** Run once to schedule the daily job. Safe to run again to reset the time. */
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'dailyDiffSummary')
    .forEach((t) => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('dailyDiffSummary')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.triggerHour)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Daily diff scheduled for ~${CONFIG.triggerHour}:00 each day.`
  );
}

/** Main entry point: diff every tracked sheet and email a summary. */
function dailyDiffSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getTrackedSheets(ss);
  const results = [];

  sheets.forEach((sheet) => {
    const current = readSheetData(sheet);
    const snapshotSheet = ss.getSheetByName(CONFIG.snapshotPrefix + sheet.getName());
    const previous = snapshotSheet ? readSheetData(snapshotSheet) : null;

    const diff = diffData(previous, current, CONFIG.keyColumns);
    if (diff.hasChanges) {
      results.push({ sheetName: sheet.getName(), diff });
    }

    saveSnapshot(ss, sheet, current);
  });

  if (results.length > 0) {
    sendSummaryEmail(results);
  } else if (CONFIG.sendEvenIfNoChanges) {
    sendNoChangeEmail();
  }
}

function getTrackedSheets(ss) {
  const all = ss.getSheets().filter((s) => !s.getName().startsWith(CONFIG.snapshotPrefix));
  if (CONFIG.sheetsToTrack.length === 0) return all;
  return all.filter((s) => CONFIG.sheetsToTrack.includes(s.getName()));
}

function readSheetData(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length === 0) return { headers: [], rows: [] };
  return { headers: values[0], rows: values.slice(1) };
}

function buildKey(headers, row, keyColumns) {
  if (!keyColumns || keyColumns.length === 0) return null;
  return keyColumns
    .map((col) => {
      const idx = headers.indexOf(col);
      return idx === -1 ? '' : String(row[idx]);
    })
    .join('|');
}

function compareRows(oldHeaders, newHeaders, oldRow, newRow) {
  const changes = [];
  newHeaders.forEach((col, idx) => {
    const oldIdx = oldHeaders.indexOf(col);
    const oldVal = oldIdx === -1 ? undefined : oldRow[oldIdx];
    const newVal = newRow[idx];
    if (String(oldVal) !== String(newVal)) {
      changes.push({ column: col, oldValue: oldVal, newValue: newVal });
    }
  });
  return changes;
}

function diffData(previous, current, keyColumns) {
  const headers = current.headers;
  const useKeys = keyColumns && keyColumns.length > 0;
  const added = [];
  const removed = [];
  const modified = [];

  if (!previous) {
    // No baseline yet (first run) — just establish one, don't report a diff.
    return { hasChanges: false, added, removed, modified, firstRun: true, headers };
  }

  if (useKeys) {
    const prevMap = new Map();
    previous.rows.forEach((r) => prevMap.set(buildKey(previous.headers, r, keyColumns), r));
    const currMap = new Map();
    current.rows.forEach((r) => currMap.set(buildKey(headers, r, keyColumns), r));

    currMap.forEach((row, key) => {
      if (!prevMap.has(key)) {
        added.push(row);
      } else {
        const changes = compareRows(previous.headers, headers, prevMap.get(key), row);
        if (changes.length > 0) modified.push({ label: key, changes });
      }
    });
    prevMap.forEach((row, key) => {
      if (!currMap.has(key)) removed.push(row);
    });
  } else {
    const maxLen = Math.max(previous.rows.length, current.rows.length);
    for (let i = 0; i < maxLen; i++) {
      const oldRow = previous.rows[i];
      const newRow = current.rows[i];
      if (oldRow === undefined) {
        added.push(newRow);
        continue;
      }
      if (newRow === undefined) {
        removed.push(oldRow);
        continue;
      }
      const changes = compareRows(previous.headers, headers, oldRow, newRow);
      if (changes.length > 0) modified.push({ label: `Row ${i + 2}`, changes });
    }
  }

  const hasChanges = added.length > 0 || removed.length > 0 || modified.length > 0;
  return { hasChanges, added, removed, modified, headers };
}

function saveSnapshot(ss, sheet, data) {
  const name = CONFIG.snapshotPrefix + sheet.getName();
  let snap = ss.getSheetByName(name);
  if (!snap) {
    snap = ss.insertSheet(name);
    snap.hideSheet();
  } else {
    snap.clear();
  }
  const values = [data.headers, ...data.rows];
  if (values.length > 0 && values[0].length > 0) {
    snap.getRange(1, 1, values.length, values[0].length).setValues(values);
  }
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function formatRow(row) {
  return row.map((v) => escapeHtml(v)).join(', ');
}

function sendSummaryEmail(results) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let html = `<h2>Daily Sheet Change Summary — ${today}</h2>`;

  results.forEach(({ sheetName, diff }) => {
    html += `<h3>${escapeHtml(sheetName)}</h3>`;

    if (diff.added.length) {
      html += `<p><b>${diff.added.length} row(s) added</b></p><ul>`;
      diff.added.forEach((row) => (html += `<li>${formatRow(row)}</li>`));
      html += '</ul>';
    }

    if (diff.removed.length) {
      html += `<p><b>${diff.removed.length} row(s) removed</b></p><ul>`;
      diff.removed.forEach((row) => (html += `<li>${formatRow(row)}</li>`));
      html += '</ul>';
    }

    if (diff.modified.length) {
      html += `<p><b>${diff.modified.length} row(s) changed</b></p><ul>`;
      diff.modified.forEach((m) => {
        const changeStr = m.changes
          .map((c) => `${escapeHtml(c.column)}: "${escapeHtml(c.oldValue)}" &rarr; "${escapeHtml(c.newValue)}"`)
          .join('; ');
        html += `<li><b>${escapeHtml(m.label)}</b> — ${changeStr}</li>`;
      });
      html += '</ul>';
    }
  });

  MailApp.sendEmail({
    to: CONFIG.recipientEmail,
    subject: `Daily Sheet Change Summary — ${today}`,
    htmlBody: html,
  });
}

function sendNoChangeEmail() {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  MailApp.sendEmail({
    to: CONFIG.recipientEmail,
    subject: `Daily Sheet Change Summary — ${today}`,
    htmlBody: `<p>No changes detected today (${today}).</p>`,
  });
}
