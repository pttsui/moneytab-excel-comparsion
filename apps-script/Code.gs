/**
 * Daily Sheet Change Summary (standalone / read-only source)
 *
 * For use when you only have VIEW access to the spreadsheet you want to
 * watch (so a bound script can't be installed on it). This is a standalone
 * script: it opens the source spreadsheet by ID for reading only, and
 * keeps its own snapshots in a separate spreadsheet that it creates in
 * your Drive (which you fully own, so writing to it is never a problem).
 *
 * Setup: see README.md in this folder.
 */

const CONFIG = {
  // The spreadsheet you want to watch. You only need view access to it.
  sourceSpreadsheetId: '1VfQboGGRzDX2FbqWBtSes0oCVPBmW9mcNuGxLj37dkw',

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

  // Name of the spreadsheet this script creates (in your own Drive) to
  // store yesterday's snapshot for comparison. Created automatically on
  // first run; its ID is then remembered in this script's properties.
  snapshotSpreadsheetName: 'Moneytab Daily Diff Snapshots',

  // Hour of day (0-23, script timezone) the daily trigger fires.
  triggerHour: 7,

  // Send an email even on days with no changes.
  sendEvenIfNoChanges: false,
};

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

  Logger.log(`Daily diff scheduled for ~${CONFIG.triggerHour}:00 each day.`);
}

/** Main entry point: diff every tracked sheet and email a summary. */
function dailyDiffSummary() {
  const sourceSs = SpreadsheetApp.openById(CONFIG.sourceSpreadsheetId);
  const snapshotSs = getOrCreateSnapshotSpreadsheet();
  const sheets = getTrackedSheets(sourceSs);
  const results = [];

  sheets.forEach((sheet) => {
    const current = readSheetData(sheet);
    const snapshotSheet = snapshotSs.getSheetByName(sheet.getName());
    const previous = snapshotSheet ? readSheetData(snapshotSheet) : null;

    const diff = diffData(previous, current, CONFIG.keyColumns);
    if (diff.hasChanges) {
      results.push({ sheetName: sheet.getName(), diff });
    }

    saveSnapshot(snapshotSs, sheet.getName(), current);
  });

  if (results.length > 0) {
    sendSummaryEmail(results);
  } else if (CONFIG.sendEvenIfNoChanges) {
    sendNoChangeEmail();
  }
}

/**
 * Returns the snapshot spreadsheet, creating it in your Drive the first
 * time this runs and remembering its ID for next time.
 */
function getOrCreateSnapshotSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty('SNAPSHOT_SPREADSHEET_ID');
  if (existingId) {
    try {
      return SpreadsheetApp.openById(existingId);
    } catch (e) {
      // Fall through and recreate if it was deleted/moved.
    }
  }

  const ss = SpreadsheetApp.create(CONFIG.snapshotSpreadsheetName);
  props.setProperty('SNAPSHOT_SPREADSHEET_ID', ss.getId());
  return ss;
}

function getTrackedSheets(ss) {
  const all = ss.getSheets();
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

function saveSnapshot(snapshotSs, sheetName, data) {
  let snap = snapshotSs.getSheetByName(sheetName);
  if (!snap) {
    snap = snapshotSs.insertSheet(sheetName);
    removeDefaultPlaceholderSheet(snapshotSs);
  } else {
    snap.clear();
  }
  const values = [data.headers, ...data.rows];
  if (values.length > 0 && values[0].length > 0) {
    snap.getRange(1, 1, values.length, values[0].length).setValues(values);
  }
}

/** Removes the blank "Sheet1" that SpreadsheetApp.create() adds by default. */
function removeDefaultPlaceholderSheet(ss) {
  const sheets = ss.getSheets();
  if (sheets.length <= 1) return;
  const placeholder = sheets.find(
    (s) => s.getName() === 'Sheet1' && s.getLastRow() === 0 && s.getLastColumn() === 0
  );
  if (placeholder) ss.deleteSheet(placeholder);
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
