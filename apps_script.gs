// ============================================================
// CONFIG — set these once before deploying
// ============================================================
const SECRET_TOKEN = 'o3topuss'; // Must match index.html
const SHEET_NAME   = 'Sheet1';                  // Change if your tab has a different name
// ============================================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.token !== SECRET_TOKEN) {
      return respond({ success: false, error: 'Unauthorized' });
    }

    if (data.action === 'check')  return respond(checkOverlaps(data));
    if (data.action === 'submit') return respond(submitEntry(data));

    return respond({ success: false, error: 'Unknown action' });
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

function doGet(e) {
  if ((e.parameter || {}).token === SECRET_TOKEN) {
    return respond({ success: true, status: 'ok' });
  }
  return respond({ success: false, error: 'Unauthorized' });
}

// ── helpers ──────────────────────────────────────────────────

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function toSec(n, unit) {
  return unit === 'min' ? parseFloat(n) * 60 : parseFloat(n);
}

function hmsToSec(t) {
  const p = String(t).split(':').map(Number);
  return (p[0] || 0) * 3600 + (p[1] || 0) * 60 + (p[2] || 0);
}

// ── overlap check ─────────────────────────────────────────────

function checkOverlaps(data) {
  const tz   = Session.getScriptTimeZone();
  const rows = getSheet().getDataRange().getValues();

  const aStart = hmsToSec(data.time);
  const aEnd   = aStart + toSec(data.duration, data.unit);

  const conflicts = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] || !r[1]) continue;

    const rDate = r[0] instanceof Date
      ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd')
      : String(r[0]);

    if (rDate !== data.date) continue;

    const bStart = hmsToSec(String(r[1]));
    const bEnd   = bStart + toSec(r[2], String(r[3]));

    if (aStart < bEnd && aEnd > bStart) {
      conflicts.push({
        row:      i + 1,
        date:     rDate,
        time:     String(r[1]),
        duration: String(r[2]),
        unit:     String(r[3]),
        addedBy:  String(r[4] || '')
      });
    }
  }

  return { success: true, conflicts };
}

// ── submit entry ──────────────────────────────────────────────

function submitEntry(data) {
  const sheet = getSheet();
  sheet.appendRow([
    data.date,      // A: Date
    data.time,      // B: Time
    data.duration,  // C: Duration
    data.unit,      // D: Unit
    data.name,      // E: Added By
    ''              // F: Edit History
  ]);

  // Sort by date (col A) then time (col B), keeping header in place
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
      .sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
  }

  return { success: true };
}

// ── response helper ───────────────────────────────────────────

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
