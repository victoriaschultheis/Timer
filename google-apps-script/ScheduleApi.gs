/**
 * SailGP Timer — live schedule JSON API (reads spreadsheet directly, no pub CSV).
 *
 * SETUP
 * 1. Open your Google Sheet → Extensions → Apps Script
 * 2. Paste this file (replace default Code.gs contents)
 * 3. Set SCHEDULE_SHEET_GID below to match your tab (from the sheet URL: gid=…)
 * 4. Deploy → New deployment → Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web app URL (ends in /exec) into index-live.html → SCHEDULE_API_URL
 *
 * Sheet layout (same as index.html):
 *   Q2  event date    V2  venue timezone
 *   Q3+ times, R3+ names  (land schedule)
 *   A1–C23 on-water race markers (optional, for race countdown / up next)
 */

const SCHEDULE_SHEET_GID = 2126807151;

const ROW_META = 2;
const COL_DATE = 17;   // Q
const COL_VENUE = 22;  // V

const ROW_LAND_START = 3;
const COL_LAND_TIME = 17;  // Q
const COL_LAND_NAME = 18;  // R

const ROW_RACE_START = 1;
const ROW_RACE_END = 23;
const COL_RACE_TIME = 1;   // A
const COL_RACE_NAME = 3;   // C

function doGet() {
  try {
    const sheet = getScheduleSheet_();
    const payload = {
      version: 1,
      updatedAt: new Date().toISOString(),
      sheetGid: sheet.getSheetId(),
      eventDate: cellDisplay_(sheet, ROW_META, COL_DATE),
      venue: cellDisplay_(sheet, ROW_META, COL_VENUE),
      land: readMarkerRows_(sheet, ROW_LAND_START, sheet.getLastRow(), COL_LAND_TIME, COL_LAND_NAME),
      races: readMarkerRows_(sheet, ROW_RACE_START, ROW_RACE_END, COL_RACE_TIME, COL_RACE_NAME),
    };
    return jsonResponse_(payload);
  } catch (err) {
    return jsonResponse_({
      version: 1,
      error: String(err && err.message ? err.message : err),
      updatedAt: new Date().toISOString(),
    });
  }
}

function getScheduleSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SCHEDULE_SHEET_GID) return sheets[i];
  }
  throw new Error('Schedule tab not found (gid ' + SCHEDULE_SHEET_GID + '). Update SCHEDULE_SHEET_GID.');
}

function cellDisplay_(sheet, row, col) {
  return String(sheet.getRange(row, col).getDisplayValue() || '').trim();
}

function readMarkerRows_(sheet, startRow, endRow, timeCol, nameCol) {
  const out = [];
  const last = Math.min(endRow, sheet.getLastRow());
  for (let r = startRow; r <= last; r++) {
    const timeStr = cellDisplay_(sheet, r, timeCol);
    const label = timeStr.toLowerCase();
    if (label === 'utc time difference' || label === 'utc time') break;
    if (!timeStr) continue;
    out.push({
      time: timeStr,
      name: cellDisplay_(sheet, r, nameCol),
    });
  }
  return out;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
