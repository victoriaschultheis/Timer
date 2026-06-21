/**
 * SailGP Timer — live schedule JSON API (reads spreadsheet directly, no pub CSV).
 *
 * SETUP
 * 1. Open your Google Sheet → Extensions → Apps Script
 * 2. Add this as ScheduleApi.gs (keep your onEdit in Code.gs)
 * 3. Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone)
 * 4. Copy the /exec URL into index-live.html → SCHEDULE_API_URL
 *
 * TAB SELECTION (pick one)
 * - Default: SCHEDULE_SHEET_GID below (from tab URL …#gid=43449268)
 * - Per request: ?gid=43449268 or ?sheet=SUN%20NY on the /exec URL
 * - index-live.html can pass SCHEDULE_SHEET_GID on each poll
 *
 * SHEET LAYOUT (per-day tabs, e.g. SUN NY, SUN HFX)
 *   Row 1:  B1 event date (e.g. 21/06/2026)   E1 venue (e.g. New York)
 *   Row 2+:  A = time (HH:MM)                 C = task / event name
 *   Rows whose name matches “Race 1…” go to races[] as well as land[].
 */

const SCHEDULE_SHEET_GID = 43449268;

const ROW_META = 1;
const COL_DATE = 2;   // B1
const COL_VENUE = 5;  // E1

const ROW_SCHEDULE_START = 2;
const COL_TIME = 1;   // A
const COL_NAME = 3;   // C (B is often blank; falls back to B if C empty)

function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var sheet = getScheduleSheet_(params.gid, params.sheet);
    var all = readMarkerRows_(sheet, ROW_SCHEDULE_START, sheet.getLastRow(), COL_TIME, COL_NAME);
    var races = all.filter(function(row) {
      return /\brace\s*\d/i.test(String(row.name || ''));
    });
    var payload = {
      version: 2,
      updatedAt: new Date().toISOString(),
      sheetGid: sheet.getSheetId(),
      sheetName: sheet.getName(),
      eventDate: cellDisplay_(sheet, ROW_META, COL_DATE),
      venue: cellDisplay_(sheet, ROW_META, COL_VENUE),
      land: all,
      races: races,
    };
    return jsonResponse_(payload);
  } catch (err) {
    return jsonResponse_({
      version: 2,
      error: String(err && err.message ? err.message : err),
      updatedAt: new Date().toISOString(),
    });
  }
}

function getScheduleSheet_(gidParam, nameParam) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  if (nameParam) {
    var byName = ss.getSheetByName(String(nameParam).trim());
    if (byName) return byName;
    throw new Error('Schedule tab not found (name "' + nameParam + '").');
  }

  var gid = gidParam ? parseInt(String(gidParam), 10) : SCHEDULE_SHEET_GID;
  if (isNaN(gid)) gid = SCHEDULE_SHEET_GID;

  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  throw new Error('Schedule tab not found (gid ' + gid + '). Update SCHEDULE_SHEET_GID or pass ?gid= / ?sheet=.');
}

function cellDisplay_(sheet, row, col) {
  return String(sheet.getRange(row, col).getDisplayValue() || '').trim();
}

function readMarkerRows_(sheet, startRow, endRow, timeCol, nameCol) {
  var out = [];
  var last = Math.min(endRow, sheet.getLastRow());
  for (var r = startRow; r <= last; r++) {
    var timeStr = cellDisplay_(sheet, r, timeCol);
    var label = timeStr.toLowerCase();
    if (label === 'utc time difference' || label === 'utc time') break;
    if (!timeStr) continue;
    var name = cellDisplay_(sheet, r, nameCol);
    if (!name && nameCol !== timeCol + 1) {
      name = cellDisplay_(sheet, r, timeCol + 1);
    }
    out.push({ time: timeStr, name: name });
  }
  return out;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
