/**
 * SailGP Timer — live schedule JSON API (reads spreadsheet directly, no pub CSV).
 *
 * SETUP
 * 1. Open your Google Sheet → Extensions → Apps Script
 * 2. Add this as ScheduleApi.gs (keep your onEdit in Code.gs)
 * 3. Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone)
 * 4. Copy the /exec URL into index.html → SCHEDULE_API_URL
 * 5. After ANY script change: Deploy → Manage deployments → Edit → New version → Deploy
 *
 * TAB SELECTION
 * - Default: SCHEDULE_SHEET_GID below (from tab URL …#gid=43449268)
 * - Per request: ?gid=43449268 or ?sheet=SUN%20NY on the /exec URL
 *
 * SHEET LAYOUT (TIMER SCHEDULE tab)
 *   Row 1:  "Date" | event date | "Venue" | venue name (e.g. New York)
 *           Also accepts B1 date / D1 or E1 venue if labels differ.
 *   Row 2+: A = time (HH:MM)   C = task / event name (B used if C empty)
 *   Rows whose name matches “Race 1…” go to races[] as well as land[].
 */

const SCHEDULE_SHEET_GID = 43449268;

const ROW_META = 1;
const COL_DATE = 2;   // B1 fallback
const COL_VENUE = 4;  // D1 fallback (next to "Venue" label in C1)

const ROW_SCHEDULE_START = 2;
const COL_TIME = 1;   // A
const COL_NAME = 3;   // C (B is often blank; falls back to B if C empty)

function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var sheet = getScheduleSheet_(params.gid, params.sheet);
    var meta = readMeta_(sheet);
    var all = readMarkerRows_(sheet, ROW_SCHEDULE_START, sheet.getLastRow(), COL_TIME, COL_NAME);
    var races = all.filter(function(row) {
      return /\brace\s*\d/i.test(String(row.name || ''));
    });
    var payload = {
      version: 2,
      updatedAt: new Date().toISOString(),
      sheetGid: sheet.getSheetId(),
      sheetName: sheet.getName(),
      eventDate: meta.eventDate,
      venue: meta.venue,
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

/**
 * Prefer labeled cells on row 1 ("Date" / "Venue"), then fall back to B1 / D1 / E1.
 */
function readMeta_(sheet) {
  var eventDate = '';
  var venue = '';
  for (var c = 1; c <= 12; c++) {
    var label = cellDisplay_(sheet, ROW_META, c).toLowerCase();
    if (!eventDate && (label === 'date' || label === 'event date')) {
      eventDate = cellDisplay_(sheet, ROW_META, c + 1);
    }
    if (!venue && (label === 'venue' || label === 'timezone' || label === 'time zone' || label === 'location' || label === 'tz')) {
      venue = cellDisplay_(sheet, ROW_META, c + 1);
    }
  }
  if (!eventDate) eventDate = cellDisplay_(sheet, ROW_META, COL_DATE);
  if (!venue) {
    venue = cellDisplay_(sheet, ROW_META, COL_VENUE);
    if (!venue) venue = cellDisplay_(sheet, ROW_META, 5); // E1
  }
  return { eventDate: eventDate, venue: venue };
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
