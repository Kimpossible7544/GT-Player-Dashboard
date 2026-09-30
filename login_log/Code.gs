// Google Apps Script that records GT dashboard logins in the sheet it is bound
// to. Deploy as a web app (Execute as: Me, Who has access: Anyone) and paste
// the /exec URL into GT_LOG_URL in gt_log.js.
//   POST {action:'log', ...}  -> append a login row
//   GET ?action=admin&key=... -> JSON of recent logins (Login Log page)
//   GET ?action=csv&key=...   -> CSV of all logins (Excel Data -> From Web)
var ADMIN_PASSWORD = '2026GT';
var LOGIN_SHEET = 'Logins';
var LOGIN_HEADERS = ['Time', 'Player', 'Roster ID', 'Access', 'Page', 'Device'];
var MAX_LOGINS = 1000;

function loginSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(LOGIN_SHEET) || ss.insertSheet(LOGIN_SHEET);
  if (sh.getLastRow() === 0) sh.appendRow(LOGIN_HEADERS);
  return sh;
}

function rows_(limit) {
  var sh = loginSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var start = limit ? Math.max(2, last - limit + 1) : 2;
  return sh.getRange(start, 1, last - start + 1, LOGIN_HEADERS.length).getValues().reverse();
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.action === 'log') {
    loginSheet_().appendRow([
      new Date(),
      String(data.name || ''),
      String(data.id || ''),
      String(data.access || ''),
      String(data.page || ''),
      String(data.device || '')
    ]);
  }
  return ContentService.createTextOutput('ok');
}

function doGet(e) {
  var p = e.parameter;
  if (p.key !== ADMIN_PASSWORD) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid password.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (p.action === 'csv') {
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var lines = [LOGIN_HEADERS].concat(rows_(0).map(function (r) {
      var t = r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd HH:mm:ss') : r[0];
      return [t].concat(r.slice(1));
    })).map(function (r) {
      return r.map(function (v) { return '"' + String(v).replace(/"/g, '""') + '"'; }).join(',');
    });
    return ContentService.createTextOutput(lines.join('\r\n')).setMimeType(ContentService.MimeType.CSV);
  }

  var logins = rows_(MAX_LOGINS).map(function (r) {
    return {
      time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: String(r[1]), id: String(r[2]), access: String(r[3]),
      page: String(r[4]), device: String(r[5])
    };
  });
  return ContentService.createTextOutput(JSON.stringify({ logins: logins }))
    .setMimeType(ContentService.MimeType.JSON);
}
