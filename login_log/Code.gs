// Google Apps Script backing the GT login log and master ID list. Bind it to a
// Google Sheet, deploy as a web app (Execute as: Me, Who has access: Anyone),
// and paste the /exec URL into GT_LOG_URL in gt_log.js.
var ADMIN_PASSWORD = '2026GT';
var LOGIN_SHEET = 'Logins';
var MASTER_SHEET = 'Master IDs';
var LOGIN_HEADERS = ['Time', 'Player', 'Roster ID', 'Access', 'Page', 'Device'];
var MAX_LOGINS = 1000;

function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function masters_() {
  var sh = sheet_(MASTER_SHEET, ['Roster ID', 'Player']);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) { return { id: String(r[0]).trim(), name: String(r[1] || '') }; });
}

function logins_() {
  var sh = sheet_(LOGIN_SHEET, LOGIN_HEADERS);
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  var start = Math.max(2, sh.getLastRow() - MAX_LOGINS + 1);
  return sh.getRange(start, 1, sh.getLastRow() - start + 1, LOGIN_HEADERS.length).getValues()
    .reverse()
    .map(function (r) {
      return {
        time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        name: String(r[1]), id: String(r[2]), access: String(r[3]),
        page: String(r[4]), device: String(r[5])
      };
    });
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.action === 'log') {
    sheet_(LOGIN_SHEET, LOGIN_HEADERS).appendRow([
      new Date(),
      String(data.name || ''),
      String(data.id || ''),
      String(data.access || ''),
      String(data.page || ''),
      String(data.device || '')
    ]);
  }
  return json_({ ok: true });
}

function doGet(e) {
  var p = e.parameter;
  if (p.action === 'masters') {
    return json_({ masters: masters_().map(function (m) { return m.id; }) });
  }
  if (p.key !== ADMIN_PASSWORD) return json_({ error: 'Invalid password.' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var id = String(p.id || '').trim();
    var sh = sheet_(MASTER_SHEET, ['Roster ID', 'Player']);
    if (p.action === 'grant' && id) {
      var exists = masters_().some(function (m) { return m.id === id; });
      if (!exists) sh.appendRow([id, String(p.name || '')]);
    } else if (p.action === 'revoke' && id) {
      var ids = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
      for (var i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]).trim() === id) sh.deleteRow(i + 2);
      }
    }
    return json_({ logins: logins_(), masters: masters_() });
  } finally {
    lock.releaseLock();
  }
}
