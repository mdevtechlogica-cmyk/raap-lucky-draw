const SHEET_NAME = 'Leads';

// Stops text like "=SUM(...)" from being run as a formula
const safe = v => /^[=+\-@]/.test(String(v || '')) ? "'" + v : (v || '');

function cellText(v) {
  if (v instanceof Date) return v.toISOString();
  const s = v == null ? '' : String(v);
  return s.charAt(0) === "'" ? s.slice(1) : s;
}

function readLeads() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getDataRange().getValues().slice(1).filter(function (r) {
    return r.join('') !== '';
  }).map(function (r) {
    return {
      at: cellText(r[0]),
      token: cellText(r[1]),
      name: cellText(r[2]),
      phone: cellText(r[3]),
      org: cellText(r[4]),
      remarks: cellText(r[5]),
      prod: cellText(r[6])
    };
  });
}

function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return jsonOut(readLeads());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Registered at', 'Ticket', 'Name', 'Phone', 'Company', 'Remarks', 'Interested in']);
    }
    const rows = Math.max(sh.getLastRow() - 1, 1);
    const phones = sh.getRange(2, 4, rows, 1).getValues().flat().map(cellText);
    if (phones.indexOf(String(d.phone)) === -1) {
      sh.appendRow([new Date(d.at), d.token, safe(d.name), "'" + d.phone,
                    safe(d.org), safe(d.remarks), d.prod]);
    }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
