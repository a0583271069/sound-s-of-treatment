// Receives the finished treatment dictation from the web page and writes it to the Sheet.
// Setup: Extensions > Apps Script (in the clinic Sheet), paste this, then
// Project settings > Script properties > add PASSWORD = <your password>.
// Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.

var TZ = 'Asia/Jerusalem';
var T_HEAD = ['received_at', 'session_id', 'patient', 'lang', 'raw_text', 'findings_count'];
var F_HEAD = ['received_at', 'session_id', 'patient', 'n', 'root', 'result', 'direction', 'direction_he', 'fixed'];

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doGet() { return out_({ ok: true, service: 'clinic-dictation' }); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var d = JSON.parse(e.postData.contents);
    var pw = PropertiesService.getScriptProperties().getProperty('PASSWORD');
    if (!pw || String(d.pw) !== pw) return out_({ ok: false, error: 'סיסמה שגויה' });
    var session = String(d.session || '');
    if (!session) return out_({ ok: false, error: 'חסר מזהה' });

    var t = sheet_('TRANSCRIPTS', T_HEAD);
    var last = t.getLastRow();
    if (last > 1) {
      var ids = t.getRange(2, 2, last - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === session) return out_({ ok: true, duplicate: true });
    }

    var now = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
    var patient = String(d.patient || '');
    var fs = d.findings || [];

    var r = t.getLastRow() + 1;
    t.getRange(r, 1, 1, 6).setNumberFormat('@');
    t.getRange(r, 1, 1, 6).setValues([[now, session, patient, String(d.lang || ''), String(d.text || ''), fs.length]]);

    if (fs.length) {
      var f = sheet_('FINDINGS', F_HEAD);
      var rows = fs.map(function (x) {
        return [now, session, patient, x.n, String(x.root), String(x.result), String(x.dir || ''), String(x.dirHe || ''), String(x.fixed || '')];
      });
      var fr = f.getLastRow() + 1;
      f.getRange(fr, 1, rows.length, F_HEAD.length).setNumberFormat('@');
      f.getRange(fr, 1, rows.length, F_HEAD.length).setValues(rows);
    }
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: 'שגיאת שרת' });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
