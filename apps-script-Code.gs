// Receives the treatment dictation from the web page and writes it to the Sheet.
// Setup: in the clinic Sheet (new Google account): Extensions > Apps Script, paste this, then
// Project settings > Script properties > add PASSWORD = <your password>.
// Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
// Weekly backup: run installBackupTrigger() once from the editor (it asks for Drive permission).

var TZ = 'Asia/Jerusalem';
var CHUNK = 45000; // Google Sheets allows 50,000 characters per cell
var T_HEAD = ['received_at', 'updated_at', 'session_id', 'patient', 'lang', 'findings_count',
  'raw_text', 'raw_text_2', 'raw_text_3', 'raw_text_4', 'timeline', 'timeline_2', 'treatment_id'];
var F_HEAD = ['received_at', 'session_id', 'patient', 'n', 'root', 'result', 'direction', 'direction_he', 'fixed', 'phrase', 'muscle', 'treatment_id'];
var BACKUP_FOLDER = 'גיבויי קליניקה';
var BACKUP_KEEP = 8;

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
  } else if (String(sh.getRange(1, head.length).getValue()) !== head[head.length - 1]) {
    // an older version of the sheet: add the missing header cell(s)
    sh.getRange(1, 1, 1, head.length).setValues([head]);
  }
  return sh;
}

function parts_(str, n) {
  var res = [];
  for (var i = 0; i < n; i++) res.push(str.substr(i * CHUNK, CHUNK));
  return res;
}

function fmtTime_(sec) {
  sec = Math.max(0, Number(sec) || 0);
  var m = Math.floor(sec / 60), s = sec % 60;
  return (m < 10 ? '0' : '') + m + ':' + (s < 10 ? '0' : '') + s;
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

    var now = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
    var patient = String(d.patient || '');
    var tid = String(d.treatment_id || '');
    var fs = d.findings || [];
    var partial = d.partial === true;
    var text = String(d.text || '');
    var tl = (d.segments || []).map(function (s) { return fmtTime_(s[0]) + ' ' + String(s[1]); }).join('\n');
    var tp = parts_(tl, 2), rp = parts_(text, 4);

    var t = sheet_('TRANSCRIPTS', T_HEAD);
    var last = t.getLastRow(), row = 0;
    if (last > 1) {
      var ids = t.getRange(2, 3, last - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === session) { row = i + 2; break; }
    }
    var nFind = fs.length;
    if (row) {
      // update the same session (autosave or final send)
      if (partial) nFind = Number(t.getRange(row, 6).getValue()) || 0;
      t.getRange(row, 2, 1, 1).setNumberFormat('@').setValue(now);
      t.getRange(row, 4, 1, 3).setNumberFormat('@').setValues([[patient, String(d.lang || ''), nFind]]);
      t.getRange(row, 7, 1, 6).setNumberFormat('@').setValues([rp.concat(tp)]);
      if (tid) t.getRange(row, 13, 1, 1).setNumberFormat('@').setValue(tid);
    } else {
      row = last + 1;
      var vals = [now, now, session, patient, String(d.lang || ''), nFind].concat(rp, tp, [tid]);
      t.getRange(row, 1, 1, vals.length).setNumberFormat('@').setValues([vals]);
    }

    if (!partial) {
      var f = sheet_('FINDINGS', F_HEAD);
      var fl = f.getLastRow();
      if (fl > 1) {
        var fids = f.getRange(2, 2, fl - 1, 1).getValues();
        for (var k = fids.length - 1; k >= 0; k--) if (String(fids[k][0]) === session) f.deleteRow(k + 2);
      }
      if (fs.length) {
        var rows = fs.map(function (x) {
          return [now, session, patient, x.n, String(x.root), String(x.result), String(x.dir || ''), String(x.dirHe || ''), String(x.fixed || ''), String(x.phrase || ''), String(x.muscle || ''), tid];
        });
        var fr = f.getLastRow() + 1;
        f.getRange(fr, 1, rows.length, F_HEAD.length).setNumberFormat('@');
        f.getRange(fr, 1, rows.length, F_HEAD.length).setValues(rows);
      }
    }
    return out_({ ok: true, chars: text.length });
  } catch (err) {
    return out_({ ok: false, error: 'שגיאת שרת' });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// ---- weekly backup of this whole spreadsheet into a Drive folder -------------
function weeklyBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var it = DriveApp.getFoldersByName(BACKUP_FOLDER);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER);
  file.makeCopy(ss.getName() + ' - גיבוי ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), folder);
  var list = [], files = folder.getFiles();
  while (files.hasNext()) { var f = files.next(); list.push({ f: f, t: f.getDateCreated().getTime() }); }
  list.sort(function (a, b) { return b.t - a.t; });
  for (var i = BACKUP_KEEP; i < list.length; i++) list[i].f.setTrashed(true);
}

// Run once by hand: creates the weekly trigger (Sunday around 03:00).
function installBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'weeklyBackup') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('weeklyBackup').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
  weeklyBackup(); // first backup right away, also asks for Drive permission
}
