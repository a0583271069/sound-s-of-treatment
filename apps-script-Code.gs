// Receives the treatment dictation from the web page and writes it to the Sheet.
// Setup: in the clinic Sheet (new Google account): Extensions > Apps Script, paste this, then
// Project settings > Script properties > add PASSWORD = <your password>.
// Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
// Weekly backup: run installBackupTrigger() once from the editor (it asks for Drive permission).

var TZ = 'Asia/Jerusalem';
var CHUNK = 45000; // Google Sheets allows 50,000 characters per cell
var T_HEAD = ['received_at', 'updated_at', 'session_id', 'patient', 'lang', 'findings_count',
  'raw_text', 'raw_text_2', 'raw_text_3', 'raw_text_4', 'timeline', 'timeline_2', 'treatment_id'];
var F_HEAD = ['received_at', 'session_id', 'patient', 'n', 'root', 'result', 'direction', 'direction_he', 'fixed', 'phrase', 'muscle', 'treatment_id', 'side', 'muscle_id', 'applied'];
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
          return [now, session, patient, x.n, String(x.root), String(x.result), String(x.dir || ''), String(x.dirHe || ''), String(x.fixed || ''), String(x.phrase || ''), String(x.muscle || ''), tid, String(x.side || ''), String(x.muscle_id || ''), ''];
        });
        var fr = f.getLastRow() + 1;
        f.getRange(fr, 1, rows.length, F_HEAD.length).setNumberFormat('@');
        f.getRange(fr, 1, rows.length, F_HEAD.length).setValues(rows);
      }
    }
    var ap = { done: 0, waiting: 0 };
    if (!partial && tid) {
      try { ap = applyPending_(tid); } catch (e3) { ap = { done: 0, waiting: -1 }; }
      try { ensureApplyTrigger_(); } catch (e4) {}
    }
    return out_({ ok: true, chars: text.length, done: ap.done, waiting: ap.waiting });
  } catch (err) {
    return out_({ ok: false, error: 'שגיאת שרת' });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// ---- write the dictated findings into the rows of the treatment -----------------
// The app creates one NEURO_RESULTS row per nerve test and one MUSCLE_RESULTS row per muscle when a
// treatment is saved. A finding is written into those rows only after ALL of them have reached the Sheet
// (the tablet uploads them one by one). Until then the finding stays in FINDINGS with an empty "applied"
// cell, and applyPending (every 5 minutes) tries again.
var RES_OK = 'עובד', RES_BAD = 'לא עובד', FIX_YES = 'תוקן', FIX_NO = 'לא תוקן';

function stamp_() { return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss'); }

function headIdx_(sh) {
  var last = Math.max(1, sh.getLastColumn()), h = sh.getRange(1, 1, 1, last).getValues()[0], m = {};
  for (var i = 0; i < h.length; i++) m[String(h[i]).trim()] = i + 1;
  return m;
}

function countFilled_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return 0;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues(), n = 0;
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() !== '') n++;
  return n;
}

// rows of one treatment in a results sheet: {keyValue: {row, notes}}
function rowsOf_(name, keyCol, tid) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  var res = { sh: sh, idx: null, map: {}, count: 0 };
  if (!sh || sh.getLastRow() < 2) return res;
  var idx = headIdx_(sh);
  res.idx = idx;
  if (!idx.Treatment_ID || !idx[keyCol]) return res;
  var width = sh.getLastColumn();
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][idx.Treatment_ID - 1]).trim() !== tid) continue;
    res.map[String(v[i][idx[keyCol] - 1]).trim()] = { row: i + 2, notes: idx.notes ? String(v[i][idx.notes - 1] || '') : '' };
    res.count++;
  }
  return res;
}

function neuroTargets_(root) { return root === 'L5' ? ['L5-EHL', 'L5-EHB'] : [root]; }

// plan: {rowNo: {right:{res,fix}, left:{res,fix}, notes:[...]}}
function plan_(plan, rowNo, side, res, fix, explicit) {
  var p = plan[rowNo] || (plan[rowNo] = { notes: [] });
  if (!explicit && p[side]) return; // a vague "ok" never overrides something said about that side
  p[side] = { res: res, fix: fix };
}

function writePlan_(r, plan) {
  if (!r.sh || !r.idx) return;
  var idx = r.idx, sh = r.sh;
  Object.keys(plan).forEach(function (k) {
    var row = Number(k), p = plan[k];
    ['right', 'left'].forEach(function (sd) {
      if (!p[sd]) return;
      if (idx['res_' + sd]) sh.getRange(row, idx['res_' + sd]).setValue(p[sd].res);
      if (idx['fix_' + sd]) sh.getRange(row, idx['fix_' + sd]).setValue(p[sd].fix);
    });
    if (p.notes.length && idx.notes) {
      var cur = '';
      Object.keys(r.map).forEach(function (key) { if (r.map[key].row === row) cur = r.map[key].notes; });
      sh.getRange(row, idx.notes).setValue((cur ? cur + '\n' : '') + p.notes.join('\n'));
    }
  });
}

// returns {done, waiting}: findings written now / findings of that treatment still waiting for their rows
function applyPending_(onlyTid) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = sheet_('FINDINGS', F_HEAD);
  var fl = f.getLastRow(), none = { done: 0, waiting: 0 };
  if (fl < 2) return none;
  var fv = f.getRange(2, 1, fl - 1, F_HEAD.length).getValues();
  var byTid = {};
  for (var i = 0; i < fv.length; i++) {
    var tid = String(fv[i][11]).trim();
    if (!tid || String(fv[i][14]).trim() !== '') continue;
    if (onlyTid && tid !== onlyTid) continue;
    (byTid[tid] = byTid[tid] || []).push(i);
  }
  var needN = countFilled_('NEURO_TESTS'), needM = countFilled_('MUSCLES');
  var total = { done: 0, waiting: 0 };
  Object.keys(byTid).forEach(function (tid) {
    var nr = rowsOf_('NEURO_RESULTS', 'test_id', tid), mr = rowsOf_('MUSCLE_RESULTS', 'muscle_id', tid);
    var nReady = needN > 0 && nr.count >= needN, mReady = needM > 0 && mr.count >= needM;
    var nPlan = {}, mPlan = {}, doneRows = [], waiting = 0;
    byTid[tid].forEach(function (i) {
      var x = fv[i];
      var result = String(x[5]), fixed = String(x[8]), dir = String(x[6]), phrase = String(x[9]);
      var side = String(x[12]).trim(), mid = String(x[13]).trim(), mname = String(x[10]).trim(), root = String(x[4]).trim();
      if (!side) side = dir === 'bendR' ? 'right' : dir === 'bendL' ? 'left' : ''; // findings from the older page version
      if (!mid && mname === 'פיריפורמיס') mid = 'M23';
      var isM = !!mid;
      if (isM ? !mReady : !nReady) { waiting++; return; }
      var res = result === 'ok' ? RES_OK : RES_BAD;
      var fix = res === RES_BAD ? (fixed === 'fixed' ? FIX_YES : fixed === 'notfixed' ? FIX_NO : '') : '';
      var keys = isM ? [mid] : neuroTargets_(root), rows = isM ? mr.map : nr.map, pl = isM ? mPlan : nPlan;
      keys.forEach(function (k) {
        var rr = rows[k];
        if (!rr) return;
        if (side) plan_(pl, rr.row, side, res, fix, true);
        else if (res === RES_OK) { plan_(pl, rr.row, 'right', res, '', false); plan_(pl, rr.row, 'left', res, '', false); }
        else (pl[rr.row] || (pl[rr.row] = { notes: [] })).notes.push('⚠ הוכתב בלי צד: ' + phrase);
      });
      doneRows.push(i);
    });
    if (nReady || mReady) {
      writePlan_(nr, nPlan);
      writePlan_(mr, mPlan);
    }
    var ts = stamp_();
    doneRows.forEach(function (i) { f.getRange(i + 2, 15).setNumberFormat('@').setValue(ts); });
    total.done += doneRows.length; total.waiting += waiting;
  });
  return total;
}

// every 5 minutes: finish findings whose rows had not reached the Sheet yet
function applyPending() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return;
  try { applyPending_(''); } finally { try { lock.releaseLock(); } catch (e) {} }
}

function ensureApplyTrigger_() {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'applyPending'; });
  if (!has) ScriptApp.newTrigger('applyPending').timeBased().everyMinutes(5).create();
}

// Run once by hand if you want the 5-minute trigger without waiting for the first dictation.
function installApplyTrigger() { ensureApplyTrigger_(); applyPending_(''); }

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
