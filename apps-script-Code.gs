// Receives the treatment dictation from the web page and writes it to the Sheet.
// Setup: in the clinic Sheet (new Google account): Extensions > Apps Script, paste this, then
// Project settings > Script properties > add PASSWORD = <your password>.
// Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
// Weekly backup: run installBackupTrigger() once from the editor (it asks for Drive permission).
// Daily backup: installDailyBackup(); nightly integrity check: installDoctor() (run each once from the editor).

var TZ = 'Asia/Jerusalem';
var CHUNK = 45000; // Google Sheets allows 50,000 characters per cell
var T_HEAD = ['received_at', 'updated_at', 'session_id', 'patient', 'lang', 'findings_count',
  'raw_text', 'raw_text_2', 'raw_text_3', 'raw_text_4', 'timeline', 'timeline_2', 'treatment_id', 'recognized', 'ignored'];
var F_HEAD = ['received_at', 'session_id', 'patient', 'n', 'root', 'result', 'direction', 'direction_he', 'fixed', 'phrase', 'muscle', 'treatment_id', 'side', 'muscle_id', 'applied', 'how'];
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

// ---- photo page: list of treatments and the photos (password protected, read only) --------------------
function listData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('TREATMENTS');
  if (!sh || sh.getLastRow() < 2) return { ok: true, treatments: [] };
  var v = sh.getDataRange().getDisplayValues(), h = v[0], res = [];
  var cols = ['Treatment_ID', 'PATIENT_ID', 'TREATMENT_NUMBER', 'DATE', 'PHOTO_FRONT', 'PHOTO_LEFT', 'PHOTO_BACK', 'PHOTO_RIGHT', 'AFTER_FRONT', 'AFTER_LEFT', 'AFTER_BACK', 'AFTER_RIGHT'];
  var ix = cols.map(function (c) { return h.indexOf(c); });
  for (var i = 1; i < v.length; i++) {
    var o = {};
    cols.forEach(function (c, k) { o[c] = ix[k] >= 0 ? String(v[i][ix[k]]) : ''; });
    if (o.Treatment_ID) res.push(o);
  }
  var names = {};
  var ps = ss.getSheetByName('PATIENTS');
  if (ps && ps.getLastRow() > 1) {
    var pv = ps.getDataRange().getDisplayValues(), ph = pv[0];
    var ci = ph.indexOf('PATIENT_ID'), cn = -1;
    ph.forEach(function (x, k) { if (cn < 0 && /name|שם/i.test(x)) cn = k; });
    if (ci >= 0 && cn >= 0) for (var j = 1; j < pv.length; j++) names[String(pv[j][ci])] = String(pv[j][cn]);
  }
  return { ok: true, treatments: res, names: names };
}

function photo_(file) {
  var base = file.split('/').pop();
  if (!base) return { ok: false, error: 'no file' };
  var it = DriveApp.getFilesByName(base);
  if (!it.hasNext()) return { ok: false, error: 'not found' };
  var b = it.next().getBlob();
  return { ok: true, mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
}

// ---- complaint / change + pain scale: two extra columns at the end of TREATMENTS (added automatically) --------------
var T_EXTRA = ['COMPLAINT_CHANGE', 'PAIN_BEFORE'];
function ensureTreatmentCols_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TREATMENTS');
  if (!sh) return;
  var last = Math.max(1, sh.getLastColumn()), h = sh.getRange(1, 1, 1, last).getValues()[0].map(function (x) { return String(x).trim(); });
  T_EXTRA.forEach(function (c) {
    if (h.indexOf(c) < 0) { sh.getRange(1, h.length + 1).setValue(c); h.push(c); }
  });
}

// ---- the previous treatment of the same patient: what is still open (red) and what was fixed (purple) ------------
function prevData_(tid) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('TREATMENTS');
  if (!sh || sh.getLastRow() < 2 || !tid) return { ok: true, items: [] };
  var v = sh.getDataRange().getDisplayValues(), h = v[0].map(function (x) { return String(x).trim(); });
  var ci = h.indexOf('Treatment_ID'), pi = h.indexOf('PATIENT_ID'), ni = h.indexOf('TREATMENT_NUMBER');
  var cc = h.indexOf('COMPLAINT_CHANGE'), pb = h.indexOf('PAIN_BEFORE');
  var cur = null, rows = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i];
    if (String(r[ci]).trim() === tid) cur = r;
    rows.push(r);
  }
  if (!cur) return { ok: true, items: [] };
  var pat = String(cur[pi]).trim(), num = Number(cur[ni]) || 0, best = null;
  rows.forEach(function (r) {
    if (String(r[pi]).trim() !== pat || String(r[ci]).trim() === tid) return;
    var n = Number(r[ni]) || 0;
    if (n < num && (!best || n > (Number(best[ni]) || 0))) best = r;
  });
  if (!best) return { ok: true, items: [], number: 0 };
  var ptid = String(best[ci]).trim(), items = [], CORR = correctionNames_();
  function scan(name, keyCol, kind, hasHow) {
    var sh2 = ss.getSheetByName(name);
    if (!sh2 || sh2.getLastRow() < 2) return;
    var d = sh2.getDataRange().getValues(), ix = headIdx_(sh2);
    for (var j = 1; j < d.length; j++) {
      if (String(d[j][ix.Treatment_ID - 1]).trim() !== ptid) continue;
      ['right', 'left'].forEach(function (sd) {
        var res = String(d[j][ix['res_' + sd] - 1] || ''), fix = String(d[j][ix['fix_' + sd] - 1] || ''), how = hasHow ? String(d[j][ix['how_' + sd] - 1] || '') : '';
        if (res !== RES_BAD) return;
        items.push({ kind: kind, id: String(d[j][ix[keyCol] - 1]).trim(), side: sd, state: fix === FIX_YES ? 'fixed' : 'open', how: how });
      });
    }
  }
  scan('MUSCLE_RESULTS', 'muscle_id', 'M', true);
  scan('NEURO_RESULTS', 'test_id', 'N', false);
  return { ok: true, number: Number(best[ni]) || 0, complaint: cc >= 0 ? String(best[cc]) : '', pain: pb >= 0 ? String(best[pb]) : '', items: items };
}

function doGet() { return out_({ ok: true, service: 'clinic-dictation' }); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var d = JSON.parse(e.postData.contents);
    var pw = PropertiesService.getScriptProperties().getProperty('PASSWORD');
    if (!pw || String(d.pw) !== pw) return out_({ ok: false, error: 'סיסמה שגויה' });
    if (d.action === 'list') return out_(listData_());
    if (d.action === 'prev') return out_(prevData_(String(d.treatment_id || '').trim()));
    if (d.action === 'photo') return out_(photo_(String(d.file || '')));
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
      if (!partial) t.getRange(row, 14, 1, 2).setNumberFormat('@').setValues([[String(d.recognized || '').substr(0, CHUNK), String(d.ignored || '').substr(0, CHUNK)]]);
    } else {
      row = last + 1;
      var vals = [now, now, session, patient, String(d.lang || ''), nFind].concat(rp, tp, [tid, String(d.recognized || '').substr(0, CHUNK), String(d.ignored || '').substr(0, CHUNK)]);
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
          return [now, session, patient, x.n, String(x.root), String(x.result), String(x.dir || ''), String(x.dirHe || ''), String(x.fixed || ''), String(x.phrase || ''), String(x.muscle || ''), tid, String(x.side || ''), String(x.muscle_id || ''), '', String(x.how || '')];
        });
        var fr = f.getLastRow() + 1;
        f.getRange(fr, 1, rows.length, F_HEAD.length).setNumberFormat('@');
        f.getRange(fr, 1, rows.length, F_HEAD.length).setValues(rows);
      }
    }
    var ap = { done: 0, waiting: 0 };
    if (!partial && tid) {
      try { generateRows_(); } catch (e5) {}
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
// Every treatment has one NEURO_RESULTS row per nerve test and one MUSCLE_RESULTS row per muscle (created by
// generateRows_ below). A finding is written into those rows only after ALL of them exist in the Sheet.
// Until then the finding stays in FINDINGS with an empty "applied" cell, and tick (every minute) tries again.
var RES_OK = 'עובד', RES_BAD = 'לא עובד', FIX_YES = 'תוקן', FIX_NO = 'לא תוקן', RES_SKIP = 'לא נבדק';

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
function plan_(plan, rowNo, side, res, fix, explicit, how) {
  var p = plan[rowNo] || (plan[rowNo] = { notes: [] });
  if (!explicit && p[side]) return; // a vague "ok" never overrides something said about that side
  p[side] = { res: res, fix: fix, how: how || '' };
}

// what the muscle was fixed with: R1..R7 -> the names in the CORRECTIONS sheet (the same text the app offers)
function correctionNames_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CORRECTIONS'), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  var idx = headIdx_(sh);
  if (!idx.correction_id || !idx.name_he) return m;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (var i = 0; i < v.length; i++) m[String(v[i][idx.correction_id - 1]).trim()] = String(v[i][idx.name_he - 1]).trim();
  return m;
}

function writePlan_(r, plan) {
  if (!r.sh || !r.idx) return;
  var idx = r.idx, sh = r.sh;
  Object.keys(plan).forEach(function (k) {
    var row = Number(k), p = plan[k];
    ['right', 'left'].forEach(function (sd) {
      if (!p[sd]) return;
      if (p[sd].res === RES_SKIP && idx['res_' + sd] && String(sh.getRange(row, idx['res_' + sd]).getValue()).trim() !== '') return; // "not examined" never overwrites a result already there
      if (idx['res_' + sd]) sh.getRange(row, idx['res_' + sd]).setValue(p[sd].res);
      if (idx['fix_' + sd]) sh.getRange(row, idx['fix_' + sd]).setValue(p[sd].fix);
      if (idx['how_' + sd]) sh.getRange(row, idx['how_' + sd]).setValue(p[sd].how);
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
  var needN = countFilled_('NEURO_TESTS'), needM = countFilled_('MUSCLES'), CORR = correctionNames_();
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
      if (!mid && mname === 'פיריפורמיס') mid = '33';
      var isM = !!mid;
      if (isM ? !mReady : !nReady) { waiting++; return; }
      var res = result === 'ok' ? RES_OK : result === 'skip' ? RES_SKIP : RES_BAD;
      var fix = res === RES_BAD ? (fixed === 'fixed' ? FIX_YES : fixed === 'notfixed' ? FIX_NO : '') : '';
      var how = '';
      // reasons (R1..R7) go with a fixed muscle that did not work; "שיחררתי" (R9) goes with a muscle that works too strongly
      if (isM && (fix === FIX_YES || res === RES_OK)) {
        how = String(x[15] || '').split(',').map(function (id) { return CORR[id.trim()] || ''; }).filter(String).join(' , ');
      }
      var keys = isM ? [mid] : neuroTargets_(root), rows = isM ? mr.map : nr.map, pl = isM ? mPlan : nPlan;
      keys.forEach(function (k) {
        var rr = rows[k];
        if (!rr) return;
        if (side) plan_(pl, rr.row, side, res, fix, true, how);
        else if (res === RES_OK || res === RES_SKIP) { plan_(pl, rr.row, 'right', res, '', false); plan_(pl, rr.row, 'left', res, '', false); }
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

// ---- the empty result rows of a new treatment, created here instead of on the tablet --------------
// The tablet used to upload them one by one (66 rows, several minutes). Now this script adds all of them
// in one go, within a minute of the treatment reaching the Sheet. A treatment that already has rows in
// a results sheet is never touched, so nothing is doubled.
function genId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 8); }

function keysOf_(sheetName, keyCol) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName), res = [];
  if (!sh || sh.getLastRow() < 2) return res;
  var idx = headIdx_(sh), c = idx[keyCol] || 1;
  var v = sh.getRange(2, c, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) { var k = String(v[i][0]).trim(); if (k) res.push(k); }
  return res;
}

function addRows_(resultSheet, keyCol, keys, tid) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(resultSheet);
  if (!sh || !keys.length) return 0;
  var idx = headIdx_(sh);
  if (!idx.result_id || !idx.Treatment_ID || !idx[keyCol]) return 0;
  var width = sh.getLastColumn(), now = new Date(), rows = [];
  keys.forEach(function (k) {
    var r = new Array(width);
    for (var i = 0; i < width; i++) r[i] = '';
    r[idx.result_id - 1] = genId_();
    r[idx.Treatment_ID - 1] = tid;
    r[idx[keyCol] - 1] = k;
    if (idx.created_at) r[idx.created_at - 1] = now;
    rows.push(r);
  });
  var start = sh.getLastRow() + 1;
  sh.getRange(start, idx[keyCol], rows.length, 1).setNumberFormat('@'); // the muscle ids are plain numbers (11-65): keep them as text
  sh.getRange(start, 1, rows.length, width).setValues(rows);
  if (idx.created_at) sh.getRange(start, idx.created_at, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  return rows.length;
}

// returns how many treatments got their rows now
function generateRows_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tsh = ss.getSheetByName('TREATMENTS');
  if (!tsh || tsh.getLastRow() < 2) return 0;
  var tids = keysOf_('TREATMENTS', 'Treatment_ID');
  if (!tids.length) return 0;
  // which test / muscle each treatment already has, so a later addition to
  // NEURO_TESTS or MUSCLES is filled in afterwards instead of being missed
  var nHave = {}, mHave = {};
  [['NEURO_RESULTS', 'test_id', nHave], ['MUSCLE_RESULTS', 'muscle_id', mHave]].forEach(function (p) {
    var sh = ss.getSheetByName(p[0]);
    if (!sh || sh.getLastRow() < 2) return;
    var idx = headIdx_(sh);
    if (!idx.Treatment_ID || !idx[p[1]]) return;
    var n = sh.getLastRow() - 1;
    var tv = sh.getRange(2, idx.Treatment_ID, n, 1).getValues();
    var kv = sh.getRange(2, idx[p[1]], n, 1).getValues();
    for (var i = 0; i < n; i++) {
      var t = String(tv[i][0]).trim();
      if (t) (p[2][t] = p[2][t] || {})[String(kv[i][0]).trim()] = true;
    }
  });
  var nKeys = keysOf_('NEURO_TESTS', 'test_id'), mKeys = keysOf_('MUSCLES', 'muscle_id'), made = 0;
  tids.forEach(function (tid) {
    var did = false;
    var missN = nKeys.filter(function (k) { return !(nHave[tid] && nHave[tid][k]); });
    var missM = mKeys.filter(function (k) { return !(mHave[tid] && mHave[tid][k]); });
    if (missN.length && addRows_('NEURO_RESULTS', 'test_id', missN, tid)) did = true;
    if (missM.length && addRows_('MUSCLE_RESULTS', 'muscle_id', missM, tid)) did = true;
    if (did) made++;
  });
  return made;
}

// ---- the patient of a form answer, even if the form trigger failed ----------
// The Google Form always writes its row. onFormSubmit then creates the patient in
// PATIENTS - but if that trigger fails (a bad character in the code, a Google
// server error), the answer is in the Sheet and the patient is nowhere, and
// nothing says so. So every minute we also check it from here: any form answer
// whose id is missing from PATIENTS gets its row now. A patient cannot be lost.
function formSheet_() {
  var out = null;
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (sh) {
    if (out || sh.getLastColumn() < 2) return;
    var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      .map(function (x) { return String(x).trim(); });
    if (h.indexOf('ת.ז') >= 0 && h.indexOf('שם מלא') >= 0) out = sh;
  });
  return out;
}

// returns how many patients were created now
function backfillPatients_() {
  var fsh = formSheet_();
  if (!fsh || fsh.getLastRow() < 2) return 0;
  var psh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PATIENTS');
  if (!psh || psh.getLastRow() < 1) return 0;

  var fh = fsh.getRange(1, 1, 1, fsh.getLastColumn()).getValues()[0]
    .map(function (x) { return String(x).trim(); });
  var fId = fh.indexOf('ת.ז'), fName = fh.indexOf('שם מלא'), fPhone = fh.indexOf('טלפון');
  if (fId < 0) return 0;

  var pd = psh.getDataRange().getValues();
  var ph = pd[0].map(function (x) { return String(x).replace(/\s/g, ''); });
  var cId = ph.indexOf('PATIENT_ID'), cName = ph.indexOf('PATIENT_NAME'),
      cPhone = ph.indexOf('PHONE'), cFile = ph.indexOf('HEALTH_FILE'),
      cCreated = ph.indexOf('CREATED_AT');
  if (cId < 0) return 0;

  // id column as plain text, so an identity number starting with 0 keeps its 0
  psh.getRange(2, cId + 1, Math.max(psh.getMaxRows() - 1, 1), 1).setNumberFormat('@');

  var have = {};
  for (var i = 1; i < pd.length; i++) {
    var k = String(pd[i][cId]).trim();
    if (k) have[k] = true;
  }

  var fv = fsh.getRange(2, 1, fsh.getLastRow() - 1, fsh.getLastColumn()).getValues();
  var width = psh.getLastColumn(), added = 0;
  for (var j = 0; j < fv.length; j++) {
    var id = String(fv[j][fId]).trim();
    if (!id || have[id]) continue;
    have[id] = true;
    var row = new Array(width);
    for (var c = 0; c < width; c++) row[c] = '';
    row[cId] = id;
    if (cName >= 0 && fName >= 0) row[cName] = String(fv[j][fName]).trim();
    if (cPhone >= 0 && fPhone >= 0) row[cPhone] = fixPhone_(fv[j][fPhone]);
    if (cFile >= 0) row[cFile] = 'uploaded';
    if (cCreated >= 0) row[cCreated] = (fv[j][0] instanceof Date) ? fv[j][0] : new Date();
    psh.appendRow(row);
    added++;
  }
  return added;
}

// ---- Drive folders of the clinic -----------------------------------------------
// One main folder with four folders inside, and under "patients" one folder per patient
// (named by PATIENT_ID, never by name) with one folder per treatment inside it.
// Nothing is ever moved or deleted here: folders are only created when they are missing.
var CLINIC_ROOT = 'קליניקה - ניסים קטנים';
var CLINIC_SUB = { app: 'אפליקציה', settings: 'הגדרות', sheets: 'טבלאות', patients: 'מטופלים' };

function childFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// the folder by its saved id; created (and saved) when it is missing or was trashed
function savedFolder_(key, make) {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty(key), f = null;
  if (id) { try { f = DriveApp.getFolderById(id); if (f.isTrashed()) f = null; } catch (e) { f = null; } }
  if (!f) { f = make(); p.setProperty(key, f.getId()); }
  return f;
}

function clinicRoot_() {
  return savedFolder_('FOLDER_ROOT', function () {
    var it = DriveApp.getRootFolder().getFoldersByName(CLINIC_ROOT);
    return it.hasNext() ? it.next() : DriveApp.createFolder(CLINIC_ROOT);
  });
}

function clinicFolders_() {
  var root = clinicRoot_(), out = { root: root };
  Object.keys(CLINIC_SUB).forEach(function (k) {
    out[k] = savedFolder_('FOLDER_' + k.toUpperCase(), function () { return childFolder_(root, CLINIC_SUB[k]); });
  });
  return out;
}

// the folder of one treatment: patients / <PATIENT_ID> / "טיפול 01 - 2026-11-12" (created when missing)
function treatFolder_(patients, pid, number, d) {
  var num = Number(number) || 0;
  var day = Object.prototype.toString.call(d) === '[object Date]' ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd')
    : (String(d).match(/^(\d{2})\/(\d{2})\/(\d{4})/) ? String(d).replace(/^(\d{2})\/(\d{2})\/(\d{4}).*$/, '$3-$2-$1') : '');
  var name = 'טיפול ' + (num < 10 ? '0' : '') + num + (day ? ' - ' + day : '');
  return childFolder_(childFolder_(patients, pid), name);
}

// COPIES (never moves) the photos of the treatments into their treatment folder. AppSheet shows the photos from
// its own folder, so the originals must stay where they are. A photo already copied is skipped, so running
// this again changes nothing. days = only treatments of the last N days (0 = all); returns the copies made now.
var PHOTO_COLS = ['PHOTO_FRONT', 'PHOTO_LEFT', 'PHOTO_BACK', 'PHOTO_RIGHT', 'AFTER_FRONT', 'AFTER_LEFT', 'AFTER_BACK', 'AFTER_RIGHT'];
function copyPhotos_(days) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TREATMENTS');
  if (!sh || sh.getLastRow() < 2) return 0;
  var v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return String(x).trim(); });
  var pi = h.indexOf('PATIENT_ID'), ni = h.indexOf('TREATMENT_NUMBER'), di = h.indexOf('DATE');
  var cols = PHOTO_COLS.map(function (c) { return h.indexOf(c); }).filter(function (i) { return i >= 0; });
  if (pi < 0 || !cols.length) return 0;
  var patients = clinicFolders_().patients, made = 0, MAX = 40;
  for (var i = v.length - 1; i >= 1 && made < MAX; i--) {
    var pid = String(v[i][pi]).trim();
    if (!pid) continue;
    if (days > 0) { var t = parseStamp_(di >= 0 ? v[i][di] : ''); if (t && Date.now() - t > days * 86400000) continue; }
    var dest = null;
    for (var c = 0; c < cols.length && made < MAX; c++) {
      var base = String(v[i][cols[c]]).split('/').pop().trim();
      if (!base) continue;
      if (!dest) dest = treatFolder_(patients, pid, ni >= 0 ? v[i][ni] : 0, di >= 0 ? v[i][di] : '');
      if (dest.getFilesByName(base).hasNext()) continue;
      var it = DriveApp.getFilesByName(base);
      if (!it.hasNext()) continue; // the photo is not in Drive yet: tried again on a later run
      it.next().makeCopy(base, dest);
      made++;
    }
  }
  return made;
}

// a folder for every treatment that does not have one yet; returns how many were created now
function treatmentFolders_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TREATMENTS');
  if (!sh || sh.getLastRow() < 2) return 0;
  var p = PropertiesService.getScriptProperties();
  var done = String(p.getProperty('FOLDERS_DONE') || '').split(',').filter(String), have = {};
  done.forEach(function (t) { have[t] = true; });
  var v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return String(x).trim(); });
  var ci = h.indexOf('Treatment_ID'), pi = h.indexOf('PATIENT_ID'), ni = h.indexOf('TREATMENT_NUMBER'), di = h.indexOf('DATE');
  if (ci < 0 || pi < 0) return 0;
  var todo = [];
  for (var i = 1; i < v.length; i++) {
    var tid = String(v[i][ci]).trim(), pat = String(v[i][pi]).trim();
    if (tid && pat && !have[tid]) todo.push(v[i]);
  }
  if (!todo.length) return 0;
  var patients = clinicFolders_().patients, made = 0;
  todo.forEach(function (r) {
    treatFolder_(patients, String(r[pi]).trim(), ni >= 0 ? r[ni] : 0, di >= 0 ? r[di] : '');
    done.push(String(r[ci]).trim());
    made++;
  });
  // a script property holds about 9,000 characters: keep the newest ids, an older one is only re-checked (never doubled)
  while (done.join(',').length > 8000) done.shift();
  p.setProperty('FOLDERS_DONE', done.join(','));
  return made;
}

// RUN THIS ONCE by hand after pasting: creates the folders (it asks for Drive permission) and shows where they are.
function installClinicFolders() {
  var f = clinicFolders_(), n = treatmentFolders_(), p = copyPhotos_(0);
  status_('folders', true, '');
  Logger.log(f.root.getUrl() + ' | treatment folders created now: ' + n + ' | photos copied now: ' + p);
  if (p >= 40) Logger.log('There may be more photos: run installClinicFolders again.');
}

// ---- every minute ------------------------------------------------------------
// create the rows of new treatments, finish findings that were waiting for rows,
// and pick up any patient the form trigger missed.
// A minute trigger runs 1440 times a day and Google allows 90 minutes of trigger
// time a day in total. Doing the full work every time would spend about three
// times that and the triggers would be cut off for the rest of the day - the app
// would simply stop, quietly. So a run that has nothing to do ends in a second:
// it only compares the length of the sheets with what it saw last time. Every
// tenth run does the full work anyway, so nothing can stay stuck.
function idle_() {
  var p = PropertiesService.getScriptProperties();
  var n = Number(p.getProperty('TICK_N') || 0) + 1;
  p.setProperty('TICK_N', String(n));
  if (n % 10 === 0) return false;
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sig = [];
  ['TREATMENTS', 'NEURO_RESULTS', 'MUSCLE_RESULTS', 'NEURO_TESTS', 'MUSCLES', 'FINDINGS', 'PATIENTS']
    .forEach(function (name) {
      var sh = ss.getSheetByName(name);
      sig.push(sh ? sh.getLastRow() : 0);
    });
  var f = formSheet_();
  sig.push(f ? f.getLastRow() : 0);
  var now = sig.join('.');
  if (now === p.getProperty('TICK_SIG')) return true;
  p.setProperty('TICK_SIG', now);
  return false;
}

function tick() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return;
  try {
    guard_('tick', function () {
      if (idle_()) return;
      ensureTreatmentCols_();
      backfillPatients_();
      generateRows_();
      applyPending_('');
      // the folders must never stop the rest: a Drive error is only recorded on the STATUS sheet
      try { treatmentFolders_(); copyPhotos_(14); status_('folders', true, ''); } catch (ef) { status_('folders', false, (ef && ef.stack) ? ef.stack : String(ef)); }
    });
  }
  finally { try { lock.releaseLock(); } catch (e) {} }
}
function applyPending() { tick(); } // old trigger name, kept so an old trigger still works

function ensureApplyTrigger_() {
  var has = false;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'tick') has = true;
    else if (fn === 'applyPending') ScriptApp.deleteTrigger(t); // replaced by the 1-minute trigger
  });
  if (!has) ScriptApp.newTrigger('tick').timeBased().everyMinutes(1).create();
}

// RUN THIS ONCE by hand in the editor (choose installApplyTrigger, press Run): it creates the 1-minute trigger.
function installApplyTrigger() { ensureApplyTrigger_(); generateRows_(); applyPending_(''); }

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

// ---- daily backup: a second folder with the last 14 daily copies ----------------
// The weekly copies above stay as they are. Run installDailyBackup() once by hand.
var DAILY_FOLDER = 'גיבויי קליניקה - יומי';
var DAILY_KEEP = 14;

function dailyBackup() {
  guard_('dailyBackup', function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var file = DriveApp.getFileById(ss.getId());
    var it = DriveApp.getFoldersByName(DAILY_FOLDER);
    var folder = it.hasNext() ? it.next() : DriveApp.createFolder(DAILY_FOLDER);
    var name = ss.getName() + ' - יומי ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    var same = folder.getFilesByName(name);
    if (!same.hasNext()) file.makeCopy(name, folder);   // at most one copy per day
    var list = [], files = folder.getFiles();
    while (files.hasNext()) { var f = files.next(); list.push({ f: f, t: f.getDateCreated().getTime() }); }
    list.sort(function (a, b) { return b.t - a.t; });
    for (var i = DAILY_KEEP; i < list.length; i++) list[i].f.setTrashed(true);
  });
}

function installDailyBackup() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'dailyBackup') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('dailyBackup').timeBased().everyDays(1).atHour(2).create();
  dailyBackup(); // first copy right away, also asks for the Drive permission
}

// ---- nightly integrity check ("doctor") -------------------------------------------
// Only reads and reports, never changes or deletes data. Result goes to the STATUS sheet (row "doctor");
// an e-mail is sent only when something is wrong.
var DOCTOR_SKIP = [];   // treatment ids the check should leave out (none)
var STUCK_MIN = 60;               // a finding with an empty "applied" cell for longer than this is stuck

function doctor() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), problems = [];
  var tids = {}, tlist = keysOf_('TREATMENTS', 'Treatment_ID');
  tlist.forEach(function (t) { tids[t] = true; });
  var parts = [
    { name: 'NEURO_RESULTS', key: 'test_id', expect: keysOf_('NEURO_TESTS', 'test_id') },
    { name: 'MUSCLE_RESULTS', key: 'muscle_id', expect: keysOf_('MUSCLES', 'muscle_id') }
  ];
  parts.forEach(function (p) {
    var sh = ss.getSheetByName(p.name);
    if (!sh || sh.getLastRow() < 2) { if (tlist.length) problems.push(p.name + ': ures lap'); return; }
    var idx = headIdx_(sh);
    if (!idx.Treatment_ID || !idx[p.key]) { problems.push(p.name + ': hianyzo oszlop'); return; }
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    var per = {}, orphan = {};
    for (var i = 0; i < v.length; i++) {
      var t = String(v[i][idx.Treatment_ID - 1]).trim(), k = String(v[i][idx[p.key] - 1]).trim();
      if (!t && !k) continue;
      if (!tids[t]) { orphan[t || '(ures)'] = (orphan[t || '(ures)'] || 0) + 1; continue; }
      (per[t] = per[t] || {})[k] = ((per[t] || {})[k] || 0) + 1;
    }
    Object.keys(orphan).forEach(function (t) { problems.push(p.name + ': ' + orphan[t] + ' arva sor (kezeles nincs): ' + t); });
    tlist.forEach(function (t) {
      if (DOCTOR_SKIP.indexOf(t) >= 0) return;
      var have = per[t] || {}, miss = 0, dup = [];
      p.expect.forEach(function (k) { if (!have[k]) miss++; });
      Object.keys(have).forEach(function (k) { if (have[k] > 1) dup.push(k); });
      if (miss) problems.push(p.name + ' ' + t + ': ' + miss + ' hianyzo sor / ' + p.expect.length);
      if (dup.length) problems.push(p.name + ' ' + t + ': dupla ' + p.key + ': ' + dup.slice(0, 5).join(','));
    });
  });
  // findings that never reached their result row
  var fs = ss.getSheetByName('FINDINGS');
  if (fs && fs.getLastRow() >= 2) {
    var fi = headIdx_(fs), fv = fs.getRange(2, 1, fs.getLastRow() - 1, fs.getLastColumn()).getValues(), stuck = 0;
    for (var j = 0; j < fv.length; j++) {
      var ap = fi.applied ? String(fv[j][fi.applied - 1]).trim() : '';
      var ft = fi.treatment_id ? String(fv[j][fi.treatment_id - 1]).trim() : '';
      if (ap || !ft) continue;
      var at = fi.received_at ? parseStamp_(fv[j][fi.received_at - 1]) : 0;
      if (!at || Date.now() - at > STUCK_MIN * 60000) stuck++;
    }
    if (stuck) problems.push('FINDINGS: ' + stuck + ' lelet nincs alkalmazva (applied ures)');
  }
  if (problems.length) {
    var msg = problems.join(' | ');
    status_('doctor', false, msg);
    alert_('doctor', 'M.T.R.S.M. - az ejszakai ellenorzes hibat talalt',
      'Az ellenorzes semmit nem javit es nem torol, csak jelez.\nIdo: ' + stamp_() + '\n\n' + problems.join('\n') +
      '\n\nReszletek a STATUS lapon (doctor sor).');
  } else {
    status_('doctor', true, '');
    try { PropertiesService.getScriptProperties().deleteProperty('alert_doctor'); } catch (e) {}
  }
  return problems;
}

// RUN THIS ONCE by hand: nightly check around 04:00, and runs it right away.
function installDoctor() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'doctor') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('doctor').timeBased().everyDays(1).atHour(4).create();
  Logger.log(JSON.stringify(doctor()));
}

// ---- watchdog: a failure must never stay silent -----------------------------
// One STATUS row per watched function, overwritten on every run, plus an e-mail
// on the first failure. Without this a single bad character stops the whole
// pipeline (no patients, no result rows, no dictation) and nothing says so.
var STATUS_SHEET = 'STATUS';
var STATUS_HEAD = ['what', 'when', 'state', 'message'];
var ALERT_GAP_H = 6;   // at most one e-mail per this many hours for the same function
var STALE_MIN = 30;    // tick is considered stuck after this many minutes without a run

// writes the current state of one function; never throws
function status_(name, ok, msg) {
  try {
    var sh = sheet_(STATUS_SHEET, STATUS_HEAD);
    var last = sh.getLastRow(), row = 0;
    if (last >= 2) {
      var col = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < col.length; i++) if (String(col[i][0]).trim() === name) { row = i + 2; break; }
    }
    if (!row) row = Math.max(last, 1) + 1;
    // plain text, or the Sheet turns the stamp into a date value and watchdog cannot read it back
    sh.getRange(row, 1, 1, 4).setNumberFormat('@')
      .setValues([[name, stamp_(), ok ? 'OK' : 'ERROR', ok ? '' : String(msg).slice(0, 500)]]);
  } catch (e) {}
}

// e-mails the owner, but not more often than ALERT_GAP_H for the same function; never throws
function alert_(name, subject, body) {
  try {
    var p = PropertiesService.getScriptProperties(), k = 'alert_' + name;
    var now = Date.now(), prev = Number(p.getProperty(k) || 0);
    if (now - prev < ALERT_GAP_H * 3600000) return;
    p.setProperty(k, String(now));
    var to = Session.getEffectiveUser().getEmail();
    if (to) MailApp.sendEmail(to, subject, body);
  } catch (e) {}
}

// runs fn, records what happened, and shouts on a failure
function guard_(name, fn) {
  try {
    fn();
    try { PropertiesService.getScriptProperties().deleteProperty('alert_' + name); } catch (e) {}
    status_(name, true, '');
  } catch (err) {
    var msg = (err && err.stack) ? err.stack : String(err);
    status_(name, false, msg);
    alert_(name, 'M.T.R.S.M. - a szkript hibara fut (' + name + ')',
      'A hatterfeldolgozas hibara futott, ezert az app nem toltodik ki magatol.\n\n' +
      'Fuggveny: ' + name + '\nIdo: ' + stamp_() + '\n\n' + msg +
      '\n\nAz aktualis allapot a tablazat STATUS lapjan latszik.');
    throw err; // stays visible in the execution log as well
  }
}

// daily: notices if the 1-minute trigger stopped firing altogether
function watchdog() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET);
  var seen = null, state = '';
  if (sh && sh.getLastRow() >= 2) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === 'tick') { seen = v[i][1]; state = String(v[i][2]); }
  }
  var t = parseStamp_(seen);
  if (!t || Date.now() - t > STALE_MIN * 60000) {
    var age = t ? Math.round((Date.now() - t) / 60000) + ' perc' : 'nem olvashato';
    alert_('watchdog', 'M.T.R.S.M. - a percenkenti feldolgozas all',
      'A tick fuggveny utolso jelzese: ' + (seen ? stampOf_(t) : 'nincs') + ' (' + (state || '-') + '), kora: ' + age + '.\n' +
      'Ennel regebbi, mint ' + STALE_MIN + ' perc, tehat a percenkenti aktivalo valoszinuleg leallt.\n' +
      'Apps Script > Aktivalok, es futtasd ujra az installApplyTrigger fuggvenyt.');
  }
}

// the stamp back to milliseconds; 0 if it cannot be read. The cell may hold the
// text 'dd/MM/yyyy HH:mm:ss' or, in an older sheet, a real date value.
function parseStamp_(s) {
  if (s == null || s === '') return 0;
  if (Object.prototype.toString.call(s) === '[object Date]') return s.getTime();
  var m = String(s).match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], +m[6]).getTime();
  var p = Date.parse(String(s));
  return isNaN(p) ? 0 : p;
}

function stampOf_(ms) { return Utilities.formatDate(new Date(ms), TZ, 'dd/MM/yyyy HH:mm:ss'); }

// RUN THIS ONCE by hand after pasting: creates the daily watchdog trigger and asks for the e-mail permission.
function installWatchdog() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'watchdog') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('watchdog').timeBased().everyDays(1).atHour(7).create();
  status_('tick', true, '');
  alert_('install', 'M.T.R.S.M. - a hibaor bekapcsolt',
    'Ez a proba-level azt jelzi, hogy a hibaor mukodik es tud levelet kuldeni.\nIdo: ' + stamp_());
}

// ---- one-time: the muscle ids M01..M56 become the plain numbers 11..65 (the order of the MUSCLES sheet) --
// Run once from the editor (migrateMuscleIds). Safe to run again: only cells that still hold an "M.." id are changed.
function migrateMuscleIds() {
  var order = ['M01','M02','M03','M04','M05','M06','M07','M08','M09','M10','M11','M12','M13','M14','M15','M16','M17','M18','M19','M20','M21','M22','M23','M24','M25','M27','M28','M29','M30','M31','M32','M33','M34','M56','M35','M36','M37','M38','M39','M40','M41','M42','M43','M44','M45','M46','M47','M48','M49','M50','M51','M52','M53','M54','M55'];
  var map = {}; order.forEach(function (id, i) { map[id] = String(i + 11); });
  var ss = SpreadsheetApp.getActiveSpreadsheet(), rep = [];
  ['MUSCLES', 'MUSCLE_RESULTS', 'FINDINGS'].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) { rep.push(name + ': -'); return; }
    var c = headIdx_(sh).muscle_id;
    if (!c) { rep.push(name + ': nincs muscle_id oszlop'); return; }
    var rng = sh.getRange(2, c, sh.getLastRow() - 1, 1), v = rng.getValues(), n = 0;
    var out = v.map(function (r) { var x = String(r[0]).trim(); if (map[x]) { n++; return [map[x]]; } return [r[0]]; });
    rng.setNumberFormat('@'); rng.setValues(out);
    rep.push(name + ': ' + n + ' cella atirva');
  });
  Logger.log(rep.join(' | '));
}
