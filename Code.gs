/**
 * AVA PRACTICE OPERATING SYSTEM — GOOGLE APPS SCRIPT API (Code.gs)
 * Aggarwal Vikram & Associates, Chartered Accountants
 *
 * This project is now used ONLY as a JSON API. The dashboard UI lives in a
 * separate index.html deployed on Vercel, which calls this Web App's URL.
 *
 * SETUP:
 *   1. Paste this file into your Apps Script project (bound to the sheet, or
 *      standalone with the sheet's ID — see SHEET_ID note below if standalone).
 *   2. Deploy → New deployment → type "Web app".
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      Click Deploy, then copy the URL ending in /exec.
 *   3. Paste that URL into APPS_SCRIPT_URL near the top of index.html's
 *      <script> tag, then deploy index.html on Vercel.
 *
 *   Whenever you edit this file, you must create a NEW VERSION under
 *   "Manage deployments" (pencil icon → Version: New version → Deploy) for
 *   the live /exec URL to pick up your changes.
 */

var SHEET_NAME = 'Sheet1';
var CACHE_SECONDS = 120;

var COL_FLOW      = 14; // O
var COL_FMS       = 15; // P
var COL_CHECKLIST = 18; // S
var COL_FORM      = 19; // T
var PROCESS_RES_COLS = [COL_FLOW, COL_FMS, COL_CHECKLIST, COL_FORM];

/**
 * Web App entry point. Always returns JSON: { depts: [...], synced: "..." }.
 * Query params:
 *   ?refresh=1   forces a fresh read, bypassing the cache.
 */
function doGet(e) {
  var params = (e && e.parameter) || {};
  var force = params.refresh === '1' || params.refresh === 'true';

  var data;
  try {
    data = getData(force);
  } catch (err) {
    data = { error: true, message: String(err && err.message || err) };
  }

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function headerMap_(headers) {
  var defs = [
    { type: 'Dashboard',      test: /dashboard/i },
    { type: 'Flow Chart',     test: /flow\s*chart/i },
    { type: 'F.M.S.',         test: /f\.?\s*m\.?\s*s\.?/i },
    { type: 'Checklist',      test: /checklist/i },
    { type: 'Google Form',    test: /google\s*form/i },
    { type: 'Google Doc',     test: /google\s*doc/i },
    { type: 'Training Video', test: /training/i },
    { type: 'Drive',          test: /drive\s*link/i }
  ];
  var map = {};
  headers.forEach(function (h, c) {
    if (PROCESS_RES_COLS.indexOf(c) > -1) return;
    var t = String(h || '').trim();
    if (!t) return;
    for (var i = 0; i < defs.length; i++) {
      if (defs[i].test.test(t)) { map[c] = defs[i].type; return; }
    }
  });
  return map;
}

function linkOf_(sh, rich, formulas, vals, r, c) {
  if (c < 0 || c >= vals[r].length) return null;

  var rt = rich[r][c];
  if (rt) {
    var u = rt.getLinkUrl();
    if (u) return u;
    var runs = rt.getRuns();
    for (var i = 0; i < runs.length; i++) {
      var ru = runs[i].getLinkUrl();
      if (ru) return ru;
    }
  }

  var f = String(formulas[r][c] || '');
  if (/HYPERLINK\s*\(/i.test(f)) {
    var m = f.match(/HYPERLINK\s*\(\s*(?:"([^"]+)"|'([^']+)'|(\$?[A-Za-z]{1,3}\$?[0-9]+))/i);
    if (m) {
      if (m[1]) return m[1];
      if (m[2]) return m[2];
      if (m[3]) {
        try {
          var refA1 = m[3].replace(/\$/g, '');
          var refVal = String(sh.getRange(refA1).getDisplayValue() || '').trim();
          if (/^https?:\/\//i.test(refVal)) return refVal;
        } catch (e) { /* invalid ref, ignore */ }
      }
    }
  }

  var v = String(vals[r][c] || '').trim();
  if (/^https?:\/\//i.test(v)) return v;

  return null;
}

function classify_(base, url) {
  if (base === 'Google Form' && url.indexOf('docs.google.com/document') > -1) return 'Google Doc';
  if (base === 'Google Form' && url.indexOf('drive.google.com/drive/folders') > -1) return 'Drive Folder';
  if (base === 'Flow Chart' && url.indexOf('script.google.com') > -1) return 'Doer App';
  return base;
}

function processRes_(sh, rich, formulas, vals, r, col, fallbackLabel) {
  var url = linkOf_(sh, rich, formulas, vals, r, col);
  if (!url) return null;
  var text = String(vals[r][col] || '').trim();
  var label = (text && !/^https?:\/\//i.test(text)) ? text : fallbackLabel;
  return { url: url, label: label };
}

function getData(forceRefresh) {
  var cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    var hit = cache.get('POS_DATA');
    if (hit) return JSON.parse(hit);
  }

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), 25);
  if (lastRow < 2) return { depts: [], synced: new Date().toISOString() };

  var rng      = sh.getRange(1, 1, lastRow, lastCol);
  var vals     = rng.getDisplayValues();
  var rich     = rng.getRichTextValues();
  var formulas = rng.getFormulas();

  var headers = vals[0];
  var LINKS = headerMap_(headers);

  function colByHeader(re) {
    for (var c = 0; c < headers.length; c++) {
      if (re.test(String(headers[c] || ''))) return c;
    }
    return -1;
  }
  var C_DEPT   = colByHeader(/department/i);
  var C_PROC   = colByHeader(/list of all processes|^process(es)?$/i);
  var C_PURP   = colByHeader(/purpose/i);
  var C_RESULT = colByHeader(/result/i);
  var C_ACCESS = colByHeader(/ownership\s*access/i);
  var C_STATUS = colByHeader(/sheet\s*status|^status$/i);
  var C_PC     = colByHeader(/process\s*coordinator/i);
  var C_SOLVER = colByHeader(/problem\s*solver/i);
  var C_DOER   = colByHeader(/doer\s*name/i);
  var C_FREQ   = colByHeader(/^freq/i);
  if (C_DEPT < 0)   C_DEPT = 1;
  if (C_PROC < 0)   C_PROC = 3;
  if (C_PURP < 0)   C_PURP = 4;
  if (C_RESULT < 0) C_RESULT = 5;
  if (C_ACCESS < 0) C_ACCESS = 6;
  if (C_STATUS < 0) C_STATUS = 7;
  if (C_PC < 0)     C_PC = 8;
  if (C_SOLVER < 0) C_SOLVER = 9;
  if (C_DOER < 0)   C_DOER = 10;
  if (C_FREQ < 0)   C_FREQ = 16;

  var depts = [], cur = null;

  for (var r = 1; r < lastRow; r++) {
    var deptName = String(vals[r][C_DEPT] || '').trim();
    if (deptName) {
      if (cur) depts.push(cur);
      cur = {
        name: deptName,
        purpose: String(vals[r][C_PURP] || '').trim(),
        status:  String(vals[r][C_STATUS] || '').trim(),
        pc:      String(vals[r][C_PC] || '').trim(),
        solver:  String(vals[r][C_SOLVER] || '').trim(),
        access:  String(vals[r][C_ACCESS] || '').trim(),
        freq: '',
        processes: [], resources: [], _seen: {}
      };
    }
    if (!cur) continue;

    var proc = String(vals[r][C_PROC] || '').trim();
    if (proc) {
      var res = {};
      var flow      = processRes_(sh, rich, formulas, vals, r, COL_FLOW, 'Flow');
      var fms       = processRes_(sh, rich, formulas, vals, r, COL_FMS, 'F.M.S.');
      var checklist = processRes_(sh, rich, formulas, vals, r, COL_CHECKLIST, 'Checklist');
      var form      = processRes_(sh, rich, formulas, vals, r, COL_FORM, 'Google Form');
      if (flow)      res.flow = flow;
      if (fms)       res.fms = fms;
      if (checklist) res.checklist = checklist;
      if (form)      res.form = form;

      cur.processes.push({
        name: proc,
        result: String(vals[r][C_RESULT] || '').trim(),
        doer:   String(vals[r][C_DOER] || '').trim(),
        res: res
      });
    }

    for (var c in LINKS) {
      c = Number(c);
      var url = linkOf_(sh, rich, formulas, vals, r, c);
      if (!url || cur._seen[url]) continue;
      cur._seen[url] = true;
      var text = String(vals[r][c] || '').trim();
      var type = classify_(LINKS[c], url);
      cur.resources.push({
        type: type,
        label: (text && !/^https?:\/\//i.test(text)) ? text : type,
        url: url
      });
    }

    var fq = String(vals[r][C_FREQ] || '').trim();
    if (fq && !cur.freq) cur.freq = fq;
  }
  if (cur) depts.push(cur);
  depts.forEach(function (d) { delete d._seen; });

  var out = { depts: depts, synced: new Date().toISOString() };
  try { cache.put('POS_DATA', JSON.stringify(out), CACHE_SECONDS); } catch (e) {}
  return out;
}

function diagnose() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  Logger.log('Reading tab: "%s" | Last row: %s | Last col: %s',
             sh.getName(), sh.getLastRow(), sh.getLastColumn());
  var lastRow = sh.getLastRow(), lastCol = Math.max(sh.getLastColumn(), 25);
  var rng = sh.getRange(1, 1, lastRow, lastCol);
  var vals = rng.getDisplayValues(), rich = rng.getRichTextValues(), formulas = rng.getFormulas();

  var LINKS = headerMap_(vals[0]);
  Logger.log('Department-level resource columns identified: %s', JSON.stringify(
    Object.keys(LINKS).map(function (c) {
      return String.fromCharCode(65 + Number(c)) + ' = ' + LINKS[c];
    })));

  var perProcCounts = { 'Flow (O)': 0, 'F.M.S. (P)': 0, 'Checklist (S)': 0, 'Google Form (T)': 0 };
  var cols = { 'Flow (O)': COL_FLOW, 'F.M.S. (P)': COL_FMS, 'Checklist (S)': COL_CHECKLIST, 'Google Form (T)': COL_FORM };
  for (var key in cols) {
    for (var r = 1; r < lastRow; r++) {
      if (linkOf_(sh, rich, formulas, vals, r, cols[key])) perProcCounts[key]++;
    }
  }
  Logger.log('Per-process links detected (fixed columns O/P/S/T): %s', JSON.stringify(perProcCounts));

  var counts = {};
  for (var c in LINKS) {
    c = Number(c); counts[LINKS[c] + ' (' + String.fromCharCode(65 + c) + ')'] = 0;
    for (var r2 = 1; r2 < lastRow; r2++) {
      if (linkOf_(sh, rich, formulas, vals, r2, c)) counts[LINKS[c] + ' (' + String.fromCharCode(65 + c) + ')']++;
    }
  }
  Logger.log('Department-level links detected per column: %s', JSON.stringify(counts));

  var data = getData(true);
  Logger.log('Departments: %s | Total dept-level resources: %s | Total processes: %s', data.depts.length,
    data.depts.reduce(function (a, d) { return a + d.resources.length; }, 0),
    data.depts.reduce(function (a, d) { return a + d.processes.length; }, 0));
}
