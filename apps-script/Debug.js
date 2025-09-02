function health() {
  var out = {
    ENV: env_(),
    SHEET_ID: conf_('SHEET_ID'),
    DOCS_WEBHOOK: !!conf_('DOCS_WEBHOOK'),
    LEADERS_WEBHOOK: !!conf_('LEADERS_WEBHOOK'),
    DEST_FOLDER_ID: conf_('DEST_FOLDER_ID')
  };
  Logger.log(JSON.stringify(out, null, 2));
  try { ss_().getSheetByName('Logs').appendRow([new Date(),'INFO','health', JSON.stringify(out)]); } catch (e) {}
  return out;
}

function debug_export_pdf_once(){
  // Use your known-good template & destination (prod ones if possible)
  var tpl = CONFIG.TEMPLATES.EVENT_RECORD;
  var destFolder = DriveApp.getFolderById(CONFIG.DEST_FOLDER_ID);

  // 1) Make a tiny copy & add a marker so we know we wrote into it
  var copy = DriveApp.getFileById(tpl).makeCopy('DEBUG Export Test (merge)', destFolder);
  var copyId = copy.getId();
  var d = DocumentApp.openById(copyId);
  d.getBody().appendParagraph('~debug export marker~');
  d.saveAndClose();

  Utilities.sleep(400); // settle

  // 2) Try fast path (getAs)
  try {
    var blob1 = DriveApp.getFileById(copyId).getAs('application/pdf'); // typical failure spot
    var f1 = destFolder.createFile(blob1).setName('DEBUG getAs.pdf');
    Logger.log('getAs: OK ' + f1.getId());
  } catch (e1) {
    Logger.log('getAs: FAIL ' + e1);
  }

  // 3) Try HTTP export fallback
  try {
    var url = 'https://docs.google.com/document/d/' + copyId + '/export?format=pdf';
    var blob2 = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    }).getBlob();
    blob2.setName('DEBUG http.pdf');
    var f2 = destFolder.createFile(blob2);
    Logger.log('HTTP: OK ' + f2.getId());
  } catch (e2) {
    Logger.log('HTTP: FAIL ' + e2);
  }

  try { DriveApp.getFileById(copyId).setTrashed(true); } catch(_){}
}


// Debug.gs
function diagLedgerBasics() {
  const ss = SpreadsheetApp.getActive();
  const tabName = CONFIG.TABS.POSITIVE_POINTS || CONFIG.TABS.POSITIVE; // ✅ correct keys
  const sh = ss.getSheetByName(tabName);
  if (!sh) throw new Error('Missing ' + tabName + ' sheet per CONFIG.');
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h||'').trim());
  Logger.log(JSON.stringify({ sheet: sh.getName(), lastRow, lastCol,
                              headerCount: headers.filter(Boolean).length, headers }));
}



function debugGraceDecisionForRow(rowIndex) {
  var rowIndex = 64;
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Events'); // adjust if needed
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h||'').trim());
  function val(name){ var i=headers.indexOf(name); return i>=0 ? sh.getRange(rowIndex, i+1).getValue() : ''; }

  var checks = {
    Employee: val('Employee'),
    Points: val('Points'),
    GraceTier: val('Grace Tier') || val('Grace Tier (Req)') || val('Grace Tier Name'),
    GraceReason: val('Grace Reason'),
    GracedBy: val('Graced By'),
    GraceApplied: val('Grace Applied'),
    LedgerRow: val('Grace Ledger Row')
  };
  Logger.log('EVENTS row ' + rowIndex + ' checks: ' + JSON.stringify(checks));

  // quick opinions
  if (!checks.Employee) Logger.log('BLOCK: missing Employee');
  if (!Number(checks.Points)) Logger.log('BLOCK: Points not numeric/positive');
  if (!checks.GraceTier) Logger.log('BLOCK: missing Grace Tier');
  if (!checks.GraceReason) Logger.log('BLOCK: missing Grace Reason');
  if (!checks.GracedBy) Logger.log('BLOCK: missing Graced By');
  if (checks.GraceApplied) Logger.log('BLOCK: already graced');
}


function debug_dumpEventHeaders(){
  const s = sh_(CONFIG.TABS.EVENTS);
  const h = headers_(s);
  sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'DEBUG', 'dumpEventHeaders', JSON.stringify(h)]);
  return h;
}

function debugPeekCredits(employee, tierName){
  var pp = ensurePositivePointsColumns_();
  var s = pp.sheet, i = pp.idx;

  var last = s.getLastRow();
  if (last < 2){ Logger.log('no ledger rows'); return; }

  var n = last - 1;
  var emp = s.getRange(2, i.Employee,   n, 1).getValues();
  var typ = s.getRange(2, i.CreditType, n, 1).getValues();
  var apr = i.Approved ? s.getRange(2, i.Approved, n, 1).getValues() : null;
  var con = s.getRange(2, i.Consumed,   n, 1).getValues();

  var wantEmp  = String(employee||'').trim();
  var wantType = String(tierName||'').trim().toLowerCase();

  for (var r=0; r<n; r++){
    Logger.log('#%s emp=%s type=%s approved=%s consumed=%s',
      r+2,
      String(emp[r][0]||'').trim(),
      String(typ[r][0]||'').trim(),
      (apr ? JSON.stringify(apr[r][0]) : '(n/a)'),
      JSON.stringify(con[r][0])
    );
  }
  var avail = countAvailableCredits_(wantEmp, tierName);
  Logger.log('preflight avail=%s rows=%s', avail.count, avail.rows.join(','));
}


function debugApplyGraceRow(){
  ensurePolicyLoaded_();
  var row = 11;                       // <-- set your row
  var ok = applyGraceForEventRow(row, 'Director Debug', 'Manual test');
  Logger.log('applyGraceForEventRow(%s) -> %s', row, ok);
}

function debugWhichSpreadsheet(){
  var a = null; try { a = SpreadsheetApp.getActiveSpreadsheet(); } catch(_){}
  Logger.log('Active ID=%s Name=%s', a ? a.getId() : '(none)', a ? a.getName() : '(none)');
  Logger.log('CONFIG.SHEET_ID=%s', (CONFIG && CONFIG.SHEET_ID) || '(none)');
  var s = ss_();
  Logger.log('ss_() -> ID=%s Name=%s', s.getId(), s.getName());
}

peekCreditsVerbose('TestA','Moderate');

function peekCreditsVerbose(employee, tierName){
  var pp = ensurePositivePointsColumns_();
  var s = pp.sheet, i = pp.idx;
  var last = s.getLastRow();
  if (last < 2){ Logger.log('Ledger empty'); return; }

  function norm(s){ return String(s||'').replace(/\s+/g,' ').trim().toLowerCase(); }
  function isStrictTrue(v){
    if (v === true) return true;
    var t = String(v||'').trim().toLowerCase();
    return t === 'true' || t === 'yes' || t === 'y' || t === '1';
  }
  function isApproved(v){
    var str = String(v||'').trim();
    if (str === '') return true;    // legacy-friendly
    return isStrictTrue(v);
  }

  var n = last - 1;
  var emp = s.getRange(2, i.Employee,   n, 1).getValues();
  var typ = s.getRange(2, i.CreditType, n, 1).getValues();
  var apr = i.Approved ? s.getRange(2, i.Approved, n, 1).getValues() : null;
  var con = s.getRange(2, i.Consumed,   n, 1).getValues();

  var wantEmp  = norm(employee);
  var wantType = norm(tierName);

  var passed = [], failed = [];
  for (var r=0; r<n; r++){
    var rowIndex = r+2;
    var E = emp[r][0], T = typ[r][0], A = apr ? apr[r][0] : '(n/a)', C = con[r][0];

    var reasons = [];
    if (norm(E) !== wantEmp) reasons.push('emp');
    if (norm(T) !== wantType) reasons.push('type');
    if (!(apr ? isApproved(A) : true)) reasons.push('not approved');
    if (isStrictTrue(C)) reasons.push('consumed');

    if (!reasons.length) passed.push(rowIndex);
    else failed.push({row: rowIndex, reasons: reasons.join(', '), E:E, T:T, A:A, C:C});
  }
  Logger.log('PASS rows: %s', passed.join(',') || '(none)');
  failed.forEach(function(x){
    Logger.log('FAIL row %s — reasons=[%s] emp=%s type=%s approved=%s consumed=%s',
      x.row, x.reasons, x.E, x.T, x.A, x.C);
  });

  var avail = countAvailableCredits_(employee, tierName);
  Logger.log('Summary preflight avail=%s rows=%s', avail.count, (avail.rows||[]).join(','));
}


function debugLedgerAvailability(){
  ensurePolicyLoaded_();
  var employee = 'TestA';  // <— set
  var eventPts = 3;        // <— set

  var tier = selectCreditTierByPoints(eventPts);
  Logger.log('Needed tier: %s', tier && tier.name);

  var pp = ensurePositivePointsColumns_();
  var s = pp.sheet, idx = pp.idx, data = s.getDataRange().getValues();
  var cEmp = idx.Employee-1, cType = idx.CreditType-1, cAppr = idx.Approved-1, cCon = idx.Consumed-1;

  var want = String(tier && tier.name || '').trim().toLowerCase();
  var avail = 0, rows = [];
  for (var r = 1; r < data.length; r++){
    var row = data[r];
    var emp = String(row[cEmp]||'').trim();
    var typ = String(row[cType]||'').trim().toLowerCase();
    var okA = (idx.Approved ? isTruthy_(row[cAppr]) : true);
    var okC = !isTruthy_(row[cCon]);
    if (emp===employee && typ===want && okA && okC){ avail++; rows.push(r+1); }
  }
  Logger.log('Available %s credits for %s: %s (rows: %s)', (tier&&tier.name)||'—', employee, avail, rows.join(','));
}


function sanityCredits(){
  ensurePolicyLoaded_();
  Logger.log(JSON.stringify(CONFIG.POLICY.CREDITS));
  Logger.log('Tier for 3 pts: ' + ((selectCreditTierByPoints(3)||{}).name || 'none'));
}

function debugGraceRow(){
  const row = 11; // <-- SET THE ROW NUMBER YOU WANT TO INSPECT

  const sh = sh_(CONFIG.TABS.EVENTS);
  if (!sh) { Logger.log('Events sheet missing'); return; }

  const headers = headers_(sh);
  const idx = n => headers.indexOf(n);
  const val = n => { const i = idx(n); return i === -1 ? '' : sh.getRange(row, i+1).getValue(); };

  const policy = CONFIG.POLICY || {};
  const grace  = policy.GRACE || {};
  const windowDays = Number(grace.eligibleWindowDays || policy.ROLLING_DAYS || 180);

  const employee  = String(val(CONFIG.COLS.Employee)).trim();
  const points    = Number(val(CONFIG.COLS.Points) || 0);
  const active    = !!val(CONFIG.COLS.Active);
  const redLine   = idx(CONFIG.COLS.RedLine) !== -1 ? !!val(CONFIG.COLS.RedLine) : false;
  const eventType = String(val(CONFIG.COLS.EventType) || '').toLowerCase();

  const incRaw  = val(CONFIG.COLS.IncidentDate);
  const incDate = (incRaw instanceof Date) ? incRaw : (incRaw ? new Date(incRaw) : null);

  const already = String(val(CONFIG.COLS.GraceApplied)).toUpperCase()==='YES' || val(CONFIG.COLS.GraceApplied)===true;

  const problems = [];
  if (grace.enabled === false) problems.push('Grace disabled by config');
  if (!employee) problems.push('No Employee');
  if (!active) problems.push('Active = FALSE');
  if ((grace.excludeRedLines !== false) && redLine) problems.push('Red-line excluded');
  if (!incDate) problems.push('IncidentDate not a Date');
  else {
    const cutoff = new Date(); cutoff.setHours(0,0,0,0); cutoff.setDate(cutoff.getDate() - Math.abs(windowDays));
    if (incDate < cutoff) problems.push(`Outside window (${windowDays} days)`);
  }
  if (already) problems.push('Already graced');
  if (eventType.includes('milestone') || eventType.includes('positive') ||
      eventType === String((grace.eventTypeName || 'Grace')).toLowerCase()) {
    problems.push('Non-disciplinary row');
  }

  // Tier selection
  const tiers = (CONFIG.POLICY && CONFIG.POLICY.CREDITS) || [];
  let tier = null;
  for (let k = 0; k < tiers.length; k++){
    if (points <= Number(tiers[k].maxPoints || 0)) { tier = tiers[k]; break; }
  }
  if (!tier) problems.push('No tier for '+points+' pts');

  // Credit availability (trim/case-insensitive)
  let creditSummary = 'N/A';
  if (!problems.length && tier){
    const ps = sh_(CONFIG.TABS.POSITIVE_POINTS);
    if (!ps){ problems.push('PositivePoints tab missing'); }
    else {
      const pv = ps.getDataRange().getValues();
      if (pv.length < 2) problems.push('No PositivePoints rows');
      else {
        const ph = pv[0].map(x => String(x || '').trim());
        const pi = n => ph.indexOf(n);
        const colEmp = pi('Employee'), colType = pi('Credit Type'), colApproved = pi('Approved?'), colConsumed = pi('Consumed?');

        const requireApproval = !((CONFIG.POLICY && CONFIG.POLICY.GRACE && CONFIG.POLICY.GRACE.requireApproval) === false);
        const wantType = String(tier.name || '').trim().toLowerCase();

        let available = 0;
        for (let r = 1; r < pv.length; r++){
          const rowv = pv[r];
          const emp  = String(rowv[colEmp]  || '').trim();
          const type = String(rowv[colType] || '').trim().toLowerCase();
          const appr = colApproved === -1 ? true : (rowv[colApproved] === true || String(rowv[colApproved]).toUpperCase() === 'TRUE');
          const cons = colConsumed === -1 ? false : (rowv[colConsumed] === true || String(rowv[colConsumed]).toUpperCase() === 'TRUE');
          if (emp === employee && type === wantType && !cons && (!requireApproval || appr)) available++;
        }
        creditSummary = `${available} ${tier.name} credit(s) available`;
        if (!available) problems.push(`No available ${tier.name} credit for ${employee}`);
      }
    }
  }

  if (problems.length){
    Logger.log('❌ Row %s blocked: %s', row, problems.join(' | '));
  } else {
    Logger.log('✅ Row %s precheck OK. Tier=%s; %s', row, tier.name, creditSummary);
  }
  Logger.log('Info — employee=%s, points=%s, active=%s, eventType=%s, date=%s',
             employee, points, active, eventType, incDate);
}



function debug_inspectWriteupCell(row){
  if (!row || isNaN(Number(row))) {
    sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'ERROR', 'inspectWriteupCell', 'missing or invalid row arg', String(row)]);
    throw new Error('inspectWriteupCell requires a numeric row argument, e.g. debug_inspectWriteupCell(50501)');
  }
  const s = sh_(CONFIG.TABS.EVENTS);
  const headers = headers_(s);
  const headerName = (CONFIG.COLS.WriteupPDF || 'Write-Up PDF');
  const colIndex = headers.indexOf(headerName) + 1; // 1-based

  sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'DEBUG', 'inspectWriteupCell', 'row', row, 'headerName', headerName, 'colIndex', colIndex]);

  if (!colIndex || colIndex <= 0) {
    sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'ERROR', 'inspectWriteupCell', 'header missing', headerName]);
    throw new Error('Header missing: ' + headerName);
  }

  const cell = s.getRange(Number(row), colIndex);
  const val = cell.getValue();
  const formula = cell.getFormula();
  let protectionInfo = 'UNDETERMINED';
  try {
    const protections = s.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(p => {
      const r = p.getRange();
      return r.getRow() <= row && r.getLastRow() >= row && r.getColumn() <= colIndex && r.getLastColumn() >= colIndex;
    });
    protectionInfo = protections.length ? 'PROTECTED' : 'UNPROTECTED';
  } catch(e){ protectionInfo = 'PROTECTION_CHECK_ERROR'; }

  sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'DEBUG', 'inspectWriteupCell', 'value', String(val), 'formula', String(formula), 'protection', protectionInfo]);
  return { headerName, colIndex, value: val, formula, protection: protectionInfo };
}

function debug_writeWriteupLinkManual(row, pdfId){
  if (!row || isNaN(Number(row))) throw new Error('writeWriteupLinkManual requires numeric row, e.g. debug_writeWriteupLinkManual(50501, "PDF_ID")');
  if (!pdfId) throw new Error('writeWriteupLinkManual requires a pdfId param, e.g. debug_writeWriteupLinkManual(50501, "1xOw5j...")');

  const s = sh_(CONFIG.TABS.EVENTS);
  const headers = headers_(s);
  const headerName = CONFIG.COLS.WriteupPDF || 'Write-Up PDF';
  const colIndex = headers.indexOf(headerName) + 1;
  if (!colIndex || colIndex <= 0) {
    sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'ERROR', 'writeWriteupLinkManual', 'header missing', headerName]);
    throw new Error('Header missing: ' + headerName);
  }

  const url = 'https://drive.google.com/file/d/' + pdfId + '/view';
  try {
    const rich = SpreadsheetApp.newRichTextValue().setText('View PDF').setLinkUrl(url).build();
    s.getRange(Number(row), colIndex).setRichTextValue(rich);
    sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'INFO', 'writeWriteupLinkManual', 'wrote rich link', row, pdfId, url]);
  } catch (e) {
    logError('writeWriteupLinkManual', e, {row, pdfId});
    try { s.getRange(Number(row), colIndex).setValue('View PDF — ' + url); } catch(_) {}
    sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'WARN', 'writeWriteupLinkManual', 'fallback wrote plain url', row, pdfId, url]);
  }
}

function sanityCheck(){
  const items = [
    ['inferMilestoneTemplate_', typeof inferMilestoneTemplate_],
    ['templateIdForMilestoneName_', typeof templateIdForMilestoneName_],
    ['createMilestonePdf_', typeof createMilestonePdf_],
    ['createEventRecordPdf_', typeof createEventRecordPdf_],
  ];
  try { sh_(CONFIG.TABS.LOGS).appendRow([new Date(),'INFO','sanityCheck_', JSON.stringify(items)]); } catch(e){}
  Logger.log(items);
}

function timeLog_(tag, t) {
  const now = Date.now();
  if (t && typeof t.start !== 'undefined') {
    const diff = now - t.start;
    Logger.log(`TIMING: ${tag} => ${diff} ms`);
    try { sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'TIMING', tag, diff]); } catch(e){}
    return { start: now };
  } else {
    return { start: now };
  }
}

/** example wrapper to profile a single submit run */
function profile_oneSubmit(row) {
  const timer = timeLog_('profile_start');
  timeLog_('afterSubmitPoints_ start', timer);
  afterSubmitPoints_(row);
  timeLog_('afterSubmitPoints_ end', timer);
  // the functions called inside afterSubmitPoints_ should also be instrumented similarly
}

function runManualProb(){ runManualProbCheck(6); }

function runDebugRow() { dumpEventRow(6); } 

/** Dump a single event row (safe): usage dumpEventRow(7) **/
function dumpEventRow(row) {
  try {
    if (!row || isNaN(row)) {
      Logger.log('dumpEventRow: missing or invalid row number. Use listRecentEvents(n) to find rows.');
      return;
    }
    const s = sh_(CONFIG.TABS.EVENTS);
    const hdr = headers_(s);
    const lastCol = hdr.length;
    const rowVals = s.getRange(Number(row), 1, 1, lastCol).getValues()[0];
    const o = asRowObject_(hdr, rowVals);
    Logger.log('dumpEventRow row=%s -> %s', row, JSON.stringify(o));
  } catch (err) {
    Logger.log('dumpEventRow ERROR: ' + String(err));
  }
}

function debug_listFormQuestions() {
  const form = FormApp.openById(CONFIG.FORM_ID);
  form.getItems().forEach(it => {
    Logger.log('Title="%s" | Type=%s', it.getTitle(), it.getType());
  });
}

function runListRecentEvents() { listRecentEvents(20); }
/** List the last N event rows with index and a summary — usage: listRecentEvents(20) **/
function listRecentEvents(n) {
  try {
    n = Number(n) || 20;
    const s = sh_(CONFIG.TABS.EVENTS);
    const hdr = headers_(s);
    const lastRow = s.getLastRow();
    const start = Math.max(2, lastRow - n + 1);
    const rows = s.getRange(start, 1, Math.max(0, lastRow - start + 1), hdr.length).getValues();
    Logger.log('listRecentEvents: rows %s..%s (showing up to %s rows)', start, lastRow, n);
    rows.forEach((r, i) => {
      const idx = start + i;
      const summary = {
        row: idx,
        Employee: r[hdr.indexOf(CONFIG.COLS.Employee)] || '',
        IncidentDate: r[hdr.indexOf(CONFIG.COLS.IncidentDate)] || '',
        EventType: r[hdr.indexOf(CONFIG.COLS.EventType)] || '',
        Points: r[hdr.indexOf(CONFIG.COLS.Points)] || '',
        Active: r[hdr.indexOf(CONFIG.COLS.Active)] || '',
        Nullify: r[hdr.indexOf(CONFIG.COLS.Nullify)] || ''
      };
      Logger.log(JSON.stringify(summary));
    });
  } catch (err) {
    Logger.log('listRecentEvents ERROR: ' + String(err));
  }
}

/** Wrapper to manually run the probation check and also print immediate execution logs: runManualProbCheck(ROW) **/
function runManualProbCheck(row) {
  try {
    if (!row || isNaN(row)) {
      Logger.log('runManualProbCheck: missing or invalid row number. Use listRecentEvents(n) to find rows.');
      return;
    }
    Logger.log('runManualProbCheck: running checkProbationFailureForEmployee_ for row ' + row);
    const res = checkProbationFailureForEmployee_(Number(row));
    Logger.log('runManualProbCheck: result => ' + String(res));
  } catch (err) {
    Logger.log('runManualProbCheck ERROR: ' + String(err));
  }
}

function runFindLastPopulatedEventRow() {
  try {
    const s = sh_(CONFIG.TABS.EVENTS);
    const hdr = headers_(s);
    const empIdx = hdr.indexOf(CONFIG.COLS.Employee);
    const dateIdx = hdr.indexOf(CONFIG.COLS.IncidentDate);
    if (empIdx === -1 && dateIdx === -1) {
      Logger.log('runFindLastPopulatedEventRow: required headers missing (Employee/IncidentDate).');
      return;
    }

    const data = s.getDataRange().getValues();
    let foundRow = null;
    for (let r = data.length - 1; r >= 1; r--) { // skip header (r=0)
      const row = data[r];
      const emp = empIdx >= 0 ? String(row[empIdx] || '').trim() : '';
      const inc = dateIdx >= 0 ? String(row[dateIdx] || '').trim() : '';
      if (emp || inc) { foundRow = r + 1; break; }
    }

    if (!foundRow) {
      Logger.log('runFindLastPopulatedEventRow: no populated event rows found (only header?).');
      return;
    }

    Logger.log('runFindLastPopulatedEventRow: last populated row = ' + foundRow);

    // Print a small neighborhood for quick inspection
    const start = Math.max(2, foundRow - 5);
    const end = Math.min(data.length, foundRow + 5);
    Logger.log('Showing rows %s..%s', start, end);
    for (let rr = start; rr <= end; rr++) {
      const rowVals = s.getRange(rr, 1, 1, hdr.length).getValues()[0];
      const summary = {
        row: rr,
        Employee: rowVals[empIdx] || '',
        IncidentDate: rowVals[dateIdx] || '',
        EventType: rowVals[hdr.indexOf(CONFIG.COLS.EventType)] || '',
        Points: rowVals[hdr.indexOf(CONFIG.COLS.Points)] || '',
        Active: rowVals[hdr.indexOf(CONFIG.COLS.Active)] || '',
        Nullify: rowVals[hdr.indexOf(CONFIG.COLS.Nullify)] || ''
      };
      Logger.log(JSON.stringify(summary));
    }

  } catch (err) {
    Logger.log('runFindLastPopulatedEventRow ERROR: ' + String(err));
  }
}

/** dump last ~50 LOGS rows (we'll inspect checkProbationFailure entries) **/
function dumpRecentLogs() {
  try {
    const s = sh_(CONFIG.TABS.LOGS);
    const last = s.getLastRow();
    if (last < 2) { Logger.log('dumpRecentLogs: no log rows found'); return; }
    const start = Math.max(2, last - 50 + 1);
    const vals = s.getRange(start, 1, Math.max(0, last - start + 1), s.getLastColumn()).getValues();
    Logger.log('dumpRecentLogs: rows %s..%s', start, last);
    vals.forEach((r, i) => {
      // print timestamp, level, tag, message and context if present
      Logger.log(JSON.stringify({ row: start + i, ts: r[0], level: r[1], tag: r[2], msg: r[3], rest: r.slice(4) }));
    });
  } catch (err) {
    Logger.log('dumpRecentLogs ERROR: ' + String(err));
  }
}

function diag_checkCreateMilestonePdf() {
  try {
    // 1) arity of createMilestonePdf_
    const arity = (typeof createMilestonePdf_ === 'function') ? createMilestonePdf_.length : 'NOT_DEFINED';
    Logger.log('createMilestonePdf_ defined? %s, arity=%s', typeof createMilestonePdf_ === 'function', arity);

    // 2) PF template id from config
    const pfTpl = (CONFIG && CONFIG.TEMPLATES && CONFIG.TEMPLATES.PROBATION_FAILURE) ? CONFIG.TEMPLATES.PROBATION_FAILURE : null;
    Logger.log('CONFIG.TEMPLATES.PROBATION_FAILURE = %s', pfTpl);

    // 3) If infer function exists, try to preview inferred template for a sample row (pick a real milestone row number)
    if (typeof inferMilestoneTemplate_ === 'function') {
      // change rowNumber to a real milestone row you have (or keep 2..)
      const sampleRow = 6; // adjust if needed
      try {
        const inferred = inferMilestoneTemplate_(sampleRow);
        Logger.log('inferMilestoneTemplate_(%s) => %s', sampleRow, inferred);
      } catch (e) {
        Logger.log('inferMilestoneTemplate_ threw: %s', e && e.message || e);
      }
    } else {
      Logger.log('inferMilestoneTemplate_ not defined in project');
    }

  } catch (err) {
    Logger.log('diag_checkCreateMilestonePdf failed: %s', err && err.message || err);
  }
}

function diag_probation_vs_tier(triggerRow) {
  triggerRow = triggerRow || 8; // change to the actual trigger row you used
  loadPolicyFromSheet_();
  const s = sh_(CONFIG.TABS.EVENTS);
  const ctx = rowCtx_(s, triggerRow);
  const emp = ctx.get(CONFIG.COLS.Employee);
  const rolling = getRollingPointsForEmployee_(emp);
  Logger.log('diag: employee=%s raw=%s pos=%s rolling=%s effective=%s', emp, rolling.raw, rolling.pos, rolling.rolling, rolling.effective);
  Logger.log('diag: computed tier = %s (level1=%s, level2=%s, level3=%s)',
             tierForPoints_(rolling.effective), CONFIG.POLICY.MILESTONES.LEVEL_1, CONFIG.POLICY.MILESTONES.LEVEL_2, CONFIG.POLICY.MILESTONES.LEVEL_3);
}

/**
 * Scan the LOGS sheet for createMilestonePdf_ templateChoice lines and return rows
 * showing whether a forced template was provided.
 */
function scanCreateMilestoneTemplateChoiceLogs() {
  const logs = sh_(CONFIG.TABS.LOGS);
  const hdr = headers_(logs);
  const rows = logs.getDataRange().getValues();
  const out = [];
  for (let r = 1; r < rows.length; r++) { // skip header
    const row = rows[r];
    const tag = String(row[2] || '').trim();        // col C = tag
    const msg = String(row[3] || '').trim();        // col D = msg
    if (tag === 'createMilestonePdf_' && msg === 'templateChoice') {
      // rest columns follow; format we've been writing:
      // [ts, level, tag, msg, rowIndex, forced=..., inferred=..., chosen=...]
      const rowIndex = row[4];
      const forcedRaw = row[5] || '';
      const inferredRaw = row[6] || '';
      const chosenRaw = row[7] || '';
      out.push({ sheetRow: r+1, eventRow: rowIndex, forced: String(forcedRaw), inferred: String(inferredRaw), chosen: String(chosenRaw), rawRow: row });
    }
  }
  Logger.log('scanCreateMilestoneTemplateChoiceLogs: found %s entries', out.length);
  out.forEach(o => Logger.log('row %s -> eventRow=%s forced=%s inferred=%s chosen=%s', o.sheetRow, o.eventRow, o.forced, o.inferred, o.chosen));
  return out;
}

/**
 * List recent createMilestonePdf_ logs (templateChoice + wrote pdf link) for quick inspection.
 * Call with a number, e.g. listRecentCreateMilestoneLogs(50)
 */
function listRecentCreateMilestoneLogs(limit) {
  limit = Number(limit || 200);
  const logs = sh_(CONFIG.TABS.LOGS);
  const rows = logs.getDataRange().getValues();
  const hits = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const ts = row[0];
    const level = row[1];
    const tag = String(row[2] || '').trim();
    const msg = String(row[3] || '').trim();
    if (tag && tag.indexOf('createMilestonePdf_') === 0) {
      hits.push({ i: r+1, ts, level, tag, msg, rest: row.slice(4) });
    }
  }
  const tail = hits.slice(-limit);
  Logger.log('listRecentCreateMilestoneLogs: showing %s entries (most recent last)', tail.length);
  tail.forEach(h => Logger.log('%s | %s | %s | %s | %s', h.i, h.ts, h.tag, h.msg, JSON.stringify(h.rest)));
  return tail;
}

function debug_checkCreateMilestonePdfImpl() {
  try {
    const has = (typeof createMilestonePdf_ === 'function');
    Logger.log('createMilestonePdf_ defined? %s', has);
    if (has) {
      Logger.log('arity=%s', createMilestonePdf_.length);
    }
    return { defined: has, arity: has ? createMilestonePdf_.length : 0 };
  } catch (err) {
    Logger.log('error: %s', String(err));
    return null;
  }
}

function debug_dumpConfigTemplatesAndPolicy() {
  try {
    const out = {
      TEMPLATES: Object.assign({}, CONFIG.TEMPLATES || {}),
      POLICY: Object.assign({}, CONFIG.POLICY || {}),
      POSITIVE_POINTS_CAP: CONFIG.POSITIVE_POINTS_CAP || null,
      SHEET_ID: CONFIG.SHEET_ID || null
    };
    Logger.log(JSON.stringify(out, null, 2));
    return out;
  } catch (err) {
    Logger.log('error: %s', String(err));
    return null;
  }
}

function run_debug_inferForRow() {
  const ROW_TO_INSPECT = 8; // <- change this to the row you want to inspect
  return debug_inferForRow(ROW_TO_INSPECT);
}

function debug_inferForRow(row) {
  try {
    row = Number(row) || 0;
    if (!row) throw new Error('Pass a numeric row index, e.g. debug_inferForRow(8)');
    const s = sh_(CONFIG.TABS.EVENTS);
    const ctx = rowCtx_(s, row);
    Logger.log('Row %s Milestone=%s EventType=%s', row, ctx.get(CONFIG.COLS.Milestone), ctx.get(CONFIG.COLS.EventType));
    const out = (typeof inferMilestoneTemplate_ === 'function') ? inferMilestoneTemplate_(row) : { error: 'inferMilestoneTemplate_ not defined' };
    Logger.log(JSON.stringify(out, null, 2));
    return out;
  } catch (err) {
    Logger.log('error: %s', String(err));
    return null;
  }
}

function TEST_diagnoseConsequence() {
  // change row and action to match the event that produced the broken PDF
  var rowNumber = 2;
  var actionText = 'Milestone — 5 pts (2-Day Suspension)';
  var result = diagnoseAndSimulateConsequenceMerge(rowNumber, actionText);
  Logger.log(JSON.stringify(result, null, 2));
}

// Quick runner for diagnosing the Event Record template merge for a specific Events row.
// Edit rowNumber to the Events row you want to test, then run TEST_diagnoseEventPdf from the editor.
function TEST_diagnoseEventPdf(){
  var rowNumber = 43; // ← change this to the row that created the blank PDF
  // If you prefer to call the diagnoser directly, it expects: diagnoseTemplateMergeForRow(docId, row)
  var docId = CONFIG.TEMPLATES.EVENT_RECORD;
  // If you don't have diagnoseTemplateMergeForRow, use the alternate diagnoser we created earlier:
  var result = diagnoseTemplateMergeForRow(docId, rowNumber);
  Logger.log(JSON.stringify(result, null, 2));
}


/**
 * diagnoseAndSimulateConsequenceMerge(row, action)
 * - Builds the same payload createConsequencePdf_ would use.
 * - Reads template tokens from the chosen template.
 * - Returns/logs: templateId, tokensFound, payloadKeys, missingTokens, and a simulated merged preview (body header footer snippets).
 *
 * Usage:
 *   var result = diagnoseAndSimulateConsequenceMerge(43, 'Milestone — 5 pts (2-Day Suspension)');
 *   Logger.log(JSON.stringify(result, null, 2));
 */
function diagnoseAndSimulateConsequenceMerge(row, action){
  try {
    // Build payload same as createConsequencePdf_
    var events = sh_(CONFIG.TABS.EVENTS);
    var ctx = rowCtx_(events, row);
    var employee = (ctx.get(CONFIG.COLS.Employee) || '').toString();
    var dateRaw = ctx.get(CONFIG.COLS.IncidentDate) || ctx.get(CONFIG.COLS.Timestamp) || '';
    var dateObj = toDate_(dateRaw) || null;
    var policy = (ctx.get(CONFIG.COLS.RelevantPolicy) || ctx.get(CONFIG.COLS.Policy) || '').toString();
    var infraction = (ctx.get(CONFIG.COLS.Infraction) || '').toString();
    var points = (ctx.get(CONFIG.COLS.Points) || '').toString();

    var payload = {
      Employee: employee,
      IncidentDate: dateObj ? formatDate_(dateObj) : (dateRaw || ''),
      Policy: policy,
      RelevantPolicy: policy,
      Infraction: infraction,
      Points: points,
      Action: action || ''
    };

    // compact aliases (match createConsequencePdf_ behavior)
    Object.keys(Object.assign({}, payload)).forEach(function(k){
      if (!k) return;
      var alias = String(k).replace(/[\s\-_]+/g,'').replace(/[^\w]/g,'');
      if (alias && !payload.hasOwnProperty(alias)) payload[alias] = payload[k];
    });

    // also include Events header values into payload (and compact aliases)
    var hdrs = headers_(events);
    for (var i=0;i<hdrs.length;i++){
      var h = hdrs[i];
      if (!h) continue;
      if (!payload.hasOwnProperty(h)) payload[h] = ctx.get(h);
      var compact = String(h).replace(/[\s\-_]+/g,'').replace(/[^\w]/g,'');
      if (compact && !payload.hasOwnProperty(compact)) payload[compact] = payload[h];
    }

    // Determine templateId same as createConsequencePdf_
    var templateId = null;
    if (/(\b5\b|\b5pt\b|level\s*1|milestone\s*5)/i.test(action)) templateId = CONFIG.TEMPLATES.MILESTONE_5;
    else if (/(\b10\b|\b10pt\b|1-?week|one\s*week|level\s*2|milestone\s*10)/i.test(action)) templateId = CONFIG.TEMPLATES.MILESTONE_10;
    else if (/termination/i.test(action)) templateId = CONFIG.TEMPLATES.TERMINATION || CONFIG.TEMPLATES.MILESTONE_15;
    else if (/policy-?protected|policy protected/i.test(action)) templateId = CONFIG.TEMPLATES.POLICY_PROTECTED;
    else if (/probation|failure/i.test(action)) templateId = CONFIG.TEMPLATES.PROBATION_FAILURE;

    if (!templateId) {
      var errMsg = 'No template matched for action: ' + action;
      logError && logError('diagnoseAndSimulateConsequenceMerge_noTemplate', errMsg, {row: row, action: action});
      return {error: errMsg};
    }

    // Read tokens from the template doc
    var doc = DocumentApp.openById(templateId);
    var bodyText = (doc.getBody && doc.getBody().getText) ? doc.getBody().getText() : '';
    var headerText = (doc.getHeader && doc.getHeader()) ? doc.getHeader().getText() : '';
    var footerText = (doc.getFooter && doc.getFooter()) ? doc.getFooter().getText() : '';
    var docText = bodyText + '\n' + headerText + '\n' + footerText;
    var tokRe = /\{\{\s*([^}]+?)\s*\}\}/g;
    var found = []; var m;
    while ((m = tokRe.exec(docText)) !== null){
      var t = String(m[1]).trim();
      if (found.indexOf(t) === -1) found.push(t);
    }

    // Matching logic (same as mergeToPdf_): case-insensitive and space/underscore/hyphen insensitive
    function findKeyForToken(token){
      var keys = Object.keys(payload);
      var lowTok = token.toLowerCase().replace(/\s+/g,'').replace(/[_\-]/g,'');
      // exact case-insensitive
      for (var i=0;i<keys.length;i++){
        if (String(keys[i]).toLowerCase() === token.toLowerCase()) return keys[i];
      }
      // relaxed match
      for (var j=0;j<keys.length;j++){
        var k = String(keys[j]).toLowerCase().replace(/\s+/g,'').replace(/[_\-]/g,'');
        if (k === lowTok) return keys[j];
      }
      return null;
    }

    var missing = [], mapping = {};
    found.forEach(function(tok){
      var key = findKeyForToken(tok);
      if (key !== null) mapping[tok] = { key: key, value: String(payload[key]) };
      else { missing.push(tok); mapping[tok] = null; }
    });

    // Simulate replacements in text (case-insensitive token pattern)
    function simulateReplace(text){
      return text.replace(/\{\{\s*([^}]+?)\s*\}\}/g, function(full, inner){
        var tkn = String(inner||'').trim();
        var key = findKeyForToken(tkn);
        if (key !== null) return String(payload[key] === undefined || payload[key] === null ? '' : payload[key]);
        return ''; // simulate blank for missing
      });
    }

    var preview = {
      bodyPreview: simulateReplace(bodyText).slice(0, 4000),
      headerPreview: simulateReplace(headerText).slice(0, 2000),
      footerPreview: simulateReplace(footerText).slice(0, 2000)
    };

    // Write a compact diagnostic row to Logs sheet if present
    try{
      var logsName = (CONFIG && CONFIG.TABS && CONFIG.TABS.LOGS) ? CONFIG.TABS.LOGS : null;
      if (logsName) {
        var logsSh = ss_().getSheetByName(logsName);
        if (logsSh) {
          logsSh.appendRow([new Date(), 'DIAG', 'diagnoseConsequenceMerge', 'row:'+row, 'template:'+templateId, 'missing:'+JSON.stringify(missing)]);
        }
      }
    }catch(_){}

    var result = {
      row: row,
      action: action,
      templateId: templateId,
      tokensFound: found,
      payloadKeys: Object.keys(payload),
      mapping: mapping,
      missingTokens: missing,
      preview: preview
    };

    Logger.log(JSON.stringify(result, null, 2));
    return result;

  } catch(err){
    logError && logError('diagnoseAndSimulateConsequenceMerge_top', err, {row:row, action: action});
    throw err;
  }
}

// If missing, paste this minimal diagnoser (variant of earlier function)
function diagnoseTemplateMergeForRow(docId, row){
  var s = sh_(CONFIG.TABS.EVENTS);
  var hdrs = headers_(s);
  var vals = s.getRange(row, 1, 1, hdrs.length).getDisplayValues()[0];
  var rowObj = asRowObject_(hdrs, vals);
  // Build canonical data (same keys createEventRecordPdf_ uses)
  var data = {
    Employee: rowObj[CONFIG.COLS.Employee] || rowObj.Employee || '',
    IncidentDate: rowObj[CONFIG.COLS.IncidentDate] || rowObj.IncidentDate || '',
    Lead: rowObj[CONFIG.COLS.Lead] || rowObj.Lead || '',
    Policy: rowObj[CONFIG.COLS.RelevantPolicy] || rowObj.Policy || '',
    Infraction: rowObj[CONFIG.COLS.Infraction] || rowObj.Infraction || '',
    Description: rowObj[CONFIG.COLS.NotesReviewer] || rowObj.IncidentDescription || rowObj.Description || '',
    CorrectiveActions: rowObj[CONFIG.COLS.CorrectiveActions] || rowObj.CorrectiveActions || '',
    TeamMemberStatement: rowObj[CONFIG.COLS.TeamMemberStatement] || rowObj.TeamMemberStatement || '',
    Points: rowObj[CONFIG.COLS.Points] || rowObj.Points || ''
  };
  // add header keys
  hdrs.forEach(function(k){ if (k && !data.hasOwnProperty(k)) data[k] = rowObj[k]||''; });
  // create compact aliases
  Object.keys(Object.assign({}, data)).forEach(function(k){ if (!k) return; var alias=String(k).replace(/[\s\-_]+/g,'').replace(/[^\w]/g,''); if(alias&&!data.hasOwnProperty(alias)) data[alias]=data[k]; });
  // read tokens in doc
  var doc = DocumentApp.openById(docId);
  var text = (doc.getBody?doc.getBody().getText(): '') + '\n' + ((doc.getHeader&&doc.getHeader())?doc.getHeader().getText():'') + '\n' + ((doc.getFooter&&doc.getFooter())?doc.getFooter().getText():'');
  var tokRe = /\{\{\s*([^}]+?)\s*\}\}/g, found=[], m;
  while((m=tokRe.exec(text))!==null){ var t=String(m[1]).trim(); if(found.indexOf(t)===-1) found.push(t); }
  var payloadKeys = Object.keys(data);
  var missing = [];
  found.forEach(function(tok){
    var lowTok = tok.toLowerCase().replace(/\s+/g,'').replace(/[_\-]/g,'');
    var matched = payloadKeys.some(function(k){ var kn=k.toLowerCase().replace(/\s+/g,'').replace(/[_\-]/g,''); return kn===lowTok || k.toLowerCase()===tok.toLowerCase(); });
    if(!matched) missing.push(tok);
  });
  var result = { templateId: docId, tokensFound: found, payloadKeys: payloadKeys, missingTokens: missing };
  Logger.log(JSON.stringify(result, null, 2));
  try{ if (CONFIG && CONFIG.TABS && CONFIG.TABS.LOGS){ var logs = ss_().getSheetByName(CONFIG.TABS.LOGS); if (logs) logs.appendRow([new Date(),'DIAG','diagnoseTemplateMergeForRow','row:'+row,'missing:'+JSON.stringify(missing)]); } }catch(e){}
  return result;
}




function TEST_inferMilestone_row(){
  var row = 3; // e.g. 123
  var out = inferMilestoneTemplate_(row);
  Logger.log(JSON.stringify(out, null, 2));
}

function TEST_createMilestonePdf_row(){
  var id = createMilestonePdf_(3, null);
  Logger.log('pdfId=' + id);
}

// --- Which checkProbationFailureForEmployee_ is actually loaded?
function pfWhich() {
  if (typeof checkProbationFailureForEmployee_ !== 'function') { Logger.log('PF: not defined'); return; }
  var src = checkProbationFailureForEmployee_.toString();
  var tag =
    src.indexOf('probation flag not active') > -1 ? 'STRICT_FLAG_VERSION' :
    src.indexOf('probation not active from tier2') > -1 ? 'TIER2_ONLY_VERSION' :
    'UNKNOWN_VERSION';
  Logger.log('PF version = ' + tag + ' (chars=' + src.length + ')');
  Logger.log(src.split('\n')[0]);
}

// --- Dry-run PF on a row (no side effects)
function pfDryRun(row) {
  var row = 7;
  var res = checkProbationFailureForEmployee_(row);
  Logger.log('PF dry-run row='+row+' -> ' + JSON.stringify(res));
}

// --- Inspect what the *flag* detector thinks, and why
function probeProbationFlag(employee, beforeRow) {
  var employee = 'Brown, KiChuana';
  if (!employee) { Logger.log('need employee'); return; }
  var s = sh_(CONFIG.TABS.EVENTS), hdr = headers_(s), map = headerIndexMap_(s);
  var cEmp=map[CONFIG.COLS.Employee], cMil=map[CONFIG.COLS.Milestone],
      cWhen=map[CONFIG.COLS.MilestoneDate]||map[CONFIG.COLS.Timestamp],
      cEff=map[CONFIG.COLS.PointsRollingEffective]||0,
      cRoll=map[CONFIG.COLS.PointsRolling]||0,
      cProb=map[CONFIG.COLS.Probation_Active]||hdr.indexOf('Probation_Active')+1;

  var last = Math.min(s.getLastRow(), beforeRow ? beforeRow-1 : 9e9);
  for (var r=2; r<=last; r++){
    var emp = s.getRange(r, cEmp).getDisplayValue().trim();
    if (emp !== employee) continue;
    var label = cMil ? String(s.getRange(r, cMil).getDisplayValue()).trim() : '';
    if (!label) continue;
    var eff = cEff ? Number(s.getRange(r, cEff).getValue()) :
             (cRoll ? Number(s.getRange(r, cRoll).getValue()) : NaN);
    var probCell = cProb ? s.getRange(r, cProb).getValue() : '';
    var probTruthy = (probCell === true || probCell === 1 ||
                      (typeof probCell === 'string' && probCell.trim().toLowerCase() === 'true') ||
                      String(probCell) === '1');

    var mapped = milestoneTextToTier_(label);
    var ptsTier = isNaN(eff) ? null : tierForPoints_(eff);

    Logger.log(JSON.stringify({
      row:r, label:label, mappedTier:mapped, ptsTier:ptsTier,
      eff:isNaN(eff)?null:eff, Probation_Active: probCell, probTruthy: probTruthy
    }));
  }
}

// --- Quick read of the strict flag function itself
function flagWhich() {
  if (typeof isProbationActiveForEmployee_ !== 'function') { Logger.log('flag fn missing'); return; }
  var src = isProbationActiveForEmployee_.toString();
  Logger.log('flag fn chars=' + src.length + '; requires Probation_Active? ' + (src.indexOf('Probation_Active')>-1));
}


function auditConfigVsSheetHeaders(){
  var s = sh_(CONFIG.TABS.EVENTS);
  var hdrs = headers_(s);
  var hdrSet = {};
  hdrs.forEach(function(h){ hdrSet[String(h||'').trim()] = true; });

  var issues = [];
  Object.keys(CONFIG.COLS).forEach(function(k){
    var name = CONFIG.COLS[k];
    if (!name) {
      issues.push('CONFIG.COLS.'+k+' is undefined/null');
      return;
    }
    if (!hdrSet[name]) {
      // Try to suggest a close match (loose compare)
      var want = norm_(name);
      var guess = hdrs
        .map(function(h){ return {h:h, score: sim_(want, norm_(h))}; })
        .sort(function(a,b){ return b.score - a.score; })[0];
      var hint = guess && guess.score >= 0.6 ? ('; did you mean "'+guess.h+'"?') : '';
      issues.push('Missing header "'+name+'" for CONFIG.COLS.'+k+hint);
    }
  });

  if (issues.length){
    Logger.log('CONFIG/Headers audit:\n' + issues.join('\n'));
  } else {
    Logger.log('CONFIG/Headers audit: ✅ all mapped headers exist');
  }

  // helpers
  function norm_(s){ return String(s||'').toLowerCase().replace(/[^a-z0-9]/g,''); }
  function sim_(a,b){
    // quick Jaccard on 3-grams for rough suggestion
    function grams(x){ var g=new Set(); for (var i=0;i<x.length-2;i++) g.add(x.slice(i,i+3)); return g; }
    var A=grams(a), B=grams(b);
    var inter=0; A.forEach(function(x){ if (B.has(x)) inter++; });
    var uni = new Set([...A,...B]).size || 1;
    return inter/uni;
  }
}

var __rowCtxOriginal = null;

function enableCtxGetGuard(on){
  if (on && !__rowCtxOriginal){
    __rowCtxOriginal = rowCtx_;
    rowCtx_ = function(sh, row){
      var ctx = __rowCtxOriginal(sh, row);
      var map = headerIndexMap_(sh);
      var origGet = ctx.get;
      ctx.get = function(h){
        if (!h){
          var where = (new Error('ctx.get(undefined)')).stack;
          try{ logError && logError('CTX_GET_UNDEFINED', { row: row, stack: String(where).slice(0,500) }); }catch(_){}
          throw new Error('ctx.get(undefined)');
        }
        if (!map[h]){
          var where2 = (new Error('ctx.get(missing_header:'+h+')')).stack;
          try{ logError && logError('CTX_GET_MISSING', { row: row, header: h, stack: String(where2).slice(0,500) }); }catch(_){}
          // still throw so you see it during tests
          throw new Error('ctx header not found: ' + h);
        }
        return origGet.call(ctx, h);
      };
      return ctx;
    };
    Logger.log('Ctx guard: ENABLED');
  } else if (!on && __rowCtxOriginal){
    rowCtx_ = __rowCtxOriginal;
    __rowCtxOriginal = null;
    Logger.log('Ctx guard: DISABLED');
  }
}


function scanMilestoneRowsForCtxFaults(){
  var s = sh_(CONFIG.TABS.EVENTS);
  var map = headerIndexMap_(s);
  var cMil = map[CONFIG.COLS.Milestone];
  if (!cMil){ Logger.log('No Milestone column found.'); return; }

  var last = s.getLastRow();
  var values = s.getRange(2,1, Math.max(0,last-1), Math.max(1, s.getLastColumn())).getValues();
  var bad = [];

  for (var i=0;i<values.length;i++){
    var r = i+2;
    var label = String(values[i][cMil-1]||'').trim();
    if (!label) continue;
    try{
      var ctx = rowCtx_(s, r);
      var ok = isMilestoneRowActive_(ctx); // will throw if it uses bad headers
    }catch(e){
      bad.push({row:r, label:label, error:String(e).slice(0,200)});
    }
  }
  if (bad.length){
    Logger.log('Active-check faults:\n' + bad.map(function(b){ return 'row '+b.row+' label="'+b.label+'" -> '+b.error; }).join('\n'));
  }else{
    Logger.log('Active-check sweep: ✅ no ctx header faults');
  }
}


function test_FindLatest_NoPointsFallback(){
  var emps = uniqueEmployees_();
  emps.forEach(function(emp){
    var latest = findLatestMilestoneForEmployee_(emp);
    if (!latest) return;
    if (latest.labelAmbiguous){
      Logger.log('AMBIGUOUS label for '+emp+' at row '+latest.row+': "'+latest.milestone+'"');
    }
    if (latest.tier === null){
      Logger.log('NO TIER (label did not map) for '+emp+' -> "'+latest.milestone+'"');
    }
  });
}

function uniqueEmployees_(){
  var s = sh_(CONFIG.TABS.EVENTS);
  var map = headerIndexMap_(s);
  var cEmp = map[CONFIG.COLS.Employee];
  var last = s.getLastRow();
  if (!cEmp || last<2) return [];
  var vals = s.getRange(2, cEmp, last-1, 1).getValues().map(function(r){ return String(r[0]||'').trim(); });
  var set = {}; vals.forEach(function(v){ if(v) set[v]=1; });
  return Object.keys(set);
}

function testDocsApiAvailable() {
  try {
    const doc = DocumentApp.create('docs-api-test-' + Date.now());
    const docId = doc.getId();
    // Try to read document metadata via the Docs advanced service
    const meta = Docs.Documents.get(docId);
    Logger.log('Docs API OK: pageSize width = %s', meta.documentStyle && meta.documentStyle.pageSize && meta.documentStyle.pageSize.width && meta.documentStyle.pageSize.width.magnitude);
    DriveApp.getFileById(docId).setTrashed(true);
  } catch (e) {
    Logger.log('Docs API not available or permission denied: %s', String(e));
  }
}