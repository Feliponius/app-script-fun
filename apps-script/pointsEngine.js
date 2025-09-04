/* =========================
 *  pointsEngine.gs  (cleaned)
 *  - Label-first milestone finder + scan + probation checker
 * ========================= */


/** True if there is an existing milestone row for employee with PendingStatus = 'Pending'. */
function hasPendingMilestoneForEmployee_(employee){
  try{
    if (!employee) return false;
    var s = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(s);
    var last = s.getLastRow();
    if (last < 2) return false;

    var cEmp  = map[CONFIG.COLS.Employee];
    var cMil  = map[CONFIG.COLS.Milestone] || 0;
    var cPend = map[CONFIG.COLS.PendingStatus] || map['Pending Status'] || 0;
    if (!cEmp || !cMil || !cPend) return false;

    var maxCol = Math.max(cEmp, cMil, cPend);
    var values = s.getRange(2, 1, last - 1, maxCol).getValues();
    for (var i = 0; i < values.length; i++){
      var row = values[i];
      if (String(row[cEmp-1] || '').toString().trim() !== String(employee).toString().trim()) continue;
      if (!String(row[cMil-1] || '').trim()) continue;
      var pend = String(row[cPend-1] || '').toLowerCase();
      if (pend === 'pending') return true;
    }
    return false;
  }catch(e){
    try{ logError && logError('hasPendingMilestoneForEmployee_', e, {employee:employee}); }catch(_){}
    return false;
  }
}

/**
 * Append a milestone row based on an existing event row and a label.
 * - Does NOT write PointsRolling(Effective) (sheet-owned).
 * - Leaves Consequence Director blank (no auto-claim).
 */
function appendMilestoneRow_(triggerRow, milestoneName, rollingPoints){
  return withBackoff_('appendMilestoneRow_', function(){
    try{
      var s = sh_(CONFIG.TABS.EVENTS);
      if (!triggerRow || triggerRow < 2) throw new Error('appendMilestoneRow_ missing triggerRow');

      // helpers
      var hdrs = headers_(s);
      function reload(){ hdrs = headers_(s); }
      function findHeader(name){ var idx = hdrs.indexOf(name); return idx >= 0 ? (idx+1) : 0; }
      function ensureCol(name, aliases){
        var c = findHeader(name);
        if (c) return c;
        (aliases||[]).some(function(a){ c = findHeader(a); return !!c; });
        if (c) return c;
        // create at end
        var newCol = s.getLastColumn() + 1;
        s.getRange(1, newCol).setValue(name);
        reload();
        return newCol;
      }

      // Make sure downstream-required columns exist (prevent "header not found")
      var cTimestamp       = ensureCol(CONFIG.COLS.Timestamp, ['Timestamp']);
      var cEmployee        = ensureCol(CONFIG.COLS.Employee, ['Employee']);
      var cIncidentDate    = ensureCol(CONFIG.COLS.IncidentDate, ['IncidentDate','Incident Date','Date']);
      var cEventType       = ensureCol(CONFIG.COLS.EventType, ['EventType','Event Type']);
      var cMilestone       = ensureCol(CONFIG.COLS.Milestone, ['Milestone']);
      var cMilestoneDate   = ensureCol(CONFIG.COLS.MilestoneDate, ['MilestoneDate','Milestone Date']);
      var cLinkedEventId   = ensureCol(CONFIG.COLS.Linked_Event_ID || CONFIG.COLS.LinkedEventID || 'Linked Event ID', ['LinkedEventID','Linked Event Row','LinkedEventRow']);
      var cPendingStatus   = ensureCol(CONFIG.COLS.PendingStatus || 'Pending Status', ['PendingStatus']);
      var cNullify         = ensureCol(CONFIG.COLS.Nullify || 'Nullify', ['Nullified','Null']);
      var cPdfLink         = ensureCol(CONFIG.COLS.PdfLink || 'PdfLink', ['Write-Up PDF','PDF Link']);
      // optional clears (no creation if you don’t want them forced)
      var cConsequencePdf  = findHeader(CONFIG.COLS.ConsequencePdf || 'Consequence PDF');
      var cLead            = findHeader(CONFIG.COLS.Lead || 'Lead');
      var cConsequenceDir  = findHeader(CONFIG.COLS.ConsequenceDirector || 'Consequence Director');

      // read trigger context
      var triggerCtx = rowCtx_(s, triggerRow);
      var employee   = String(triggerCtx.get(CONFIG.COLS.Employee) || '').trim();
      if (!employee) return null;
      var incDate    = triggerCtx.get(CONFIG.COLS.IncidentDate) || new Date();

      // build a blank row object keyed by header text
      var rowObj = {};
      for (var i=0;i<hdrs.length;i++){ if (hdrs[i]) rowObj[hdrs[i]] = ''; }

      // fill milestone fields
      rowObj[hdrs[cTimestamp-1]]     = new Date();
      rowObj[hdrs[cEmployee-1]]      = employee;
      rowObj[hdrs[cIncidentDate-1]]  = incDate;
      rowObj[hdrs[cEventType-1]]     = 'Milestone';
      rowObj[hdrs[cMilestone-1]]     = milestoneName || 'Milestone';
      rowObj[hdrs[cMilestoneDate-1]] = new Date();
      rowObj[hdrs[cLinkedEventId-1]] = triggerRow;

      // explicit blanks/safety
      if (cLead)           rowObj[hdrs[cLead-1]]          = '';
      if (cConsequenceDir) rowObj[hdrs[cConsequenceDir-1]]= '';
      if (cPdfLink)        rowObj[hdrs[cPdfLink-1]]       = '';
      if (cConsequencePdf) rowObj[hdrs[cConsequencePdf-1]]= '';
      // mark pending + ensure Nullify exists and is false
      rowObj[hdrs[cPendingStatus-1]] = 'Pending';
      rowObj[hdrs[cNullify-1]]       = false;

      // append
      var newRow = appendEventsRow_(rowObj, {}); // helper returns new row index
      if (!newRow) throw new Error('appendEventsRow_ failed');

      // header hygiene (if you rely on these later)
      try{ ensureRollingEffectiveHeader_(s); }catch(_){}

      // audit
      try{
        logAudit && logAudit(
          'milestone_create',
          'Milestone appended (Pending)',
          newRow,
          { byRow: triggerRow, employee: employee, name: rowObj[hdrs[cMilestone-1]], effectiveAtTrigger: Number(rollingPoints||0) }
        );
      }catch(_){}

      // notify
      try{
        if (typeof notifyLeadersMilestonePending_ === 'function'){
          notifyLeadersMilestonePending_(newRow);
        }
      }catch(nerr){
        try{ logError && logError('appendMilestoneRow_notify_pending', nerr, {row:newRow}); }catch(_){}
      }

      return newRow;

    }catch(err){
      try{ logError && logError('appendMilestoneRow_top', err, {triggerRow:triggerRow, milestoneName:milestoneName, rollingPoints:rollingPoints}); }catch(_){}
      throw err;
    }
  }, 4, 250);
}




function cancelPendingMilestonesForSource_(eventRow){
  try{
    var s = sh_(CONFIG.TABS.EVENTS);
    var hdrs = headers_(s) || [];
    var last = s.getLastRow(); if (last < 2) return 0;

    // normalize header match
    function norm(h){ return String(h||'').toLowerCase().replace(/[\s_\-]+/g,''); }
    function findCol(cands){
      cands = [].concat(cands||[]).filter(Boolean);
      var all = hdrs.map((h,i)=>({i:i+1, raw:h, k:norm(h)}));
      for (var c=0;c<cands.length;c++){
        var want = norm(cands[c]);
        for (var j=0;j<all.length;j++) if (all[j].k === want) return all[j].i;
      }
      return 0;
    }

    // columns (tolerant)
    var cEvt  = findCol([ (CONFIG.COLS&&CONFIG.COLS.EventType)||'EventType', 'Event Type' ]);
    var cPend = findCol([ (CONFIG.COLS&&CONFIG.COLS.PendingStatus)||'Pending Status', 'PendingStatus' ]);
    var cNull = findCol([ (CONFIG.COLS&&CONFIG.COLS.Nullify)||'Nullify' ]);
    var cPdf  = findCol([ (CONFIG.COLS&&CONFIG.COLS.PdfLink)||'Write-Up PDF', 'WriteUpPDF','PDF Link','PdfLink' ]);
    // link can be either of these
    var cLink = findCol([
      (CONFIG.COLS&&CONFIG.COLS.Linked_Event_ID)||'Linked_Event_ID',
      (CONFIG.COLS&&CONFIG.COLS.LinkedEventID)||'LinkedEventID',
      (CONFIG.COLS&&CONFIG.COLS.LinkedEventRow)||'Linked Event Row',
      'Linked Event ID'
    ]);

    // observability
    logInfo_ && logInfo_('cancelPendingMilestonesForSource_bindings', {
      eventRow: eventRow, cEvt: cEvt, cPend: cPend, cNull: cNull, cPdf: cPdf, cLink: cLink,
      headersPreview: hdrs.slice(0,30)
    });

    if (!cEvt || !cLink || !cNull) return 0;

    var maxC = Math.max(cEvt, cPend, cNull, cPdf, cLink);
    var vals = s.getRange(2, 1, last-1, maxC).getValues();

    var n = 0;
    for (var i = 0; i < vals.length; i++){
      var rowN = i + 2;
      if (String(vals[i][cEvt-1]||'').trim().toLowerCase() !== 'milestone') continue;

      var linkVal = Number(vals[i][cLink-1]); // handles "90" or 90
      if (linkVal !== Number(eventRow)) continue;

      var hasPdf = !!(cPdf && String(vals[i][cPdf-1]||'').trim());
      var status = hasPdf ? 'Voided (graced trigger)' : 'Cancelled (graced trigger)';

      if (cPend) s.getRange(rowN, cPend).setValue(status);
      s.getRange(rowN, cNull).setValue(true);
      n++;
    }

    logInfo_ && logInfo_('cancelPendingMilestonesForSource_result', { eventRow: eventRow, affected: n });
    if (n && typeof logAudit === 'function') try{ logAudit('GraceCascade','Cancelled/Voided linked milestones', eventRow, {count:n}); }catch(_){}
    return n;

  } catch(e){
    logError && logError('cancelPendingMilestonesForSource_', e, { eventRow: eventRow });
    return 0;
  }
}




/**
 * scanAndHandleMilestones_
 * - Label-first logic, aborts when prior milestone label is ambiguous (logs hint).
 * - Uses checkProbationFailureForEmployee_ to decide Probation Failure (strict Tier-2 source).
 */
function scanAndHandleMilestones_(row){
  return withLocks_('scanAndHandleMilestones', function(){
    try{
      if (typeof loadPolicyFromSheet_ === 'function') { try{ loadPolicyFromSheet_(); }catch(_){ } }

      var s = sh_(CONFIG.TABS.EVENTS);
      var map = headerIndexMap_(s);
      var ctx = null;
      try { ctx = rowCtx_(s, row); } catch(_) { ctx = null; }

      var employee = String(ctx ? (ctx.get(CONFIG.COLS.Employee) || '') : '').trim();
      if (!employee) { logInfo_ && logInfo_('scanAndHandleMilestones_', 'no employee for row ' + row); return; }

      SpreadsheetApp.flush(); Utilities.sleep(200);

      // Determine effective points (prefer Effective -> Rolling -> derive)
      var effCol = map[CONFIG.COLS.PointsRollingEffective] || 0;
      var rollCol = map[CONFIG.COLS.PointsRolling] || 0;
      var effective = NaN;
      try{
        if (effCol) {
          var v = s.getRange(row, effCol).getValue();
          if (v !== '' && v !== null && !isNaN(v)) effective = Number(v);
        }
        if (!isFinite(effective) && rollCol) {
          var v2 = s.getRange(row, rollCol).getValue();
          if (v2 !== '' && v2 !== null && !isNaN(v2)) effective = Number(v2);
        }
      }catch(_){}

      if (!isFinite(effective)) {
        var prior = 0;
        try { prior = Number((typeof getRollingPointsForEmployee_ === 'function') ? getRollingPointsForEmployee_(employee) : 0) || 0; } catch(_){ prior = 0; }
        var delta = _derivePointsDeltaForRow_(s, map, row);
        effective = Math.max(0, prior + Number(delta || 0));
        logInfo_ && logInfo_('scanAndHandleMilestones_', 'derived effective', { row:row, employee:employee, prior:prior, delta:delta, effective:effective });
      }

      // Skip if a milestone is already pending
      if (typeof hasPendingMilestoneForEmployee_ === 'function' && hasPendingMilestoneForEmployee_(employee)) {
        logInfo_ && logInfo_('scanAndHandleMilestones_', 'skipping: pending milestone already exists for ' + employee, { row: row });
        return;
      }

      var latest = (typeof findLatestMilestoneForEmployee_ === 'function') ? findLatestMilestoneForEmployee_(employee, { beforeRow: row }) : null;
      logInfo_ && logInfo_('scanAndHandleMilestones_latest', {
        row: row, employee: employee,
        latest: latest && { row: latest.row, milestone: latest.milestone, tier: latest.tier, labelAmbiguous: latest.labelAmbiguous, active: latest.active, pointsRolling: latest.pointsRolling }
      });

      // Abort if latest label is ambiguous
      if (latest && latest.labelAmbiguous) {
        try {
          var hints = (typeof getMilestonePatternHints_ === 'function') ? getMilestonePatternHints_() : [];
          var hintText = Array.isArray(hints) && hints.length
            ? hints.map(function(h){ return 'Tier' + (h.tier||'?') + ': ' + ((h.examples||[]).slice(0,3).join(', ')||'<examples>'); }).join(' | ')
            : '';
          logError && logError('scanAndHandleMilestones_labelAmbiguous', {
            row: row, employee: employee, foundLabel: latest.milestone, latestRow: latest.row, hintText: hintText
          });
          try {
            if (CONFIG && CONFIG.TABS && CONFIG.TABS.LOGS) {
              var logsSh = sh_(CONFIG.TABS.LOGS);
              if (logsSh && typeof logsSh.appendRow === 'function') {
                logsSh.appendRow([new Date(), 'WARN', 'scanAndHandleMilestones_labelAmbiguous', 'row:'+row, employee, String(latest.milestone||'').slice(0,200), hintText]);
              }
            }
          } catch(_){}
        } catch(_){}
        return;
      }

      // ---- PF gate (run BEFORE tier progression)
      try {
        var pfCheck = (typeof checkProbationFailureForEmployee_ === 'function')
          ? checkProbationFailureForEmployee_(row)
          : { shouldTrigger: false };

        if (pfCheck && pfCheck.shouldTrigger) {
          var pfLabel = (CONFIG && CONFIG.POLICY && CONFIG.POLICY.MILESTONE_NAMES && CONFIG.POLICY.MILESTONE_NAMES.PROBATION_FAILURE) || 'Probation Failure';
          var pfRow = null;
          try {
            pfRow = appendMilestoneRow_(row, pfLabel, pfCheck.eff || effective);
          } catch (apErr) {
            logError && logError('scanAndHandleMilestones_appendProbationFailureErr', apErr, { row: row, employee: employee });
            pfRow = null;
          }
          if (pfRow) {
            // >>> ADD: set flags strictly for this label (PF = TRUE, do NOT start probation)
            try { if (typeof setProbationFlagsForLabel_ === 'function') setProbationFlagsForLabel_(CONFIG.TABS.EVENTS, pfRow, pfLabel); } catch(_){}
            logAudit && logAudit('Probation Failure appended', row, {
              employee: employee, milestoneRow: pfRow, effective: pfCheck.eff || effective, threshold: pfCheck.threshold
            });
          } else {
            logInfo_ && logInfo_('scanAndHandleMilestones_', 'probation failure append returned null', { row: row, employee: employee });
          }
          return; // short-circuit after PF
        } else {
          logInfo_ && logInfo_('scanAndHandleMilestones_', {
            row: row, employee: employee,
            pfCheck: pfCheck && { shouldTrigger: pfCheck.shouldTrigger, reason: pfCheck.reason, latest: pfCheck.latest && { row: pfCheck.latest.row, tier: pfCheck.latest.tier, active: pfCheck.latest.active } }
          });
        }
      } catch (epf) {
        logError && logError('scanAndHandleMilestones_pfGateErr', epf, { row: row, employee: employee });
      }

      var currentTier = tierForPoints_(effective);
      if (!currentTier) { logInfo_ && logInfo_('scanAndHandleMilestones_', 'no tier for effective=' + effective + ' row=' + row); return; }

      // Skip if an active same-tier milestone exists
      var alreadyHasActiveSameTier = false;
      try {
        if (typeof hasActiveMilestoneOfTier_ === 'function')
          alreadyHasActiveSameTier = !!hasActiveMilestoneOfTier_(employee, currentTier, { beforeRow: row });
        else if (latest && typeof latest.tier !== 'undefined' && latest.tier !== null)
          alreadyHasActiveSameTier = (Number(latest.tier) === Number(currentTier)) && !!latest.active;
      } catch (e) {
        try { logError && logError('scanAndHandleMilestones_hasActiveCheckErr', e, { row: row, employee: employee }); } catch(_){}
        alreadyHasActiveSameTier = false;
      }
      if (alreadyHasActiveSameTier) {
        logInfo_ && logInfo_('scanAndHandleMilestones_skip_existingActive', { row: row, employee: employee, tier: currentTier });
        return;
      }

      // Prior achieved tier from labeled milestone only
      var achievedTier = (latest && latest.tier != null) ? Number(latest.tier || 0) : 0;

      if (currentTier <= achievedTier) {
        logInfo_ && logInfo_('scanAndHandleMilestones_', 'no progression; currentTier=' + currentTier + ' achieved=' + achievedTier);
        return;
      }

      var label = 'Level ' + currentTier;
      var names = (CONFIG && CONFIG.POLICY) ? CONFIG.POLICY.MILESTONE_NAMES : null;
      if (names){
        if (currentTier === 1 && names.LEVEL_1) label = names.LEVEL_1;
        if (currentTier === 2 && names.LEVEL_2) label = names.LEVEL_2;
        if (currentTier === 3 && names.LEVEL_3) label = names.LEVEL_3;
      }

      logInfo_ && logInfo_('scanAndHandleMilestones_', 'triggering milestone', { row: row, employee: employee, label: label, effective: effective, currentTier: currentTier });

      // Normal ladder append (exactly one)
      try {
        var newRow = null;
        var lvl3 = (CONFIG && CONFIG.POLICY && CONFIG.POLICY.MILESTONES && CONFIG.POLICY.MILESTONES.LEVEL_3) || 15;
        var lvl2 = (CONFIG && CONFIG.POLICY && CONFIG.POLICY.MILESTONES && CONFIG.POLICY.MILESTONES.LEVEL_2) || 10;
        var lvl1 = (CONFIG && CONFIG.POLICY && CONFIG.POLICY.MILESTONES && CONFIG.POLICY.MILESTONES.LEVEL_1) || 5;

        if (Number(effective) >= Number(lvl3)) newRow = appendMilestoneRow_(row, (names && names.LEVEL_3) || 'MILESTONE_15', effective);
        else if (Number(effective) >= Number(lvl2)) newRow = appendMilestoneRow_(row, (names && names.LEVEL_2) || 'MILESTONE_10', effective);
        else if (Number(effective) >= Number(lvl1)) newRow = appendMilestoneRow_(row, (names && names.LEVEL_1) || 'MILESTONE_5', effective);

        if (!newRow) {
          logInfo_ && logInfo_('scanAndHandleMilestones_', 'appendMilestoneRow_ returned null — not creating milestone', { row: row, employee: employee });
          return;
        }

        // >>> ADD: set probation flags based on the chosen label (only Tier-2 “probation” starts probation)
        try { if (typeof setProbationFlagsForLabel_ === 'function') setProbationFlagsForLabel_(CONFIG.TABS.EVENTS, newRow, label); } catch(_){}

        logAudit && logAudit('Milestone appended', row, { milestoneRow: newRow, employee: employee, effective: effective });
        return;

      } catch (apErr2) {
        logError && logError('scanAndHandleMilestones_appendErr2', apErr2, { row: row, employee: employee });
        return;
      }

    } catch (e){
      try { logError('scanAndHandleMilestones_', e, { row: row }); } catch(_){}
    }
  });
}




/**
 * checkProbationFailureForEmployee_(triggerRow)
 * - Requires an ACTIVE Tier-2 milestone BEFORE triggerRow to consider triggering Probation Failure
 * - Returns { shouldTrigger: bool, eff, threshold, reason, latest }
 */
function checkProbationFailureForEmployee_(triggerRow){
  try {
    var s = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(s);

    // Employee @ triggerRow
    var cEmp = map[CONFIG.COLS.Employee];
    if (!cEmp) {
      logInfo_ && logInfo_('checkProbationFailureForEmployee_', { triggerRow: triggerRow, reason: 'no Employee column' });
      return { shouldTrigger: false, reason: 'no employee col' };
    }
    var employee = String(s.getRange(triggerRow, cEmp).getDisplayValue() || '').trim();
    if (!employee) {
      logInfo_ && logInfo_('checkProbationFailureForEmployee_', { triggerRow: triggerRow, reason: 'no employee' });
      return { shouldTrigger: false, reason: 'no employee' };
    }

    // Effective points at the trigger row
    var eff = 0;
    var cEff  = map[CONFIG.COLS.PointsRollingEffective] || 0;
    var cRoll = map[CONFIG.COLS.PointsRolling] || 0;
    try {
      if (cEff) {
        var v = s.getRange(triggerRow, cEff).getValue();
        if (v !== '' && v !== null && !isNaN(v)) eff = Number(v);
      }
      if (!isFinite(eff) || eff === 0) {
        if (cRoll) {
          var v2 = s.getRange(triggerRow, cRoll).getValue();
          if (v2 !== '' && v2 !== null && !isNaN(v2)) eff = Number(v2);
        }
      }
      if (!isFinite(eff)) eff = 0;
    } catch(_) { eff = 0; }

    // Threshold (default 14)
    var pfThreshold = Number(
      (CONFIG && CONFIG.POLICY && CONFIG.POLICY.PROBATION_FAILURE) ||
      14
    );

    // STRICT: requires explicit probation flag to be active
    var probActive = false;
    try { probActive = !!isProbationActiveForEmployee_(employee, new Date()); } catch(_){ probActive = false; }

    // Optional: latest milestone purely for LOG CONTEXT (no logic depends on it)
    var latest = null;
    try { if (typeof findLatestMilestoneForEmployee_ === 'function') latest = findLatestMilestoneForEmployee_(employee, { beforeRow: triggerRow }); } catch(_){}

    if (!probActive) {
      logInfo_ && logInfo_('checkProbationFailureForEmployee_', {
        triggerRow: triggerRow, employee: employee, eff: eff, threshold: pfThreshold,
        outcome: 'no trigger (probation flag not active)',
        latest: latest && { row: latest.row, milestone: latest.milestone, tier: latest.tier, active: latest.active }
      });
      return { shouldTrigger: false, eff: eff, threshold: pfThreshold, reason: 'probation flag not active', latest: latest };
    }

    if (Number(eff) >= Number(pfThreshold)) {
      logInfo_ && logInfo_('checkProbationFailureForEmployee_', {
        triggerRow: triggerRow, employee: employee, eff: eff, threshold: pfThreshold,
        outcome: 'trigger',
        latest: latest && { row: latest.row, milestone: latest.milestone, tier: latest.tier, active: latest.active }
      });
      return { shouldTrigger: true, eff: eff, threshold: pfThreshold, reason: 'trigger', latest: latest };
    }

    logInfo_ && logInfo_('checkProbationFailureForEmployee_', {
      triggerRow: triggerRow, employee: employee, eff: eff, threshold: pfThreshold,
      outcome: 'threshold not met',
      latest: latest && { row: latest.row, milestone: latest.milestone, tier: latest.tier, active: latest.active }
    });
    return { shouldTrigger: false, eff: eff, threshold: pfThreshold, reason: 'threshold not met', latest: latest };

  } catch (err) {
    try { logError && logError('checkProbationFailureForEmployee_topErr', err, { triggerRow: triggerRow }); } catch(_){}
    return { shouldTrigger: false, reason: 'error' };
  }
}

/** =========================
 * GRACE SHIMS — pointsEngine delegates to grace.gs
 * Keep signatures intact so legacy calls don't crash.
 * ========================= */

// Forwarder to grace.gs's applyGraceToEvent
function applyGraceForEventRow(eventRow, directorName, reason, opts){
  try{
    if (typeof applyGraceToEvent !== 'function') return false;
    var eventsTab = (CONFIG && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
    var sh = SpreadsheetApp.getActive().getSheetByName(eventsTab);
    if (!sh || !eventRow || eventRow < 2) return false;

    var h = (typeof headerMap === 'function') ? headerMap(sh) : headerIndexMap_(sh);
    var empCol = h['Employee'] || (CONFIG.COLS && h[CONFIG.COLS.Employee]) || 0;
    var employee = empCol ? String(sh.getRange(eventRow, empCol).getDisplayValue()||'').trim() : '';

    // requiredTier is optional; fall back to 'minor'
    var reqTier = 'minor';
    var reqCol = h['Required Tier'] || h['Required Grace Tier'] || 0;
    if (reqCol){
      var v = sh.getRange(eventRow, reqCol).getDisplayValue();
      reqTier = normalizeTier ? normalizeTier(v) : (String(v||'').toLowerCase());
    }

    return applyGraceToEvent({ sheet: sh, rowIndex: eventRow, employee: employee, requiredTier: reqTier, headers: h }, reqTier);
  }catch(_){ return false; }
}

// Legacy consumer (employee+tier) — deprecated; NO-OP that reports false
function consumeCredit(employee, tierName, directorName, appliedEventRow, reason, opts){
  // pointsEngine no longer owns consumption; use grace.gs selection/consumption instead
  try{ Logger.log('consumeCredit (pointsEngine) is deprecated; use grace.gs'); }catch(_){}
  return null;
}

// Universal credit helpers — delegate to grace.gs if available, else benign defaults
function countUniversalCredits_(employee){
  try{
    if (typeof listAvailableCreditsForEmployee === 'function' && typeof isUniversalCreditType === 'function'){
      var rows = listAvailableCreditsForEmployee(employee) || [];
      var total = 0, used = [];
      for (var i=0;i<rows.length;i++){
        if (isUniversalCreditType(rows[i].type)) { total += Number(rows[i].points||0); used.push(rows[i].row); }
      }
      return { count: used.length, rows: used };
    }
  }catch(_){}
  return { count: 0, rows: [] };
}

function countAvailableCredits_(employee, tierName){
  try{
    if (typeof listAvailableCreditsForEmployee === 'function'){
      var need = (typeof normalizeTier === 'function') ? normalizeTier(tierName) : String(tierName||'').toLowerCase();
      var rows = listAvailableCreditsForEmployee(employee) || [];
      var out = [];
      for (var i=0;i<rows.length;i++){
        var have = (typeof normalizeTier === 'function') ? normalizeTier(rows[i].tier || rows[i].type) : String(rows[i].tier||rows[i].type||'').toLowerCase();
        var ok = (typeof isTierAtLeast === 'function') ? isTierAtLeast(have, need) : (have === need);
        if (ok) out.push(rows[i].row);
      }
      return { count: out.length, rows: out };
    }
  }catch(_){}
  return { count: 0, rows: [] };
}

// Build Grace event rows — handled in grace.gs now; return an empty row shape
function buildGraceEventRow(events, originalRow, employee, directorName, reason, tierName){
  // Return a blank row (same width) to avoid caller errors if still used
  var cols = events ? events.getLastColumn() : 0;
  var row = [];
  for (var c=0;c<cols;c++) row.push('');
  return row;
}

// Finding/voiding linked milestones — handled elsewhere now; keep harmless behavior

// Legacy type checker — prefer grace.gs's isUniversalCreditType
function isUniversalCreditType_(creditType){
  if (typeof isUniversalCreditType === 'function') return isUniversalCreditType(creditType);
  var s = String(creditType||'').toLowerCase();
  return s === 'all' || s === 'universal' || s.indexOf('all-points') !== -1 || s.indexOf('all points') !== -1 || s.indexOf('universal credit') !== -1;
}

/**
 * handlePerformanceOnSubmit
 * Called when a Performance Issue event is submitted
 * Sets initial performance case stage and pending status
 */
function handlePerformanceOnSubmit(rowIndex) {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(events);
    var ctx = rowCtx_(events, rowIndex);
    
    // Set initial performance case stage
    var perfStageCol = map[CONFIG.COLS.PerfCaseStage] || map['Perf Case Stage'] || 0;
    if (perfStageCol) {
      events.getRange(rowIndex, perfStageCol).setValue('Open');
    }
    
    // Set initial pending status
    var pendingCol = map[CONFIG.COLS.PendingStatus] || map['Pending Status'] || 0;
    if (pendingCol) {
      events.getRange(rowIndex, pendingCol).setValue('Pending GA Assignment');
    }
    
    logInfo_ && logInfo_('handlePerformanceOnSubmit', { row: rowIndex });
  } catch (err) {
    logError && logError('handlePerformanceOnSubmit', err, { row: rowIndex });
  }
}

/**
 * checkPerformanceDeadlines
 * Scans all performance cases and checks for overdue deadlines
 * Returns count of overdue cases found
 */
function checkPerformanceDeadlines() {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(events);
    var last = events.getLastRow();
    if (last < 2) return 0;
    
    var cEventType = map[CONFIG.COLS.EventType] || map['EventType'] || 0;
    var cPerfStage = map[CONFIG.COLS.PerfCaseStage] || map['Perf Case Stage'] || 0;
    var cPerfDeadline = map[CONFIG.COLS.PerfDeadline] || map['Perf Deadline'] || 0;
    var cEmployee = map[CONFIG.COLS.Employee] || map['Employee'] || 0;
    var cPendingStatus = map[CONFIG.COLS.PendingStatus] || map['Pending Status'] || 0;
    
    if (!cEventType || !cPerfStage || !cPerfDeadline || !cEmployee) {
      logInfo_ && logInfo_('checkPerformanceDeadlines', 'missing required columns');
      return 0;
    }
    
    var maxCol = Math.max(cEventType, cPerfStage, cPerfDeadline, cEmployee, cPendingStatus);
    var values = events.getRange(2, 1, last - 1, maxCol).getValues();
    var overdueCount = 0;
    var today = new Date();
    today.setHours(0, 0, 0, 0); // Start of day for comparison
    
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var rowNum = i + 2;
      
      // Check if this is a performance issue
      var eventType = String(row[cEventType - 1] || '').trim().toLowerCase();
      if (eventType !== 'performance issue') continue;
      
      // Check if case is still open
      var stage = String(row[cPerfStage - 1] || '').trim();
      if (stage !== 'Open' && stage !== 'GA Open') continue;
      
      // Check deadline
      var deadline = row[cPerfDeadline - 1];
      if (!deadline) continue;
      
      var deadlineDate = new Date(deadline);
      if (isNaN(deadlineDate.getTime())) continue;
      
      deadlineDate.setHours(0, 0, 0, 0);
      
      if (deadlineDate < today) {
        // Case is overdue
        overdueCount++;
        
        // Update status to overdue
        if (cPendingStatus) {
          events.getRange(rowNum, cPendingStatus).setValue('Overdue');
        }
        
        // Log the overdue case
        var employee = String(row[cEmployee - 1] || '').trim();
        logInfo_ && logInfo_('checkPerformanceDeadlines_overdue', {
          row: rowNum,
          employee: employee,
          deadline: deadline,
          stage: stage
        });
      }
    }
    
    logInfo_ && logInfo_('checkPerformanceDeadlines', { overdueCount: overdueCount });
    return overdueCount;
    
  } catch (err) {
    logError && logError('checkPerformanceDeadlines', err);
    return 0;
  }
}

/**
 * shouldEscalatePerf
 * Determines if a performance case should be escalated based on overdue status
 * Returns true if case should be escalated, false otherwise
 */
function shouldEscalatePerf(employee) {
  try {
    if (!employee) return false;
    
    var events = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(events);
    var last = events.getLastRow();
    if (last < 2) return false;
    
    var cEventType = map[CONFIG.COLS.EventType] || map['EventType'] || 0;
    var cEmployee = map[CONFIG.COLS.Employee] || map['Employee'] || 0;
    var cPerfStage = map[CONFIG.COLS.PerfCaseStage] || map['Perf Case Stage'] || 0;
    var cPerfDeadline = map[CONFIG.COLS.PerfDeadline] || map['Perf Deadline'] || 0;
    var cPendingStatus = map[CONFIG.COLS.PendingStatus] || map['Pending Status'] || 0;
    
    if (!cEventType || !cEmployee || !cPerfStage || !cPerfDeadline) {
      return false;
    }
    
    var maxCol = Math.max(cEventType, cEmployee, cPerfStage, cPerfDeadline, cPendingStatus);
    var values = events.getRange(2, 1, last - 1, maxCol).getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Find the most recent performance issue for this employee
    var latestPerfCase = null;
    var latestRowNum = 0;
    
    for (var i = values.length - 1; i >= 0; i--) {
      var row = values[i];
      var rowNum = i + 2;
      var emp = String(row[cEmployee - 1] || '').trim();
      
      if (emp.toLowerCase() !== employee.toLowerCase()) continue;
      
      var eventType = String(row[cEventType - 1] || '').trim().toLowerCase();
      if (eventType !== 'performance issue') continue;
      
      // Found the most recent performance issue
      latestPerfCase = row;
      latestRowNum = rowNum;
      break;
    }
    
    if (!latestPerfCase) return false;
    
    // Check if case is still active
    var stage = String(latestPerfCase[cPerfStage - 1] || '').trim();
    if (stage !== 'Open' && stage !== 'GA Open') return false;
    
    // Check if deadline is overdue
    var deadline = latestPerfCase[cPerfDeadline - 1];
    if (!deadline) return false;
    
    var deadlineDate = new Date(deadline);
    if (isNaN(deadlineDate.getTime())) return false;
    
    deadlineDate.setHours(0, 0, 0, 0);
    
    var isOverdue = deadlineDate < today;
    
    // Check if already marked as overdue
    var pendingStatus = String(latestPerfCase[cPendingStatus - 1] || '').trim();
    var isAlreadyOverdue = pendingStatus.toLowerCase() === 'overdue';
    
    // Should escalate if overdue and not already marked as such
    var shouldEscalate = isOverdue && !isAlreadyOverdue;
    
    if (shouldEscalate) {
      // Update status to overdue
      if (cPendingStatus) {
        events.getRange(latestRowNum, cPendingStatus).setValue('Overdue');
      }
      
      logInfo_ && logInfo_('shouldEscalatePerf_escalated', {
        employee: employee,
        row: latestRowNum,
        deadline: deadline,
        stage: stage
      });
    }
    
    return shouldEscalate;
    
  } catch (err) {
    logError && logError('shouldEscalatePerf', err, { employee: employee });
    return false;
  }
}

/**
 * handlePerformanceIssueSubmit
 * Called when a Performance Issue event is submitted
 * Creates PERF_ISSUE PDF and updates count
 */
function handlePerformanceIssueSubmit(rowIndex) {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(events);
    var ctx = rowCtx_(events, rowIndex);
    
    var employee = String(ctx.get(CONFIG.COLS.Employee) || '').trim();
    if (!employee) return;
    
    // Get current Performance Issue count for this employee
    var currentCount = getPerformanceIssueCount_(employee, rowIndex);
    var newCount = currentCount + 1;
    
    // Update count in current row
    var countCol = map[CONFIG.COLS.PerfIssueCount] || map['Perf Issue Count'] || 0;
    if (countCol) {
      events.getRange(rowIndex, countCol).setValue(newCount);
    }
    
    // Check if this is the 2nd Performance Issue (trigger Growth Plan)
    if (newCount === 2) {
      triggerGrowthPlanConsequence_(employee, rowIndex);
    }
    
    // Check if this is a Performance Issue while on reduction status
    var reductionStatus = getCurrentReductionStatus_(employee, rowIndex);
    if (reductionStatus === 'Indefinite') {
      // Trigger Greater Reduction
      triggerGreaterReductionConsequence_(employee, rowIndex);
    } else if (reductionStatus === 'Greater') {
      // Trigger Performance Failure Termination
      triggerPerformanceFailureTermination_(employee, rowIndex);
    }
    
    logInfo_ && logInfo_('handlePerformanceIssueSubmit', { 
      row: rowIndex, 
      employee: employee, 
      count: newCount,
      reductionStatus: reductionStatus
    });
    
  } catch (err) {
    logError && logError('handlePerformanceIssueSubmit', err, { row: rowIndex });
  }
}

/**
 * getPerformanceIssueCount_
 * Gets the current count of Performance Issues for an employee
 */
function getPerformanceIssueCount_(employee, beforeRow) {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(events);
    var last = events.getLastRow();
    if (last < 2) return 0;
    
    var cEmployee = map[CONFIG.COLS.Employee] || map['Employee'] || 0;
    var cEventType = map[CONFIG.COLS.EventType] || map['EventType'] || 0;
    var cActive = map[CONFIG.COLS.Active] || map['Active'] || 0;
    
    if (!cEmployee || !cEventType) return 0;
    
    var maxCol = Math.max(cEmployee, cEventType, cActive);
    var values = events.getRange(2, 1, last - 1, maxCol).getValues();
    var count = 0;
    
    for (var i = 0; i < values.length; i++) {
      var rowNum = i + 2;
      if (beforeRow && rowNum >= beforeRow) break;
      
      var row = values[i];
      var emp = String(row[cEmployee - 1] || '').trim();
      var eventType = String(row[cEventType - 1] || '').trim().toLowerCase();
      var active = cActive ? (row[cActive - 1] === true || String(row[cActive - 1] || '').toLowerCase() === 'true') : true;
      
      if (emp.toLowerCase() === employee.toLowerCase() && 
          eventType === 'performance issue' && 
          active) {
        count++;
      }
    }
    
    return count;
    
  } catch (err) {
    logError && logError('getPerformanceIssueCount_', err, { employee: employee });
    return 0;
  }
}

/**
 * getCurrentReductionStatus_
 * Gets the current reduction status for an employee
 */
function getCurrentReductionStatus_(employee, beforeRow) {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(events);
    var last = events.getLastRow();
    if (last < 2) return 'None';
    
    var cEmployee = map[CONFIG.COLS.Employee] || map['Employee'] || 0;
    var cReductionStatus = map[CONFIG.COLS.PerfReductionStatus] || map['Perf Reduction Status'] || 0;
    var cActive = map[CONFIG.COLS.Active] || map['Active'] || 0;
    
    if (!cEmployee || !cReductionStatus) return 'None';
    
    var maxCol = Math.max(cEmployee, cReductionStatus, cActive);
    var values = events.getRange(2, 1, last - 1, maxCol).getValues();
    
    // Find the most recent reduction status
    for (var i = values.length - 1; i >= 0; i--) {
      var rowNum = i + 2;
      if (beforeRow && rowNum >= beforeRow) continue;
      
      var row = values[i];
      var emp = String(row[cEmployee - 1] || '').trim();
      var status = String(row[cReductionStatus - 1] || '').trim();
      var active = cActive ? (row[cActive - 1] === true || String(row[cActive - 1] || '').toLowerCase() === 'true') : true;
      
      if (emp.toLowerCase() === employee.toLowerCase() && active && status) {
        return status;
      }
    }
    
    return 'None';
    
  } catch (err) {
    logError && logError('getCurrentReductionStatus_', err, { employee: employee });
    return 'None';
  }
}

/**
 * triggerGrowthPlanConsequence_
 * Creates a Growth Plan consequence after 2 Performance Issues
 * Sets up for director claiming
 */
function triggerGrowthPlanConsequence_(employee, triggerRow) {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(events);
    
    // Create Growth Plan consequence row
    var consequenceData = {
      Employee: employee,
      EventType: 'Performance Milestone',        // Changed from 'Consequence'
      IncidentDate: new Date(),                  // Added: Set to current date
      Infraction: 'Growth Plan - Performance Improvement Required',
      ConsequenceDirector: '', // Leave blank for manual assignment
      PendingStatus: 'Pending Director Assignment'
    };
    
    // Add to Events sheet
    var newRow = appendEventsRow_(consequenceData);
    if (!newRow) return;
    
    // Set Performance tracking fields
    var ctx = rowCtx_(events, newRow);
    var today = new Date();
    
    if (map[CONFIG.COLS.PerfGrowthPlanDate] || map['Perf Growth Plan Date']) {
      var dateCol = map[CONFIG.COLS.PerfGrowthPlanDate] || map['Perf Growth Plan Date'];
      events.getRange(newRow, dateCol).setValue(today);
    }
    
    if (map[CONFIG.COLS.PerfReductionStatus] || map['Perf Reduction Status']) {
      var statusCol = map[CONFIG.COLS.PerfReductionStatus] || map['Perf Reduction Status'];
      events.getRange(newRow, statusCol).setValue('Growth Plan');
    }
    
    logInfo_ && logInfo_('triggerGrowthPlanConsequence', { 
      employee: employee, 
      triggerRow: triggerRow, 
      consequenceRow: newRow 
    });
    
  } catch (err) {
    logError && logError('triggerGrowthPlanConsequence_', err, { 
      employee: employee, 
      triggerRow: triggerRow 
    });
  }
}

/**
 * triggerGreaterReductionConsequence_
 * Creates Greater Reduction consequence for Performance Issue while on Indefinite Reduction
 * Sets up for director claiming
 */
function triggerGreaterReductionConsequence_(employee, triggerRow) {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(events);
    
    // Create Greater Reduction consequence row
    var consequenceData = {
      Employee: employee,
      EventType: 'Consequence',
      Infraction: 'Greater Reduction of Hours - Performance Issue on Indefinite Reduction',
      ConsequenceDirector: '', // Leave blank for manual assignment
      PendingStatus: 'Pending Director Assignment'
    };
    
    // Add to Events sheet
    var newRow = appendEventsRow_(consequenceData);
    if (!newRow) return;
    
    // Set Performance tracking fields
    var statusCol = map[CONFIG.COLS.PerfReductionStatus] || map['Perf Reduction Status'] || 0;
    if (statusCol) {
      events.getRange(newRow, statusCol).setValue('Greater');
    }
    
    logInfo_ && logInfo_('triggerGreaterReductionConsequence', { 
      employee: employee, 
      triggerRow: triggerRow, 
      consequenceRow: newRow 
    });
    
  } catch (err) {
    logError && logError('triggerGreaterReductionConsequence_', err, { 
      employee: employee, 
      triggerRow: triggerRow 
    });
  }
}

/**
 * triggerPerformanceFailureTermination_
 * Creates Performance Failure Termination for Performance Issue while on Greater Reduction
 * Sets up for director claiming
 */
function triggerPerformanceFailureTermination_(employee, triggerRow) {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(events);
    
    // Create Performance Failure Termination row
    var consequenceData = {
      Employee: employee,
      EventType: 'Consequence',
      Infraction: 'Performance Failure Termination - Performance Issue on Greater Reduction',
      ConsequenceDirector: '', // Leave blank for manual assignment
      PendingStatus: 'Pending Director Assignment'
    };
    
    // Add to Events sheet
    var newRow = appendEventsRow_(consequenceData);
    if (!newRow) return;
    
    logInfo_ && logInfo_('triggerPerformanceFailureTermination', { 
      employee: employee, 
      triggerRow: triggerRow, 
      consequenceRow: newRow 
    });
    
  } catch (err) {
    logError && logError('triggerPerformanceFailureTermination_', err, { 
      employee: employee, 
      triggerRow: triggerRow 
    });
  }
}

/**
 * handlePerformanceEdit
 * Handles performance-related edits in the Events sheet
 * Returns true if handled, false otherwise
 */
function handlePerformanceEdit(e) {
  try {
    if (!e || !e.range || !e.range.getSheet) return false;
    
    var sh = e.range.getSheet();
    var eventsTab = (CONFIG && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
    if (sh.getName() !== eventsTab) return false;
    
    var row = e.range.getRow();
    if (row < 2) return false;
    
    var col = e.range.getColumn();
    var hdrs = headers_(sh);
    var colName = hdrs[col - 1] || '';
    var ctx = rowCtx_(sh, row);
    var map = headerIndexMap_(sh);
    
    // Handle Consequence Director claiming
    if (colName === CONFIG.COLS.ConsequenceDirector || colName === 'Consequence Director') {
      var newDirector = e.value;
      var oldDirector = e.oldValue;
      
      if (newDirector && newDirector !== oldDirector) {
        var eventType = String(ctx.get(CONFIG.COLS.EventType) || '').trim().toLowerCase();
        var infraction = String(ctx.get(CONFIG.COLS.Infraction) || '').trim();
        
        // Check if this is a Growth Plan consequence
        if (eventType === 'consequence' && infraction.toLowerCase().indexOf('growth plan') !== -1) {
          // Growth Plan claimed - check if deadline is set
          var deadline = ctx.get(CONFIG.COLS.PerfGrowthPlanDeadline) || ctx.get('Perf Growth Plan Deadline');
          if (deadline) {
            // Both claimed and deadline set - update status
            var pendingCol = map[CONFIG.COLS.PendingStatus] || map['Pending Status'] || 0;
            if (pendingCol) {
              sh.getRange(row, pendingCol).setValue('Pending Growth Plan Decision');
            }
            
            // Generate Growth Plan PDF
            try {
              var pdfId = createConsequencePdf_(row, 'Growth Plan');
              
              // Write PDF link to sheet
              if (pdfId) {
                var pdfHdr = CONFIG.COLS.PdfLink || 'Write-Up PDF';
                var pdfUrl = 'https://drive.google.com/file/d/' + pdfId + '/view';
                setRichLinkSafe(sh, row, pdfHdr, 'View PDF', pdfUrl);
              }
              
              logInfo_ && logInfo_('handlePerformanceEdit_growth_plan_claimed', { 
                row: row, 
                director: newDirector, 
                deadline: deadline,
                pdfId: pdfId
              });
            } catch (pdfErr) {
              logError && logError('handlePerformanceEdit_growth_plan_pdf_err', pdfErr, { 
                row: row, 
                director: newDirector 
              });
            }
          } else {
            // Claimed but no deadline yet - update status
            var pendingCol = map[CONFIG.COLS.PendingStatus] || map['Pending Status'] || 0;
            if (pendingCol) {
              sh.getRange(row, pendingCol).setValue('Pending Deadline Assignment');
            }
          }
        }
        
        // Check if this is a Greater Reduction consequence
        else if (eventType === 'consequence' && infraction.toLowerCase().indexOf('greater reduction') !== -1) {
          // Greater Reduction claimed - generate PDF immediately
          var pendingCol = map[CONFIG.COLS.PendingStatus] || map['Pending Status'] || 0;
          if (pendingCol) {
            sh.getRange(row, pendingCol).setValue('Completed');
          }
          
          try {
            var pdfId = createConsequencePdf_(row, 'Greater Reduction');
            logInfo_ && logInfo_('handlePerformanceEdit_greater_reduction_claimed', { 
              row: row, 
              director: newDirector,
              pdfId: pdfId
            });
          } catch (pdfErr) {
            logError && logError('handlePerformanceEdit_greater_reduction_pdf_err', pdfErr, { 
              row: row, 
              director: newDirector 
            });
          }
        }
        
        // Check if this is a Performance Failure Termination consequence
        else if (eventType === 'consequence' && infraction.toLowerCase().indexOf('performance failure termination') !== -1) {
          // Termination claimed - generate PDF immediately
          var pendingCol = map[CONFIG.COLS.PendingStatus] || map['Pending Status'] || 0;
          if (pendingCol) {
            sh.getRange(row, pendingCol).setValue('Completed');
          }
          
          try {
            var pdfId = createConsequencePdf_(row, 'Performance Failure Termination');
            logInfo_ && logInfo_('handlePerformanceEdit_termination_claimed', { 
              row: row, 
              director: newDirector,
              pdfId: pdfId
            });
          } catch (pdfErr) {
            logError && logError('handlePerformanceEdit_termination_pdf_err', pdfErr, { 
              row: row, 
              director: newDirector 
            });
          }
        }
      }
    }
    
    // Handle Growth Plan deadline entry
    if (colName === CONFIG.COLS.PerfGrowthPlanDeadline || colName === 'Perf Growth Plan Deadline') {
      var deadline = e.value;
      if (deadline && deadline !== e.oldValue) {
        var director = String(ctx.get(CONFIG.COLS.ConsequenceDirector) || '').trim();
        
        if (director) {
          // Both claimed and deadline set - update status
          var pendingCol = map[CONFIG.COLS.PendingStatus] || map['Pending Status'] || 0;
          if (pendingCol) {
            sh.getRange(row, pendingCol).setValue('Pending Growth Plan Decision');
          }
          
          // Generate Growth Plan PDF
          try {
            var pdfId = createConsequencePdf_(row, 'Growth Plan');
            
            // Write PDF link to sheet
            if (pdfId) {
              var pdfHdr = CONFIG.COLS.PdfLink || 'Write-Up PDF';
              var pdfUrl = 'https://drive.google.com/file/d/' + pdfId + '/view';
              setRichLinkSafe(sh, row, pdfHdr, 'View PDF', pdfUrl);
            }
            
            logInfo_ && logInfo_('handlePerformanceEdit_growth_plan_deadline', { 
              row: row, 
              deadline: deadline,
              director: director,
              pdfId: pdfId
            });
          } catch (pdfErr) {
            logError && logError('handlePerformanceEdit_growth_plan_pdf_err', pdfErr, { 
              row: row, 
              deadline: deadline,
              director: director
            });
          }
        }
      }
    }
    
    // Handle Growth Plan decision
    if (colName === CONFIG.COLS.PerfGrowthPlanDecision || colName === 'Perf Growth Plan Decision') {
      var decision = String(e.value || '').trim().toLowerCase();
      var oldDecision = String(e.oldValue || '').trim().toLowerCase();
      
      if (decision && decision !== oldDecision) {
        var employee = String(ctx.get(CONFIG.COLS.Employee) || '').trim();
        
        if (decision === 'success' || decision === 'passed') {
          // Return to good standing - reset count
          handleGrowthPlanSuccess_(employee, row);
        } else if (decision === 'failure' || decision === 'failed') {
          // Indefinite reduction
          handleGrowthPlanFailure_(employee, row);
        }
        
        logInfo_ && logInfo_('handlePerformanceEdit_growth_plan_decision', { 
          row: row, 
          employee: employee, 
          decision: decision 
        });
      }
    }
    
    return true;
    
  } catch (err) {
    logError && logError('handlePerformanceEdit', err, { 
      row: e.range ? e.range.getRow() : 'unknown',
      col: e.range ? e.range.getColumn() : 'unknown'
    });
    return false;
  }
}

/**
 * handleGrowthPlanSuccess_
 * Handles successful completion of Growth Plan
 * Creates a new "Return to Good Standing" event
 */
function handleGrowthPlanSuccess_(employee, row) {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var today = new Date();
    
    // Create new "Return to Good Standing" event row
    var successEventData = {
      Employee: employee,
      EventType: 'Return to Good Standing',
      IncidentDate: today,
      Infraction: 'Growth Plan Completed Successfully',
      Points: 0,
      PendingStatus: 'Completed'
    };
    
    // Add the new event row
    var newRow = appendEventsRow_(successEventData);
    if (newRow) {
      // Manually create PDF since this bypasses the form submit flow
      try {
        var pdfId = null;
        if (typeof createConsequencePdf_ === 'function') {
          pdfId = createConsequencePdf_(newRow, 'Return to Good Standing');
        } else if (typeof createEventRecordPdf_ === 'function') {
          pdfId = createEventRecordPdf_(newRow);
        }
        
        if (pdfId) {
          // Write PDF link to the new row
          var pdfHdr = CONFIG.COLS.PdfLink || 'Write-Up PDF';
          var pdfUrl = 'https://drive.google.com/file/d/' + pdfId + '/view';
          setRichLinkSafe(events, newRow, pdfHdr, 'View PDF', pdfUrl);
        }
        
        logInfo_ && logInfo_('handleGrowthPlanSuccess_pdf_created', { 
          newRow: newRow,
          employee: employee,
          pdfId: pdfId
        });
      } catch (pdfErr) {
        logError && logError('handleGrowthPlanSuccess_pdf_err', pdfErr, { 
          newRow: newRow,
          employee: employee
        });
      }
      
      logInfo_ && logInfo_('handleGrowthPlanSuccess_event_created', { 
        originalRow: row,
        newRow: newRow,
        employee: employee
      });
    }
    
    // Update original Growth Plan row status
    var ctx = rowCtx_(events, row);
    ctx.set(CONFIG.COLS.PendingStatus, 'Completed - Success Event Created');
    
    // Reset Performance Issue count for this employee
    resetPerformanceIssueCount_(employee);
    
  } catch (err) {
    logError && logError('handleGrowthPlanSuccess_', err, { employee: employee, row: row });
  }
}

/**
 * handleGrowthPlanFailure_
 * Handles failed Growth Plan (creates indefinite reduction event)
 */
function handleGrowthPlanFailure_(employee, row) {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var today = new Date();
    
    // Create new "Indefinite Reduction" event row
    var failureEventData = {
      Employee: employee,
      EventType: 'Performance Milestone',
      IncidentDate: today,
      Infraction: 'Growth Plan Failed - Indefinite Hour Reduction',
      Points: 0,
      ConsequenceDirector: '', // Leave blank for manual assignment
      PendingStatus: 'Pending Director Assignment'
    };
    
    // Add the new event row
    var newRow = appendEventsRow_(failureEventData);
    if (newRow) {
      // Set performance tracking fields on the new row
      var newCtx = rowCtx_(events, newRow);
      
      // Set reduction status to Indefinite
      if (CONFIG.COLS.PerfReductionStatus) {
        newCtx.set(CONFIG.COLS.PerfReductionStatus, 'Indefinite');
      }
      
      // Manually create PDF since this bypasses the form submit flow
      try {
        var pdfId = null;
        if (typeof createConsequencePdf_ === 'function') {
          pdfId = createConsequencePdf_(newRow, 'Indefinite Reduction');
        } else if (typeof createEventRecordPdf_ === 'function') {
          pdfId = createEventRecordPdf_(newRow);
        }
        
        if (pdfId) {
          // Write PDF link to the new row
          var pdfHdr = CONFIG.COLS.PdfLink || 'Write-Up PDF';
          var pdfUrl = 'https://drive.google.com/file/d/' + pdfId + '/view';
          setRichLinkSafe(events, newRow, pdfHdr, 'View PDF', pdfUrl);
        }
        
        logInfo_ && logInfo_('handleGrowthPlanFailure_pdf_created', { 
          newRow: newRow,
          employee: employee,
          pdfId: pdfId
        });
      } catch (pdfErr) {
        logError && logError('handleGrowthPlanFailure_pdf_err', pdfErr, { 
          newRow: newRow,
          employee: employee
        });
      }
      
      logInfo_ && logInfo_('handleGrowthPlanFailure_event_created', { 
        originalRow: row,
        newRow: newRow,
        employee: employee
      });
    }
    
    // Update original Growth Plan row status
    var ctx = rowCtx_(events, row);
    ctx.set(CONFIG.COLS.PendingStatus, 'Completed - Failure Event Created');
    
  } catch (err) {
    logError && logError('handleGrowthPlanFailure_', err, { employee: employee, row: row });
  }
}

/**
 * triggerGreaterReductionConsequence_
 * Creates Greater Reduction consequence for Performance Issue while on Indefinite Reduction
 * Sets up for director claiming
 */
function triggerGreaterReductionConsequence_(employee, triggerRow) {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(events);
    
    // Create Greater Reduction consequence row
    var consequenceData = {
      Employee: employee,
      EventType: 'Consequence',
      Infraction: 'Greater Reduction of Hours - Performance Issue on Indefinite Reduction',
      ConsequenceDirector: '', // Leave blank for manual assignment
      PendingStatus: 'Pending Director Assignment'
    };
    
    // Add to Events sheet
    var newRow = appendEventsRow_(consequenceData);
    if (!newRow) return;
    
    // Set Performance tracking fields
    var statusCol = map[CONFIG.COLS.PerfReductionStatus] || map['Perf Reduction Status'] || 0;
    if (statusCol) {
      events.getRange(newRow, statusCol).setValue('Greater');
    }
    
    logInfo_ && logInfo_('triggerGreaterReductionConsequence', { 
      employee: employee, 
      triggerRow: triggerRow, 
      consequenceRow: newRow 
    });
    
  } catch (err) {
    logError && logError('triggerGreaterReductionConsequence_', err, { 
      employee: employee, 
      triggerRow: triggerRow 
    });
  }
}

/**
 * triggerPerformanceFailureTermination_
 * Creates Performance Failure Termination for Performance Issue while on Greater Reduction
 * Sets up for director claiming
 */
function triggerPerformanceFailureTermination_(employee, triggerRow) {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(events);
    
    // Create Performance Failure Termination row
    var consequenceData = {
      Employee: employee,
      EventType: 'Consequence',
      Infraction: 'Performance Failure Termination - Performance Issue on Greater Reduction',
      ConsequenceDirector: '', // Leave blank for manual assignment
      PendingStatus: 'Pending Director Assignment'
    };
    
    // Add to Events sheet
    var newRow = appendEventsRow_(consequenceData);
    if (!newRow) return;
    
    logInfo_ && logInfo_('triggerPerformanceFailureTermination', { 
      employee: employee, 
      triggerRow: triggerRow, 
      consequenceRow: newRow 
    });
    
  } catch (err) {
    logError && logError('triggerPerformanceFailureTermination_', err, { 
      employee: employee, 
      triggerRow: triggerRow 
    });
  }
}

/**
 * resetPerformanceIssueCount_
 * Resets Performance Issue count for an employee after successful Growth Plan
 */
function resetPerformanceIssueCount_(employee) {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var map = headerIndexMap_(events);
    var last = events.getLastRow();
    if (last < 2) return;
    
    var cEmployee = map[CONFIG.COLS.Employee] || map['Employee'] || 0;
    var cEventType = map[CONFIG.COLS.EventType] || map['EventType'] || 0;
    var cInfraction = map[CONFIG.COLS.Infraction] || map['Infraction'] || 0;
    var cNullify = map[CONFIG.COLS.Nullify] || map['Nullify'] || 0;
    var cCount = map[CONFIG.COLS.PerfIssueCount] || map['Perf Issue Count'] || 0;
    
    if (!cEmployee || !cEventType || !cNullify) return;
    
    var maxCol = Math.max(cEmployee, cEventType, cInfraction, cNullify, cCount);
    var values = events.getRange(2, 1, last - 1, maxCol).getValues();
    var nullifiedCount = 0;
    var nullifiedGrowthPlans = 0;
    
    for (var i = 0; i < values.length; i++) {
      var rowNum = i + 2;
      var row = values[i];
      var emp = String(row[cEmployee - 1] || '').trim();
      var eventType = String(row[cEventType - 1] || '').trim().toLowerCase();
      var infraction = cInfraction ? String(row[cInfraction - 1] || '').trim().toLowerCase() : '';
      var isNullified = row[cNullify - 1] === true || String(row[cNullify - 1] || '').toLowerCase() === 'true';
      
      if (emp.toLowerCase() === employee.toLowerCase() && !isNullified) {
        
        // Nullify Performance Issues
        if (eventType === 'performance issue') {
          events.getRange(rowNum, cNullify).setValue(true);
          nullifiedCount++;
        }
        
        // Nullify Growth Plan consequences
        else if (eventType === 'performance milestone' && infraction.indexOf('growth plan') !== -1) {
          events.getRange(rowNum, cNullify).setValue(true);
          nullifiedGrowthPlans++;
        }
      }
    }
    
    // Reset count to 0 for all rows of this employee
    // (This ensures the Perf Issue Count column shows 0 even before the next Performance Issue)
    for (var j = 0; j < values.length; j++) {
      var rowNum2 = j + 2;
      var row2 = values[j];
      var emp2 = String(row2[cEmployee - 1] || '').trim();
      
      if (emp2.toLowerCase() === employee.toLowerCase() && cCount) {
        events.getRange(rowNum2, cCount).setValue(0);
      }
    }
    
    logInfo_ && logInfo_('resetPerformanceIssueCount_', { 
      employee: employee, 
      nullifiedPerformanceIssues: nullifiedCount,
      nullifiedGrowthPlans: nullifiedGrowthPlans
    });
    
  } catch (err) {
    logError && logError('resetPerformanceIssueCount_', err, { employee: employee });
  }
}















