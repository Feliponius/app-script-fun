function activeEmail_() {
  // Legacy: returns the effective user email (blank for personal accounts in many cases)
  // return (Session.getActiveUser() && Session.getActiveUser().getEmail()) || '';
  return ''; // disabled in legacy mode
}

function protectGraceSystemColumns(){
  var s = sh_(CONFIG.TABS.EVENTS), map = headerIndexMap_(s);
  var cols = [CONFIG.COLS.GraceApplied, CONFIG.COLS.GraceTier, CONFIG.COLS.GraceLedgerRow, CONFIG.COLS.LinkedEventRow, CONFIG.COLS.Nullify]
    .map(h=>map[h]).filter(Boolean);
  cols.forEach(function(c){ var p = s.protect().setRange(s.getRange(2, c, s.getMaxRows()-1, 1)); p.removeEditors(p.getEditors()); p.setWarningOnly(true); });
}

function addCreditTypeValidation(){
  var pp = ensurePositivePointsColumns_(), s = pp.sheet, i = pp.idx;
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Minor','Moderate','Major'], true).setAllowInvalid(false).build();
  s.getRange(2, i.CreditType, Math.max(1, s.getMaxRows()-1)).setDataValidation(rule);
}

function isDirector_(email) {
  // Legacy: would check against CONFIG.DIRECTORS list
  // return CONFIG.DIRECTORS
  //   .map(e => e.toLowerCase())
  //   .includes(String(email || '').toLowerCase());
  return true; // legacy mode — always true
}

function assertDirector_() {
  // Legacy: would throw if not a director
  // const email = activeEmail_();
  // if (!isDirector_(email)) throw new Error('Unauthorized: Director access required. (' + email + ')');
  // return email;
  return ''; // legacy mode — always "passes"
}

// Legacy-disabled: Prevent non-directors from changing policy settings in Config tab
function onEditConfigGuard_(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== CONFIG.TABS.CONFIG) return;

  // Legacy:
  // const editorEmail = activeEmail_();
  // if (!isDirector_(editorEmail)) {
  //   e.range.setValue(e.oldValue || '');
  //   SpreadsheetApp.getActive().toast('Only directors can change policy settings.', 'Access Denied', 5);
  // }

  // Current: do nothing — control access via Google Sheets protected ranges
  return;
}

// Legacy-disabled: Handle director-only consequence actions in Events tab
function onDirectorEdit_(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== CONFIG.TABS.EVENTS) return;

  const colName = headers_(sh)[e.range.getColumn() - 1];
  const editedValue = e.value;

  // Legacy:
  // const editor = activeEmail_();
  // if (!isDirector_(editor)) {
  //   e.range.setValue(e.oldValue || '');
  //   SpreadsheetApp.getActive().toast('Only directors can set milestones.', 'Access Denied', 5);
  //   return;
  // }

  if (colName === CONFIG.COLS.Milestone && editedValue) {
    const rowIndex = e.range.getRow();
    generateConsequencePdf_(rowIndex, editedValue);
    logAudit(`Milestone set: ${editedValue}`, rowIndex);
  }
}

// Improved onEdit claim handler — chooses milestone template and generates PDF.
// Paste this in the same file that contains your onEdit handlers (e.g., guards.gs).
function onEditMilestoneClaim_(e) {
  try {
    if (!e || !e.range || !e.range.getSheet) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== (CONFIG && CONFIG.TABS && CONFIG.TABS.EVENTS)) return;

    // Ignore header-row edits (ARRAYFORMULA/header changes)
    var rowIndex = e.range.getRow();
    if (rowIndex === 1) return;

    var headersList = headers_(sh);
    var colName = headersList[e.range.getColumn() - 1];
    if (colName !== (CONFIG && CONFIG.COLS && CONFIG.COLS.ConsequenceDirector)) return; // only act on director claims

    var ctx = null;
    try { ctx = rowCtx_(sh, rowIndex); } catch (_) { ctx = null; }

    // --- Build PDF header candidates (defensive) ---
    var pdfHeaderCandidates = [
      (CONFIG && CONFIG.COLS && CONFIG.COLS.PdfLink),
      (CONFIG && CONFIG.COLS && CONFIG.COLS.WriteUpPDF),
      (CONFIG && CONFIG.COLS && CONFIG.COLS.WriteupPDF),
      'Write-Up PDF',
      'WriteupPDF'
    ].filter(Boolean);

    // --- Read existing Writeup (before) safely via sheet ranges (avoid ctx.get throwing) ---
    var writeupBefore = '';
    for (var i = 0; i < pdfHeaderCandidates.length; i++) {
      var h = pdfHeaderCandidates[i];
      var idx = headersList.indexOf(h);
      if (idx !== -1) {
        try { writeupBefore = String(sh.getRange(rowIndex, idx + 1).getValue() || '').trim(); } catch (_) { writeupBefore = ''; }
        if (writeupBefore) break;
      }
    }

    // Director name (value from the edit, or fallback to cell)
    var directorName = '';
    if (e && typeof e.value !== 'undefined' && e.value !== null) directorName = String(e.value).trim();
    if (!directorName && ctx) {
      try { directorName = String(ctx.get(CONFIG.COLS.ConsequenceDirector) || '').trim(); } catch (_) { directorName = ''; }
    }
    if (!directorName) return;            // nothing to do if blank

    // Only proceed if this row is Pending and has no PDF already
    var pending = '';
    try { if (ctx) pending = String(ctx.get(CONFIG.COLS.PendingStatus) || '').trim(); } catch(_) { pending = ''; }
    if (pending.toLowerCase() !== 'pending') return;
    if (writeupBefore !== '') return;

    // Ensure Lead placeholder is the claiming director (best-effort)
    try { if (ctx) ctx.set(CONFIG.COLS.Lead, directorName); } catch (ignore) {}

    // -------------------- Generate PDF (claim path) --------------------
    var pdfId = null;
    try {
      var milestoneLabel = '';
      try { milestoneLabel = ctx ? String(ctx.get(CONFIG.COLS.Milestone) || '').trim() : ''; } catch(_) { milestoneLabel = ''; }

      var pfTplId = (CONFIG && CONFIG.TEMPLATES && CONFIG.TEMPLATES.PROBATION_FAILURE) || null;
      var pfRe = /probation.*fail|probation.*failure|probation failure|termination|terminated/i;

      if (pfTplId && pfRe.test(milestoneLabel)) {
        try {
          pdfId = createMilestonePdf_(rowIndex, pfTplId);
        } catch (errForced) {
          logError && logError('createMilestonePdf_forcedPF_failed', errForced, { row: rowIndex, tpl: pfTplId });
          pdfId = null;
        }
      }

      if (!pdfId) {
        try { pdfId = createMilestonePdf_(rowIndex); } catch (errInf) {
          logError && logError('createMilestonePdf_infer_failed', errInf, { row: rowIndex });
          pdfId = null;
        }
      }

      if (!pdfId) {
        try {
          pdfId = createEventRecordPdf_(rowIndex);
        } catch (errEvt) {
          logError && logError('createEventRecordPdf_failed', errEvt, { row: rowIndex });
          pdfId = null;
        }
      }
    } catch (errPdfFlow) {
      logError && logError('onEditMilestoneClaim_pdfFlow', errPdfFlow, { row: rowIndex });
      pdfId = null;
    }

    // Re-check pdf cell defensively (use sheet ranges)
    var writeupAfter = '';
    for (var j = 0; j < pdfHeaderCandidates.length; j++) {
      var hh = pdfHeaderCandidates[j];
      var idx2 = headersList.indexOf(hh);
      if (idx2 !== -1) {
        try { writeupAfter = String(sh.getRange(rowIndex, idx2 + 1).getValue() || '').trim(); } catch (_) { writeupAfter = ''; }
        if (writeupAfter) break;
      }
    }

    // If createEventRecordPdf_ returned an id but the sheet still has no link, write a friendly rich link
    if (pdfId && !writeupAfter) {
      try {
        var url = (String(pdfId).indexOf('http') === 0) ? pdfId : ('https://drive.google.com/file/d/' + String(pdfId) + '/view');
        var writeHeader = pdfHeaderCandidates.length ? pdfHeaderCandidates[0] : ((CONFIG && CONFIG.COLS && CONFIG.COLS.PdfLink) || 'Write-Up PDF');
        setRichLinkSafe(sh, rowIndex, writeHeader, 'View PDF', url);
        writeupAfter = 'View PDF';
      } catch (wrErr) {
        logError && logError('onEditMilestoneClaim_setRichLinkSafe', wrErr, { row: rowIndex, pdfId: pdfId });
      }
    }

    // If PDF exists or was created, mark Completed and — optionally — enable probation
    if (pdfId || writeupAfter) {
      try { if (ctx) ctx.set(CONFIG.COLS.PendingStatus, 'Completed'); } catch (ignore) {}

      logAudit && logAudit('MilestoneClaim', 'Director ' + directorName + ' claimed milestone row ' + rowIndex + ' — pdfId=' + (pdfId||'none'), rowIndex, { pdfId: pdfId });

      // --- enable probation when Tier 2 milestone is claimed (guarded) ---
      try {
        // Do NOT auto-write probation fields unless explicitly allowed by the feature flag
        if (!scriptShouldWriteProbation_()) {
          logInfo_ && logInfo_('onEditMilestoneClaim_skipProbation', { row: rowIndex, employee: (ctx ? (ctx.get(CONFIG.COLS.Employee)||'') : ''), milestone: (ctx ? (ctx.get(CONFIG.COLS.Milestone)||'') : '') });
        } else {
          var employee = '';
          try { employee = ctx ? String(ctx.get(CONFIG.COLS.Employee) || '').trim() : ''; } catch(_) { employee = ''; }

          // compute effective rolling points (prefer row ctx values, then sheet scan)
          var effective = 0;
          try {
            effective = Number(ctx ? (ctx.get(CONFIG.COLS.PointsRollingEffective) || ctx.get(CONFIG.COLS.PointsRolling) || 0) : 0);
          } catch(_) { effective = 0; }
          if (!effective) {
            try { effective = Number(getRollingPointsForEmployee_(employee, { beforeRow: rowIndex }) || 0); } catch(_) { effective = Number(effective || 0); }
          }

          var computedTier = tierForPoints_(effective);

          // consider Tier 2+ by computedTier OR explicit configured names OR a clear "Level 2/3" label
          var confNameLevel2 = (CONFIG && CONFIG.POLICY && CONFIG.POLICY.MILESTONE_NAMES && CONFIG.POLICY.MILESTONE_NAMES.LEVEL_2) || '';
          var confNameLevel3 = (CONFIG && CONFIG.POLICY && CONFIG.POLICY.MILESTONE_NAMES && CONFIG.POLICY.MILESTONE_NAMES.LEVEL_3) || '';
          var labelTrim = String(milestoneLabel || '').trim();
          var labelMatchesLevel2or3 = (confNameLevel2 && String(confNameLevel2).trim() === labelTrim) || (confNameLevel3 && String(confNameLevel3).trim() === labelTrim);
          var isTier2OrHigher = (Number(computedTier) >= 2) || labelMatchesLevel2or3 || /level\s*[2-9]/i.test(labelTrim);

          if (isTier2OrHigher) {
            // set probation fields (script is allowed because scriptShouldWriteProbation_() returned true)
            var suspensionDays = _policyNumber_ ? _policyNumber_('SUSPENSION_L2_DAYS', 7) : 7;
            var probationDays = _policyNumber_ ? _policyNumber_('PROBATION_DAYS', 30) : 30;
            var startDate = addDaysLocal_ ? addDaysLocal_(new Date(), suspensionDays) : addDaysLocal_(new Date(), suspensionDays); // fallback same name; keep consistent
            var endDate = addDaysLocal_ ? addDaysLocal_(startDate, probationDays) : startDate;

            try {
              if (ctx && CONFIG.COLS.ProbationActive) ctx.set(CONFIG.COLS.ProbationActive, true);
              if (ctx && CONFIG.COLS.ProbationStart) ctx.set(CONFIG.COLS.ProbationStart, startDate);
              if (ctx && CONFIG.COLS.ProbationEnd) ctx.set(CONFIG.COLS.ProbationEnd, endDate);
              if (ctx && CONFIG.COLS.PerfNoPickup) ctx.set(CONFIG.COLS.PerfNoPickup, true);
              if (ctx && CONFIG.COLS.PerfNoPickupEnd) ctx.set(CONFIG.COLS.PerfNoPickupEnd, endDate);
              if (ctx && CONFIG.COLS.Per_NoPickup_Active) ctx.set(CONFIG.COLS.Per_NoPickup_Active, true);
              logAudit && logAudit('ProbationAutoSet', 'Probation flags auto-set on claim for ' + employee, rowIndex, { start: formatDate_(startDate), end: formatDate_(endDate), effective: effective });
            } catch (setErr) {
              logError && logError('onEditMilestoneClaim_setProbation', setErr, { rowIndex: rowIndex, employee: employee });
            }
          }
        }
      } catch (autoProbErr) {
        logError && logError('onEditMilestoneClaim_autoProbationOuter', autoProbErr, { rowIndex: rowIndex });
      }

    } else {
      logAudit && logAudit('MilestoneClaimFailed', 'Director ' + directorName + ' claimed milestone row ' + rowIndex + ' — PDF generation failed', rowIndex, {});
    }

    // Optional: update linked trigger row audit if present
    try {
      var linked = '';
      try { linked = ctx ? ctx.get(CONFIG.COLS.Linked_Event_ID) : ''; } catch(_) { linked = ''; }
      if (linked) logAudit && logAudit('MilestoneClaimLinked', 'Milestone ' + rowIndex + ' claimed by ' + directorName, Number(linked), {});
    } catch (_) { /* ignore */ }

  } catch (err) {
    logError && logError('onEditMilestoneClaim_', err, { e: e });
  }
}





/**
 * onEditNullify_
 * Lightweight onEdit handler to react when the "Nullify" column is edited.
 * - Recomputes effective rolling points for the row's employee
 * - Writes PointsRolling (Effective) if available (fallback to PointsRolling)
 * - Calls scanAndHandleMilestones_ to refresh milestone decisions immediately
 *
 * Install as an onEdit trigger or call from your global onEdit(e) if you have one.
 */
function onEditNullify(e){
  try {
    if (!e || !e.range || !e.range.getSheet) return;
    const sh = e.range.getSheet();
    if (String(sh.getName()).trim() !== String(CONFIG.TABS.EVENTS).trim()) return; // only act on Events

    const hdrs = headers_(sh);                 // <-- array of header names
    const col   = e.range.getColumn();
    const row   = e.range.getRow();
    const colName = hdrs[col - 1] || '';

    // Only react to edits on the Nullify column
    if (colName !== CONFIG.COLS.Nullify) return;

    // Avoid no-op writes (e.oldValue can be undefined on some edit types)
    if (typeof e.oldValue !== 'undefined' && e.oldValue === e.value) return;

    const ctx = rowCtx_(sh, row);
    const employee = ctx.get(CONFIG.COLS.Employee);
    if (!employee) {
      logAudit && logAudit(`onEditNullify: skipped — no employee on row ${row}`, row);
      return;
    }

    // Compute rolling/effective using your helper (supports number or {effective} return)
    const roll = getRollingPointsForEmployee_(employee);
    let effective = "";
    if (typeof roll === 'number') {
      effective = Number(roll || 0);
    } else if (roll && typeof roll === 'object') {
      effective = Number(roll.effective || roll.rolling || 0);
    }

    // Write PointsRolling(Effective) safely
    try {
      const effHeader = CONFIG.COLS.PointsRollingEffective;
      const effIdx = effHeader ? hdrs.indexOf(effHeader) : -1;  // <-- use hdrs, not headers_
      if (effIdx !== -1) {
        ctx.set(effHeader, effective);
      } else if (hdrs.indexOf(CONFIG.COLS.PointsRolling) !== -1) {
        ctx.set(CONFIG.COLS.PointsRolling, effective);
      } else {
        // Neither column exists — log once so you can add the header
        logError && logError('onEditNullify: no PointsRolling(Effective) header', null, { row, employee });
      }
    } catch (writeErr) {
      logError && logError('onEditNullify setEffective', writeErr, { row, employee });
    }

    // Re-run milestone scan for this row
    try {
      scanAndHandleMilestones_(row);
    } catch (scanErr) {
      logError && logError('onEditNullify scanAndHandleMilestones_', scanErr, { row, employee });
    }

    logAudit && logAudit(`Nullify toggled for ${employee} (row ${row}) — effective=${effective}`, row);
  } catch (err) {
    logError && logError('onEditNullify', err, { e });
  }
}






