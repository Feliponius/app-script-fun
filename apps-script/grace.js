/** =========================
 *  grace.gs
 *  Positive Points & Grace engine
 *  (camelCase only)
 *  =========================
 *
 *  Public entrypoints you’ll call:
 *   - handleGraceEdit(e)                      // wire at the top of onAnyEdit
 *   - recordPositiveCreditFromEvent(rowIndex) // call after appending Events row or on edits
 *   - applyGraceToEvent(eventRowCtx, requiredTier?) // internal use by handleGraceEdit
 *
 *  Useful helpers (public):
 *   - listAvailableCreditsForEmployee(employee)
 *   - consumeCredit(ledgerRow, usedOnRow, notesOpt)
 *   - normalizeTier / tierRank / isTierAtLeast
 *   - isUniversalCreditType / countUniversalCredits
 *   - resolveTierForAction(actionText, fallbackCellText)
 */

/** ===== Basic utils (no-throw) ===== */
function safeSheetByName(name) {
  try { return SpreadsheetApp.getActive().getSheetByName(name); } catch (e) { return null; }
}
function safeVal(v) { return v == null ? '' : v; }

/** Header map: { "Header Name": 1-based column } */
function headerMap(sheet) {
  if (!sheet) return {};
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  var vals = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  var map = {};
  for (var c = 0; c < vals.length; c++) map[String(vals[c]).trim()] = c + 1;
  return map;
}

function headerMapCI(sheet){
  var map = {};
  if (!sheet) return map;
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return map;
  var vals = sheet.getRange(1,1,1,lastCol).getValues()[0] || [];
  for (var c=0;c<vals.length;c++){
    var raw = String(vals[c]||'').trim();
    if (!raw) continue;
    map[raw] = c+1;                 // exact
    map[raw.toLowerCase()] = c+1;   // case-insensitive
  }
  return map;
}
function getCol(hmap, names){
  names = Array.isArray(names) ? names : [names];
  for (var i=0;i<names.length;i++){
    var n = names[i]; if (!n) continue;
    var exact = hmap[n]; if (exact) return exact;
    var ci = hmap[String(n).toLowerCase()]; if (ci) return ci;
  }
  return 0;
}


/** Safe row getter by header name(s). */
function rowGet(rowArr, hmap, names, defaultValue) {
  names = Array.isArray(names) ? names : [names];
  for (var i = 0; i < names.length; i++) {
    var col = hmap[names[i]];
    if (col) return safeVal(rowArr[col - 1]);
  }
  return defaultValue;
}

/** Returns the first empty body row (row >= 2) or lastRow+1 if none are empty. */
/** Returns the first open body row. Prefers blank EventType; falls back to truly blank rows (''/null/FALSE). */
function nextOpenRow_(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 2;

  var hmap = headerMap(sh);
  var cEvt = hmap['EventType'] || hmap['Event Type'] || 0;

  // Prefer the first row whose EventType is blank
  if (cEvt) {
    var vals = sh.getRange(2, cEvt, lastRow - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i][0];
      if (v === '' || v === null) return i + 2;
    }
  }

  // Fallback: find the first row where all cells are blank-like (''/null/FALSE)
  var lastCol = Math.max(1, sh.getLastColumn());
  var rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (var r = 0; r < rows.length; r++) {
    var empty = true;
    for (var c = 0; c < rows[r].length; c++) {
      var cell = rows[r][c];
      if (!(cell === '' || cell === null || cell === false)) { empty = false; break; }
    }
    if (empty) return r + 2;
  }

  return lastRow + 1;
}


/** Ensure and write PositivePoints -> "Grace Row" for the ledger record used. */
function setGraceRowOnLedger_(ledgerRow, graceChildRow) {
  var sh = getPositivePointsSheet();
  if (!sh || !ledgerRow || ledgerRow < 2) return false;
  var hmap = headerMap(sh);
  var cGraceRow = hmap['Grace Row'] || 0;
  if (!cGraceRow) {
    sh.insertColumnAfter(sh.getLastColumn());
    var nc = sh.getLastColumn();
    sh.getRange(1, nc).setValue('Grace Row');
    hmap = headerMap(sh);
    cGraceRow = hmap['Grace Row'] || 0;
  }
  if (!cGraceRow) return false;
  sh.getRange(ledgerRow, cGraceRow).setValue(graceChildRow);
  return true;
}


/** ===== Tier utils ===== */
function normalizeTier(t) {
  var s = String(t || '').trim().toLowerCase();
  
  // Handle universal credit types first
  if (isUniversalCreditType(s)) return 'universal';
  
  if (s === 'min' || s === 'minor') return 'minor';
  if (s === 'mod' || s === 'moderate') return 'moderate';
  if (s === 'maj' || s === 'major') return 'major';
  return '';
}

function tierRank(t) {
  switch (normalizeTier(t)) {
    case 'universal': return 4;  // Highest tier - can cover any requirement
    case 'major': return 3;
    case 'moderate': return 2;
    case 'minor': return 1;
    default: return 0;
  }
}

function isTierAtLeast(haveTier, needTier) {
  return tierRank(haveTier) >= tierRank(needTier);
}
function displayTierLabel(t) {
  var s = normalizeTier(t);
  if (s === 'minor') return 'Minor';
  if (s === 'moderate') return 'Moderate';
  if (s === 'major') return 'Major';
  return '';
}


// Robust shim: maps your PositivePoints headers and only creates sane, non-empty headers if needed.
function ensurePositivePointsColumns_(){
  var sh = getPositivePointsSheet();
  if (!sh) throw new Error('PositivePoints sheet not found');

  var hdr = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0] || [];

  function find(names){
    names = Array.isArray(names) ? names : [names];
    for (var i = 0; i < names.length; i++){
      var name = String(names[i] || '').trim();
      if (!name) continue;
      var j = hdr.indexOf(name);
      if (j !== -1) return j + 1; // 1-based
    }
    return 0;
  }

  function ensure(name){
    // Never allow blank header titles
    var canonical = String(name || '').trim();
    if (!canonical) canonical = 'Col ' + (hdr.length + 1);

    var col = hdr.indexOf(canonical) + 1;
    if (col) return col;

    sh.insertColumnAfter(sh.getLastColumn());
    var newCol = sh.getLastColumn();
    sh.getRange(1, newCol).setValue(canonical);
    hdr = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0] || [];
    return newCol;
  }

  // Your current headers (from your sheet):
  // Timestamp | IncidentDate | Employee | Credit Type | Approved? | Consumed? | Consumed On | Consumed By
  // | Applied Event Row | Notes | Earned Row | Credit Reason | Earned Event Row

  var idx = {
    Employee:     find(['Employee'])                         || ensure('Employee'),
    CreditType:   find(['Credit Type','Type'])               || ensure('Credit Type'),
    Approved:     find(['Approved?','Approved'])             || 0,
    Consumed:     find(['Consumed?','Used?','Consumed'])     || ensure('Consumed?'),
    ConsumedOn:   find(['Consumed On','Used On Date'])       || ensure('Consumed On'),
    ConsumedBy:   find(['Consumed By'])                      || 0,
    AppliedEvent: find(['Applied Event Row','Used On Row'])  || ensure('Applied Event Row'),
    Notes:        find(['Notes','Note'])                     || ensure('Notes'),
    // “Earned Row” is an alias some rows use; keep both in play
    EarnedEvent:  find(['Earned Event Row','Source Event Row','From Event Row','EarnedFromRow','Earned Row']) || 0,
    Reason:       find(['Credit Reason','Earned Reason','Reason']) || 0
  };

  return { sheet: sh, hdr: hdr, idx: idx };
}


/** ===== Universal program-type classification (caps/analytics) ===== */
function isUniversalCreditType(creditType) {
  var s = String(creditType || '').trim().toLowerCase();
  if (s === 'all' || s === 'universal') return true;
  if (s.indexOf('all-points') !== -1) return true;
  if (s.indexOf('all points') !== -1) return true;
  if (s.indexOf('universal credit') !== -1) return true;
  // add store-specific synonyms here if needed
  return false;
}
function countUniversalCredits(rows) {
  var total = 0;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var type = r.type || '';
    var val = Number(r.points || 0);
    if (isUniversalCreditType(type)) total += (isNaN(val) ? 0 : val);
  }
  return total;
}

/** ===== PositivePoints ledger I/O ===== */
function getPositivePointsSheet() {
  var name = (typeof CONFIG !== 'undefined' && CONFIG.TABS && CONFIG.TABS.POSITIVE_POINTS)
    ? CONFIG.TABS.POSITIVE_POINTS
    : 'PositivePoints';
  return safeSheetByName(name);
}

/** Load unused credits for an employee (normalized for selection). */
function listAvailableCreditsForEmployee(employee) {
  var sh = getPositivePointsSheet();
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var hmap = headerMap(sh);
  var data = sh.getRange(2,1,lastRow-1, sh.getLastColumn()).getValues();
  var out = [];

  var cEmp   = hmap['Employee'];
  var cType  = hmap['Credit Type'] || hmap['Type'];
  var cPts   = hmap['Points'] || hmap['Value'] || 0;
  var cUsed  = hmap['Consumed?'] || hmap['Used?'] || hmap['Consumed'] || hmap['Used'];

  for (var i=0;i<data.length;i++){
    var row = data[i];
    var emp = cEmp ? row[cEmp-1] : '';
    if (String(emp).trim() !== String(employee).trim()) continue;

    var usedRaw = cUsed ? String(row[cUsed-1]||'').toLowerCase() : '';
    if (usedRaw === 'true' || usedRaw === 'y' || usedRaw === '1') continue;

    var typeRaw = cType ? String(row[cType-1]||'') : '';
    var tier    = normalizeTier(typeRaw); // derive tier from Credit Type (e.g., "Moderate" / "Minor")
    var pts     = cPts ? Number(row[cPts-1]||0) : 0;

    out.push({ employee: emp, tier: tier, type: typeRaw, points: isNaN(pts)?0:pts, row: i+2 });
  }
  return out;
}


/** Mark a ledger credit as used and stamp metadata. */
function consumeCredit(ledgerRow, usedOnRow, notesOpt) {
  var sh = getPositivePointsSheet();
  if (!sh || !ledgerRow || ledgerRow < 2) return false;

  var hmap = headerMap(sh);

  // Only create a column when a non-empty name is provided and it's missing.
  function ensure(name) {
    if (!name) return 0;                       // no-op for falsy names
    if (hmap[name]) return hmap[name];         // already exists
    sh.insertColumnAfter(sh.getLastColumn());  // add at far right
    var nc = sh.getLastColumn();
    sh.getRange(1, nc).setValue(name);
    hmap = headerMap(sh);                      // refresh header map
    return hmap[name] || 0;
  }

  var cConsumed   = hmap['Consumed?'] || hmap['Used?'] || hmap['Consumed'] || hmap['Used'];
  var cConsumedOn = hmap['Consumed On'] || hmap['Used On Date'];
  var cConsumedBy = hmap['Consumed By'] || 0;
  var cAppliedRow = hmap['Applied Event Row'] || hmap['Used On Row'] || 0;
  var cNotes      = hmap['Notes'] || 0;

  // Ensure required columns exist (no accidental insert when already present)
  cConsumed   = cConsumed   || ensure('Consumed?');
  cConsumedOn = cConsumedOn || ensure('Consumed On');
  if (!cAppliedRow) cAppliedRow = ensure('Applied Event Row');
  if (!cNotes)      cNotes      = ensure('Notes');

  var maxCol = Math.max(cConsumed, cConsumedOn, cConsumedBy || 0, cAppliedRow, cNotes);
  var rng = sh.getRange(ledgerRow, 1, 1, maxCol);
  var rowVals = rng.getValues()[0];

  rowVals[cConsumed - 1]   = true;
  rowVals[cConsumedOn - 1] = new Date();
  if (cConsumedBy) rowVals[cConsumedBy - 1] = (Session.getActiveUser && Session.getActiveUser().getEmail) ? Session.getActiveUser().getEmail() : '';
  rowVals[cAppliedRow - 1] = usedOnRow || '';
  rowVals[cNotes - 1]      = (notesOpt != null ? notesOpt : rowVals[cNotes - 1]);

  rng.setValues([rowVals]);
  return true;
}



/** Choose the smallest sufficient credit by tier, then earliest row. */
function pickSufficientCredit(availableCredits, requiredTier) {
  var needRank = tierRank(requiredTier);
  var candidates = (availableCredits || []).filter(function (c) {
    return tierRank(c.tier) >= needRank;
  });

  candidates.sort(function (a, b) {
    var rdiff = tierRank(a.tier) - tierRank(b.tier);
    if (rdiff !== 0) return rdiff;
    return (a.row || 0) - (b.row || 0);
  });

  return candidates.length ? candidates[0] : null;
}

/** Apply grace to an Events row (writes Grace fields + consumes credit). */
function applyGraceToEvent(eventRowCtx, requiredTier) {
  if (!eventRowCtx || !eventRowCtx.sheet || !eventRowCtx.rowIndex) return false;

  var needTier = normalizeTier(requiredTier || eventRowCtx.requiredTier || 'minor');
  var credits  = listAvailableCreditsForEmployee(eventRowCtx.employee);
  var chosen   = pickSufficientCredit(credits, needTier);
  if (!chosen) {
    try { Logger.log(JSON.stringify({ msg: 'applyGraceToEvent_noCredit', needTier: needTier, employee: eventRowCtx.employee })); } catch (e) {}
    return false;
  }

  // 1) consume selected credit in PositivePoints (stamps Consumed?, Consumed On, Applied Event Row)
  if (!consumeCredit(chosen.row, eventRowCtx.rowIndex, 'Applied to grace event')) return false;

  // 2) write grace flags back to the parent Events row (no Points touches)
  var sh   = eventRowCtx.sheet;
  var hmap = eventRowCtx.headers || headerMap(sh);

  function colEnsure(name) {
    if (hmap[name]) return hmap[name];
    sh.insertColumnAfter(sh.getLastColumn());
    var newCol = sh.getLastColumn();
    sh.getRange(1, newCol).setValue(name);
    hmap = headerMap(sh);
    return hmap[name];
  }

  var appliedCol = colEnsure('Grace Applied');
  var tierCol    = colEnsure('Grace Tier');
  var ledgerCol  = colEnsure('Grace Ledger Row');
  var maxCol     = Math.max(appliedCol, tierCol, ledgerCol);

  var rowVals = sh.getRange(eventRowCtx.rowIndex, 1, 1, maxCol).getValues()[0];
  rowVals[appliedCol - 1] = true;
  rowVals[tierCol - 1]    = chosen.tier; // actual tier used
  rowVals[ledgerCol - 1]  = chosen.row;  // PositivePoints ledger row consumed
  sh.getRange(eventRowCtx.rowIndex, 1, 1, maxCol).setValues([rowVals]);

  // 3) append a synthetic "Grace" child row for audit history (NO points changes)
  var childRow = appendGraceEventRow_(eventRowCtx, chosen);

  // 4) write the child row back onto the ledger as "Grace Row"
  setGraceRowOnLedger_(chosen.row, childRow);

  return true;
}



/** remove Grace and Grace event row */
function unapplyGraceFromEvent(eventRowCtx) {
  try {
    // ---- Guard ----
    if (!eventRowCtx || !eventRowCtx.sheet || !eventRowCtx.rowIndex) return false;

    var sh  = eventRowCtx.sheet;
    var row = Number(eventRowCtx.rowIndex);

    // Resolve header map (prefer headerIndexMap_ if available)
    var hmap = eventRowCtx.headers
            || (typeof headerIndexMap_ === 'function' ? headerIndexMap_(sh)
               : (typeof headerMap === 'function' ? headerMap(sh) : null));
    if (!hmap) throw new Error('header map not available');

    // Alias-safe column resolver for Events
    function evCol() {
      var names = Array.prototype.slice.call(arguments).flat().filter(Boolean);
      for (var i = 0; i < names.length; i++) {
        var c = hmap[names[i]];
        if (c) return c;
      }
      return 0;
    }

    // --- Events columns (alias tolerant) ---
    var cGrace        = evCol(CONFIG && CONFIG.COLS && CONFIG.COLS.Grace,        'Grace');
    var cGraceReason  = evCol(CONFIG && CONFIG.COLS && CONFIG.COLS.GraceReason,  'Grace Reason','GraceReason');
    var cGracedBy     = evCol(CONFIG && CONFIG.COLS && CONFIG.COLS.GracedBy,     'Graced By','GracedBy');
    var cLedgerRow    = evCol(CONFIG && CONFIG.COLS && CONFIG.COLS.GraceLedgerRow,'Grace Ledger Row','GraceLedgerRow','Linked Grace Row');
    var cGraceApplied = evCol(CONFIG && CONFIG.COLS && CONFIG.COLS.GraceApplied, 'Grace Applied','GraceApplied');

    // ---------- PositivePoints sheet + columns ----------
    // Try robust resolver first; then fall back to tab name or helper.
    var ppSheet = null, ppHdr = null;
    try {
      if (typeof ensurePositivePointsColumns_ === 'function') {
        var ppMeta = ensurePositivePointsColumns_(); // { sheet, idx } in your codebase
        ppSheet = ppMeta && ppMeta.sheet;
        ppHdr   = ppMeta && ppMeta.idx ? ppMeta.idx : null; // may contain direct indexes
      }
    } catch(_) {}

    if (!ppSheet) {
      try {
        if (typeof getPositivePointsSheet === 'function') ppSheet = getPositivePointsSheet();
      } catch(_) {}
    }
    if (!ppSheet && CONFIG && CONFIG.TABS && CONFIG.TABS.POSITIVE_POINTS) {
      try { ppSheet = sh_(CONFIG.TABS.POSITIVE_POINTS); } catch(_) {}
    }

    // Build a PP column resolver that works whether we have an index map (idx) or header map.
    var ppMap = null;
    if (!ppHdr && ppSheet) {
      ppMap = (typeof headerIndexMap_ === 'function' ? headerIndexMap_(ppSheet)
             : (typeof headerMap === 'function' ? headerMap(ppSheet) : null));
    }

    function ppCol() {
      var names = Array.prototype.slice.call(arguments).flat().filter(Boolean);
      // If ensurePositivePointsColumns_ provided numeric indexes, prefer that
      if (ppHdr) {
        for (var i = 0; i < names.length; i++) {
          var key = names[i];
          if (ppHdr[key]) return ppHdr[key];
        }
      }
      // Fallback: header-based mapping
      if (ppMap) {
        for (var j = 0; j < names.length; j++) {
          var c = ppMap[names[j]];
          if (c) return c;
        }
      }
      return 0;
    }

    // ---------- Find linked ledger row (if any) ----------
    var ledgerRow = 0;
    if (cLedgerRow) {
      ledgerRow = Number(sh.getRange(row, cLedgerRow).getValue()) || 0;
    }

    // Fallback search by "Applied Event Row" if no stored ledger pointer
    if (!ledgerRow && ppSheet) {
      var cAppliedEvent = ppCol('Applied Event Row','Used On Row','AppliedEventRow','AppliesToRow');
      if (cAppliedEvent) {
        var last = ppSheet.getLastRow();
        if (last >= 2) {
          var vals = ppSheet.getRange(2, 1, last - 1, Math.max(cAppliedEvent, ppSheet.getLastColumn())).getValues();
          for (var i = 0; i < vals.length; i++) {
            if (Number(vals[i][cAppliedEvent - 1]) === row) { ledgerRow = i + 2; break; }
          }
        }
      }
    }

    // ---------- Clear ledger "consumed" state & notes ----------
    if (ppSheet && ledgerRow) {
      var cConsumed   = ppCol('Consumed?','Used?','Consumed');
      var cConsumedOn = ppCol('Consumed On','Used On Date','ConsumedOn');
      var cConsumedBy = ppCol('Consumed By','ConsumedBy','Used By');
      var cAppliedRow = ppCol('Applied Event Row','Used On Row','AppliedEventRow','AppliesToRow');
      var cGraceRow   = ppCol('Grace Row','Linked Grace Row','GraceRow');
      var cNotes      = ppCol('Notes','Credit Reason','Earned Reason','Reason');

      if (cConsumed)   ppSheet.getRange(ledgerRow, cConsumed).setValue(false);
      if (cConsumedOn) ppSheet.getRange(ledgerRow, cConsumedOn).setValue('');
      if (cConsumedBy) ppSheet.getRange(ledgerRow, cConsumedBy).setValue('');
      if (cAppliedRow) ppSheet.getRange(ledgerRow, cAppliedRow).setValue('');
      if (cGraceRow)   ppSheet.getRange(ledgerRow, cGraceRow).setValue('');
      if (cNotes)      ppSheet.getRange(ledgerRow, cNotes).setValue('');
    }

    // ---------- Clear Events-side grace fields ----------
    if (cGrace)        sh.getRange(row, cGrace).setValue(false);
    if (cGraceApplied) sh.getRange(row, cGraceApplied).setValue(false);
    if (cGraceReason)  sh.getRange(row, cGraceReason).setValue('');
    if (cGracedBy)     sh.getRange(row, cGracedBy).setValue('');
    if (cLedgerRow)    sh.getRange(row, cLedgerRow).setValue('');

    // ---------- Delete any synthetic “Grace” child rows linked to this parent ----------
    try { if (typeof deleteGraceRowsForSource_ === 'function') deleteGraceRowsForSource_(row); }
    catch (err) { try { logError && logError('deleteGraceRowsForSource_', err, { parent: row }); } catch(_){} }

    // ---------- Audit ----------
    try {
      logAudit && logAudit(
        'grace_unapplied_clear_fields',
        row,
        { cleared: ['Grace','Grace Applied','Grace Reason','Graced By','Grace Ledger Row'], ledgerRow: ledgerRow || null }
      );
    } catch(_) {}

    return true;

  } catch (e) {
    try { logError && logError('unapplyGraceFromEvent', e, {}); } catch(_) {}
    return false;
  }
}




/** onAnyEdit hook for Grace. Returns true if it handled the edit. */
function handleGraceEdit(e){
  try{
    if (!e || !e.range || !e.range.getSheet) return false;

    var sh = e.range.getSheet();
    var eventsTab = (CONFIG && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
    if (String(sh.getName()).trim() !== String(eventsTab).trim()) return false;

    var rowIndex = e.range.getRow();
    if (rowIndex === 1) return false; // ignore header

    var hdrs = headers_(sh);                 // array of header names
    var colIndex = e.range.getColumn();
    var colName  = hdrs[colIndex - 1] || '';

    // We only care about Grace / Grace Reason / Graced By
    var watch = ['Grace','Grace Reason','Graced By'];
    if (watch.indexOf(colName) === -1) return false;

    // Build row context
    var hmap = headerMap(sh);
    var row  = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];

    function getBy(names, def){
      names = Array.isArray(names) ? names : [names];
      for (var i=0;i<names.length;i++){
        var c = hmap[names[i]];
        if (c) return row[c-1];
      }
      return def;
    }
    function truthy(v){
      if (v === true) return true;
      var s = String(v||'').trim().toLowerCase();
      return (s === 'true' || s === 'yes' || s === 'y' || s === '1' || s === 'checked');
    }

    var employee = String(getBy(['Employee'],'')).trim();
    if (!employee) return false;

    var graceVal = getBy(['Grace'],'');
    var applyNow = truthy(graceVal);

    // Debounce to avoid re-entry loops when we write back
    var cache = (CacheService && CacheService.getScriptCache) ? CacheService.getScriptCache() : null;
    var key   = cache ? ('grace-row-' + rowIndex) : null;
    if (cache && cache.get(key)) return true;

    // Required tier: pull from a helper column if you have it, else default to 'minor'
    var reqTier  = (typeof normalizeTier === 'function')
      ? normalizeTier(getBy(['Required Tier','Required Grace Tier'],'minor'))
      : 'minor';

    var ctx = { sheet: sh, rowIndex: rowIndex, employee: employee, requiredTier: reqTier, headers: hmap };

    if (colName === 'Grace') {
      if (applyNow) {
        var ok = (typeof applyGraceToEvent === 'function') ? applyGraceToEvent(ctx, reqTier) : false;
        logInfo_ && logInfo_('grace_apply_attempt', { row: rowIndex, employee: employee, reqTier: reqTier, ok: ok, editedCol: colName });

        // >>> cancel/void any milestone rows linked to this source event
        try {
          if (ok && typeof cancelPendingMilestonesForSource_ === 'function') {
            var n = cancelPendingMilestonesForSource_(rowIndex);
            logInfo_ && logInfo_('grace_cancel_milestones', {
              eventRow: rowIndex,
              cancelledOrVoided: n,
              note: n ? 'matched linked milestones' : 'no milestones matched; check Linked_Event_ID / EventType'
            });
            if (typeof logAudit === 'function') {
              logAudit('GraceCascade', 'Cancelled/Voided linked milestones', rowIndex, { count: n });
            }
          }
        } catch (err) {
          logError && logError('grace_cancel_milestones_err', err, { eventRow: rowIndex });
        }

        if (cache && ok) cache.put(key, '1', 3);
      } else {
        // (existing UNDO path)
        var undone = false;
        if (typeof unapplyGraceFromEvent === 'function') {
          try { undone = unapplyGraceFromEvent(ctx); } catch(_){ undone = false; }
        }
        try { if (typeof deleteGraceRowsForSource_ === 'function') deleteGraceRowsForSource_(rowIndex); } catch(_){}
        logInfo_ && logInfo_('grace_unapply_attempt', { row: rowIndex, employee: employee, editedCol: colName, undone: !!undone });
        if (cache) cache.put(key, '1', 3);
      }
      return true;
    }

    // If we got here, we watched a Grace-related column but not the checkbox itself
    return false;

  } catch(err){
    logError && logError('handleGraceEdit', err, { where: 'top' });
    return false;
  }
}




/** ===== Optional PositiveMap (Action → Tier) ===== */
function normalizeActionKey(s) {
  return String(s || '').toLowerCase().replace(/[^\p{L}\p{N}\s]/gu, '').replace(/\s+/g, ' ').trim();
}
function lookupTierFromMap(actionText) {
  try {
    if (!actionText) return '';
    var tab = (typeof CONFIG !== 'undefined' && CONFIG.TABS && CONFIG.TABS.POSITIVE_MAP) ? CONFIG.TABS.POSITIVE_MAP : 'PositiveMap';
    var sh = safeSheetByName(tab);
    if (!sh) return '';
    var hdr = headerMap(sh);
    var cAction = hdr['Action'] || hdr['Positive Action'] || hdr['PositiveAction'] || hdr['Name'] || hdr['Title'];
    var cTier   = hdr['Tier']   || hdr['Credit Type']     || hdr['Type']          || hdr['Level'];
    if (!cAction || !cTier) return '';
    var last = sh.getLastRow();
    if (last < 2) return '';
    var want = normalizeActionKey(actionText);
    var vals = sh.getRange(2, 1, last - 1, Math.max(cAction, cTier)).getValues();
    for (var i = 0; i < vals.length; i++) {
      var act = normalizeActionKey(vals[i][cAction - 1]);
      if (act && act === want) return normalizeTier(vals[i][cTier - 1]);
    }
    return '';
  } catch (e) { return ''; }
}
/** Prefer map → fallback cell text → 'moderate'. */
function resolveTierForAction(actionText, fallbackCellText) {
  return lookupTierFromMap(actionText) || normalizeTier(fallbackCellText) || 'moderate';
}

/** ===== Append a Positive Credit to the ledger from an Events row ===== */
function recordPositiveCreditFromEvent(rowIndex) {
  var eventsTab = (typeof CONFIG !== 'undefined' && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
  var evSh = SpreadsheetApp.getActive().getSheetByName(eventsTab);
  if (!evSh || rowIndex < 2) return false;

  // debounce
  var cache = (CacheService && CacheService.getScriptCache) ? CacheService.getScriptCache() : null;
  var cacheKey = cache ? ('pp-ledger-row-' + rowIndex) : null;
  if (cache && cache.get(cacheKey)) return false;

  var evHdrCI = headerMapCI(evSh);
  var row = evSh.getRange(rowIndex,1,1,evSh.getLastColumn()).getValues()[0];
  function getVal(names){ var c=getCol(evHdrCI,names); return c?row[c-1]:''; }

  var eventType = String(getVal(['EventType','Documentation Type','Event Type']) || '').toLowerCase();
  if (!/^(positive(\s*(point|credit))?|positive)$/.test(eventType)) return false;

  var employee  = String(getVal('Employee') || '').trim();
  if (!employee) return false;

  var actionText = getVal(['PositiveAction','Positive Action','Positive Point','PositiveActionType']) || '';
  var tierWord   = resolveTierForAction(actionText, getVal(['Points','Tier','Credit Type','Type','Level']) || '');
  var creditTypeToStore = (tierWord === 'minor' || tierWord === 'moderate' || tierWord === 'major')
    ? (tierWord.charAt(0).toUpperCase() + tierWord.slice(1))  // "Minor"/"Moderate"/"Major"
    : (String(tierWord || actionText || '').trim());           // e.g., "All-Points" or action label

  var ppSh = getPositivePointsSheet();
  if (!ppSh) return false;
  var ppHdr = headerMap(ppSh);

  function colEnsure(name){
    if (ppHdr[name]) return ppHdr[name];
    ppSh.insertColumnAfter(ppSh.getLastColumn());
    var nc = ppSh.getLastColumn(); ppSh.getRange(1,nc).setValue(name);
    ppHdr = headerMap(ppSh); return ppHdr[name];
  }

  var cTimestamp    = colEnsure('Timestamp');
  var cIncidentDate = ppHdr['IncidentDate'] || colEnsure('IncidentDate'); // your header is without space
  var cEmployee     = colEnsure('Employee');
  var cCreditType   = ppHdr['Credit Type'] || colEnsure('Credit Type');
  var cApproved     = ppHdr['Approved?'] || ppHdr['Approved'] || null;
  var cConsumed     = ppHdr['Consumed?'] || ppHdr['Used?'] || null;
  var cAppliedRow   = ppHdr['Applied Event Row'] || colEnsure('Applied Event Row');
  var cReason       = ppHdr['Credit Reason'] || ppHdr['Earned Reason'] || ppHdr['Reason'] || ppHdr['Notes'] || null;

  var newRow = Math.max(2, ppSh.getLastRow() + 1);
  var writeMax = Math.max(cTimestamp, cIncidentDate, cEmployee, cCreditType, cApproved || 0, cConsumed || 0, cAppliedRow, cReason || 0);
  var vals = ppSh.getRange(newRow, 1, 1, writeMax).getValues()[0];

  vals[cTimestamp - 1]    = new Date();
  vals[cIncidentDate - 1] = getVal(['IncidentDate','Incident Date','Timestamp']) || new Date();
  vals[cEmployee - 1]     = employee;
  vals[cCreditType - 1]   = creditTypeToStore;   // store the label here
  if (cApproved)          vals[cApproved - 1] = true;
  if (cConsumed)          vals[cConsumed - 1] = false;
  vals[cAppliedRow - 1]   = rowIndex;
  if (cReason)            vals[cReason - 1] = getVal(['IncidentDescription','Notes / Reviewer','Notes Reviewer']) || actionText || '';

  ppSh.getRange(newRow, 1, 1, writeMax).setValues([vals]);
  if (cache) cache.put(cacheKey, '1', 3);
  try { Logger.log(JSON.stringify({msg:'positive_ledger_append', row:rowIndex, ledgerRow:newRow, employee:employee, creditType:creditTypeToStore})); } catch(_){}
  return true;
}


/** Append a synthetic "Grace" child row in Events linking back to the source row.
 *  Does NOT touch Points/Final/Rolling columns. */
function appendGraceEventRow_(eventRowCtx, chosen) {
  var sh = eventRowCtx.sheet;
  var hmap = headerMap(sh);

  function colEnsure(name){
    if (hmap[name]) return hmap[name];
    sh.insertColumnAfter(sh.getLastColumn());
    var nc = sh.getLastColumn();
    sh.getRange(1, nc).setValue(name);
    hmap = headerMap(sh);
    return hmap[name];
  }

  // Only the fields needed to document the grace action
  var cEventType      = colEnsure('EventType');
  var cIncidentDate   = colEnsure('IncidentDate');
  var cEmployee       = colEnsure('Employee');
  var cGracedBy       = hmap['Graced By']        || colEnsure('Graced By');
  var cGraceReason    = hmap['Grace Reason']     || colEnsure('Grace Reason');
  var cGraceTier      = hmap['Grace Tier']       || colEnsure('Grace Tier');
  var cGraceLedgerRow = hmap['Grace Ledger Row'] || colEnsure('Grace Ledger Row');
  var cLinkedEventRow = hmap['Linked Event Row'] || colEnsure('Linked Event Row');

  // Pull a couple values from the parent row (incident date, graced by, reason)
  var parentVals  = sh.getRange(eventRowCtx.rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
  var incDate     = hmap['IncidentDate'] ? parentVals[hmap['IncidentDate']-1] : new Date();
  var gracedByVal = cGracedBy    ? parentVals[cGracedBy-1]    : '';
  var graceReason = cGraceReason ? parentVals[cGraceReason-1] : '';

  // Next open row (first empty), not necessarily the bottom
  var newRow = nextOpenRow_(sh);
  var maxCol = Math.max(cEventType, cIncidentDate, cEmployee, cGracedBy||0, cGraceReason||0,
                        cGraceTier, cGraceLedgerRow, cLinkedEventRow);

  var out = sh.getRange(newRow, 1, 1, maxCol).getValues()[0];

  out[cEventType-1]      = 'Grace';
  out[cIncidentDate-1]   = incDate || new Date();
  out[cEmployee-1]       = eventRowCtx.employee;
  if (cGracedBy)         out[cGracedBy-1]       = gracedByVal;
  if (cGraceReason)      out[cGraceReason-1]    = graceReason;
  out[cGraceTier-1]      = chosen.tier;               // Minor/Moderate/Major
  out[cGraceLedgerRow-1] = chosen.row;                // PositivePoints ledger row consumed
  out[cLinkedEventRow-1] = eventRowCtx.rowIndex;      // link back to parent

  // No Points / Final / Rolling updates here — we leave those untouched
  sh.getRange(newRow, 1, 1, maxCol).setValues([out]);

  try { logInfo_ && logInfo_('grace_child_row_appended', { parent: eventRowCtx.rowIndex, child: newRow, ledgerRow: chosen.row }); } catch(_){}
  return newRow;
}


function deleteGraceRowsForSource_(eventRow) {
  try {
    var eventsTab = (CONFIG && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
    var sh = SpreadsheetApp.getActive().getSheetByName(eventsTab);
    if (!sh || eventRow < 2) return 0;

    // Case-insensitive header map
    var hmap = (typeof headerMapCI === 'function') ? headerMapCI(sh) : headerMap(sh);

    // Local helper that does NOT shadow global getCol
    function getColLocal(names) {
      if (typeof getCol === 'function') return getCol(hmap, names);
      names = Array.isArray(names) ? names : [names];
      for (var i = 0; i < names.length; i++) {
        var n = names[i];
        var c = hmap[n] || hmap[String(n).toLowerCase()];
        if (c) return c;
      }
      return 0;
    }

    var cEvt  = getColLocal(['EventType','Event Type']);
    var cLink = getColLocal(['Linked Event Row','LinkedEventRow','Linked Event ID','LinkedEventID']);
    if (!cEvt || !cLink) return 0;

    var last = sh.getLastRow();
    if (last < 2) return 0;

    var vals = sh.getRange(2, 1, last - 1, Math.max(cEvt, cLink)).getValues();
    var toDelete = [];
    for (var i = 0; i < vals.length; i++) {
      var absRow = i + 2;
      var evt  = String(vals[i][cEvt - 1] || '').trim().toLowerCase();
      var link = Number(vals[i][cLink - 1] || 0);
      if (evt === 'grace' && link === Number(eventRow)) toDelete.push(absRow);
    }

    // delete bottom-up
    for (var j = toDelete.length - 1; j >= 0; j--) {
      sh.deleteRow(toDelete[j]);
    }

    // Single, clean log line
    if (typeof logInfo_ === 'function') {
      logInfo_('grace_child_rows_deleted', JSON.stringify({
        parent: Number(eventRow),
        removed: toDelete.length,
        rows: toDelete
      }));
    } else {
      Logger.log(JSON.stringify({
        msg: 'grace_child_rows_deleted',
        parent: Number(eventRow),
        removed: toDelete.length,
        rows: toDelete
      }));
    }

    return toDelete.length; // ✅ return actual count
  } catch (err) {
    if (typeof logError === 'function') {
      logError('deleteGraceRowsForSource_', err, { eventRow: eventRow });
    } else {
      Logger.log('deleteGraceRowsForSource_ error: ' + (err && err.message ? err.message : err));
    }
    return 0;
  }
}

