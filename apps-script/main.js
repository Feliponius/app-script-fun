

function _isPositiveCreditEventText_(evt){
  var s = String(evt || '').trim().toLowerCase();
  return (
    s === 'positive' ||
    s === 'positive credit' ||
    /^positive\s*(point|credit)\b/.test(s) ||
    s === 'positive point removal'
  );
}

// Replace/override the original onFormSubmit with a locks+backoff wrapper.
// If you renamed your original to _onFormSubmit_impl, this will call it.
// If you prefer not to rename, create a wrapper that calls your actual implementation name.
function onFormSubmit(e){
  return withLocks_('onFormSubmit_master', function(){
    return withBackoff_('onFormSubmit_master', function(){
      try {
        try {
          var who = (Session && Session.getEffectiveUser && Session.getEffectiveUser().getEmail) ? Session.getEffectiveUser().getEmail() : 'unknown';
          logInfo_ && logInfo_('onFormSubmit_start', { by: who, hasEvent: !!e });
        } catch(_) {}

        if (typeof _onFormSubmit_impl === 'function') {
          return _onFormSubmit_impl(e);
        } else if (typeof onFormSubmit_main === 'function') {
          return onFormSubmit_main(e);
        } else {
          throw new Error('No underlying onFormSubmit implementation found. Rename your current handler to _onFormSubmit_impl or onFormSubmit_main.');
        }
      } catch (err) {
        logError && logError('onFormSubmit_master_err', err, {hasEvent: !!e});
        throw err;
      }
    }, 5, 500);
  }, 30000);
}






// 2) SINGLE dispatcher you point your installable trigger at.
function onAnyEdit(e){
  return withLocks_('onAnyEdit', function(){
    try{
      if (!e || !e.range || !e.range.getSheet) return;

      var sh = e.range.getSheet();
      var eventsTab = (CONFIG && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
      var onEvents = (sh.getName && sh.getName() === eventsTab);

      // 1) Grace (apply/undo) — short-circuit if handled
      if (typeof handleGraceEdit === 'function') {
        try { if (handleGraceEdit(e)) return; } catch(err){ logError && logError('handleGraceEdit', err); }
      }

      // 1.5) Row visibility (Active/Grace) — run early so UI feels instant
      if (typeof handleVisibilityOnEdit_ === 'function') {
        try { handleVisibilityOnEdit_(e); } catch(err){ logError && logError('handleVisibilityOnEdit_', err); } // <-- fixed label
      }

      // 2) Other handlers (defensive)
      if (typeof onEditNullify === 'function')         try{ onEditNullify(e); }         catch(err){ logError && logError('onEditNullify', err); }
      if (typeof onDirectorEdit_ === 'function')       try{ onDirectorEdit_(e); }       catch(err){ logError && logError('onDirectorEdit_', err); }
      if (typeof onEditMilestoneClaim_ === 'function') try{ onEditMilestoneClaim_(e);}  catch(err){ logError && logError('onEditMilestoneClaim_', err); }
      if (typeof onEditConfigGuard_ === 'function')    try{ onEditConfigGuard_(e); }    catch(err){ logError && logError('onEditConfigGuard_', err); }

      // 3) Create Milestone PDF — only when the Consequence Director cell changed
      try {
        if (typeof onEdit_CreatePdfWhenDirectorSet === 'function' && onEvents) {
          var hdrs = headers_(sh);
          var colName = hdrs[e.range.getColumn() - 1] || '';
          var directorHeader = (CONFIG && CONFIG.COLS && CONFIG.COLS.ConsequenceDirector) || 'Consequence Director';

          if (colName === directorHeader) {
            var newVal = (typeof e.value !== 'undefined') ? e.value : sh.getActiveCell().getDisplayValue();
            var oldVal = (typeof e.oldValue === 'undefined') ? '' : e.oldValue;
            if (newVal && newVal !== oldVal) {
              var cache = (CacheService && CacheService.getScriptCache) ? CacheService.getScriptCache() : null;
              var key = cache ? ('pdf-director-' + e.range.getRow()) : null;
              if (!cache || !cache.get(key)) {
                logInfo_ && logInfo_('onAnyEdit_calling_CreatePdfWhenDirectorSet', { row: e.range.getRow(), director: newVal });
                try { onEdit_CreatePdfWhenDirectorSet(e); } catch (pdfErr) { logError && logError('onEdit_CreatePdfWhenDirectorSet', pdfErr); }
                if (cache) cache.put(key, '1', 8);
              }
            }
          }
        }
      } catch (errPdf){ logError && logError('onAnyEdit_pdf_wrapper', errPdf); }

      // 4) Optional: notify leaders when claim column changes
      try {
        if (typeof onEdit_NotifyLeadersOnClaim_ === 'function' && onEvents) {
          var hdrs2 = headers_(sh);
          var colName2 = hdrs2[e.range.getColumn() - 1] || '';
          var claimHeader = (CONFIG && CONFIG.COLS && CONFIG.COLS.LeaderClaimed) || 'Leader Claimed';
          if (colName2 === claimHeader) {
            try { onEdit_NotifyLeadersOnClaim_(e); } catch (nErr) { logError && logError('onEdit_NotifyLeadersOnClaim_', nErr); }
          }
        }
      } catch (errN){ logError && logError('onAnyEdit_notify_wrapper', errN); }

      // 5) Positive ledger on manual edits (EventType/Points/Positive columns)
      try {
        if (onEvents && e.range.getRow) {
          var hdrs3 = headers_(sh);
          var edited = hdrs3[e.range.getColumn() - 1];
          var r = e.range.getRow();
          if (r >= 2) {
            var watchCols = [
              CONFIG.COLS && CONFIG.COLS.EventType,
              CONFIG.COLS && CONFIG.COLS.Points,
              CONFIG.COLS && CONFIG.COLS.PositiveActionType,
              'Positive Point',
              'Credit Type',
              'PositiveAction',
              'Positive Action'
            ].filter(Boolean);

            if (watchCols.indexOf(edited) !== -1) {
              var cache2 = (CacheService && CacheService.getScriptCache) ? CacheService.getScriptCache() : null;
              var key2 = cache2 ? ('pp-ledger-' + r) : null;
              if (!cache2 || !cache2.get(key2)) {
                var ok = false;
                try {
                  ok = recordPositiveCreditFromEvent(r);
                  if (cache2) cache2.put(key2, '1', 3);
                } catch(err2){
                  logError && logError('positive_ledger_from_edit', err2, { row: r, edited: edited });
                }
                logInfo_ && logInfo_('positive_ledger_after_edit', { row: r, edited: edited, did: ok });
              }
            }
          }
        }
      } catch (err3) {
        logError && logError('positive_ledger_edit_wrapper', err3);
      }

    } catch (err){
      logError && logError('onAnyEdit_top', err);
    }
  });
}







