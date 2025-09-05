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

/**
 * handlePerformanceOnSubmit
 * Called when a Performance Issue event is submitted
 */
function handlePerformanceOnSubmit(rowIndex) {
  try {
    var events = sh_(CONFIG.TABS.EVENTS);
    var ctx = rowCtx_(events, rowIndex);
    
    // Set initial performance case stage
    if (CONFIG.COLS.PerfCaseStage) {
      ctx.set(CONFIG.COLS.PerfCaseStage, 'Open');
    }
    
    // Set initial pending status
    if (CONFIG.COLS.PendingStatus) {
      ctx.set(CONFIG.COLS.PendingStatus, 'Pending GA Assignment');
    }
    
    logInfo_ && logInfo_('handlePerformanceOnSubmit', { row: rowIndex });
  } catch (err) {
    logError && logError('handlePerformanceOnSubmit', err, { row: rowIndex });
  }
}

/**
 * handlePerformanceEdit
 * Handles performance-related edits in the Events sheet
 * Returns true if handled, false otherwise
 */
function handlePerformanceEditMain(e) {
  Logger.log('*** MAIN.JS VERSION OF handlePerformanceEdit CALLED ***');
  try {
    Logger.log('=== PERFORMANCE HANDLER DEBUG ===');
    
    if (!e || !e.range || !e.range.getSheet) {
      Logger.log('Performance: Missing e/range/sheet, returning false');
      return false;
    }
    
    var sh = e.range.getSheet();
    var eventsTab = (CONFIG && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
    if (sh.getName() !== eventsTab) {
      Logger.log('Performance: Not Events tab (' + sh.getName() + '), returning false');
      return false;
    }
    
    var row = e.range.getRow();
    if (row < 2) {
      Logger.log('Performance: Header row, returning false');
      return false;
    }
    
    var col = e.range.getColumn();
    var hdrs = headers_(sh);
    var colName = hdrs[col - 1] || '';
    var ctx = rowCtx_(sh, row);
    
    Logger.log('Performance: Row=' + row + ', Col=' + col + ', ColName="' + colName + '"');
    
    // Check if this is a performance-related edit
    var eventType = String(ctx.get(CONFIG.COLS.EventType) || '').trim();
    Logger.log('Performance: EventType="' + eventType + '"');
    Logger.log('Performance: EventType.toLowerCase()="' + eventType.toLowerCase() + '"');
    Logger.log('Performance: Checking if "' + eventType.toLowerCase() + '" !== "performance issue"');
    
    if (eventType.toLowerCase() !== 'performance issue') {
      Logger.log('Performance: Not performance issue, returning false');
      return false;
    }
    
    Logger.log('Performance: IS performance issue, checking columns...');
    
    // Only return true if we actually handle this specific column
    var handledColumns = [
      CONFIG.COLS.PerfDeadline, 
      CONFIG.COLS.PerfGADecision,
      CONFIG.COLS.ConsequenceDirector  // For performance consequence claiming
    ];
    
    Logger.log('Performance: Handled columns: ' + JSON.stringify(handledColumns));
    Logger.log('Performance: Current column: "' + colName + '"');
    Logger.log('Performance: Column in handled list: ' + (handledColumns.indexOf(colName) !== -1));
    
    if (handledColumns.indexOf(colName) === -1) {
      Logger.log('Performance: Column not handled, returning false');
      return false;
    }
    
    Logger.log('Performance: Column IS handled, checking cache...');
    
    var cache = (CacheService && CacheService.getScriptCache) ? CacheService.getScriptCache() : null;
    var key = cache ? ('perf-' + colName.toLowerCase().replace(/\s+/g, '-') + '-row-' + row) : null;
    if (cache && cache.get(key)) {
      Logger.log('Performance: Cache hit for key "' + key + '", returning true');
      return true;
    }
    
    Logger.log('Performance: No cache, proceeding with logic...');
    
    // Handle GA open + deadline entry
    if (colName === CONFIG.COLS.PerfDeadline) {
      var deadline = e.value;
      if (deadline && deadline !== e.oldValue) {
        // GA is now open, update stage and status
        if (CONFIG.COLS.PerfCaseStage) {
          ctx.set(CONFIG.COLS.PerfCaseStage, 'GA Open');
        }
        if (CONFIG.COLS.PendingStatus) {
          ctx.set(CONFIG.COLS.PendingStatus, 'Pending GA Decision');
        }
        if (CONFIG.COLS.PerfGADate) {
          ctx.set(CONFIG.COLS.PerfGADate, new Date());
        }
        
        if (cache && key) cache.put(key, '1', 10); // Set cache
        
        logInfo_ && logInfo_('handlePerformanceEdit_ga_opened', { row: row, deadline: deadline });
        Logger.log('Performance: Handled GA deadline, returning true');
        return true;
      }
    }
    
    // Handle GA decision (success/failure)
    if (colName === CONFIG.COLS.PerfGADecision) {
      var decision = String(e.value || '').trim().toLowerCase();
      var oldDecision = String(e.oldValue || '').trim().toLowerCase();
      
      if (decision && decision !== oldDecision) {
        var action = '';
        var templateId = '';
        
        if (decision === 'success' || decision === 'passed') {
          action = 'Return to Good Standing';
          templateId = CONFIG.TEMPLATES.PERF_GA_SUCCESS;
          
          // Update performance case stage
          if (CONFIG.COLS.PerfCaseStage) {
            ctx.set(CONFIG.COLS.PerfCaseStage, 'GA Success');
          }
          if (CONFIG.COLS.PendingStatus) {
            ctx.set(CONFIG.COLS.PendingStatus, 'Completed');
          }
          
        } else if (decision === 'failure' || decision === 'failed') {
          // Check if this is a greater reduction case
          var infraction = String(ctx.get(CONFIG.COLS.Infraction) || '').toLowerCase();
          if (infraction.indexOf('greater reduction') !== -1) {
            action = 'Greater Reduction';
            templateId = CONFIG.TEMPLATES.PER_GREATER_REDUCTION;
          } else {
            action = 'Performance Failure Termination';
            templateId = CONFIG.TEMPLATES.PERF_TERMINATION;
          }
          
          // Update performance case stage
          if (CONFIG.COLS.PerfCaseStage) {
            ctx.set(CONFIG.COLS.PerfCaseStage, 'GA Failure');
          }
          if (CONFIG.COLS.PendingStatus) {
            ctx.set(CONFIG.COLS.PendingStatus, 'Pending Termination');
          }
        }
        
        // Create consequence PDF for the decision
        if (action && templateId) {
          try {
            var pdfId = createConsequencePdf_(row, action);
            logInfo_ && logInfo_('handlePerformanceEdit_decision_pdf', { 
              row: row, 
              decision: decision, 
              action: action, 
              pdfId: pdfId 
            });
          } catch (pdfErr) {
            logError && logError('handlePerformanceEdit_decision_pdf_err', pdfErr, { 
              row: row, 
              decision: decision, 
              action: action 
            });
          }
        }
        
        // Update Perf NoPickup dates if applicable
        if (decision === 'failure' || decision === 'failed') {
          var today = new Date();
          var endDate = new Date(today);
          endDate.setDate(endDate.getDate() + 30); // 30 days from today
          
          if (CONFIG.COLS.Perf_NoPickup) {
            ctx.set(CONFIG.COLS.Perf_NoPickup, today);
          }
          if (CONFIG.COLS.Perf_NoPickup_End) {
            ctx.set(CONFIG.COLS.Perf_NoPickup_End, endDate);
          }
          if (CONFIG.COLS.Per_NoPickup_Active) {
            ctx.set(CONFIG.COLS.Per_NoPickup_Active, true);
          }
        }
        
        if (cache && key) cache.put(key, '1', 10); // Set cache
        Logger.log('Performance: Handled GA decision, returning true');
        return true;
      }
    }

    // Handle Consequence Director claiming
    if (colName === CONFIG.COLS.ConsequenceDirector) {
      var newDirector = e.value;
      var oldDirector = e.oldValue;
      
      if (newDirector && newDirector !== oldDirector) {
        var infraction = String(ctx.get(CONFIG.COLS.Infraction) || '').trim();
        
        if (infraction.toLowerCase().indexOf('growth plan') !== -1) {
          // Handle Growth Plan claiming logic
          // ... (transfer the growth plan logic from the deleted function)
        }
        // Add other performance consequence types as needed
        
        if (cache && key) cache.put(key, '1', 10);
        return true;
      }
    }
    
    Logger.log('Performance: No action taken, returning false');
    return false;
    
  } catch (err) {
    Logger.log('Performance: Error - ' + err);
    logError && logError('handlePerformanceEdit', err, { 
      row: e.range ? e.range.getRow() : 'unknown',
      col: e.range ? e.range.getColumn() : 'unknown'
    });
    return false;
  }
}


// 2) SINGLE dispatcher you point your installable trigger at.
function onAnyEdit(e){
  return withLocks_('onAnyEdit', function(){
    try{
      if (!e || !e.range || !e.range.getSheet) return;

      var sh = e.range.getSheet();
      var eventsTab = (CONFIG && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
      var onEvents = (sh.getName && sh.getName() === eventsTab);
      
      var row = e.range.getRow();
      var col = e.range.getColumn();
      var hdrs = headers_(sh);
      var colName = hdrs[col - 1] || '';
      
      Logger.log('=== onAnyEdit DEBUG ===');
      Logger.log('Sheet: ' + sh.getName() + ', Row: ' + row + ', Col: ' + col + ', ColName: "' + colName + '"');
      Logger.log('Value: "' + (e.value || '') + '", OldValue: "' + (e.oldValue || '') + '"');
      Logger.log('onEvents: ' + onEvents);

      // 1) Grace (apply/undo) — short-circuit ONLY if Grace column edited
      if (typeof handleGraceEdit === 'function') {
        Logger.log('Calling handleGraceEdit...');
        try { 
          var graceResult = handleGraceEdit(e);
          Logger.log('handleGraceEdit returned: ' + graceResult);
          if (graceResult) {
            Logger.log('=== SHORT-CIRCUITED BY GRACE ===');
            return;
          }
        } catch(err){ 
          Logger.log('handleGraceEdit error: ' + err);
          logError && logError('handleGraceEdit', err); 
        }
      }

      // 2) Performance edit handling — short-circuit ONLY if Performance column edited
      if (typeof handlePerformanceEditMain === 'function') {
        Logger.log('Calling handlePerformanceEditMain...');
        try { 
          var perfResult = handlePerformanceEditMain(e);
          Logger.log('handlePerformanceEditMain returned: ' + perfResult);
          if (perfResult) {
            Logger.log('=== SHORT-CIRCUITED BY PERFORMANCE ===');
            return;
          }
        } catch(err){ 
          Logger.log('handlePerformanceEditMain error: ' + err);
          logError && logError('handlePerformanceEditMain', err); 
        }
      }

      // 3) Row visibility (non-blocking)
      if (typeof handleVisibilityOnEdit_ === 'function') {
        Logger.log('Calling handleVisibilityOnEdit_...');
        try { handleVisibilityOnEdit_(e); } catch(err){ logError && logError('handleVisibilityOnEdit_', err); }
      }

      // 4) Milestone claiming and other handlers
      Logger.log('Calling other handlers...');
      if (typeof onEditNullify === 'function')         try{ onEditNullify(e); }         catch(err){ logError && logError('onEditNullify', err); }
      if (typeof onDirectorEdit_ === 'function')       try{ onDirectorEdit_(e); }       catch(err){ logError && logError('onDirectorEdit_', err); }
      
      if (typeof onEditMilestoneClaim_ === 'function') {
        Logger.log('Calling onEditMilestoneClaim_...');
        try{ onEditMilestoneClaim_(e);}  catch(err){ logError && logError('onEditMilestoneClaim_', err); }
      }
      
      if (typeof onEditConfigGuard_ === 'function')    try{ onEditConfigGuard_(e); }    catch(err){ logError && logError('onEditConfigGuard_', err); }

      // 5) PDF generation path
      try {
        if (typeof onEdit_CreatePdfWhenDirectorSet === 'function' && onEvents) {
          var directorHeader = (CONFIG && CONFIG.COLS && CONFIG.COLS.ConsequenceDirector) || 'Consequence Director';
          if (colName === directorHeader) {
            Logger.log('Director column edit detected, checking PDF generation...');
            var newVal = (typeof e.value !== 'undefined') ? e.value : sh.getActiveCell().getDisplayValue();
            var oldVal = (typeof e.oldValue === 'undefined') ? '' : e.oldValue;
            if (newVal && newVal !== oldVal) {
              var cache = (CacheService && CacheService.getScriptCache) ? CacheService.getScriptCache() : null;
              var key = cache ? ('pdf-director-' + e.range.getRow()) : null;
              if (!cache || !cache.get(key)) {
                Logger.log('Calling onEdit_CreatePdfWhenDirectorSet...');
                logInfo_ && logInfo_('onAnyEdit_calling_CreatePdfWhenDirectorSet', { row: e.range.getRow(), director: newVal });
                try { onEdit_CreatePdfWhenDirectorSet(e); } catch (pdfErr) { logError && logError('onEdit_CreatePdfWhenDirectorSet', pdfErr); }
                if (cache) cache.put(key, '1', 8);
              } else {
                Logger.log('PDF generation skipped due to cache');
              }
            }
          }
        }
      } catch (errPdf){ logError && logError('onAnyEdit_pdf_wrapper', errPdf); }

      Logger.log('=== onAnyEdit COMPLETE ===');

    } catch(err){ 
      Logger.log('onAnyEdit top-level error: ' + err);
      logError && logError('onAnyEdit', err); 
    }
  });
}







