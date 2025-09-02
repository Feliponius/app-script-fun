/*****************************************
 * trigger.gs — installable triggers for onAnyEdit architecture
 *****************************************/

/** Open target spreadsheet (strict) */
function _openTargetSs_() {
  if (!CONFIG || !CONFIG.SHEET_ID) {
    throw new Error('CONFIG.SHEET_ID missing; set it in Config.gs');
  }
  return SpreadsheetApp.openById(CONFIG.SHEET_ID);
}

/** Utility: does a trigger already exist? */
function _hasTrigger_(funcName, eventType) {
  return ScriptApp.getProjectTriggers().some(function(t){
    return t.getHandlerFunction() === funcName && t.getEventType() === eventType;
  });
}

/** Ensure ONE installable onEdit trigger for onAnyEdit(e) */
function ensureOnAnyEditTrigger_() {
  // remove any stray onEdit triggers that are NOT onAnyEdit
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getEventType() === ScriptApp.EventType.ON_EDIT &&
        t.getHandlerFunction() !== 'onAnyEdit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  if (_hasTrigger_('onAnyEdit', ScriptApp.EventType.ON_EDIT)) return;

  ScriptApp.newTrigger('onAnyEdit')
    .forSpreadsheet(_openTargetSs_())
    .onEdit()
    .create();
}

/** Ensure installable onFormSubmit trigger */
function ensureOnFormSubmitTrigger_() {
  if (_hasTrigger_('onFormSubmit', ScriptApp.EventType.ON_FORM_SUBMIT)) return;
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(_openTargetSs_())
    .onFormSubmit()
    .create();
}

/** Ensure a daily clock trigger (idempotent) */
function ensureDailyClockTrigger_(funcName, hour) {
  var exists = ScriptApp.getProjectTriggers().some(function(t){
    return t.getHandlerFunction() === funcName &&
           t.getEventType() === ScriptApp.EventType.CLOCK;
  });
  if (exists) return;

  ScriptApp.newTrigger(funcName)
    .timeBased()
    .atHour(Number(hour) || 3)    // default 3am local script TZ
    .everyDays(1)
    .inTimezone(Session.getScriptTimeZone())
    .create();
}

/** Optional helper: remove *all* other onEdit triggers (paranoid cleanup) */
function migrateTriggersToOnAnyEdit_() {
  var triggers = ScriptApp.getProjectTriggers();
  var kept = false;
  triggers.forEach(function(t){
    if (t.getEventType() === ScriptApp.EventType.ON_EDIT) {
      if (t.getHandlerFunction() === 'onAnyEdit' && !kept) {
        kept = true; // keep the first onAnyEdit trigger
      } else {
        ScriptApp.deleteTrigger(t);
      }
    }
  });
  if (!kept) {
    ensureOnAnyEditTrigger_();
  }
}

/** One‑shot: install all required triggers (safe to re‑run) */
function setup() {
  ensureOnAnyEditTrigger_();
  ensureOnFormSubmitTrigger_();

  // Nightly maintenance (only if implemented)
  if (typeof nightlyMaintenance_ === 'function') {
    ensureDailyClockTrigger_('nightlyMaintenance_', CONFIG && CONFIG.NIGHTLY_HOUR);
  }
  // Nightly form sync (only if implemented)
  if (typeof refreshFormLists_ === 'function') {
    ensureDailyClockTrigger_('refreshFormLists_', CONFIG && CONFIG.NIGHTLY_HOUR);
  }
}

/** Nice to have: keep triggers healthy whenever the spreadsheet opens */
function onOpen() {
  try { setup(); } catch (e) { try { Logger.log('onOpen/setup error: ' + e); } catch(_){} }
}

/** Optional: inspect current triggers in Logs */
function debugListTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    Logger.log('%s — %s', t.getHandlerFunction(), t.getEventType());
  });
}
