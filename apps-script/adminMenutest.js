/*****************************
 * Employee / Admin Utilities
 * Consolidated for template-preserving Employee History PDFs
 *
 * Expectations:
 * - CONFIG.TABS.EVENTS points to your Events sheet
 * - CONFIG.DEST_FOLDER_ID (optional) destination folder for PDFs
 * - CONFIG.TEMPLATES.EMP_HISTORY (optional) Google Doc template id
 * - Optional helpers: sh_(tabName), logError(), logAudit()
 *****************************/

/* ===========================
 * Lookups & Sidebar helpers
 * =========================== */

/**
 * getEmployeeList()
 * Returns a deduplicated, sorted list of employee names from the Events sheet.
 */
function getEmployeeList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventsTabName = (typeof CONFIG !== 'undefined' && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
    const sheet = (typeof sh_ === 'function') ? sh_(eventsTabName) : (ss.getSheetByName(eventsTabName) || ss.getSheetByName('Events'));
    if (!sheet) {
      Logger.log('getEmployeeList: Events sheet not found: %s', eventsTabName);
      return [];
    }

    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    if (lastCol === 0 || lastRow <= 1) return [];

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
    const headerCandidates = ['Employee', 'Employee Name', 'Name'];
    let colIdx = -1;
    for (const h of headerCandidates) {
      const idx = headers.indexOf(h);
      if (idx !== -1) { colIdx = idx; break; }
    }
    if (colIdx === -1) {
      Logger.log('getEmployeeList: "Employee" header not found on %s; headers=%s', sheet.getName(), JSON.stringify(headers));
      return [];
    }

    const range = sheet.getRange(2, colIdx + 1, Math.max(0, lastRow - 1), 1);
    const values = range.getValues().map(r => String(r[0] || '').trim()).filter(v => v);

    return Array.from(new Set(values)).sort((a, b) => a.localeCompare(b));
  } catch (e) {
    Logger.log('getEmployeeList ERROR: ' + String(e));
    return [];
  }
}

/**
 * getEmployeeSuggestions(query, limit)
 * Case-insensitive substring matches from getEmployeeList.
 */
function getEmployeeSuggestions(query, limit) {
  try {
    query = String(query || '').trim().toLowerCase();
    limit = Number(limit) || 50;
    const all = getEmployeeList();
    if (!query) return all.slice(0, limit);
    return all.filter(name => name.toLowerCase().includes(query)).slice(0, limit);
  } catch (e) {
    Logger.log('getEmployeeSuggestions error: ' + String(e));
    return [];
  }
}

/**
 * clientLookupEmployeeSafe(name)
 * Returns a safe envelope (resultJson string).
 */
function clientLookupEmployeeSafe(name) {
  try {
    Logger.log('clientLookupEmployeeSafe called with: "%s"', name);
    const result = lookupEmployee(name);

    let resultJson;
    try {
      resultJson = JSON.stringify(result);
      Logger.log('clientLookupEmployeeSafe: stringify OK (len=%s)', resultJson.length);
    } catch (err) {
      Logger.log('clientLookupEmployeeSafe: JSON.stringify FAILED: %s', String(err));
      try {
        const summary = {
          type: typeof result,
          isNull: result === null,
          keys: (result && typeof result === 'object' && !Array.isArray(result)) ? Object.keys(result).slice(0, 20) : null,
          note: 'Original object not JSON-serializable; returning summary'
        };
        resultJson = JSON.stringify(summary);
      } catch (err2) {
        resultJson = JSON.stringify({ error: 'could_not_serialize', message: String(err2) });
      }
    }

    return {
      ok: true,
      calledWith: String(name || ''),
      scriptId: ScriptApp.getScriptId(),
      ts: new Date().toISOString(),
      resultJson: resultJson
    };
  } catch (e) {
    Logger.log('clientLookupEmployeeSafe ERROR: ' + String(e));
    return { ok: false, error: String(e) };
  }
}

/**
 * clientLookupEmployee(name)
 * Wrapper returning raw result (useful for logs).
 */
function clientLookupEmployee(name) {
  try {
    Logger.log('clientLookupEmployee called with: "%s"', name);
    const result = lookupEmployee(name);
    try { Logger.log('clientLookupEmployee: result = %s', JSON.stringify(result)); } catch (e) { Logger.log('clientLookupEmployee: stringify failed: ' + String(e)); }
    return {
      wrapper: true,
      calledWith: String(name || ''),
      resultType: typeof result,
      resultIsNull: result === null,
      resultIsUndefined: typeof result === 'undefined',
      result: result
    };
  } catch (err) {
    Logger.log('clientLookupEmployee ERROR: ' + String(err));
    return { wrapper: true, error: String(err) };
  }
}

/**
 * Minimal probe for connectivity.
 */
function debugPing(name) {
  try {
    Logger.log('debugPing called with: "%s"', name);
    return {
      ok: true,
      name: String(name || ''),
      ts: new Date().toISOString(),
      scriptId: ScriptApp.getScriptId(),
      ssUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl()
    };
  } catch (err) {
    Logger.log('debugPing ERROR: ' + String(err));
    return { ok: false, error: String(err) };
  }
}

/* ===========================
 * Menu
 * =========================== */
function onOpen() { 
  try { applyEventsAutoHide_(); } catch(err) { try { logError && logError('applyEventsAutoHide_onOpen', err); } catch(_){} }
  showMenuSafe();
  try { addAdminBackfillMenu_(); } catch(_) {}
}

function showMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Admin Tools')
    .addSubMenu(
      ui.createMenu('🔧 System Tools')
        .addItem('Refresh Form Lists Now', 'refreshFormLists')
        .addItem('Restore Validations Only', 'restoreEventsValidations')
        .addItem('Clear Last Row', 'clearLastRow')
        .addItem('Back Fill Pdfs','addAdminBackfillMenu_')
        .addItem('Apply Auto-Hide Now', 'applyEventsAutoHide_')
        .addItem('🚨 Reset Events Board (TEST ONLY)', 'resetEventsBoard')
    )
    .addSubMenu(
      ui.createMenu('👤 Employee Tools')
        .addItem('Employee Lookup', 'employeeLookup')
    )
    .addSubMenu(
      ui.createMenu('📊 Reports & Monitoring')
        .addItem('Probation Watchlist', 'probationWatchlist')
    )
    .addToUi();
}

/* ===========================
 * Employee tools
 * =========================== */

function employeeLookup() {
  const html = HtmlService.createHtmlOutputFromFile('employeeLookup')
    .setTitle('Employee Lookup')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * lookupEmployee(name)
 * Returns structured result with recent events & milestone info.
 */
function lookupEmployee(name) {
  try {
    Logger.log('lookupEmployee called with: %s', name);
    const s = (typeof sh_ === 'function') ? sh_(CONFIG.TABS.EVENTS) : SpreadsheetApp.getActiveSpreadsheet().getSheetByName((CONFIG && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events');
    if (!s) return { error: 'Events sheet not found' };

    const data = s.getDataRange().getValues();
    if (!data || data.length < 2) return { error: 'Events sheet appears empty' };

    const headers = data[0].map(h => (h || '').toString().trim());
    const rows = data.slice(1);

    const COL_EMPLOYEE       = headers.indexOf('Employee');
    const COL_DATE           = headers.indexOf('IncidentDate');
    const COL_EVENTTYPE      = headers.indexOf('EventType');
    const COL_INFRACTION     = headers.indexOf('Infraction');
    const COL_POINTS         = headers.indexOf('Points');
    const COL_ROLLING_TOTAL  = headers.indexOf('PointsRolling (Effective)');
    const COL_MILESTONE      = headers.indexOf('Milestone');
    const COL_GRACENOTES     = headers.indexOf('GraceNotes');
    const COL_POSITIVEACTION = headers.indexOf('PositiveAction');
    const COL_MILESTONE_DATE = headers.indexOf('MilestoneDate');
    const COL_PROB_ENDS      = headers.indexOf('Probation_End');

    if (COL_EMPLOYEE === -1) return { error: 'Employee column not found' };

    const target = String(name || '').trim().toLowerCase();
    if (!target) return { error: 'No name provided' };

    const exactRows = rows.filter(r => {
      if (!r || r.length <= COL_EMPLOYEE) return false;
      const emp = r[COL_EMPLOYEE];
      if (!emp && emp !== 0) return false;
      return String(emp).trim().toLowerCase() === target;
    });

    Logger.log('Found %s exact rows for "%s"', exactRows.length, name);

    if (exactRows.length === 0) {
      const nameCandidates = rows.map(r => (r && r.length > COL_EMPLOYEE) ? r[COL_EMPLOYEE] : null).filter(Boolean);
      const uniqueNames = Array.from(new Set(nameCandidates));
      const suggestions = uniqueNames.filter(n => n.toLowerCase().includes(target)).slice(0, 20);
      if (suggestions.length) return { error: 'No exact match', suggestions };
      return { error: 'No records found' };
    }

    function parseDateVal(v) {
      if (!v && v !== 0) return null;
      if (v instanceof Date) { if (isNaN(v)) return null; return v; }
      try { const d = new Date(v); if (isNaN(d)) return null; return d; } catch (e) { return null; }
    }
    function getMilestoneTextFromRow(r) {
      let text = '';
      if (COL_GRACENOTES !== -1 && r.length > COL_GRACENOTES && r[COL_GRACENOTES]) text = String(r[COL_GRACENOTES]).trim();
      if (!text && COL_MILESTONE !== -1 && r.length > COL_MILESTONE && r[COL_MILESTONE]) text = String(r[COL_MILESTONE]).trim();
      if (!text && COL_POSITIVEACTION !== -1 && r.length > COL_POSITIVEACTION && r[COL_POSITIVEACTION]) text = String(r[COL_POSITIVEACTION]).trim();
      return text || '';
    }
    function shortenMilestone(text) {
      if (!text) return '';
      const s = String(text).trim();
      const m = s.match(/(\d+\s*[-]?\s*(Day|Week|Month|Days|Weeks|Months)\s+(Suspension|Suspensions|Probation))/i);
      if (m && m[1]) {
        return m[1].replace(/Suspensions?/i, 'Sus').replace(/Suspension/i, 'Sus').replace(/Probation/i, 'Prob').trim();
      }
      const segments = s.split(/\+| - |:|\(|\/|—/);
      let first = segments[0].trim();
      if (first.length > 18) {
        first = first.replace(/Suspension/i, 'Sus').replace(/Final Warning/i, 'FinalW').replace(/Written Warning/i, 'Warn');
        if (first.length > 18) first = first.slice(0, 18).trim() + '...';
      } else {
        first = first.replace(/Suspension/i, 'Sus');
      }
      return first;
    }

    const milestoneCandidates = [];
    for (let i = 0; i < exactRows.length; i++) {
      const r = exactRows[i];
      const ev = (COL_EVENTTYPE !== -1 && r.length > COL_EVENTTYPE) ? r[COL_EVENTTYPE] : '';
      const pts = (COL_POINTS !== -1 && r.length > COL_POINTS) ? r[COL_POINTS] : null;
      const milDateVal = (COL_MILESTONE_DATE !== -1 && r.length > COL_MILESTONE_DATE) ? r[COL_MILESTONE_DATE] : null;
      const parsedMilDate = parseDateVal(milDateVal) || null;
      const isLikelyMilestone = (ev && String(ev).toString().toLowerCase().includes('milestone')) ||
                                Boolean(parsedMilDate) ||
                                ((pts === 0 || String(pts) === '0') && ((COL_INFRACTION === -1) || !r[COL_INFRACTION] || String(r[COL_INFRACTION]).trim() === ''));
      if (isLikelyMilestone) {
        milestoneCandidates.push({ row: r, milDate: parsedMilDate, text: getMilestoneTextFromRow(r) });
      }
    }

    let chosenMilestone = null;
    if (milestoneCandidates.length) {
      milestoneCandidates.sort((a, b) => {
        if (a.milDate && b.milDate) return b.milDate - a.milDate;
        if (a.milDate && !b.milDate) return -1;
        if (!a.milDate && b.milDate) return 1;
        return 0;
      });
      chosenMilestone = milestoneCandidates[0];
    }

    const lastRow = exactRows[exactRows.length - 1];
    let milestoneText = '';
    if (chosenMilestone && chosenMilestone.text) milestoneText = String(chosenMilestone.text).trim();
    else milestoneText = getMilestoneTextFromRow(lastRow);

    const recentRows = exactRows.slice(-5);
    const recentEvents = recentRows.map(r => {
      const dateVal = (COL_DATE !== -1 && r.length > COL_DATE) ? r[COL_DATE] : null;
      const inf = (COL_INFRACTION !== -1 && r.length > COL_INFRACTION) ? r[COL_INFRACTION] : '';
      const pts = (COL_POINTS !== -1 && r.length > COL_POINTS) ? r[COL_POINTS] : 0;
      const ev = (COL_EVENTTYPE !== -1 && r.length > COL_EVENTTYPE) ? r[COL_EVENTTYPE] : '';

      const isMilestoneRow = (ev && String(ev).toLowerCase().includes('milestone')) ||
                             ((pts === 0 || String(pts) === '0') && (!inf || String(inf).trim() === '')) ||
                             (COL_MILESTONE_DATE !== -1 && r.length > COL_MILESTONE_DATE && parseDateVal(r[COL_MILESTONE_DATE]) !== null);

      const thisMilestoneText = isMilestoneRow ? getMilestoneTextFromRow(r) : '';
      const thisMilestoneShort = thisMilestoneText ? shortenMilestone(thisMilestoneText) : '';

      return {
        date: dateVal,
        infraction: inf || '',
        points: (pts === '' || pts === null) ? 0 : pts,
        isMilestone: Boolean(isMilestoneRow),
        milestoneText: thisMilestoneText || '',
        milestoneShort: thisMilestoneShort || ''
      };
    });

    let triggeringEvent = null;
    for (let i = exactRows.length - 1; i >= 0; i--) {
      const r = exactRows[i];
      const pts = (COL_POINTS !== -1 && r.length > COL_POINTS) ? r[COL_POINTS] : null;
      if (pts && Number(pts) > 0) {
        triggeringEvent = {
          date: (COL_DATE !== -1 && r.length > COL_DATE) ? r[COL_DATE] : null,
          infraction: (COL_INFRACTION !== -1 && r.length > COL_INFRACTION) ? r[COL_INFRACTION] : '',
          points: pts
        };
        break;
      }
    }

    const probationEndsRaw = (COL_PROB_ENDS !== -1 && lastRow.length > COL_PROB_ENDS) ? lastRow[COL_PROB_ENDS] : '';

    return {
      name,
      rollingTotal: lastRow[COL_ROLLING_TOTAL],
      milestone: milestoneText || '',
      milestoneShort: milestoneText ? shortenMilestone(milestoneText) : '',
      probationEnds: probationEndsRaw,
      recentEvents,
      triggeringEvent,
      diagnostic: {
        headers,
        colIndexes: {
          employee: COL_EMPLOYEE,
          date: COL_DATE,
          eventType: COL_EVENTTYPE,
          infraction: COL_INFRACTION,
          points: COL_POINTS,
          rollingTotal: COL_ROLLING_TOTAL,
          milestone: COL_MILESTONE,
          graceNotes: COL_GRACENOTES,
          positiveAction: COL_POSITIVEACTION,
          milestoneDate: COL_MILESTONE_DATE,
          probEnds: COL_PROB_ENDS
        },
        exactRowsFound: exactRows.length
      }
    };
  } catch (e) {
    try { if (typeof logError === 'function') logError('lookupEmployee', e, { name }); else Logger.log(e); } catch(ignore){}
    return { error: e.message || String(e) };
  }
}

/**
 * Entry point for sidebar menu click.
 * If active selection is an employee in Events, offers inline generation first.
 */
function employeeHistoryPdf() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheetName = sheet ? sheet.getName() : '';
    var eventsTabName = (typeof CONFIG !== 'undefined' && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
    var onEventsSheet = sheetName === eventsTabName;

    var headers = [];
    try { headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h||'').trim(); }); } catch (e) { headers = []; }

    var employeeColIndex = headers.indexOf('Employee'); // 0-based
    var activeRange = sheet.getActiveRange();
    var activeValue = '';
    if (activeRange && activeRange.getNumRows && activeRange.getNumColumns) {
      if (activeRange.getNumRows() === 1 && activeRange.getNumColumns() === 1) {
        activeValue = String(activeRange.getValue() || '').trim();
      }
    }

    if (onEventsSheet && activeValue) {
      var activeCol = activeRange.getColumn(); // 1-based
      var inEmployeeColumn = (employeeColIndex !== -1) && (activeCol === (employeeColIndex + 1));
      var allowInline = inEmployeeColumn;

      if (!allowInline && employeeColIndex !== -1) {
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          var values = sheet.getRange(2, employeeColIndex + 1, Math.max(0, lastRow - 1), 1).getValues().map(function(r){ return String(r[0] || '').trim().toLowerCase(); });
          if (values.indexOf(activeValue.toLowerCase()) !== -1) allowInline = true;
        }
      }

      if (allowInline) {
        var confirm = ui.alert('Generate Employee History PDF', 'Generate history PDF for "' + activeValue + '" now?', ui.ButtonSet.YES_NO);
        if (confirm === ui.Button.YES) {
          var result = generateEmployeeHistoryPdf(activeValue);
          if (result && result.ok) {
            ui.alert('PDF generated for ' + activeValue + '\n\nOpen: ' + (result.url || result.pdfUrl));
            if (typeof logAudit === 'function') logAudit('system','employeeHistoryPdf_inline', Session.getActiveUser ? Session.getActiveUser().getEmail() : 'unknown', { name: activeValue, pdfId: result.id || result.pdfId });
          } else {
            ui.alert('Error generating PDF: ' + (result && result.error ? result.error : 'Unknown error'));
          }
          return;
        } else {
          return;
        }
      }
    }

    var html = HtmlService.createHtmlOutputFromFile('employeeLookup').setTitle('Employee Lookup').setWidth(400);
    ui.showSidebar(html);

  } catch (e) {
    try { if (typeof logError === 'function') logError('employeeHistoryPdf', e); else Logger.log(e); } catch(ignore){}
    try { SpreadsheetApp.getUi().alert('Error in employeeHistoryPdf: ' + (e && e.message ? e.message : String(e))); } catch(ignore){}
  }
}

/* ===========================
 * Employee History PDF (Template-preserving)
 * =========================== */

function generateEmployeeHistoryPdf(name) {
  return withBackoff_('generateEmployeeHistoryPdf', function(){
    try {
      if (!name) throw new Error('No employee name provided');

      // Call the object-returning function directly
      var out = (typeof createEmployeeHistoryPDF_ === 'function')
        ? createEmployeeHistoryPDF_(name)
        : (typeof createEmployeeHistoryPdf_ === 'function' ? createEmployeeHistoryPdf_(name) : null);

      if (!out) throw new Error('No PDF generator found');

      // Accept either a string id or an object { ok, pdfId, url, pdfUrl }
      var pdfId = (typeof out === 'string') ? out : (out && out.pdfId);
      if (!pdfId) throw new Error((out && out.error) ? out.error : 'PDF generation failed (no pdfId)');

      var url = (out && out.url) ? out.url : ('https://drive.google.com/file/d/' + pdfId + '/view');

      try { if (typeof logAudit === 'function') logAudit('system','generateEmployeeHistoryPdf', null, { name: name, pdfId: pdfId }); } catch(_){}
      return { ok: true, url: url, pdfUrl: url, id: pdfId, pdfId: pdfId };
    } catch (e) {
      try { if (typeof logError === 'function') logError('generateEmployeeHistoryPdf', e, { name: name }); } catch(_){}
      return { ok: false, error: (e && e.message) ? e.message : String(e) };
    }
  }, 4, 250);
}


/* ===========================
 * Table helpers (template-preserving)
 * =========================== */

/**
 * Replace all instances of {{TOKEN}} (with optional spaces inside the braces).
 */
function replaceBodyToken_(body, token, replacement) {
  try {
    if (!body || !token) return;
    var rx = '\\{\\{\\s*' + token.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&') + '\\s*\\}\\}';
    body.replaceText(rx, replacement != null ? String(replacement) : '');
  } catch (e) {
    Logger.log('replaceBodyToken_ error: %s', String(e));
  }
}
/* ===========================
 * System tools
 * =========================== */

function nullifyEvent() { SpreadsheetApp.getUi().alert('❌ Nullify Event placeholder'); }
function nullifyMilestone() { SpreadsheetApp.getUi().alert('❌ Nullify Milestone placeholder'); }
function probationWatchlist() { SpreadsheetApp.getUi().alert('📊 Probation Watchlist placeholder'); }

function refreshFormLists() {
  try {
    const form = FormApp.openById(CONFIG.FORM_ID);

    const rubric = sh_(CONFIG.TABS.RUBRIC);
    const vals = rubric.getRange(2, 1, Math.max(rubric.getLastRow() - 1, 0), 4).getValues();
    let infractions = vals.map(r => {
      const infraction = String(r[0] || '').trim();
      const points     = String(r[1] || '').trim();
      const policy     = String(r[3] || '').trim();
      if (!infraction) return null;
      let label = '';
      if (policy) label += `[${policy}] `;
      label += infraction;
      if (points) label += ` — ${points} pts`;
      return label;
    }).filter(Boolean);
    infractions = Array.from(new Set(infractions));

    const infractionItem = form.getItems().find(it => it.getTitle() === 'Infraction');
    if (infractionItem) infractionItem.asListItem().setChoiceValues(infractions);

    const dv = sh_(CONFIG.TABS.DATA_VALIDATION);
    let employees = dv.getRange(2, 1, Math.max(dv.getLastRow() - 1, 0), 1).getValues()
      .map(r => String(r[0] || '').trim()).filter(Boolean);
    employees = Array.from(new Set(employees));

    let managers = dv.getRange(2, 2, Math.max(dv.getLastRow() - 1, 0), 1).getValues()
      .map(r => String(r[0] || '').trim()).filter(Boolean);
    managers = Array.from(new Set(managers));

    const empItem = form.getItems().find(it => it.getTitle() === 'Employee');
    if (empItem) empItem.asListItem().setChoiceValues(employees);

    const mgrItem = form.getItems().find(it => it.getTitle() === 'Lead');
    if (mgrItem) mgrItem.asListItem().setChoiceValues(managers);

    try { logAudit('system', 'refreshFormLists', null, { infractions: infractions.length, employees: employees.length, managers: managers.length }); } catch(e){}

    SpreadsheetApp.getUi().alert(`✅ Form updated!\nInfractions: ${infractions.length}\nEmployees: ${employees.length}\nLeads: ${managers.length}`);
  } catch (e) {
    try { logError('refreshFormLists', e); } catch(ignore){ Logger.log(e); }
    SpreadsheetApp.getUi().alert('❌ Error while refreshing form lists: ' + e);
  }
}

function restoreEventsValidations() {
  try {
    const s = sh_(CONFIG.TABS.EVENTS);
    const map = headerIndexMap_(s);                // header → 1-based column index
    const rows = Math.max(1, s.getMaxRows() - 1);  // from row 2 down
    const directors = (CONFIG && CONFIG.DIRECTORS) ?
      CONFIG.DIRECTORS :
      ['Philip Pixler', 'Kristie Wright', 'Quana Slaughter', 'Cristian Zavala'];

    // Helpers
    function colByName(name){
      // try exact header, then CONFIG.COLS alias if present
      return map[name] || (CONFIG && CONFIG.COLS && map[CONFIG.COLS[name]]) || 0;
    }
    function wholeColRange(col){
      return s.getRange(2, col, rows, 1);
    }

    // Build validators
    const vCheckbox = SpreadsheetApp.newDataValidation()
      .requireCheckbox().build();
    const vDirectors = SpreadsheetApp.newDataValidation()
      .requireValueInList(directors, true) // show dropdown
      .setAllowInvalid(false)
      .build();

    // Targets (header → validator)
    const targets = [
      { header: 'Grace',                 dv: vCheckbox },
      { header: 'Graced By',             dv: vDirectors },
      { header: 'Consequence Director',  dv: vDirectors },
      { header: 'Nullify',               dv: vCheckbox },
      { header: 'Nullified By',          dv: vDirectors },
    ];

    const applied = [];
    const missing = [];

    targets.forEach(t => {
      const c = colByName(t.header);
      if (c) {
        wholeColRange(c).setDataValidation(t.dv);
        applied.push(t.header);
      } else {
        missing.push(t.header);
      }
    });

    try {
      logAudit('system', 'restoreEventsValidations', null, { applied, missing });
    } catch (_) {}

    var msg = '✅ Data validations restored for: ' + (applied.join(', ') || '(none)');
    if (missing.length) msg += '\n⚠️ Missing headers (not updated): ' + missing.join(', ');
    SpreadsheetApp.getUi().alert(msg);

  } catch (e) {
    try { logError('restoreEventsValidations', e); } catch(_) {}
    SpreadsheetApp.getUi().alert('❌ Error while restoring validations: ' + e);
  }
}


function resetEventsBoard() {
  try {
    const s = sh_(CONFIG.TABS.EVENTS);
    const lastRow = s.getLastRow();
    const lastCol = s.getLastColumn();

    if (lastRow > 1) {
      s.getRange(2, 1, lastRow - 1, lastCol).clearContent().clearDataValidations();
    }
    restoreEventsValidations();

    try { logAudit('system', 'resetEventsBoard', null, { cleared: true }); } catch(e){}
    SpreadsheetApp.getUi().alert('✅ Events board wiped. Headers preserved, validations reset.');
  } catch (e) {
    try { logError('resetEventsBoard', e); } catch(ignore){ Logger.log(e); }
    SpreadsheetApp.getUi().alert('❌ Error while resetting Events board: ' + e);
  }
}

/**
 * clearLastRow()
 * Finds the last meaningful row, clears it, then reapplies validations for C/D/S/U.
 */
function clearLastRow() {
  try {
    const s = sh_(CONFIG.TABS.EVENTS);
    const lastRowReported = s.getLastRow();
    const lastCol = s.getLastColumn();

    if (lastRowReported <= 1) {
      SpreadsheetApp.getUi().alert('⚠️ No rows to clear — only headers exist.');
      return;
    }

    const startRow = 2;
    const rowCount = Math.max(0, lastRowReported - 1);
    if (rowCount === 0) {
      SpreadsheetApp.getUi().alert('⚠️ No rows to clear — only headers exist.');
      return;
    }
    const block = s.getRange(startRow, 1, rowCount, lastCol).getValues();
    const headers = s.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());

    const priorityCols = ['Timestamp', 'Employee', 'IncidentDate', 'Infraction'];

    function lastNonEmptyIndexInCol(colIndexZeroBased) {
      if (colIndexZeroBased === -1) return -1;
      for (let i = block.length - 1; i >= 0; i--) {
        const cell = block[i][colIndexZeroBased];
        if (String(cell !== null && typeof cell !== 'undefined' ? cell : '').trim() !== '') return i;
      }
      return -1;
    }

    let lastFilledIndex = -1;
    for (const colName of priorityCols) {
      const idx = headers.indexOf(colName);
      if (idx !== -1) {
        const lastIdxForCol = lastNonEmptyIndexInCol(idx);
        if (lastIdxForCol > lastFilledIndex) lastFilledIndex = lastIdxForCol;
      }
    }

    if (lastFilledIndex === -1) {
      for (let i = block.length - 1; i >= 0; i--) {
        const row = block[i];
        const hasValue = row.some(function(cell) {
          return String(cell !== null && typeof cell !== 'undefined' ? cell : '').trim() !== '';
        });
        if (hasValue) { lastFilledIndex = i; break; }
      }
    }

    if (lastFilledIndex === -1) {
      SpreadsheetApp.getUi().alert('⚠️ No filled rows found to clear (only headers or all-empty rows).');
      return;
    }

    const rowToClear = startRow + lastFilledIndex;

    s.getRange(rowToClear, 1, 1, lastCol).clearContent().clearDataValidations();

    const directors = (typeof CONFIG !== 'undefined' && CONFIG.DIRECTORS) ? CONFIG.DIRECTORS : ['Philip Pixler', 'Kristie Wright', 'Quana Slaughter', 'Cristian Zavala'];
    s.getRange(rowToClear, 3).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    s.getRange(rowToClear, 4).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(directors, true).build());
    s.getRange(rowToClear, 19).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(directors, true).build());
    s.getRange(rowToClear, 21).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes'], true).build());

    try { logAudit('system', 'clearLastRow', null, { rowCleared: rowToClear }); } catch(e){}
    SpreadsheetApp.getUi().alert(`✅ Cleared row #${rowToClear} and restored validations.`);
  } catch (e) {
    try { logError('clearLastRow', e); } catch (err) { Logger.log('clearLastRow error: %s', String(e)); }
    SpreadsheetApp.getUi().alert('❌ Error while clearing last row: ' + e);
  }
}

function headerMapCI_(sh){
  var firstRow = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var map = {};
  for (var c=0;c<firstRow.length;c++){
    var key = String(firstRow[c] || '').trim().toLowerCase();
    if (key && !map[key]) map[key] = c+1; // 1-based column
  }
  return map;
}

function handleVisibilityOnEdit_(e){
  try{
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    var tabName = (typeof CONFIG !== 'undefined' && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
    if (sh.getName() !== tabName) return;

    var h = headerMapCI_(sh);
    var cActive = h['active'] || 0;
    var cGrace  = h['grace'] || 0;
    if (!cActive || !cGrace) return;

    var editedCol = e.range.getColumn();
    if (editedCol !== cActive && editedCol !== cGrace) return;

    // Only re-evaluate if edit happened on data rows
    if (e.range.getRow() >= 2) applyEventsAutoHide_();
  } catch(err){
    // swallow — keep edits fast
  }
}

function toBool_(v){
  // Handles true/false, "TRUE"/"FALSE", 1/0, blanks
  if (v === true) return true;
  if (v === false) return false;
  if (typeof v === 'number') return v !== 0;
  var s = String(v || '').trim().toLowerCase();
  if (s === 'true' || s === 'yes' || s === 'y') return true;
  if (s === 'false' || s === 'no'  || s === 'n') return false;
  return false;
}

function batchHideRows_(sh, rows){
  if (!rows || !rows.length) return;
  rows.sort(function(a,b){ return a-b; });
  var start = rows[0], prev = rows[0];
  for (var i=1;i<=rows.length;i++){
    var cur = rows[i];
    if (cur !== prev + 1){
      sh.hideRows(start, prev - start + 1);
      start = cur;
    }
    prev = cur;
  }
}

function batchUnhideRows_(sh, rows){
  if (!rows || !rows.length) return;
  rows.sort(function(a,b){ return a-b; });
  var start = rows[0], prev = rows[0];
  for (var i=1;i<=rows.length;i++){
    var cur = rows[i];
    if (cur !== prev + 1){
      // 👇 use showRows instead of unhideRows
      sh.showRows(start, prev - start + 1);
      start = cur;
    }
    prev = cur;
  }
}

function applyEventsAutoHide_(){
  var tabName = (typeof CONFIG !== 'undefined' && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
  var sh = SpreadsheetApp.getActive().getSheetByName(tabName);
  if (!sh) return;

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  // Map headers (case-insensitive)
  var h = headerMapCI_(sh);
  var cActive = h['active'] || 0;
  var cGrace  = h['grace'] || 0;
  if (!cActive || !cGrace) return;

  // Pull data (rows 2: last)
  var numRows = lastRow - 1;
  var range = sh.getRange(2, 1, numRows, Math.max(cActive, cGrace));
  var values = range.getValues();

  // Compute which rows should be hidden
  var rowsToHide = [];
  var rowsToShow = []; // explicitly unhide
  for (var i = 0; i < values.length; i++){
    var r = values[i];
    var activeVal = r[cActive - 1];
    var graceVal  = r[cGrace  - 1];

    var isActive = toBool_(activeVal);
    var isGrace  = toBool_(graceVal);

    // Hide if NOT active and NOT grace
    if (!isActive && !isGrace){
      rowsToHide.push(i + 2); // sheet row index
    } else {
      rowsToShow.push(i + 2);
    }
  }

  // Unhide rows that should be visible (batch by contiguous ranges)
  batchUnhideRows_(sh, rowsToShow);

  // Hide expired (not Active) but not graced
  batchHideRows_(sh, rowsToHide);
}

// Added by Codex: safe menu builder with ASCII labels
function showMenuSafe() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Admin Tools')
    .addSubMenu(
      ui.createMenu('System Tools')
        .addItem('Refresh Form Lists Now', 'refreshFormLists')
        .addItem('Restore Validations Only', 'restoreEventsValidations')
        .addItem('Clear Last Row', 'clearLastRow')
        .addItem('Back Fill Pdfs', 'addAdminBackfillMenu_')
        .addItem('Apply Auto-Hide Now', 'applyEventsAutoHide_')
        .addItem('Reset Events Board (TEST ONLY)', 'resetEventsBoard')
    )
    .addSubMenu(
      ui.createMenu('Employee Tools')
        .addItem('Employee Lookup', 'employeeLookup')
    )
    .addSubMenu(
      ui.createMenu('Reports & Monitoring')
        .addItem('Probation Watchlist', 'probationWatchlist')
    )
    .addToUi();
}

// Install an installable onOpen trigger for this spreadsheet if needed
function ensureOnOpenTrigger() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) { Logger.log('ensureOnOpenTrigger: no active spreadsheet'); return 'No active spreadsheet'; }
    const existing = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction && t.getHandlerFunction() === 'onOpen');
    if (existing && existing.length) return 'onOpen trigger already installed';
    ScriptApp.newTrigger('onOpen').forSpreadsheet(ss).onOpen().create();
    return 'Installed onOpen trigger';
  } catch (e) {
    Logger.log('ensureOnOpenTrigger error: ' + String(e));
    return 'ensureOnOpenTrigger error: ' + String(e);
  }
}
