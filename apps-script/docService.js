/**
 * createEventRecordPdf_
 * - Builds a data object from row using CONFIG.COLS mapping (safe-get)
 * - Adds header-named keys and compact aliases so templates with slightly different tokens work
 * - Falls back to rubric lookup for points if needed
 * - Calls mergeToPdf_ and writes rich link into the row
 */
function createEventRecordPdf_(row){
  return withBackoff_('createEventRecordPdf', function(){
    try {
      var events = sh_(CONFIG.TABS.EVENTS);
      // Ensure pending spreadsheet writes are flushed before reading values
      try{ SpreadsheetApp.flush(); }catch(_){}

      var map = (typeof headerIndexMap_ === 'function') ? headerIndexMap_(events) : headerMap(events);
      var cEvt   = map[CONFIG.COLS.EventType] || map['EventType'] || 0;
      var cPosA  = map[CONFIG.COLS.PositiveAction] || map['PositiveAction'] || map['Positive Action'] || 0;
      var cPosPt = map['Positive Point'] || map[CONFIG.COLS.PositivePoint] || 0;

      function isPositiveEvt_(){
        function val(c){ return c ? String(events.getRange(row, c).getDisplayValue() || '').trim().toLowerCase() : ''; }
        var evt = val(cEvt), posA = val(cPosA), posPt = val(cPosPt);
        // Treat anything that reads like "Positive", "Positive Credit", or “Positive Point Removal” as non-PDF
        return /^positive/.test(evt) || /^positive/.test(posA) || /^positive/.test(posPt);
      }

      if (isPositiveEvt_()){
        try { logInfo_ && logInfo_('createEventRecordPdf_skip_positive', { row: row }); } catch(_){}
        return null; // <<< SKIP PDF
      }
      
      // Build data using centralized builder if available
      var data;
      if (typeof buildPdfDataFromRow_ === 'function') {
        data = buildPdfDataFromRow_(events, row, { generateWriteUpTitle: true, includeHeaders: true });
      } else {
        // fallback: read row display values into data
        var hdrs = headers_(events);
        var rowVals = events.getRange(row,1,1,Math.max(1,hdrs.length)).getDisplayValues()[0] || [];
        var rowObj = asRowObject_(hdrs, rowVals);
        data = {
          Employee: rowObj[CONFIG.COLS.Employee] || '',
          IncidentDate: rowObj[CONFIG.COLS.IncidentDate] || '',
          Lead: rowObj[CONFIG.COLS.Lead] || '',
          RelevantPolicy: rowObj[CONFIG.COLS.RelevantPolicy] || rowObj.Policy || '',
          Policy: rowObj[CONFIG.COLS.RelevantPolicy] || rowObj.Policy || '',
          Infraction: rowObj[CONFIG.COLS.Infraction] || '',
          Description: rowObj[CONFIG.COLS.NotesReviewer] || rowObj.IncidentDescription || '',
          CorrectiveActions: rowObj[CONFIG.COLS.CorrectiveActions] || '',
          TeamMemberStatement: rowObj[CONFIG.COLS.TeamMemberStatement] || '',
          Points: rowObj[CONFIG.COLS.Points] || '',
          PointsRolling: rowObj[CONFIG.COLS.PointsRollingEffective] || rowObj[CONFIG.COLS.PointsRolling] || ''
        };
        // add header names & aliases
        hdrs.forEach(function(h){ if (h && !data.hasOwnProperty(h)) data[h]=rowObj[h]||''; });
        Object.keys(Object.assign({}, data)).forEach(function(k){ if(!k) return; var alias = String(k).replace(/[\s\-_]+/g,'').replace(/[^\w]/g,''); if(alias && !data.hasOwnProperty(alias)) data[alias] = data[k]; });
        // ensure EventID / WriteUpTitle exist
        data.EventID = data.EventID || data.LinkedEventID || String(row);
        data.WriteUpTitle = data.WriteUpTitle || ('Write-Up: ' + (data.Employee||'unknown') + ' (' + (data.IncidentDate||'no-date') + ')');
      }

      // Log payload keys and a compact sample to your Logs sheet (if present)
      try{
        var logMsg = 'createEventRecordPdf_payload row=' + row + ' keys=' + JSON.stringify(Object.keys(data));
        logInfo_ && logInfo_('createEventRecordPdf_payload', logMsg);
        if (CONFIG && CONFIG.TABS && CONFIG.TABS.LOGS){
          var logsSh = ss_().getSheetByName(CONFIG.TABS.LOGS);
          if (logsSh) logsSh.appendRow([new Date(), 'INFO', 'createEventRecordPdf_payload', 'row:'+row, JSON.stringify(Object.keys(data)).slice(0,500)]);
        }
      }catch(_){}

      // Build filename and ensure we're using robust merge
      var filename = 'Consequence: ' + slug_(data.Employee || 'unknown') + ' — ' + (data.Action || 'Action') + ' (' + formatDate_(new Date()) + ')';

      // Merge (robust mergeToPdf_ should now be in place)
      var pdfId = mergeToPdf_(CONFIG.TEMPLATES.EVENT_RECORD, data, filename);

      // Write the link into the row with safe fallback
      var url = 'https://drive.google.com/file/d/' + pdfId + '/view';
      try {
        setRichLinkSafe(events, row, CONFIG.COLS.PdfLink || 'Write-Up PDF', 'View PDF', url);
        logInfo_ && logInfo_('createEventRecordPdf_writeLink', 'richLink written row='+row);
      } catch(e){
        try {
          var hdrMap = headerIndexMap_(events);
          var idx = hdrMap[CONFIG.COLS.PdfLink] || hdrMap['Write-Up PDF'] || (events.getLastColumn()+1);
          events.getRange(row, idx).setValue(url);
          logInfo_ && logInfo_('createEventRecordPdf_writeLink', 'plainUrl written row='+row+' col='+idx);
        } catch(ex){
          logError && logError('createEventRecordPdf_writeLink_err', ex, {row:row, pdfId:pdfId});
        }
      }

      try{ logInfo_ && logInfo_('createEventRecordPdf_done', {row:row, pdfId:pdfId}); }catch(_){}

      return pdfId;
    } catch(err){
      logError && logError('createEventRecordPdf_top', err, {row: row});
      throw err;
    }
  }, 4, 300);
}



/**
 * createConsequencePdf_
 * - row: Events sheet row index
 * - action: descriptive action string (used to choose template & filename)
 */
function createConsequencePdf_(row, action){
  return withBackoff_('createConsequencePdf', function(){
    var events = sh_(CONFIG.TABS.EVENTS);
    var ctx = rowCtx_(events, row);
    // build base payload
    var data = buildPdfDataFromRow_(events, row, { generateWriteUpTitle: true, includeHeaders: true });
    // attach action
    data.Action = action || data.Action || '';

    data.EventID = data.EventID || data.LinkedEventID || String(row);
    data.WriteUpTitle = data.WriteUpTitle ||
      ('Write-Up: ' + (data.Employee||'unknown') + ' (' + (data.IncidentDate||'no-date') + ')');

    // choose template (same logic you already had)...
    var templateId = null;
    if (/(\b5\b|\b5pt\b|level\s*1|milestone\s*5)/i.test(action)) templateId = CONFIG.TEMPLATES.MILESTONE_5;
    else if (/(\b10\b|\b10pt\b|1-?week|one\s*week|level\s*2|milestone\s*10)/i.test(action)) templateId = CONFIG.TEMPLATES.MILESTONE_10;
    else if (/termination/i.test(action)) templateId = CONFIG.TEMPLATES.TERMINATION || CONFIG.TEMPLATES.MILESTONE_15;
    else if (/policy-?protected|policy protected/i.test(action)) templateId = CONFIG.TEMPLATES.POLICY_PROTECTED;
    else if (/probation|failure/i.test(action)) templateId = CONFIG.TEMPLATES.PROBATION_FAILURE;

    if (!templateId) throw new Error('No matching template for action: '+action);

    // filename
    var filename = data.WriteUpTitle || ('Consequence: ' + slug_(data.Employee || 'unknown') + ' (' + (data.IncidentDate || 'no-date') + ')');

    // merge and write link
    var pdfId = mergeToPdf_(templateId, data, filename);
    var url = 'https://drive.google.com/file/d/'+pdfId+'/view';
    try { setRichLinkSafe(events, row, CONFIG.COLS.ConsequencePDF || 'Consequence PDF', 'View PDF', url); } catch(e){ try{ events.getRange(row, headerIndexMap_(events)[CONFIG.COLS.ConsequencePDF]||events.getLastColumn()+1).setValue(url); }catch(_){ } }
    return pdfId;
  }, 4, 300);
}

// Returns 'TRUE' if the Positive credit earned by this Events row has been consumed.
// Looks up PositivePoints where "Earned Event Row" == eventsRow and reads "Consumed?".
function _lookupConsumedForEarnedEventRow_(eventsRow){
  try{
    var pp = ensurePositivePointsColumns_(); var s = pp.sheet, i = pp.idx;
    var last = s.getLastRow(); if (last < 2) return 'FALSE';
    var width = Math.max(i.EarnedEvent, i.Consumed);
    var vals = s.getRange(2, 1, last-1, width).getValues();
    for (var r = 0; r < vals.length; r++){
      var earned = Number(vals[r][i.EarnedEvent-1] || 0);
      if (earned === Number(eventsRow)) {
        var con = vals[r][i.Consumed-1] === true ? 'TRUE' : 'FALSE';
        return con;
      }
    }
    return 'FALSE';
  }catch(e){ try{ logError && logError('_lookupConsumedForEarnedEventRow_', e, { eventsRow: eventsRow }); }catch(_){ } return 'FALSE'; }
}

function _isPositiveCreditEventText_(evt){
  var s = String(evt || '').trim().toLowerCase();
  return (
    s === 'positive' ||
    s === 'positive credit' ||
    /^positive\s*(point|credit)\b/.test(s) ||
    s === 'positive point removal'
  );
}

/**
 * createMilestonePdf_
 * - rowIndex: Events sheet row
 * - forceTemplateId: optional override template id
 *
 * Uses buildPdfDataFromRow_ when available, otherwise falls back to a minimal payload.
 */
function createMilestonePdf_(rowIndex, forceTemplateId){
  try{
    var s = sh_(CONFIG.TABS.EVENTS);
    var ctx = rowCtx_(s, rowIndex);

    // --- Guard: only create a milestone PDF when a director has claimed it ---
    // Allow when a forceTemplateId is explicitly provided (e.g., Probation Failure),
    // otherwise skip if ConsequenceDirector is empty.
    try {
      var directorName = String(ctx.get(CONFIG.COLS.ConsequenceDirector) || '').trim();
      if (!directorName && !forceTemplateId) {
        try { sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'INFO', 'createMilestonePdf_', 'skipped_no_director', rowIndex, String(forceTemplateId||'')]); } catch(_){}
        return null;
      }
    } catch (guardErr) {
      try { sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'WARN', 'createMilestonePdf_', 'guard_read_err', rowIndex, String(guardErr).slice(0,200)]); } catch(_){}
      return null;
    }

    // ---- Template inference (defensive)
    var inferred = null;
    try {
      if (typeof inferMilestoneTemplate_ === 'function') {
        inferred = inferMilestoneTemplate_(rowIndex) || null;
      }
    } catch (e) {
      try { sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'WARN', 'createMilestonePdf_', 'infer_failed', rowIndex, String(e).slice(0,200)]); } catch(_){}
      inferred = null;
    }

    // ---- Read the row label for PF safety gating
    var labelForRow = '';
    try {
      labelForRow = String(ctx.get(CONFIG.COLS.Milestone) || '').trim();
    } catch(_) {}

    // ---- Only honor a forced PF template when the label actually is Probation Failure
    var forced = forceTemplateId || null;
    var pfId = (CONFIG && CONFIG.TEMPLATES && CONFIG.TEMPLATES.PROBATION_FAILURE) || null;
    if (forced && pfId && forced === pfId && !/probation failure/i.test(labelForRow || '')) {
      // Ignore accidental PF override on non-PF rows (e.g., termination)
      try { logInfo_ && logInfo_('createMilestonePdf_force_guard_ignored', { row: rowIndex, forced: forced, label: labelForRow }); } catch(_){}
      forced = null;
    }

    // ---- Choose template: guarded force -> inferred
    var chosen = forced || (inferred && inferred.templateId) || null;
    if (!chosen) {
      logInfo_ && logInfo_('createMilestonePdf_noTemplate', { row: rowIndex, inferred: inferred, label: labelForRow });
      try { sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'INFO', 'createMilestonePdf_', 'no_template_chosen', rowIndex, JSON.stringify({label:labelForRow, inferred: inferred||{}})]); } catch(_){}
      return null;
    }

    // ---- Payload
    var payload = {
      IncidentDate: ctx.get(CONFIG.COLS.IncidentDate) || '',
      Employee: ctx.get(CONFIG.COLS.Employee) || '',
      ConsequenceDirector: String(ctx.get(CONFIG.COLS.ConsequenceDirector) || '').trim(),
      MilestoneName: labelForRow || (inferred && inferred.reason) || '',
      PointsRolling: ctx.get(CONFIG.COLS.PointsRollingEffective) || '',
      EventID: ctx.get(CONFIG.COLS.LinkedEventID) || String(rowIndex),
      WriteUpTitle: 'Write-Up: ' + (ctx.get(CONFIG.COLS.Employee) || 'unknown') +
                    ' (' + (ctx.get(CONFIG.COLS.IncidentDate) || formatDate_(new Date())) + ')'
    };


    // ---- Filename
    var outName = (inferred && inferred.fileName)
      || ('Milestone: ' + slug_(payload.Employee) + ' — ' + (payload.MilestoneName || 'Milestone') + ' (' + formatDate_(new Date()) + ')');

    // Log which template we're actually using and why
    logInfo_ && logInfo_('createMilestonePdf_preMerge', {
      row: rowIndex,
      templateId: chosen,
      forcedUsed: !!forced,
      label: labelForRow,
      outName: outName,
      payload: payload,
      inferred: inferred
    });

    // ---- Merge & link (writes to Write-Up PDF column; keep if that's your target)
    var pdfId = mergeToPdf_(chosen, payload, outName);
    if (!pdfId) {
      logError && logError('createMilestonePdf_mergeFailed', { row: rowIndex, templateId: chosen });
      return null;
    }

    var url = 'https://drive.google.com/file/d/' + pdfId + '/view';
    try {
      setRichLinkSafe(s, rowIndex, CONFIG.COLS.PdfLink || 'Write-Up PDF', 'View PDF', url);
    } catch(e){
      try {
        var map = headerIndexMap_(s);
        var col = map[CONFIG.COLS.PdfLink] || map['Write-Up PDF'] || (s.getLastColumn()+1);
        s.getRange(rowIndex, col).setValue(url);
      } catch(_){}
    }

    logInfo_ && logInfo_('createMilestonePdf_done', { row: rowIndex, pdfId: pdfId, templateId: chosen });
    return pdfId;

  } catch(err){
    logError && logError('createMilestonePdf_', err, { rowIndex: rowIndex, forceTemplateId: forceTemplateId });
    return null;
  }
}

function resolveStrictDestFolder_() {
  var id = String((CONFIG && CONFIG.DEST_FOLDER_ID) || '').trim();
  if (!id) throw new Error('CONFIG.DEST_FOLDER_ID is empty');

  // Try as Folder first
  try {
    var f = DriveApp.getFolderById(id);   // throws if not a real folder
    if (f.isTrashed && f.isTrashed()) {
      throw new Error('Destination folder is in Trash: ' + id);
    }
    // Optional breadcrumb
    try { logInfo_ && logInfo_('dest_folder_ok', { id: id }); } catch(_){}
    return f;
  } catch (eFolder) {
    // Diagnose: is this ID something else (file/shortcut)?
    var diag = { id: id };
    try {
      var fi = DriveApp.getFileById(id);  // works for files & shortcuts
      diag.mime = fi.getMimeType && fi.getMimeType();
      diag.name = fi.getName && fi.getName();
      diag.trashed = fi.isTrashed && fi.isTrashed();
    } catch (eFileProbe) {
      diag.fileProbeError = String(eFileProbe);
    }
    try { logError && logError('dest_folder_invalid', eFolder, diag); } catch(_){}

    var msg = 'DEST_FOLDER_ID is not a usable FOLDER id: ' + id;
    if (diag.mime === 'application/vnd.google-apps.shortcut') {
      msg += ' (it is a Shortcut; open the TARGET folder and use its /folders/<ID>)';
    } else if (diag.mime) {
      msg += ' (mime=' + diag.mime + ', name="' + (diag.name || '') + '")';
    }
    throw new Error(msg);
  }
}


function mergeToPdf_(docId, data, outName){
  // Version stamp so we know which build is running
  try { logInfo_ && logInfo_('mergeToPdf_version', { v: '2025-08-28T10:45:00-05:00' }); } catch(_){}
  try {
    logInfo_ && logInfo_('mergeToPdf_principal', {
      effective: String(Session.getEffectiveUser()),
      active: String(Session.getActiveUser())
    });
  } catch(_){}

  if (!docId) throw new Error('Template Doc ID not set');

  return withBackoff_('mergeToPdf', function(){
    // ── Stage A: Resolve destination folder (STRICT, no fallback) ───────────
    var folder = resolveStrictDestFolder_();
    try { logInfo_ && logInfo_('mergeToPdf_stage', { stage: 'folder_ok', dest: folder.getId() }); } catch(_){}

    // ── Stage B: Template guard (must be a Google Doc) ───────────────────────
    try {
      var tpl = DriveApp.getFileById(docId);
      var mt  = (tpl && tpl.getMimeType && tpl.getMimeType()) || '';
      logInfo_ && logInfo_('mergeToPdf_stage', { stage: 'tpl_mime', mime: mt, docId: docId });
      if (String(mt).indexOf('vnd.google-apps.document') === -1) {
        throw new Error('Template is not a Google Doc: ' + (mt || 'unknown mime'));
      }
    } catch(eMime){
      logError && logError('mergeToPdf_tpl_mime_err', eMime, { docId: docId });
      throw eMime;
    }

    // ── Stage C: Make a working copy ─────────────────────────────────────────
    var safeName = String(outName || 'Document').slice(0, 120);
    var copyId = null;
    try {
      var copy = DriveApp.getFileById(docId).makeCopy(safeName + ' (merge)', folder);
      copyId = copy.getId();
      logInfo_ && logInfo_('mergeToPdf_stage', { stage: 'copy_ok', copyId: copyId });
    } catch(eCopy){
      logError && logError('mergeToPdf_copy_err', eCopy, { docId: docId, dest: folder.getId() });
      throw new Error('Failed to make working copy of template (permissions/quota?)');
    }

    // Small settle so indexing catches up
    try { Utilities.sleep(400); } catch(_){}

    // ── Stage D: Open document + collect tokens ──────────────────────────────
    var doc, body, header, footer;
    try {
      doc = DocumentApp.openById(copyId);
      body = doc.getBody();
      try { header = (doc.getHeader && doc.getHeader()) || null; } catch(_){ header = null; }
      try { footer = (doc.getFooter && doc.getFooter()) || null; } catch(_){ footer = null; }
      logInfo_ && logInfo_('mergeToPdf_stage', { stage: 'doc_opened', copyId: copyId, hasHeader: !!header, hasFooter: !!footer });
    } catch(eOpen){
      logError && logError('mergeToPdf_open_err', eOpen, { copyId: copyId });
      throw new Error('Failed to open working copy as a Google Doc');
    }

    function escRegex(s){ return String(s).replace(/[-\/\\^$*+?.()|[\]{}]/g,'\\$&'); }
    function norm(s){ return String(s||'').toLowerCase().replace(/[\s_\-]+/g,''); }

    var keyMap = {};
    try {
      Object.keys(data||{}).forEach(function(k){
        var nk = norm(k); if (nk && !keyMap[nk]) keyMap[nk] = k;
      });
    } catch(eKeys){
      logError && logError('mergeToPdf_keys_err', eKeys, { dataKeys: Object.keys(data||{}) });
    }

    function collectTokens(container){
      if (!container) return [];
      var text = (container.getText && container.getText()) || '';
      var seen = [], m, rx = /\{\{\s*([^}]+?)\s*\}\}/g;
      while ((m = rx.exec(text)) !== null){
        var t = String(m[1]).trim();
        if (seen.indexOf(t) === -1) seen.push(t);
      }
      return seen;
    }

    var foundTokens = []
      .concat(collectTokens(body))
      .concat(collectTokens(header))
      .concat(collectTokens(footer))
      .filter(function(v, i, a){ return a.indexOf(v) === i; });

    var replacements = {};
    try {
      foundTokens.forEach(function(tok){
        var k = keyMap[norm(tok)];
        if (!k){
          Object.keys(data||{}).some(function(kk){
            if (String(kk).toLowerCase() === String(tok).toLowerCase()){ k = kk; return true; }
            return false;
          });
        }
       function fmtVal(v){
        // normalize Date objects to a short, stable format
        if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
          return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd'); // or 'MM/dd/yyyy'
        }
        return (v == null) ? '' : String(v);
      }
      replacements[tok] = fmtVal(k != null ? data[k] : null);
      });
      logInfo_ && logInfo_('mergeToPdf_stage', { stage: 'tokens_mapped', tokenCount: foundTokens.length });
    } catch(eMap){
      logError && logError('mergeToPdf_map_err', eMap, { tokens: foundTokens });
    }

    function replaceInContainer(container){
      if (!container || typeof container.getNumChildren !== 'function') return 0;
      var total = 0, n = container.getNumChildren();
      for (var i = 0; i < n; i++){
        var child = container.getChild(i);
        var type = child.getType && child.getType();
        if (type === DocumentApp.ElementType.PARAGRAPH || type === DocumentApp.ElementType.LIST_ITEM){
          total += replaceInParagraph(child.asParagraph ? child.asParagraph() : child);
        } else if (type === DocumentApp.ElementType.TABLE){
          var tbl = child.asTable();
          for (var r = 0; r < tbl.getNumRows(); r++){
            var row = tbl.getRow(r);
            for (var c = 0; c < row.getNumCells(); c++){
              total += replaceInContainer(row.getCell(c));
            }
          }
        } else if (type === DocumentApp.ElementType.TABLE_CELL){
          total += replaceInContainer(child.asTableCell());
        } else {
          if (typeof child.getNumChildren === 'function' && child.getNumChildren() > 0){
            total += replaceInContainer(child);
          }
        }
      }
      return total;
    }

    function replaceInParagraph(par){
      try{
        var runs = [], full = '';
        for (var j = 0; j < par.getNumChildren(); j++){
          var el = par.getChild(j);
          if (el.getType() === DocumentApp.ElementType.TEXT){
            var t = el.asText().getText();
            runs.push({ el: el.asText(), text: t, start: full.length, end: full.length + t.length - 1 });
            full += t;
          }
        }
        if (!full) return 0;

        var matches = [];
        Object.keys(replacements).forEach(function(tok){
          var patt = new RegExp('\\{\\{\\s*' + escRegex(tok) + '\\s*\\}\\}', 'g');
          var m;
          while ((m = patt.exec(full)) !== null){
            matches.push({ token: tok, start: m.index, end: m.index + m[0].length - 1, replacement: replacements[tok] });
          }
        });
        if (!matches.length) return 0;

        matches.sort(function(a,b){ return b.start - a.start; });

        var replacedHere = 0;
        matches.forEach(function(match){
          var start = match.start, end = match.end, repl = match.replacement;

          for (var r = runs.length - 1; r >= 0; r--){
            var run = runs[r];
            if (run.end < start || run.start > end) continue;
            var lo = Math.max(run.start, start);
            var hi = Math.min(run.end, end);
            var offStart = lo - run.start;
            var offEnd = hi - run.start;
            try { run.el.deleteText(offStart, offEnd); } catch(_){}
            var before = run.text.slice(0, offStart);
            var after  = run.text.slice(offEnd + 1);
            var delta  = (offEnd - offStart + 1);
            run.text = before + after;
            run.end -= delta;
            for (var k = r + 1; k < runs.length; k++){
              runs[k].start -= delta;
              runs[k].end   -= delta;
            }
          }

          if (repl){
            for (var r2 = 0; r2 < runs.length; r2++){
              var run2 = runs[r2];
              if (start >= run2.start && start <= run2.end + 1){
                var insOff = start - run2.start;
                try { run2.el.insertText(insOff, repl); } catch(_){}
                run2.text = run2.text.slice(0, insOff) + repl + run2.text.slice(insOff);
                var d = repl.length;
                run2.end += d;
                for (var k2 = r2 + 1; k2 < runs.length; k2++){
                  runs[k2].start += d;
                  runs[k2].end   += d;
                }
                break;
              }
            }
          }
          replacedHere++;
        });
        return replacedHere;
      }catch(e){
        try{ logError && logError('replaceInParagraph_err', e, { paraText: (par.getText && par.getText()) || '' }); }catch(_){}
        return 0;
      }
    }

    var totalRepl = 0;
    try {
      totalRepl += replaceInContainer(body);
      totalRepl += replaceInContainer(header);
      totalRepl += replaceInContainer(footer);
      logInfo_ && logInfo_('mergeToPdf_stage', { stage: 'replaced', totalRepl: totalRepl });
    } catch(eRepl){
      logError && logError('mergeToPdf_repl_err', eRepl, { copyId: copyId });
    }

    try { doc.saveAndClose(); } catch(_){}
    try { DocumentApp.openById(copyId).saveAndClose(); } catch(_){}
    try { Utilities.sleep(400); } catch(_){}

    // ── Stage E: Export (Docs HTTP first) ────────────────────────────────────
    function exportViaDocsHttp_(){
      var url = 'https://docs.google.com/document/d/' + copyId + '/export?format=pdf';
      var resp = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: false
      });
      var blob = resp.getBlob();
      blob.setName(safeName + '.pdf');
      return blob;
    }
    function exportViaGetAs_(){
      var f = DriveApp.getFileById(copyId);
      try { f.getBlob(); } catch(_){}
      return f.getAs('application/pdf');
    }

    try { logInfo_ && logInfo_('mergeToPdf_stage', { stage: 'before_export', copyId: copyId }); } catch(_){}

    var out = null, lastErr = null;
    for (var attempt = 1; attempt <= 5; attempt++){
      try {
        var pdfBlob = (attempt <= 3) ? exportViaDocsHttp_() : exportViaGetAs_();
        out = folder.createFile(pdfBlob).setName(safeName + '.pdf');
        logInfo_ && logInfo_('mergeToPdf_stage', { stage: 'export_ok', attempt: attempt, outId: out && out.getId ? out.getId() : null });
        if (out && out.getId) break;
      } catch (e) {
        lastErr = e;
        try { logError && logError('mergeToPdf_export_attempt', e, { attempt: attempt, copyId: copyId }); } catch(_){}
        Utilities.sleep(Math.min(1500 * Math.pow(2, attempt-1), 8000));
      }
    }
    if (!out) throw lastErr || new Error('PDF export failed after retries');

    try{ DriveApp.getFileById(copyId).setTrashed(true); }catch(_){}

    try{
      logInfo_ && logInfo_('mergeToPdf_done', 'out=' + out.getId()
        + ' tokens=' + JSON.stringify(foundTokens)
        + ' totalReplacements=' + totalRepl);
    }catch(_){}

    return out.getId();
  }, 4, 300);
}











/**
 * inferMilestoneTemplate_(rowIndex)
 * - Inspects the event row and returns an object { templateId, fileName, reason }
 * - Uses: milestone label, effective rolling points, CONFIG.POLICY thresholds, and CONFIG.TEMPLATES.
 */
function inferMilestoneTemplate_(rowIndex){
  try{
    var s = sh_(CONFIG.TABS.EVENTS);
    var ctx = rowCtx_(s, rowIndex);
    var map = headerIndexMap_(s);

    var milestoneLabel = String(ctx.get(CONFIG.COLS.Milestone) || '').trim();
    var emp  = String(ctx.get(CONFIG.COLS.Employee) || '').trim();
    var date = ctx.get(CONFIG.COLS.MilestoneDate) || ctx.get(CONFIG.COLS.IncidentDate) || new Date();
    var eff  = Number(ctx.get(CONFIG.COLS.PointsRollingEffective) || ctx.get(CONFIG.COLS.PointsRolling) || 0);

    // thresholds
    if (typeof loadPolicyFromSheet_ === 'function') { try { loadPolicyFromSheet_(); } catch(_) {} }
    var L1 = Number((CONFIG.POLICY && CONFIG.POLICY.MILESTONES && CONFIG.POLICY.MILESTONES.LEVEL_1) || 5);
    var L2 = Number((CONFIG.POLICY && CONFIG.POLICY.MILESTONES && CONFIG.POLICY.MILESTONES.LEVEL_2) || 10);
    var L3 = Number((CONFIG.POLICY && CONFIG.POLICY.MILESTONES && CONFIG.POLICY.MILESTONES.LEVEL_3) || 15);

    var ct = CONFIG.TEMPLATES || {};
    var chosen = null, reason = '', tier = null;

    // 1) explicit termination
    if (/termination|term/i.test(milestoneLabel)){
      chosen = ct.TERMINATION || ct.MILESTONE_15 || ct.MILESTONE_10 || null;
      reason = 'label indicated termination';
    }

    // 2) label hints
    if (!chosen){
      if (/\b5\b|5pt|2-day|2 day|two day/i.test(milestoneLabel)) {
        chosen = ct.MILESTONE_5 || null;  reason = 'label hinted 5pt/2-day';
      } else if (/\b10\b|10pt|1-week|1 week|one week|final warning/i.test(milestoneLabel)) {
        chosen = ct.MILESTONE_10 || null; reason = 'label hinted 10pt/1-week';
      } else if (/\b15\b|15pt|termination/i.test(milestoneLabel)) {
        chosen = ct.MILESTONE_15 || null; reason = 'label hinted 15pt/termination';
      }
    }

    // 3) points-based
    if (!chosen){
      tier = tierForPoints_(eff);
      if (tier === 1) { chosen = ct.MILESTONE_5  || null; reason = 'tierBased L1'; }
      else if (tier === 2) { chosen = ct.MILESTONE_10 || null; reason = 'tierBased L2'; }
      else if (tier === 3) { chosen = ct.MILESTONE_15 || null; reason = 'tierBased L3'; }
    }

    // 4) fallback
    if (!chosen){
      chosen = ct.MILESTONE || ct.EVENT_RECORD || null;
      if (chosen) reason = 'fallback to generic milestone/event template';
    }

    var prettyDate = (date && toDate_(date)) ? formatDate_(toDate_(date)) : (date ? String(date) : formatDate_(new Date()));
    var fileName   = 'Milestone: ' + slug_(emp || 'unknown') + ' — ' + (milestoneLabel || ('Level ' + (tier != null ? tier : '?'))) + ' (' + prettyDate + ')';

    try{
      logInfo_ && logInfo_('inferMilestoneTemplate_', {
        row: rowIndex, employee: emp, milestoneLabel: milestoneLabel, effective: eff,
        L1: L1, L2: L2, L3: L3, chosen: chosen ? String(chosen).slice(0,40) : null, reason: reason
      });
    }catch(_){}

    if (!chosen) return null;
    return { templateId: chosen, fileName: fileName, reason: reason, effective: eff };
  }catch(err){
    try{ logError && logError('inferMilestoneTemplate_', err, {row: rowIndex}); }catch(_){}
    return null;
  }
}


// Creates the Employee History PDF using CONFIG.TEMPLATES.EMP_HISTORY.
// Expects the template to contain a table with headers: Date | Event | Lead | Infraction | Pts | Roll | Notes | PDF
/**
 * createEmployeeHistoryPDF(name[, opts])
 * - Template-preserving Employee History PDF generator.
 * - Uses your helpers: replaceBodyToken_, findHistoryTable_, fillHistoryTable_, wirePdfLinksIntoTable (if present).
 * - Honors CONFIG.TEMPLATES.EMP_HISTORY and CONFIG.DEST_FOLDER_ID.
 *
 * Returns: { ok: true, pdfId, url } on success; { ok: false, error } on failure.
 */
function createEmployeeHistoryPDF_(name, opts) {
  opts = opts || {};
  try {
    if (!name) throw new Error('No employee name provided');

    // ----- Resolve Events sheet + headers -----
    var eventsTab = (CONFIG && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
    var s = (typeof sh_ === 'function') ? sh_(eventsTab) : SpreadsheetApp.getActiveSpreadsheet().getSheetByName(eventsTab);
    if (!s) throw new Error('Events sheet not found: ' + eventsTab);

    var data = s.getDataRange().getValues();
    if (!data || data.length < 2) throw new Error('No event rows found');

    var hdr = data[0].map(function (h) { return String(h || '').trim(); });
    function col() {
      var names = Array.prototype.slice.call(arguments).flat().filter(Boolean);
      for (var i = 0; i < names.length; i++) {
        var idx = hdr.indexOf(names[i]);
        if (idx !== -1) return idx;
      }
      return -1;
    }

    var cEmployee    = col('Employee'); if (cEmployee === -1) throw new Error('Employee column not found');
    var cDate        = col('IncidentDate','Date','Timestamp');
    var cEventType   = col('EventType');
    var cLead        = col('Lead');
    var cInfraction  = col('Infraction');
    var cPoints      = col('Points');
    var cRolling     = col('PointsRolling (Effective)','PointsRolling','FinalPoints');
    var cCorrective  = col('CorrectiveActions');
    var cTeamStmt    = col('TeamMemberStatement');
    var cGraceNotes  = col('GraceNotes');
    var cMilestone   = col('Milestone');
    var cProbEnd     = col('Probation_End','ProbationEnd');
    var cPdf         = col('Write-Up PDF','Signed_PDF_Link','WriteUpPDF');
    var cPosAction   = col('PositiveAction', 'Positive Action', 'Positive Point');
    var cGraceApplied = col('Grace Applied','GraceApplied');
    var cLinkedRow    = col('Linked Event Row','Linked_Event_ID','Linked Event Row #','Linked Row');

    var target = String(name).trim().toLowerCase();
    var rows = [];
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var emp = row[cEmployee];
      if (!emp && emp !== 0) continue;
      if (String(emp).trim().toLowerCase() !== target) continue;
      rows.push({
        sheetRowIndex: r + 1,
        dateRaw:        (cDate       !== -1) ? row[cDate]       : '',
        eventType:      (cEventType  !== -1) ? row[cEventType]  : '',
        lead:           (cLead       !== -1) ? row[cLead]       : '',
        infraction:     (cInfraction !== -1) ? row[cInfraction] : '',
        points:         (cPoints     !== -1) ? row[cPoints]     : '',
        rollingRaw:     (cRolling    !== -1) ? row[cRolling]    : '',
        corrective:     (cCorrective !== -1) ? row[cCorrective] : '',
        teamStatement:  (cTeamStmt   !== -1) ? row[cTeamStmt]   : '',
        graceNotes:     (cGraceNotes !== -1) ? row[cGraceNotes] : '',
        milestoneText:  (cMilestone  !== -1) ? row[cMilestone]  : '',
        probEndRaw:     (cProbEnd    !== -1) ? row[cProbEnd]    : '',
        posAction:      (cPosAction  !== -1) ? row[cPosAction]  : '',
        graceApplied:   (cGraceApplied !== -1) ? row[cGraceApplied] : '',
        linkedRow:      (cLinkedRow    !== -1) ? Number(row[cLinkedRow]) : NaN
      });
    }
    if (!rows.length) throw new Error('No events found for ' + name);

    // ----- Extract Write-Up links (robust) -----
    var writeUpMap = {};
    if (cPdf !== -1) {
      var haveReadLink = (typeof readLinkUrlFromCell_ === 'function');
      if (haveReadLink) {
        for (var i = 0; i < rows.length; i++) {
          var sr = rows[i].sheetRowIndex;
          var url = null;
          try { url = readLinkUrlFromCell_(s.getRange(sr, cPdf + 1)); } catch (_) {}
          writeUpMap[sr] = { url: url || null };
        }
      } else {
        var lastRow = s.getLastRow(), n = Math.max(0, lastRow - 1);
        var rich  = n ? s.getRange(2, cPdf + 1, n, 1).getRichTextValues() : [];
        var forms = n ? s.getRange(2, cPdf + 1, n, 1).getFormulas()       : [];
        var disp  = n ? s.getRange(2, cPdf + 1, n, 1).getDisplayValues()  : [];
        function fromRich(rv){ if (!rv) return null; try {
          var link = rv.getLinkUrl && rv.getLinkUrl(); if (link) return link;
          var runs = rv.getRuns ? rv.getRuns() : []; for (var k=0;k<runs.length;k++){ var u=runs[k].getLinkUrl&&runs[k].getLinkUrl(); if(u) return u; }
        }catch(_){ } return null; }
        function fromFormula(f){ if(!f) return null; var m=String(f).match(/HYPERLINK\s*\(\s*["']([^"']+)["']/i); return (m&&m[1])?m[1]:null; }
        function fromPlain(sv){ var m=String(sv||'').match(/https?:\/\/[^\s)]+/i); return m?m[0]:null; }
        for (var j = 0; j < rows.length; j++) {
          var sIdx = rows[j].sheetRowIndex, idx0 = sIdx - 2, url = null;
          if (idx0 >= 0 && n) {
            url = fromRich(rich[idx0] && rich[idx0][0]) ||
                  fromFormula(forms[idx0] && forms[idx0][0]) ||
                  fromPlain(disp[idx0] && disp[idx0][0]);
          }
          writeUpMap[sIdx] = { url: url || null };
        }
      }
    }

    // ----- Helpers -----
    var tz = Session.getScriptTimeZone() || 'UTC';
    function ymd(v){ if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v,tz,'yyyy-MM-dd');
      try{ var d=new Date(v); if(!isNaN(d)) return Utilities.formatDate(d,tz,'yyyy-MM-dd'); }catch(_){}
      var s=String(v||''); return s.length>=10 ? s.slice(0,10) : s;
    }
    function shortMilestone(text){
      if(!text) return '';
      var t=String(text);
      var m=t.match(/(\d+\s*[-]?\s*(?:Day|Days|Week|Weeks|Month|Months))\s+Suspension/i);
      if(m&&m[1]) return m[1].replace(/\s*-\s*/,'-').replace(/\s+/g,'-')+' Sus';
      if(/Suspension/i.test(t)) return 'Suspension';
      var s20=t.trim(); return s20.length<=20 ? s20 : (s20.slice(0,20).trim()+'…');
    }
    function clamp_(s,max){ s=String(s||''); return s.length>max ? (s.slice(0,max-1)+'…') : s; }

    // Positive ledger lookup for enrichment
    function lookupPositiveLedgerByEventRow_(sheetRow){
      try{
        var tab = (CONFIG && CONFIG.TABS && CONFIG.TABS.POSITIVE_POINTS) ? CONFIG.TABS.POSITIVE_POINTS : 'Positive Points';
        var ps  = sh_(tab);
        var hdr2 = headers_(ps);
        function h(names){ for (var i=0;i<names.length;i++){ var j=hdr2.indexOf(names[i]); if (j!==-1) return j; } return -1; }
        var iEarned = h(['Earned Event Row','Source Event Row','From Event Row','EarnedFromRow','Earned Row']);
        var iReason = h(['Credit Reason','Earned Reason','Reason','Notes']);
        var iType   = h(['Credit Type','Tier','Type']);
        if (iEarned === -1) return null;
        var vals = ps.getDataRange().getValues();
        for (var r=1; r<vals.length; r++){
          if (Number(vals[r][iEarned]) === Number(sheetRow)){
            return {
              tier:   (iType   !== -1) ? String(vals[r][iType]   || '').trim() : '',
              reason: (iReason !== -1) ? String(vals[r][iReason] || '').trim() : ''
            };
          }
        }
        return null;
      }catch(_){ return null; }
    }

    // ----- Rolling total, milestone, probation -----
    var rollingTotal = '';
    for (var rr = rows.length - 1; rr >= 0; rr--) {
      var v = rows[rr].rollingRaw;
      if (v !== '' && v !== null && typeof v !== 'undefined') { rollingTotal = v; break; }
    }
    var L2 = Number((CONFIG && CONFIG.POLICY && CONFIG.POLICY.MILESTONES && CONFIG.POLICY.MILESTONES.LEVEL_2) || 10);

    // Compute “current” milestone respecting grace rescinds
   function computeCurrentMilestone_(){
    // ensure we have a threshold for L2
    var L2 = Number((CONFIG && CONFIG.POLICY && CONFIG.POLICY.MILESTONES && CONFIG.POLICY.MILESTONES.LEVEL_2) || 10);

    // find last milestone-like label in this employee’s rows
    var lastMilestoneIdx = -1, label = '';
    for (var i = rows.length - 1; i >= 0; i--){
      var t = String(rows[i].milestoneText || '').trim();
      if (t){
        lastMilestoneIdx = i;
        label = t;
        break;
      }
    }
    if (lastMilestoneIdx === -1) return '';

    // If it isn’t Probation Failure, just return it.
    if (!/probation\s*failure/i.test(label)) return label;

    // ---------- Identify triggering event row (last non-milestone with positive points before PF)
    var triggerRowId = NaN;
    for (var k = lastMilestoneIdx - 1; k >= 0; k--){
      var r = rows[k];
      var isMilestoneEvt = String(r.eventType || '').toLowerCase().indexOf('milestone') !== -1;
      var pts = Number(r.points || 0);
      if (!isMilestoneEvt && pts > 0) { triggerRowId = r.sheetRowIndex; break; }
    }

    // ---------- Was that triggering event later graced?
    // Prefer explicit columns (Grace Applied + Linked Row); fallback to “grace” text sniff.
    var pfRescinded = false;
    function isTruthy(v){
      var s = String(v || '').trim().toLowerCase();
      return v === true || s === 'true' || s === 'yes' || s === 'y' || s === '1';
    }

    if (!isNaN(triggerRowId)) {
      for (var j = lastMilestoneIdx + 1; j < rows.length; j++){
        var g = rows[j];
        var applied = isTruthy(g.graceApplied);
        var linksTrigger = (g.linkedRow != null && Number(g.linkedRow) === Number(triggerRowId));
        if (applied && linksTrigger) { pfRescinded = true; break; }
      }
    } else {
      // Heuristic fallback: any later applied grace OR obvious “grace” text → consider PF rescinded
      for (var j2 = lastMilestoneIdx + 1; j2 < rows.length; j2++){
        var g2 = rows[j2];
        var applied2 = isTruthy(g2.graceApplied);
        var hay = [g2.eventType, g2.infraction, g2.milestoneText, g2.graceNotes, g2.posAction]
                  .map(function(v){ return String(v||''); }).join(' ');
        if (applied2 || /grace/i.test(hay)) { pfRescinded = true; break; }
      }
    }

    // ---------- Safety: if rolling total is now below L2, PF shouldn’t be shown
    var belowL2Now = (rollingTotal !== '' && !isNaN(Number(rollingTotal)) && Number(rollingTotal) < L2);

    if (pfRescinded || belowL2Now) {
      // If probation is still active, show a neutral "Probation" label; else clear.
      var onProb = (typeof isProbationActiveForEmployee_ === 'function') ? !!isProbationActiveForEmployee_(name) : false;
      return onProb ? 'Probation' : '';
    }
    return label;
  }

  var mostRecentMilestone = computeCurrentMilestone_();


    var mostRecentProbEnd = '';
    for (var pe = rows.length - 1; pe >= 0; pe--) {
      if (rows[pe].probEndRaw) { mostRecentProbEnd = rows[pe].probEndRaw; break; }
    }

    // ----- Build table data -----
    var expectedHeaders = ['Date','Event','Lead','Infraction','Pts','Roll','Notes','PDF'];
    var tableData = [expectedHeaders];
    for (var x = 0; x < rows.length; x++) {
      var rObj = rows[x];

      var notes = []
        .concat(rObj.corrective ? [String(rObj.corrective)] : [])
        .concat(rObj.teamStatement ? [String(rObj.teamStatement)] : [])
        .concat(rObj.graceNotes ? [String(rObj.graceNotes)] : [])
        .join(' | ');

      var inf = rObj.infraction || '';
      var isMilestoneEvt = String(rObj.eventType || '').toLowerCase().indexOf('milestone') !== -1;

      // Positive credit enrichment
      var ledger = lookupPositiveLedgerByEventRow_(rObj.sheetRowIndex);
      var isPositiveEvt =
        /^positive/i.test(String(rObj.eventType || '')) ||
        !!rObj.posAction || !!ledger;

      if (isMilestoneEvt) {
        inf = rObj.milestoneText ? shortMilestone(rObj.milestoneText) : shortMilestone(inf);
      } else if (isPositiveEvt) {
        var what = ledger && ledger.reason ? ledger.reason : (rObj.posAction || '');
        var tier = ledger && ledger.tier ? ' (' + ledger.tier + ')' : '';
        inf = 'Positive Credit' + (what ? ': ' + what : '') + tier;
        if (what) notes = clamp_((notes ? notes + ' | ' : '') + 'Reason: ' + what, 180);
      }

      tableData.push([
        ymd(rObj.dateRaw),
        String(rObj.eventType || ''),
        String(rObj.lead || ''),
        String(inf || ''),
        (rObj.points === '' || rObj.points == null) ? '' : String(rObj.points),
        (rObj.rollingRaw === '' || rObj.rollingRaw == null) ? '' : String(rObj.rollingRaw),
        clamp_(String(notes || ''), 180),
        '' // PDF column gets link wiring after fill
      ]);
    }

    // ----- Prepare Doc -----
    var now = new Date();
    var stamp = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm');
    var docTitle = name + ' — Employee History — ' + stamp;

    var tplId = (CONFIG && CONFIG.TEMPLATES && CONFIG.TEMPLATES.EMP_HISTORY) ? CONFIG.TEMPLATES.EMP_HISTORY : null;
    var destFolder = null;
    if (CONFIG && CONFIG.DEST_FOLDER_ID) {
      try { destFolder = DriveApp.getFolderById(CONFIG.DEST_FOLDER_ID); } catch (_) { destFolder = null; }
    }
    if (!destFolder) destFolder = DriveApp.getRootFolder();

    var docFile, doc;
    try {
      if (tplId) {
        docFile = DriveApp.getFileById(tplId).makeCopy(docTitle, destFolder);
        doc = DocumentApp.openById(docFile.getId());
      } else {
        doc = DocumentApp.create(docTitle);
        docFile = DriveApp.getFileById(doc.getId());
      }
    } catch (eDoc) {
      doc = DocumentApp.create(docTitle);
      docFile = DriveApp.getFileById(doc.getId());
    }

    var body = doc.getBody();

    // ----- Tokens -----
    if (typeof replaceBodyToken_ === 'function') {
      replaceBodyToken_(body, 'EMP_NAME', name);
      replaceBodyToken_(body, 'ROLLING_TOTAL', (rollingTotal !== '' ? String(rollingTotal) : 'N/A'));
      replaceBodyToken_(body, 'MILESTONE', (mostRecentMilestone || 'None'));
      replaceBodyToken_(body, 'PROBATION_ENDS', (mostRecentProbEnd ? ymd(mostRecentProbEnd) : 'N/A'));
      replaceBodyToken_(body, 'GENERATED_TS', Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss z'));
    }

    // Keep table off trailing edge of page 1, unless explicitly disabled
    if (opts.breakBeforeTable !== false) { try { body.appendPageBreak(); } catch(_) {} }

    // ----- Fill table -----
    var tableFilled = false;
    try {
      if (typeof findHistoryTable_ === 'function' && typeof fillHistoryTable_ === 'function') {
        var tbl = findHistoryTable_(body, expectedHeaders);
        if (tbl) { fillHistoryTable_(tbl, tableData, writeUpMap, rows); tableFilled = true; }
      }
    } catch (_) {}

    if (!tableFilled) {
      var appended = body.appendTable(tableData);
      if (typeof wirePdfLinksIntoTable === 'function') {
        try { wirePdfLinksIntoTable(appended, rows, writeUpMap); } catch (_) {}
      } else {
        var headerCols = appended.getRow(0).getNumCells();
        var pdfCol = Math.max(0, headerCols - 1);
        for (var irow = 1; irow < appended.getNumRows(); irow++) {
          var sIdx = rows[irow - 1].sheetRowIndex;
          var url = (writeUpMap[sIdx] && writeUpMap[sIdx].url) ? writeUpMap[sIdx].url : '';
          var cell = appended.getRow(irow).getCell(pdfCol);
          cell.clear();
          var p = cell.appendParagraph(url ? 'View PDF' : '');
          try { if (url) p.setLinkUrl(url); } catch (_) {}
        }
      }
    }

    // ----- Save & export -----
    doc.saveAndClose();
    var pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf').setName(docTitle + '.pdf');
    var pdfFile = destFolder.createFile(pdfBlob);

    try { if (tplId && docFile && docFile.getId) DriveApp.getFileById(docFile.getId()).setTrashed(true); } catch (_){}

    try {
      if (typeof logAudit === 'function') {
        try { logAudit('system','createEmployeeHistoryPDF', null, { name: name, pdfId: pdfFile.getId() }); }
        catch (_1) { try { logAudit('createEmployeeHistoryPDF','ok', null, { name: name, pdfId: pdfFile.getId() }); } catch (_2) {} }
      }
    } catch(_){}

    // return plain id (so caller builds correct URL)
    return pdfFile.getId();

  } catch (e) {
    try { if (typeof logError === 'function') logError('createEmployeeHistoryPDF', e, { name: name }); } catch (_){}
    return null;
  }
}

function _buildEmployeeHistoryTable_(employee){
  var s = sh_(CONFIG.TABS.EVENTS);
  var hdr = headers_(s), map = headerIndexMap_(s);
  var last = s.getLastRow(); if (last < 2) return { table: [['Date','Event','Lead','Infraction','Pts','Roll','Notes','PDF']], rows: [], writeMap: {} };

  function get(r, h){ var c = map[h]; return c ? s.getRange(r, c).getValue() : ''; }
  function getDisp(r, h){ var c = map[h]; return c ? s.getRange(r, c).getDisplayValue() : ''; }

  var want = String(employee || '').trim();
  var out = [['Date','Event','Lead','Infraction','Pts','Roll','Notes','PDF']];
  var rowsMeta = [];
  var writeMap = {};

  for (var r = 2; r <= last; r++){
    var emp = get(r, CONFIG.COLS.Employee); if (String(emp || '').trim() !== want) continue;

    var evt    = String(get(r, CONFIG.COLS.EventType) || '').trim();
    var milLab = String(get(r, CONFIG.COLS.Milestone) || '').trim();
    var isMil  = (evt.toLowerCase() === 'milestone') || !!milLab;
    var isGrace= (evt.toLowerCase() === ((CONFIG.POLICY && CONFIG.POLICY.GRACE && CONFIG.POLICY.GRACE.eventTypeName) || 'Grace').toLowerCase());
    var isPos  = _isPositiveCreditEventText_(evt);

    // Base fields
    var date = get(r, CONFIG.COLS.IncidentDate) || get(r, CONFIG.COLS.Timestamp) || '';
    var lead = getDisp(r, CONFIG.COLS.Lead) || '';
    var infr = getDisp(r, CONFIG.COLS.Infraction) || '';
    var pts  = Number(get(r, CONFIG.COLS.Points) || 0);
    var roll = Number(get(r, CONFIG.COLS.PointsRollingEffective) || get(r, CONFIG.COLS.PointsRolling) || 0);
    var notes= getDisp(r, CONFIG.COLS.NotesReviewer) || '';
    var pdfUrl = (function(){
      try{
        var c = map[CONFIG.COLS.PdfLink] || map['Write-Up PDF'] || 0;
        if (!c) return '';
        var rich = readLinkUrlFromCell_(s.getRange(r, c));
        return rich || String(s.getRange(r, c).getDisplayValue() || '');
      }catch(_){ return ''; }
    })();

    // Nullify (checkbox) → treat as 0 pts
    var nullified = !!(map[CONFIG.COLS.Nullify] && (get(r, CONFIG.COLS.Nullify) === true));
    if (nullified) pts = 0;

    // Per your rules:
    // - Grace events: Lead = Graced By; Notes = Grace Reason; Pts = 0; Event = "Grace"
    if (isGrace){
      var by = getDisp(r, CONFIG.COLS.GracedBy) || '';
      var why = getDisp(r, CONFIG.COLS.GraceReason) || '';
      lead = by || lead;
      notes = why || notes;
      pts = 0;
      infr = ''; // not specified to show anything here
    }

    // - Positive events: Event="Positive Points"; Infraction = PositiveAction; Notes=Consumed?; PDF blank
    if (isPos){
      var posAct = getDisp(r, CONFIG.COLS.PositiveActionType) || getDisp(r, 'PositiveAction') || getDisp(r, 'Positive Action') || '';
      infr = posAct || infr;
      notes = 'Consumed: ' + _lookupConsumedForEarnedEventRow_(r);
      pdfUrl = ''; // do not show a PDF for positive events
    }

    // - Milestones: Event=label; Pts blank; Lead=Consequence Director; Notes=Pending Status
    var eventName = evt;
    if (isMil){
      eventName = milLab || 'Milestone';
      pts = ''; // blank by request
      var cd = getDisp(r, CONFIG.COLS.ConsequenceDirector) || '';
      if (cd) lead = cd; // show director who claimed
      var pend = map[CONFIG.COLS.PendingStatus] ? getDisp(r, CONFIG.COLS.PendingStatus) : '';
      if (pend) notes = pend;
    }

    // If an original event got nullified (because of Grace), reflect director/reason too
    if (!isGrace && nullified){
      var by2 = getDisp(r, CONFIG.COLS.GracedBy) || '';
      var why2 = getDisp(r, CONFIG.COLS.GraceReason) || '';
      if (by2) lead = by2;
      if (why2) notes = why2;
    }

    // Build row
    out.push([
      date ? formatDate_(date) : '',
      eventName || evt || (isPos ? 'Positive Points' : ''),
      lead || '',
      infr || '',
      (pts === '' ? '' : String(pts || 0)),
      String(roll || 0),
      notes || '',
      pdfUrl ? 'View PDF' : ''
    ]);

    // Track urls for link pass
    if (pdfUrl){
      writeMap[r] = { url: pdfUrl };
    }
    rowsMeta.push({ sheetRowIndex: r });
  }

  // sort by date ascending
  out = [out[0]].concat(out.slice(1).sort(function(a,b){
    var ad = new Date(a[0] || 0).getTime(); var bd = new Date(b[0] || 0).getTime();
    return ad - bd;
  }));

  return { table: out, rows: rowsMeta, writeMap: writeMap };
}


// utils.gs (or docService.gs)
function readLinkUrlFromCell_(range) {
  const rich = range.getRichTextValue();
  if (!rich) return null;
  const runs = rich.getRuns();
  for (var i = 0; i < runs.length; i++) {
    const url = runs[i].getLinkUrl();
    if (url) return url;
  }
  return null;
}