/**
 * DIAGNOSTIC HELPERS — paste into project and save
 */

/** 1) Build payload exactly as PDF generator will (uses buildPdfDataFromRow_ if present) */
function _diag_buildPayloadForRow(row){
  var s = sh_(CONFIG.TABS.EVENTS);
  if (typeof buildPdfDataFromRow_ === 'function'){
    return buildPdfDataFromRow_(s, row, { generateWriteUpTitle: true, includeHeaders: true });
  } else {
    // fallback: mimic createEventRecordPdf_ behavior
    var hdrs = headers_(s);
    var vals = s.getRange(row,1,1,hdrs.length).getDisplayValues()[0];
    var rowObj = asRowObject_(hdrs, vals);
    var data = {
      Employee: rowObj[CONFIG.COLS.Employee] || '',
      IncidentDate: rowObj[CONFIG.COLS.IncidentDate] || '',
      Lead: rowObj[CONFIG.COLS.Lead] || '',
      Policy: rowObj[CONFIG.COLS.RelevantPolicy] || rowObj.Policy || '',
      Infraction: rowObj[CONFIG.COLS.Infraction] || '',
      Description: rowObj[CONFIG.COLS.NotesReviewer] || rowObj.IncidentDescription || '',
      CorrectiveActions: rowObj[CONFIG.COLS.CorrectiveActions] || '',
      TeamMemberStatement: rowObj[CONFIG.COLS.TeamMemberStatement] || '',
      Points: rowObj[CONFIG.COLS.Points] || ''
    };
    // add headers and compact aliases
    hdrs.forEach(function(k){ if (k && !data.hasOwnProperty(k)) data[k] = rowObj[k]||''; });
    Object.keys(Object.assign({}, data)).forEach(function(k){ if(!k) return; var alias=String(k).replace(/[\s\-_]+/g,'').replace(/[^\w]/g,''); if(alias && !data.hasOwnProperty(alias)) data[alias] = data[k]; });
    // EventID and WriteUpTitle
    if (!data.EventID) data.EventID = data.LinkedEventID || String(row);
    if (!data.WriteUpTitle) data.WriteUpTitle = 'Write-Up: ' + (data.Employee||'unknown') + ' (' + (data.IncidentDate||'no-date') + ')';
    return data;
  }
}

/** 2) Read tokens from a Doc template (curly-brace tokens) */
function _diag_listDocTokens(docId){
  var doc = DocumentApp.openById(docId);
  var text = (doc.getBody && doc.getBody().getText ? doc.getBody().getText() : '') + '\n' +
             ((doc.getHeader && doc.getHeader()) ? doc.getHeader().getText() : '') + '\n' +
             ((doc.getFooter && doc.getFooter()) ? doc.getFooter().getText() : '');
  var re = /\{\{\s*([^}]+?)\s*\}\}/g, found=[], m;
  while((m=re.exec(text))!==null){
    var t = String(m[1]).trim();
    if(found.indexOf(t)===-1) found.push(t);
  }
  return found;
}

/** 3) Simulate merge (string replacement) — returns preview snippets */
function _diag_simulateMergeText(templateDocId, payload){
  var doc = DocumentApp.openById(templateDocId);
  var bodyText = doc.getBody? doc.getBody().getText() : '';
  var hdrText  = (doc.getHeader && doc.getHeader()) ? doc.getHeader().getText() : '';
  var ftrText  = (doc.getFooter && doc.getFooter()) ? doc.getFooter().getText() : '';
  function findKeyForToken(token){
    var keys = Object.keys(payload||{});
    var lowTok = token.toLowerCase().replace(/\s+/g,'').replace(/[_\-]/g,'');
    for(var i=0;i<keys.length;i++){ if(String(keys[i]).toLowerCase()===token.toLowerCase()) return keys[i]; }
    for(var j=0;j<keys.length;j++){ var k = String(keys[j]).toLowerCase().replace(/\s+/g,'').replace(/[_\-]/g,''); if(k===lowTok) return keys[j]; }
    return null;
  }
  function repl(text){
    return text.replace(/\{\{\s*([^}]+?)\s*\}\}/g, function(full, inner){
      var t = String(inner||'').trim();
      var k = findKeyForToken(t);
      if (k !== null) return String(payload[k]===null||payload[k]===undefined ? '' : payload[k]);
      return ''; // simulate blank for missing tokens
    });
  }
  return {
    bodyPreview: repl(bodyText).slice(0,4000),
    headerPreview: repl(hdrText).slice(0,2000),
    footerPreview: repl(ftrText).slice(0,2000)
  };
}

/** 4) Find the intermediate "(merge)" doc copy that mergeToPdf_ creates (if present) */
function _diag_findLatestMergeCopyByOutNamePrefix(outNamePrefix){
  var folder = DriveApp.getFolderById(CONFIG.DEST_FOLDER_ID);
  var files = folder.getFiles();
  var matches = [];
  while(files.hasNext()){
    var f = files.next();
    var n = f.getName();
    if (n.indexOf(outNamePrefix + ' (merge)') === 0 || n.indexOf('(merge)') !== -1 && n.indexOf(outNamePrefix)!==-1){
      matches.push({name: n, id: f.getId(), date: f.getDateCreated()});
    }
  }
  // sort by date desc
  matches.sort(function(a,b){ return (b.date||0) - (a.date||0); });
  return matches;
}

/** 5) Inspect a doc copy's raw text to see whether tokens remain un-replaced */
function _diag_inspectDocById(docId){
  var doc = DocumentApp.openById(docId);
  return {
    id: docId,
    name: doc.getName(),
    bodyText: doc.getBody? doc.getBody().getText().slice(0,8000) : '',
    headerText: (doc.getHeader && doc.getHeader()) ? doc.getHeader().getText().slice(0,2000) : '',
    footerText: (doc.getFooter && doc.getFooter()) ? doc.getFooter().getText().slice(0,2000) : '',
    tokensRemaining: (function(){
      var txt = (doc.getBody?doc.getBody().getText():'') + '\n' + ((doc.getHeader&&doc.getHeader())?doc.getHeader().getText():'') + '\n' + ((doc.getFooter&&doc.getFooter())?doc.getFooter().getText():'');
      var re = /\{\{\s*([^}]+?)\s*\}\}/g, out=[], m;
      while((m=re.exec(txt))!==null){ var t=String(m[1]).trim(); if(out.indexOf(t)===-1) out.push(t); }
      return out;
    })()
  };
}

/** 6) Full diagnostic runner for event template (dry-run, no writes) */
function diagnoseEventPdfFull(row){
  try{
    var templateId = CONFIG.TEMPLATES.EVENT_RECORD;
    if(!templateId) throw new Error('CONFIG.TEMPLATES.EVENT_RECORD missing');

    var payload = _diag_buildPayloadForRow(row);
    var tokens = _diag_listDocTokens(templateId);
    var preview = _diag_simulateMergeText(templateId, payload);

    // compute missing tokens (same rules)
    var payloadKeys = Object.keys(payload || {});
    var missing = [];
    tokens.forEach(function(tok){
      var lowTok = tok.toLowerCase().replace(/\s+/g,'').replace(/[_\-]/g,'');
      var matched = payloadKeys.some(function(k){
        if(!k) return false;
        var kn = k.toLowerCase().replace(/\s+/g,'').replace(/[_\-]/g,'');
        return kn === lowTok || k.toLowerCase() === tok.toLowerCase();
      });
      if(!matched) missing.push(tok);
    });

    // attempt to find any intermediate "(merge)" copies with filename prefix from payload.WriteUpTitle (if present)
    var outNamePrefix = (payload && payload.WriteUpTitle) ? payload.WriteUpTitle : ('Write-Up: ' + (payload.Employee||'unknown'));
    var copies = [];
    try{ copies = _diag_findLatestMergeCopyByOutNamePrefix(outNamePrefix); }catch(_){ copies = []; }

    var result = {
      row: row,
      templateId: templateId,
      tokensFound: tokens,
      payloadKeys: payloadKeys,
      missingTokens: missing,
      simulatedPreview: preview,
      foundCopies: copies
    };

    Logger.log(JSON.stringify(result, null, 2));
    // write compact diag row
    try{ if (CONFIG && CONFIG.TABS && CONFIG.TABS.LOGS){ var logs=ss_().getSheetByName(CONFIG.TABS.LOGS); if(logs) logs.appendRow([new Date(),'DIAG','diagnoseEventPdfFull','row:'+row,'missing:'+JSON.stringify(missing),'copies:'+copies.length]); } }catch(_){}
    return result;
  }catch(err){
    logError && logError('diagnoseEventPdfFull', err, {row: row});
    throw err;
  }
}

/** 7) If you find a copy from foundCopies, inspect first match quickly */
function diagInspectFirstCopyForRow(row){
  var r = diagnoseEventPdfFull(row);
  if (!r.foundCopies || r.foundCopies.length===0) return { error: 'no copies found' };
  var id = r.foundCopies[0].id;
  return _diag_inspectDocById(id);
}

function TEST_runDiagnoseEventPdfFull(){
  var row = 2; // ← change to the Events row that produced the bad PDF
  var res = diagnoseEventPdfFull(row);
  Logger.log(JSON.stringify(res, null, 2));
}

function TEST_checkDestFolder(){
  try{
    Logger.log('DEST_FOLDER_ID=' + (CONFIG && CONFIG.DEST_FOLDER_ID));
    var f = DriveApp.getFolderById(CONFIG.DEST_FOLDER_ID);
    Logger.log('Folder name: ' + f.getName() + ' — id OK');
    // list recent 5 files for quick sanity
    var it = f.getFiles();
    var out = [];
    var i=0;
    while(it.hasNext() && i<5){ var file = it.next(); out.push({name:file.getName(), id:file.getId(), date:file.getDateCreated()}); i++; }
    Logger.log('Recent files in folder: ' + JSON.stringify(out));
    return { ok:true, folderName: f.getName(), recent: out };
  }catch(e){
    Logger.log('DEST_FOLDER error: ' + e);
    return { ok:false, err: String(e) };
  }
}

function TEST_mergeToPdfAndInspectRow_wrapper(){
  var rowToTest = 2; // <-- set this to the failing Events row
  try {
    var res = TEST_mergeToPdfAndInspectRow(rowToTest); // calls the test helper you already have
    Logger.log('TEST_mergeToPdfAndInspectRow result:\n' + JSON.stringify(res, null, 2));
  } catch (e) {
    Logger.log('TEST_mergeToPdfAndInspectRow_wrapper ERROR: ' + e);
    throw e;
  }
}

function TEST_mergeToPdfAndInspectRow(row){
  try{
    // Build payload exactly like createEventRecordPdf_ does
    var payload = _diag_buildPayloadForRow(row); // you already have this diag helper
    var tmpl = CONFIG.TEMPLATES.EVENT_RECORD;
    if(!tmpl) throw new Error('CONFIG.TEMPLATES.EVENT_RECORD missing');

    Logger.log('Calling mergeToPdf_ with template: ' + tmpl);
    var pdfId = mergeToPdf_(tmpl, payload, (payload.WriteUpTitle || ('Write-Up: ' + (payload.Employee||'unknown'))));
    Logger.log('mergeToPdf_ returned pdfId: ' + pdfId);

    // Now inspect the output file and any intermediate copy
    var folder = DriveApp.getFolderById(CONFIG.DEST_FOLDER_ID);
    // newest PDF file with the output name (ends with .pdf)
    var pdfFile = DriveApp.getFileById(pdfId);
    Logger.log('PDF file name: ' + pdfFile.getName() + ' size:' + pdfFile.getSize());

    // Look for intermediate "(merge)" doc copies with WriteUpTitle prefix
    var prefix = payload.WriteUpTitle || ('Write-Up: ' + (payload.Employee||'unknown'));
    var copies = [];
    var files = folder.getFiles();
    while(files.hasNext()){
      var f = files.next();
      if (f.getName().indexOf(prefix + ' (merge)') === 0 || (f.getName().indexOf('(merge)') !== -1 && f.getName().indexOf(prefix)!==-1) ){
        copies.push({name:f.getName(), id:f.getId(), date:f.getDateCreated()});
      }
    }
    Logger.log('Intermediate copies found: ' + JSON.stringify(copies));
    // If there is at least one copy, inspect tokens remaining on the first
    if(copies.length>0){
      var docInfo = _diag_inspectDocById(copies[0].id);
      Logger.log('First copy inspection: ' + JSON.stringify({name: docInfo.name, tokensRemaining: docInfo.tokensRemaining}));
    }
    return { pdfId: pdfId, pdfName: pdfFile.getName(), copies: copies };
  }catch(e){
    Logger.log('TEST_mergeToPdfAndInspectRow error: ' + e);
    return { err: String(e) };
  }
}


function TEST_writeLinkBack(row){
  try{
    var s = sh_(CONFIG.TABS.EVENTS);
    var hdrMap = headerIndexMap_(s);
    var pdfColName = CONFIG.COLS.PdfLink || 'Write-Up PDF';
    // Try writing a test URL into the row (won't affect merge)
    var testUrl = 'https://example.com/test-' + new Date().getTime();
    try{
      setRichLinkSafe(s, row, pdfColName, 'View PDF', testUrl);
      Logger.log('setRichLinkSafe succeeded for header: ' + pdfColName);
    }catch(e){
      Logger.log('setRichLinkSafe failed: ' + e);
      // fallback: try raw write to column index
      var idx = hdrMap[pdfColName] || hdrMap['Write-Up PDF'] || (s.getLastColumn()+1);
      s.getRange(row, idx).setValue(testUrl);
      Logger.log('Wrote fallback plain URL to column ' + idx);
    }
    return { ok: true };
  }catch(e){
    Logger.log('TEST_writeLinkBack error: ' + e);
    return { err: String(e) };
  }
}

function TEST_simulateMergeForRow(){
  var rowToTest = 2; // <-- set this to the failing Events row
  try {
    var payload = _diag_buildPayloadForRow(rowToTest); // from diagnostics you pasted earlier
    var templateId = CONFIG.TEMPLATES.EVENT_RECORD;
    if (!templateId) throw new Error('CONFIG.TEMPLATES.EVENT_RECORD not set');
    var preview = _diag_simulateMergeText(templateId, payload);
    Logger.log('Payload keys: ' + JSON.stringify(Object.keys(payload)));
    Logger.log('Preview (body slice):\n' + preview.bodyPreview);
    Logger.log('Preview (footer slice):\n' + preview.footerPreview);
    return { payloadKeys: Object.keys(payload), preview: preview };
  } catch (e) {
    Logger.log('TEST_simulateMergeForRow ERROR: ' + e);
    throw e;
  }
}

// Set rowToTest to the Events row that previously produced a blank PDF, then run this wrapper.
function TEST_mergeToPdfDebug_wrapper(){
  var rowToTest = 2; // <-- set this to the failing Events row
  try{
    // build payload exactly like createEventRecordPdf_ uses
    var payload = _diag_buildPayloadForRow(rowToTest); // diagnostic helper you already have
    var tmpl = CONFIG.TEMPLATES.EVENT_RECORD;
    if (!tmpl) throw new Error('CONFIG.TEMPLATES.EVENT_RECORD not set');

    Logger.log('Calling mergeToPdf_debug with template: ' + tmpl + ' and payload keys: ' + JSON.stringify(Object.keys(payload)));
    var res = mergeToPdf_debug(tmpl, payload, (payload.WriteUpTitle || ('Write-Up: ' + (payload.Employee||'unknown'))));
    Logger.log('mergeToPdf_debug result: ' + JSON.stringify(res, null, 2));
    // Also log intermediate doc info for quick inspection link
    try{
      var copyDoc = DocumentApp.openById(res.copyId);
      Logger.log('Intermediate doc name: ' + copyDoc.getName() + ' — id: ' + res.copyId);
    }catch(e){ Logger.log('Could not open intermediate doc: ' + e); }
    return res;
  }catch(e){
    Logger.log('TEST_mergeToPdfDebug_wrapper error: ' + e);
    throw e;
  }
}


/**
 * Debug variant of mergeToPdf_:
 * - performs the same merge but logs replaced/missing tokens
 * - DOES NOT trash the intermediate doc copy (keeps it for manual inspection)
 * - returns { pdfId, copyId, replacedTokens, missingTokens }
 */
function mergeToPdf_debug(docId, data, outName){
  if (!docId) throw new Error('Template Doc ID not set');
  return withBackoff_('mergeToPdf_debug', function(){
    var folder = DriveApp.getFolderById(CONFIG.DEST_FOLDER_ID);
    var copy = DriveApp.getFileById(docId).makeCopy(outName+' (merge)', folder);
    var copyId = copy.getId();
    var doc = DocumentApp.openById(copyId);
    var body = doc.getBody();
    var header = (doc.getHeader && doc.getHeader()) || null;
    var footer = (doc.getFooter && doc.getFooter()) || null;

    function escRegex(s){ return String(s).replace(/[-\/\\^$*+?.()|[\]{}]/g,'\\$&'); }

    // collect tokens present
    var docText = (body && body.getText? body.getText() : '') + '\n' +
                  (header && header.getText? header.getText() : '') + '\n' +
                  (footer && footer.getText? footer.getText() : '');
    var tokRe = /\{\{\s*([^}]+?)\s*\}\}/g;
    var foundTokens = [];
    var m;
    while ((m = tokRe.exec(docText)) !== null){
      var t = String(m[1]).trim();
      if (foundTokens.indexOf(t) === -1) foundTokens.push(t);
    }

    // matching helper (case & spacing tolerant)
    function findKeyForToken(token){
      if (!data) return null;
      var keys = Object.keys(data);
      var lowTok = token.toLowerCase().replace(/\s+/g,'').replace(/[_\-]/g,'');
      for (var i=0;i<keys.length;i++){
        if (String(keys[i]).toLowerCase() === token.toLowerCase()) return keys[i];
      }
      for (var j=0;j<keys.length;j++){
        var k = String(keys[j]).toLowerCase().replace(/\s+/g,'').replace(/[_\-]/g,'');
        if (k === lowTok) return keys[j];
      }
      return null;
    }

    var replacedTokens = [], missingTokens = [];

    foundTokens.forEach(function(tok){
      var key = findKeyForToken(tok);
      var replacement = '';
      if (key !== null && typeof data[key] !== 'undefined' && data[key] !== null) {
        replacement = String(data[key]);
      } else {
        // try case-insensitive fallback
        var alt = null;
        for (var kk in data){
          if (kk && kk.toLowerCase() === tok.toLowerCase()){ alt = kk; break; }
        }
        if (alt) replacement = String(data[alt]||'');
      }

      try {
        var pattern = '\\{\\{\\s*' + escRegex(tok) + '\\s*\\}\\}';
        body.replaceText(pattern, replacement);
        if (header) try{ header.replaceText(pattern, replacement); }catch(_){}
        if (footer) try{ footer.replaceText(pattern, replacement); }catch(_){}
        if (replacement !== '') replacedTokens.push(tok); else missingTokens.push(tok);
      } catch(e){
        try{ logError && logError('mergeToPdf_debug_replaceErr', e, {token: tok}); }catch(_){}
        missingTokens.push(tok);
      }
    });

    doc.saveAndClose();

    // create the PDF file
    var pdf = DriveApp.getFileById(copyId).getAs('application/pdf');
    var outFile = folder.createFile(pdf).setName(outName + '.pdf');

    // intentionally DO NOT trash the copy — keep it for inspection during debug
    // try{ DriveApp.getFileById(copyId).setTrashed(true); }catch(_){}

    try{ logInfo_ && logInfo_('mergeToPdf_debug_done', 'out='+outFile.getId() + ' replaced=' + JSON.stringify(replacedTokens) + ' missing=' + JSON.stringify(missingTokens)); }catch(_){}

    return { pdfId: outFile.getId(), copyId: copyId, replacedTokens: replacedTokens, missingTokens: missingTokens };
  }, 4, 300);
}


function TEST_triggerContext(){
  try{
    Logger.log('EffectiveUser: ' + (Session.getEffectiveUser && Session.getEffectiveUser().getEmail ? Session.getEffectiveUser().getEmail() : 'unknown'));
  }catch(e){ Logger.log('Session.getEffectiveUser failed: ' + e); }
  try{
    var f = DriveApp.getFolderById(CONFIG.DEST_FOLDER_ID);
    Logger.log('DEST folder OK: ' + f.getName());
  }catch(e){
    Logger.log('DEST folder access failed: ' + e);
  }
}

function TEST_inspectPdfFromLog(){
  var pdfId = '1wkIeaSyIK1HZIwNPE_d7e568Z8ZPpWKb'; // replace with the pdfId from your onFormSubmit log
  try {
    var f = DriveApp.getFileById(pdfId);
    Logger.log(JSON.stringify({
      id: f.getId(),
      name: f.getName(),
      size: f.getSize(),
      created: f.getDateCreated(),
      updated: f.getLastUpdated(),
      owner: (f.getOwner && f.getOwner()) ? f.getOwner().getEmail() : 'unknown'
    }, null, 2));
  } catch (e) {
    Logger.log('inspectPdf error: ' + e);
    throw e;
  }
}

/**
 * Very verbose debug variant: keeps copy, logs before/after snippets, returns copyId + pdfId.
 * Use only temporarily for debugging.
 */
function mergeToPdf_debug_verbose(docId, data, outName){
  if (!docId) throw new Error('Template Doc ID not set');
  return withBackoff_('mergeToPdf_debug_verbose', function(){
    var folder = DriveApp.getFolderById(CONFIG.DEST_FOLDER_ID);
    var copyFile = DriveApp.getFileById(docId).makeCopy(outName + ' (merge)', folder);
    var copyId = copyFile.getId();

    Logger.log('mergeToPdf_debug_verbose copyId=' + copyId + ' copyName=' + copyFile.getName());

    var doc = DocumentApp.openById(copyId);
    var body = doc.getBody();
    var header = (doc.getHeader && doc.getHeader()) || null;
    var footer = (doc.getFooter && doc.getFooter()) || null;

    // Get before snapshot (body + header + footer)
    var before = '';
    try {
      before = (body && body.getText ? body.getText() : '').slice(0, 1200) + '\n--HDR--\n' +
               ((header && header.getText) ? header.getText().slice(0,400) : '') + '\n--FTR--\n' +
               ((footer && footer.getText) ? footer.getText().slice(0,400) : '');
    } catch(e) { before = 'ERR_before:' + e; }

    Logger.log('mergeToPdf_debug_verbose BEFORE snippet:\n' + before);

    // Replacement logic (relaxed, same as hardened function)
    function escRegex(s){ return String(s).replace(/[-\/\\^$*+?.()|[\]{}]/g,'\\$&'); }
    function normalizeKey(s){ return String(s||'').toLowerCase().replace(/[\s_\-]+/g,''); }

    var docText = (body && body.getText ? body.getText() : '') + '\n' + (header && header.getText ? header.getText() : '') + '\n' + (footer && footer.getText ? footer.getText() : '');
    var tokRe = /\{\{\s*([^}]+?)\s*\}\}/g;
    var foundTokens = [], m;
    while ((m = tokRe.exec(docText)) !== null){
      var t = String(m[1]).trim();
      if (foundTokens.indexOf(t) === -1) foundTokens.push(t);
    }

    var dataKeyMap = {};
    Object.keys(data || {}).forEach(function(k){ var nk = normalizeKey(k); if (nk && !dataKeyMap[nk]) dataKeyMap[nk] = k; });

    function findKeyForToken(token){
      if (!token) return null;
      var nk = normalizeKey(token);
      if (dataKeyMap[nk]) return dataKeyMap[nk];
      var keys = Object.keys(data || {});
      for (var i=0;i<keys.length;i++) if (String(keys[i]).toLowerCase() === token.toLowerCase()) return keys[i];
      return null;
    }

    foundTokens.forEach(function(tok){
      var key = findKeyForToken(tok);
      var repl = '';
      if (key !== null && typeof data[key] !== 'undefined' && data[key] !== null) repl = String(data[key]);
      var pattern = '\\{\\{\\s*' + escRegex(tok) + '\\s*\\}\\}';
      var rx = new RegExp(pattern, 'g');
      try{ body.replaceText(rx, repl); }catch(e){ Logger.log('body.replaceText err ' + e + ' token=' + tok); }
      try{ if (header) header.replaceText(rx, repl); }catch(e){ Logger.log('header.replaceText err ' + e + ' token=' + tok); }
      try{ if (footer) footer.replaceText(rx, repl); }catch(e){ Logger.log('footer.replaceText err ' + e + ' token=' + tok); }
    });

    doc.saveAndClose();

    // after snapshot
    var after = '';
    try {
      var doc2 = DocumentApp.openById(copyId);
      var b2 = doc2.getBody();
      var h2 = (doc2.getHeader && doc2.getHeader()) || null;
      var f2 = (doc2.getFooter && doc2.getFooter()) || null;
      after = (b2 && b2.getText ? b2.getText().slice(0,1200) : '') + '\n--HDR--\n' + ((h2 && h2.getText)?h2.getText().slice(0,400):'') + '\n--FTR--\n' + ((f2 && f2.getText)?f2.getText().slice(0,400):'');
    } catch(e) { after = 'ERR_after:' + e; }

    Logger.log('mergeToPdf_debug_verbose AFTER snippet:\n' + after);

    // create pdf
    var pdfBlob = DriveApp.getFileById(copyId).getAs('application/pdf');
    var outFile = folder.createFile(pdfBlob).setName(outName + '.pdf');

    Logger.log('mergeToPdf_debug_verbose outPdfId=' + outFile.getId());

    // DO NOT trash the copy — we want to inspect it
    Logger.log('mergeToPdf_debug_verbose leaving intermediate copy in Drive for inspection: ' + copyId);

    return { copyId: copyId, pdfId: outFile.getId(), beforeSnippet: before, afterSnippet: after, foundTokens: foundTokens };
  }, 4, 300);
}

/** wrapper to call the verbose debug for a specific row — run this from editor */
function TEST_mergeToPdfDebugVerbose_wrapper(){
  var rowToTest = 3; // <--- set the row that produced the PDF you inspected
  try {
    var payload = _diag_buildPayloadForRow(rowToTest); // or buildPdfDataFromRow_(sh_(CONFIG.TABS.EVENTS), rowToTest)
    var tmpl = CONFIG.TEMPLATES.EVENT_RECORD;
    var result = mergeToPdf_debug_verbose(tmpl, payload, (payload.WriteUpTitle || ('Write-Up: ' + (payload.Employee||'unknown'))));
    Logger.log('TEST_mergeToPdfDebugVerbose_wrapper result: ' + JSON.stringify(result, null, 2));
    return result;
  } catch(e) {
    Logger.log('TEST_mergeToPdfDebugVerbose_wrapper error: ' + e);
    throw e;
  }
}

// Installs/refreshes the PointsRolling (Effective) ARRAYFORMULA anchored to IncidentDate,
// and excludes rows where Nullify = TRUE (those rows contribute 0 to the rolling sum).
// function ensureRollingEffectiveHeader(sheet){
//   try{
//     sheet = sheet || sh_(CONFIG.TABS.EVENTS);
//     var map = headerIndexMap_(sheet);
//     var headerName = CONFIG.COLS.PointsRollingEffective || 'PointsRolling (Effective)';

//     var colEff  = map[headerName];
//     var colEmp  = map[CONFIG.COLS.Employee];
//     var colDate = map[CONFIG.COLS.IncidentDate];
//     var colPts  = map[CONFIG.COLS.Points];
//     var colNull = map[CONFIG.COLS.Nullify] || map['Nullify'] || 0;

//     // required columns
//     if (!colEff || !colEmp || !colDate || !colPts) {
//       logInfo_ && logInfo_('ensureRollingEffectiveHeader_', 'required column missing; skipping header install');
//       return false;
//     }

//     // helper: convert 1-based column index to A1 letter(s)
//     function toLetter(n){
//       var s = '';
//       while (n > 0){
//         var m = (n - 1) % 26;
//         s = String.fromCharCode(65 + m) + s;
//         n = Math.floor((n - m - 1) / 26);
//       }
//       return s;
//     }

//     var L_eff  = toLetter(colEff);
//     var L_emp  = toLetter(colEmp);
//     var L_date = toLetter(colDate);
//     var L_pts  = toLetter(colPts);
//     var L_null = colNull ? toLetter(colNull) : null;

//     // row-wise ranges (relative)
//     var empRel  = L_emp + '2:' + L_emp;
//     var dateRel = L_date + '2:' + L_date;

//     // absolute ranges for FILTER
//     var empAbs  = '$' + L_emp + '$2:$' + L_emp;
//     var dateAbs = '$' + L_date + '$2:$' + L_date;
//     var ptsAbs  = '$' + L_pts + '$2:$' + L_pts;
//     var nullAbs = L_null ? ('$' + L_null + '$2:$' + L_null) : null;

//     var rollingDays = Number((CONFIG.POLICY && CONFIG.POLICY.ROLLING_DAYS) || 180);

//     // Build FILTER criteria string, include Nullify <> TRUE only if present
//     var filterCrit = ptsAbs + ', ' + empAbs + '=emp, ' + dateAbs + '>=dt-' + rollingDays + ', ' + dateAbs + '<=dt' + (nullAbs ? (', ' + nullAbs + '<>TRUE') : '');

//     // Build target formula. Careful with parentheses and commas (Sheets is picky).
//     var targetFormula =
//       '=ARRAYFORMULA({' +
//         '"' + headerName + '"; ' +
//         'IF(LEN(' + empRel + '), ' +
//           'MAP(' + empRel + ', ' + dateRel + ', ' +
//             'LAMBDA(emp, dt, IF(ISNUMBER(dt), SUM(FILTER(' + filterCrit + ')), 0))' +
//           ')' +
//         ')' +
//       '})';

//     // Compare normalized existing formula/value to avoid unnecessary rewrites
//     var headCell = sheet.getRange(1, colEff);
//     var existing = String(headCell.getFormula() || headCell.getDisplayValue() || '').replace(/\s+/g, '');
//     var want     = String(targetFormula).replace(/\s+/g, '');

//     if (existing !== want){
//       headCell.setFormula(targetFormula);
//       logInfo_ && logInfo_('ensureRollingEffectiveHeader_', 'installed/updated rolling header formula in col ' + colEff);
//     } else {
//       logInfo_ && logInfo_('ensureRollingEffectiveHeader_', 'rolling header formula already up-to-date');
//     }

//     return true;
//   }catch(e){
//     try{ logError && logError('ensureRollingEffectiveHeader_', e); }catch(_){}
//     return false;
//   }
// }

function TEST_simulateProbationFailure(employeeName, triggerRow){
  // find a synthetic trigger row (or pass real triggerRow). This function will run checkProbationFailureForEmployee_.
  try {
    var sr = triggerRow || 2;
    var res = checkProbationFailureForEmployee_(sr);
    Logger.log('simulate result: %s', JSON.stringify(res));
    return res;
  } catch(e) { Logger.log(e); return { error: String(e) }; }
}