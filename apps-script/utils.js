// utils.gs — CLEAR v2.x unified utilities (dev-ready)
// Single source of truth for spreadsheet helpers, logging, caching, policy shims,
// points/milestones helpers, and doc wiring. No top-level I/O beyond tiny caches.

// ============================================================================
// Spreadsheet selection (dev uses CONFIG.SHEET_ID for its bound Sheet)
// ============================================================================
function ss_() {
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && CONFIG && CONFIG.SHEET_ID && active.getId() === CONFIG.SHEET_ID) return active;
  } catch (_) { /* fall through */ }
  if (!CONFIG || !CONFIG.SHEET_ID) throw new Error('CONFIG.SHEET_ID missing; set it in Config.');
  return SpreadsheetApp.openById(CONFIG.SHEET_ID);
}

function sh_(name) {
  var s = ss_().getSheetByName(name);
  if (!s) throw new Error('Missing sheet: ' + name);
  return s;
}

function sheetRowUrl_(sh, row){
  try{
    var ssId = ss_().getId();
    var gid  = sh.getSheetId();
    var a1   = 'A'+row;
    return 'https://docs.google.com/spreadsheets/d/'+ssId+'/edit#gid='+gid+'&range='+a1;
  }catch(_){ return ''; }
}


// ============================================================================
// Header/Index caches (fast & safe)
// ============================================================================
var __HDR_CACHE__ = Object.create(null);
var __MAP_CACHE__ = Object.create(null);

function headers_(sheet) {
  sheet = (typeof sheet === 'string') ? sh_(sheet) : sheet;
  var key = sheet.getSheetId();
  if (__HDR_CACHE__[key]) return __HDR_CACHE__[key];
  var row = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0] || [];
  __HDR_CACHE__[key] = row;
  return row;
}

function headerIndexMap_(sheet) {
  sheet = (typeof sheet === 'string') ? sh_(sheet) : sheet;
  var key = sheet.getSheetId();
  if (__MAP_CACHE__[key]) return __MAP_CACHE__[key];
  var map = {};
  headers_(sheet).forEach(function (h, i) { if (h) map[String(h)] = i + 1; });
  __MAP_CACHE__[key] = map;
  return map;
}

function invalidateHeaderMapCache_(sheet) {
  var sh = (typeof sheet === 'string') ? sh_(sheet) : sheet;
  var key = sh.getSheetId();
  delete __HDR_CACHE__[key];
  delete __MAP_CACHE__[key];
}

// ============================================================================
// A1 helpers & tiny utils
// ============================================================================
function colIndex_(sheet, header) {
  var idx = headerIndexMap_(sheet)[header];
  if (!idx) throw new Error('Header not found: ' + header);
  return idx;
}
function columnToLetter_(col) { var s=''; while(col>0){var m=(col-1)%26; s=String.fromCharCode(65+m)+s; col=Math.floor((col-1)/26);} return s; }
function a1_(row,col){ return columnToLetter_(col)+row; }
function isTruthy_(v){ if (v===true) return true; var s=String(v||'').trim().toUpperCase(); return s==='TRUE'||s==='YES'||s==='Y'||s==='1'; }
function findHeader1_(hdr, names){ if(!hdr||!hdr.length) return 0; var L=hdr.map(function(h){return String(h||'').trim().toLowerCase();}); for(var k=0;k<names.length;k++){var i=L.indexOf(String(names[k]||'').trim().toLowerCase()); if(i!==-1) return i+1;} return 0; }
function toDate_(v){ if(!v) return null; if(Object.prototype.toString.call(v)==='[object Date]') return v; var d=new Date(v); return isNaN(d)?null:d; }
function addDaysLocal_(d, days){ var tz=Session.getScriptTimeZone()||'America/Chicago'; var base=new Date(Utilities.formatDate(new Date(d), tz, 'yyyy-MM-dd')+'T00:00:00'); base.setDate(base.getDate()+Number(days||0)); return base; }
function formatDate_(d){ return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function slug_(s){ return String(s||'').trim().replace(/\s+/g,' ').replace(/[\/:*?"<>|]/g,'').slice(0,120); }
function asRowObject_(headers, row){ var o={}; headers.forEach(function(h,i){o[h]=row[i];}); return o; }

// ============================================================================
// Logging / audit (resilient; no hard failures)
// ============================================================================
function logError(where, err, extra){
  try{
    var msg = (err && err.stack) ? err.stack : String(err);
    var tab = (CONFIG && CONFIG.TABS && CONFIG.TABS.LOGS) || 'Logs';
    var s = ss_().getSheetByName(tab);
    if (s) s.appendRow([new Date(),'ERROR',where||'',msg, JSON.stringify(extra||{})]);
    else Logger.log('ERROR '+(where||'')+': '+msg+' '+JSON.stringify(extra||{}));
  }catch(e){ try{ Logger.log('logError failed: '+e); }catch(_){} }
}
var logError_ = logError;

function logInfo_(tag, msg){
  try{
    var tab = (CONFIG && CONFIG.TABS && CONFIG.TABS.LOGS) || 'Logs';
    var s = ss_().getSheetByName(tab);
    if (s) s.appendRow([new Date(),'INFO',tag||'', (typeof msg==='string'?msg:JSON.stringify(msg)) ]);
    else Logger.log('INFO '+(tag||'')+': '+(typeof msg==='string'?msg:JSON.stringify(msg)));
  }catch(e){ try{ Logger.log('logInfo_ failed: '+e);}catch(_){} }
}

// Unified audit shim
function logAudit(){
  try{
    var actor='', action='', row='', meta={};
    var a=[].slice.call(arguments);
    if (a.length>=4){ actor=String(a[0]||''); action=String(a[1]||''); row=a[2]; meta=a[3]||{}; }
    else if (a.length===3){
      if (typeof a[1]==='number' || /^\d+$/.test(String(a[1]||''))){ actor=(Session.getActiveUser&&Session.getActiveUser().getEmail&&Session.getActiveUser().getEmail())||'script'; action=String(a[0]||''); row=a[1]; meta=a[2]||{}; }
      else { actor=String(a[0]||''); action=String(a[1]||''); meta=a[2]||{}; row=(meta.row||meta.rowIndex||meta.sheetRowIndex)||''; }
    } else if (a.length===2){
      if (a[1] && typeof a[1]==='object'){ actor=(Session.getActiveUser&&Session.getActiveUser().getEmail&&Session.getActiveUser().getEmail())||'script'; action=String(a[0]||''); meta=a[1]||{}; row=(meta.row||meta.rowIndex||meta.sheetRowIndex)||''; }
      else { actor=String(a[0]||''); action=String(a[1]||''); }
    } else if (a.length===1){ action=String(a[0]||''); }
    if (row && typeof row==='object') row=('row' in row)?row.row:'';

    var tab=(CONFIG&&CONFIG.TABS&&CONFIG.TABS.AUDIT)||'Audit';
    var s=ss_().getSheetByName(tab) || ss_().insertSheet(tab);
    if (s.getLastRow()===0) s.appendRow(['Timestamp','Actor','Action','Row','Details']);
    var details=''; try{details=JSON.stringify(meta||{});}catch(e){details=String(meta);}
    s.appendRow([new Date(), actor, action, row, details]);
  }catch(e){ try{ logError('logAudit_shim', e, {args:[].slice.call(arguments)});}catch(_){} }
}

// ============================================================================
// Locks & retry
// ============================================================================
function withLocks_(label, fn, waitMs){
  waitMs=Math.max(1000, Number(waitMs||30000));
  var sl=null, dl=null;
  try{ sl=LockService.getScriptLock(); sl.waitLock(waitMs); }catch(e){ logError('withLocks_script',e,{label:label}); return fn(); }
  try{ dl=LockService.getDocumentLock(); dl.waitLock(Math.max(500, Math.floor(waitMs/2))); }catch(e){ try{sl.releaseLock();}catch(_){} logError('withLocks_doc',e,{label:label}); throw e; }
  try{ return fn(); }
  finally{ try{dl&&dl.releaseLock();}catch(e){logError('withLocks_release_doc',e,{label:label});} try{sl&&sl.releaseLock();}catch(e){logError('withLocks_release_script',e,{label:label});} }
}

function withBackoff_(label, func, attempts, baseMs){
  attempts=Math.max(1, Number(attempts||4)); baseMs=Math.max(10, Number(baseMs||200));
  for (var i=0;i<attempts;i++){
    try{ return func(); }
    catch(err){
      var msg=String(err && (err.message||err))||'';
      var transient=/Too many simultaneous invocations|invoked too many times|Internal error|Service error|Rate Limit Exceeded|Quota exceeded|LockService/i.test(msg);
      if (!transient || i===attempts-1){ logError('withBackoff_'+label, err, {attempt:i+1, attempts:attempts}); throw err; }
      logInfo_('withBackoff_retry', label+' attempt '+(i+1)+' transient: '+msg);
      Utilities.sleep(baseMs*Math.pow(2,i));
    }
  }
}

// ============================================================================
// Row context & link helpers
// ============================================================================
function rowCtx_(sheet, row){
  var sh=(typeof sheet==='string')?sh_(sheet):sheet, map=headerIndexMap_(sh);
  return {
    row: row, sheet: sh,
    get: function(h){ var c=map[h]; if(!c) throw new Error('rowCtx_: header not found: '+h); return sh.getRange(row,c).getValue(); },
    set: function(h,v){ var c=map[h]; if(!c) throw new Error('rowCtx_: header not found: '+h); sh.getRange(row,c).setValue(v); return this; },
    setRichLink: function(h, text, url){ var c=map[h]; if(!c) throw new Error('rowCtx_: header not found: '+h); var r=sh.getRange(row,c); var rich=SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(url).build(); r.setRichTextValue(rich); return this; },
    a1: function(h){ var c=map[h]; if(!c) throw new Error('rowCtx_: header not found: '+h); return a1_(row,c); }
  };
}

function setRichLinkSafe(sheet, row, headerName, text, url){
  try{
    var sh=(typeof sheet==='string')?sh_(sheet):sheet;
    var hdrs=headers_(sh);
    var col=hdrs.indexOf(headerName)+1;
    if (!col){ col=sh.getLastColumn()+1; sh.getRange(1,col).setValue(headerName); invalidateHeaderMapCache_(sh); }
    var rich=SpreadsheetApp.newRichTextValue().setText(text||'View').setLinkUrl(url).build();
    sh.getRange(Number(row), col).setRichTextValue(rich);
    return true;
  }catch(e){
    try{ var sh2=(typeof sheet==='string')?sh_(sheet):sheet; var hdrs2=headers_(sh2); var c2=hdrs2.indexOf(headerName)+1; sh2.getRange(Number(row), c2||1).setValue((text||'View')+' — '+(url||'')); }catch(_){}
    return false;
  }
}

// ============================================================================
// Points column ARRAYFORMULA helpers
// ============================================================================
function buildPointsArrayFormula_(sheet){
  var map = headerIndexMap_(sheet);
  var ptsCol = map[CONFIG.COLS.Points], evtCol = map[CONFIG.COLS.EventType], infCol = map[CONFIG.COLS.Infraction];
  var paCol  = map[CONFIG.COLS.PositiveActionType] || map['PositiveActionType'] || 0;
  if (!ptsCol || !evtCol || !infCol || !paCol) return '';
  function L(c){return columnToLetter_(c);}
  return '=ARRAYFORMULA({"'+CONFIG.COLS.Points+'"; IF(' + L(evtCol) + '2:' + L(evtCol) + '="Positive Point Removal",' +
         'IFERROR(-XLOOKUP(' + L(paCol) + '2:' + L(paCol) + ', PositivePoints[PositiveActionType], PositivePoints[Points]),0),' +
         'IFERROR(XLOOKUP(' + L(infCol) + '2:' + L(infCol) + ', Rubric[Infraction], Rubric[Points]),0))})';
}

function ensureRollingEffectiveHeader_(sheetOrName){
  try{
    var sh = (typeof sheetOrName === 'string') ? sh_(sheetOrName) : (sheetOrName || sh_(CONFIG.TABS.EVENTS));
    var hdrs = headers_(sh), map=headerIndexMap_(sh);
    var effName=(CONFIG.COLS.PointsRollingEffective)||'PointsRolling (Effective)';
    var rollName=(CONFIG.COLS.PointsRolling)||'PointsRolling';
    var finName=(CONFIG.COLS.FinalPoints)||'FinalPoints';
    var tsName=(CONFIG.COLS.Timestamp)||'Timestamp';
    var cEff=map[effName]||0, cRoll=map[rollName]||0, cFin=map[finName]||0, cTs=map[tsName]||0;
    if (!cRoll && !cFin) return false;
    if (!cEff){ cEff=sh.getLastColumn()+1; sh.getRange(1,cEff).setValue(effName); invalidateHeaderMapCache_(sh); map=headerIndexMap_(sh); }
    function L(c){return columnToLetter_(c);}
    var gateL=cTs?L(cTs):'A';
    var rollR=cRoll?(L(cRoll)+'2:'+L(cRoll)):null, finR=cFin?(L(cFin)+'2:'+L(cFin)):null;
    var formula = rollR && finR
      ? '=ARRAYFORMULA({"'+effName+'"; IF(LEN('+gateL+'2:'+gateL+'), IF('+rollR+'<>"",'+rollR+', IF('+finR+'<>"",'+finR+', )), )})'
      : rollR
      ? '=ARRAYFORMULA({"'+effName+'"; IF(LEN('+gateL+'2:'+gateL+'), IF('+rollR+'<>"",'+rollR+', ), )})'
      : '=ARRAYFORMULA({"'+effName+'"; IF(LEN('+gateL+'2:'+gateL+'), IF('+finR+'<>"",'+finR+', ), )})';
    var head=sh.getRange(1,cEff), existing=String(head.getFormula()||'');
    if (!/ARRAYFORMULA/i.test(existing)){ head.setFormula(formula); return true; }
    var hasRoll=cRoll && existing.indexOf(L(cRoll))!==-1, hasFin=cFin && existing.indexOf(L(cFin))!==-1;
    if (!(hasRoll||hasFin)){ head.setFormula(formula); return true; }
    return false;
  }catch(e){ logError('ensureRollingEffectiveHeader_', e); return false; }
}

function updatePointsArrayHeaderIfNeeded_(sheet){
  try{
    var map=headerIndexMap_(sheet), pCol=map[CONFIG.COLS.Points]; if(!pCol) return false;
    var head=sheet.getRange(1,pCol), f=head.getFormula(); if(!f||!/ARRAYFORMULA/i.test(f)) return false;
    if (/Positive Point Removal/i.test(f)) return false;
    var nf=buildPointsArrayFormula_(sheet); if(!nf) return false;
    head.setFormula(nf); logInfo_('updatePointsArrayHeaderIfNeeded_','Upgraded Points ARRAYFORMULA.'); return true;
  }catch(e){ logError('updatePointsArrayHeaderIfNeeded_', e); return false; }
}

// ============================================================================
// Append Events row (idempotent slotting)
// ============================================================================
function appendEventsRow_(rowArrayOrObj, opts){
  opts=opts||{};
  return withBackoff_('appendEventsRow', function(){
    var s=sh_(CONFIG.TABS.EVENTS), hdrs=headers_(s), cols=Math.max(1,hdrs.length), rowArr=new Array(cols);
    if (rowArrayOrObj && Object.prototype.toString.call(rowArrayOrObj)==='[object Object]'){
      for (var i=0;i<hdrs.length;i++){ var h=hdrs[i]; rowArr[i]=rowArrayOrObj.hasOwnProperty(h)?rowArrayOrObj[h]:''; }
    } else {
      var src=Array.isArray(rowArrayOrObj)?rowArrayOrObj:[]; for (var j=0;j<cols;j++){ rowArr[j]=(j<src.length)?src[j]:''; }
    }
    var tIdx=hdrs.indexOf(CONFIG.COLS.Timestamp), last=s.getLastRow(), target=null;
    if (last>=2 && tIdx!==-1){
      var ts=s.getRange(2,tIdx+1,Math.max(1,last-1),1).getValues();
      for (var r=0;r<ts.length;r++){ if (ts[r][0]===''||ts[r][0]===null){ target=2+r; break; } }
    }
    if (!target && opts.afterRow && Number(opts.afterRow)>=1){ s.insertRowAfter(opts.afterRow); target=opts.afterRow+1; }
    if (!target){ target=Math.max(2,(last||1)+1); s.insertRowAfter(Math.max(1,last||1)); }
    if (tIdx!==-1 && (!rowArr[tIdx]||String(rowArr[tIdx]).trim()==='')) rowArr[tIdx]=new Date();
    s.getRange(target,1,1,cols).setValues([rowArr]);
    try{ logAudit('appendEventsRow','append', target, {});}catch(_){}
    return target;
  },4,250);
}

// ============================================================================
// Rolling/effective points derivation & query
// ============================================================================
function lookupPositiveTierFromAction_(actionName){
  try {
    if (!actionName) return '';
    var tab = CONFIG && CONFIG.TABS && CONFIG.TABS.POSITIVE_POINTS
      ? CONFIG.TABS.POSITIVE_POINTS
      : 'PositivePoints';
    var s = sh_(tab), hdrs=headers_(s);
    function col(names){ for (var i=0;i<names.length;i++){ var j=hdrs.indexOf(names[i]); if (j!==-1) return j+1; } return 0; }
    var cAction=col(['Positive Action','PositiveAction','Action']);
    var cType  =col(['Type','Credit Type','Tier']);
    if (!cAction || !cType) return '';
    var last=s.getLastRow(); if (last<2) return '';
    var vals=s.getRange(2,1,last-1,Math.max(cAction,cType)).getValues();
    var key=String(actionName||'').trim();
    for (var r=0;r<vals.length;r++){
      var got=String(vals[r][cAction-1]||'').trim();
      if (got===key) return String(vals[r][cType-1]||'').trim();
    }
    return '';
  }catch(e){
    logError && logError('lookupPositiveTierFromAction_', e, {action:actionName});
    return '';
  }
}


function _lookupPoints_(tabName, keyColName, keyValue, pointsColName){
  try{
    if (!tabName || !keyValue) return 0;
    var sh=sh_(tabName), map=headerIndexMap_(sh), cKey=map[keyColName], cPts=map[pointsColName]||map['Points'];
    if(!cKey||!cPts) return 0;
    var last=sh.getLastRow(); if(last<2) return 0;
    var vals=sh.getRange(2,1,last-1,Math.max(cKey,cPts)).getValues();
    var want=String(keyValue||'').replace(/\s+/g,' ').trim();
    for (var i=0;i<vals.length;i++){ var got=String(vals[i][cKey-1]||'').replace(/\s+/g,' ').trim(); if(got===want){ var p=vals[i][cPts-1]; return (p!==''&&p!=null&&!isNaN(p))?Number(p):0; } }
    return 0;
  }catch(e){ logError('_lookupPoints_', e, {tab:tabName, key:keyValue}); return 0; }
}

function _derivePointsDeltaForRow_(s, map, row){
  try{
    var ncol=map[CONFIG.COLS.Nullify]||map['Nullify']||0;
    if (ncol){ var n=s.getRange(row,ncol).getValue(); if (n===true || String(n).toUpperCase()==='TRUE') return 0; }
    var evt = map[CONFIG.COLS.EventType] ? String(s.getRange(row, map[CONFIG.COLS.EventType]).getDisplayValue()||'') : '';
    var inf = map[CONFIG.COLS.Infraction] ? String(s.getRange(row, map[CONFIG.COLS.Infraction]).getDisplayValue()||'') : '';
    var pa  = map[CONFIG.COLS.PositiveActionType] ? String(s.getRange(row, map[CONFIG.COLS.PositiveActionType]).getDisplayValue()||'') : '';
    if (/positive point removal/i.test(evt)) { var ptsPos=_lookupPoints_((CONFIG.TABS.POSITIVE||'PositivePoints'),'PositiveActionType',pa,'Points'); return ptsPos?-Number(ptsPos):0; }
    var ptsRubric=_lookupPoints_(CONFIG.TABS.RUBRIC,(CONFIG.COLS.Infraction||'Infraction'),inf,'Points'); return Number(ptsRubric||0);
  }catch(e){ logError('_derivePointsDeltaForRow_', e, {row:row}); return 0; }
}

function getRollingPointsForEmployee_(employee, opts){
  try{
    opts=opts||{}; if(!employee) return 0;
    var s=sh_(CONFIG.TABS.EVENTS), map=headerIndexMap_(s), last=s.getLastRow(); if(last<2) return 0;
    var cEmp=map[CONFIG.COLS.Employee], cEff=map[CONFIG.COLS.PointsRollingEffective]||0, cRoll=map[CONFIG.COLS.PointsRolling]||0;
    if(!cEmp||(!cEff&&!cRoll)) return 0;
    var endRow=Math.min(last, Math.max(1, Number(opts.beforeRow||0)-1) || last); if(endRow<2) return 0;
    var vals=s.getRange(2,1,endRow-1,Math.max(cEmp,cEff,cRoll)).getValues();
    for (var i=vals.length-1;i>=0;i--){
      var r=vals[i]; if (String(r[cEmp-1]||'')!==String(employee)) continue;
      var eff = cEff ? Number(r[cEff-1]||0) : NaN;
      var roll= cRoll? Number(r[cRoll-1]||0): NaN;
      var v=!isNaN(eff)?eff:(!isNaN(roll)?roll:0); return Math.max(0, Number(v||0));
    }
    return 0;
  }catch(e){ logError('getRollingPointsForEmployee_', e, {employee:employee, opts:opts}); return 0; }
}

// After-submit hook (calculate effective, run milestone/probation checks)
function afterSubmitPoints_(row){
  try{
    var s=sh_(CONFIG.TABS.EVENTS), map=headerIndexMap_(s);
    if (!row || row<2) return;
    var ctx=rowCtx_(s,row), employee=ctx.get(CONFIG.COLS.Employee)||''; if(!employee) return;
    SpreadsheetApp.flush(); Utilities.sleep(200);
    var effCol=map[CONFIG.COLS.PointsRollingEffective]||0, rollCol=map[CONFIG.COLS.PointsRolling]||0, pointsCol=map[CONFIG.COLS.Points]||0;
    var effective=NaN;
    if (effCol){ var v=s.getRange(row,effCol).getValue(); if (v!==''&&v!=null&&!isNaN(v)) effective=Number(v); }
    if (!isFinite(effective)&&rollCol){ var v2=s.getRange(row,rollCol).getValue(); if (v2!==''&&v2!=null&&!isNaN(v2)) effective=Number(v2); }
    if (!isFinite(effective)){
      var prior=getRollingPointsForEmployee_(employee,{beforeRow:row})||0;
      var delta=0; if (pointsCol){ var p=s.getRange(row,pointsCol).getValue(); delta=(p!==''&&p!=null&&!isNaN(p))?Number(p):_derivePointsDeltaForRow_(s,map,row); }
      effective=Math.max(0, prior+delta);
    }
    logAudit('afterSubmitPoints_compute', row, {employee:employee, effective:effective});
    try{ if (typeof scanAndHandleMilestones_==='function') scanAndHandleMilestones_(row); }catch(e1){ logError('afterSubmit_scanMilestones', e1, {row:row}); }
    try{ if (typeof checkProbationFailureForEmployee_==='function') checkProbationFailureForEmployee_(row); }catch(e2){ logError('afterSubmit_checkProb', e2, {row:row}); }
  }catch(err){ logError('afterSubmitPoints_top', err, {row:row}); }
}

// ============================================================================
// Probation & milestones
// ============================================================================
function tierForPoints_(pts){
  var p=Number(pts||0);
  var L1=Number((CONFIG.POLICY&&CONFIG.POLICY.MILESTONES&&CONFIG.POLICY.MILESTONES.LEVEL_1)||5);
  var L2=Number((CONFIG.POLICY&&CONFIG.POLICY.MILESTONES&&CONFIG.POLICY.MILESTONES.LEVEL_2)||10);
  var L3=Number((CONFIG.POLICY&&CONFIG.POLICY.MILESTONES&&CONFIG.POLICY.MILESTONES.LEVEL_3)||15);
  if (L3 && p>=L3) return 3; if (L2 && p>=L2) return 2; if (L1 && p>=L1) return 1; return 0;
}

function milestoneTextToTier_(text){
  if (!text) return null; var t=String(text).trim(), lower=t.toLowerCase();
  try{
    var names=(CONFIG&&CONFIG.POLICY&&CONFIG.POLICY.MILESTONE_NAMES)||{};
    if (names.LEVEL_1 && t.toLowerCase()===String(names.LEVEL_1).toLowerCase()) return 1;
    if (names.LEVEL_2 && t.toLowerCase()===String(names.LEVEL_2).toLowerCase()) return 2;
    if (names.LEVEL_3 && t.toLowerCase()===String(names.LEVEL_3).toLowerCase()) return 3;
  }catch(_){}
  if (/\b(terminate|termination|15pt|final termination)\b/i.test(lower)) return 3;
  if (/\b(10pt|1-?week|final warning|week suspension|30-?day|30 day)\b/i.test(lower)) return 2;
  if (/\b(5pt|2-?day|two day|2 day|2-?day suspension|suspension)\b/i.test(lower)) return 1;
  if (/\bprobation\b/i.test(lower)) return null;
  return null;
}

function getMilestonePatternHints_(){ return [{tier:3,examples:['termination','15pt']},{tier:2,examples:['1-Week','10pt','Final Warning','probation start']},{tier:1,examples:['2-Day','5pt']}]; }

function isMilestoneRowActive_(ctx){
  try{
    if (!ctx||typeof ctx.get!=='function') return false;
    var nullify=ctx.get((CONFIG.COLS.Nullify)||'Nullify')||'';
    var nullifiedBy=ctx.get((CONFIG.COLS.NullifiedBy)||'Nullified By')||'';
    if (String(nullify).trim()!=='') return false;
    if (String(nullifiedBy).trim()!=='') return false;
    var activeVal=ctx.get((CONFIG.COLS.Active)||'Active');
    var aStr=(typeof activeVal==='string')?activeVal.trim().toLowerCase():activeVal;
    if (activeVal===true||activeVal===1||aStr==='true'||aStr==='yes'||aStr==='y'||aStr==='1') return true;
    var ps=String(ctx.get((CONFIG.COLS.PendingStatus)||'Pending Status')||'').toLowerCase();
    if (ps.indexOf('null')!==-1||ps.indexOf('expired')!==-1||ps.indexOf('inactive')!==-1) return false;
    if (ps.indexOf('pending')!==-1||ps.indexOf('completed')!==-1||ps==='') return true;
    return true;
  }catch(e){ logError('isMilestoneRowActive_', e); return false; }
}

function hasActiveMilestoneOfTier_(employee, tier, opts){
  try{
    opts=opts||{}; var s=sh_(CONFIG.TABS.EVENTS), map=headerIndexMap_(s);
    var last=Math.min(s.getLastRow()||0, Number(opts.beforeRow||Infinity)-1); if(!employee||last<2) return false;
    var cEmp=map[CONFIG.COLS.Employee], cMil=map[CONFIG.COLS.Milestone]||0, cEvt=map[CONFIG.COLS.EventType]||0;
    var vals=s.getRange(2,1,last-1,Math.max(cEmp,cMil,cEvt,1)).getValues();
    for (var i=vals.length-1;i>=0;i--){
      var physical=i+2, row=vals[i];
      if (String(row[cEmp-1]||'').trim()!==String(employee).trim()) continue;
      var evtType=cEvt?String(row[cEvt-1]||'').toLowerCase().trim():'';
      var milText=cMil?String(row[cMil-1]||'').trim():'';
      if (!(evtType==='milestone'||(milText&&milText.length))) continue;
      var ctx=rowCtx_(s, physical), active=isMilestoneRowActive_(ctx);
      var theTier=milestoneTextToTier_(milText)||null;
      if (!theTier){ var eff=Number(ctx.get(CONFIG.COLS.PointsRollingEffective)||ctx.get(CONFIG.COLS.PointsRolling)||0); theTier=tierForPoints_(eff); }
      if (theTier===tier && active) return true;
    }
    return false;
  }catch(e){ logError('hasActiveMilestoneOfTier_', e, {employee:employee, tier:tier, opts:opts}); return false; }
}

function findLatestMilestoneForEmployee_(employee, opts){
  try{
    opts=opts||{}; var beforeRow=Number(opts.beforeRow||Infinity); if(!employee) return null;
    var s=sh_(CONFIG.TABS.EVENTS), map=headerIndexMap_(s), last=s.getLastRow(); if(last<2) return null;
    var cEmp=map[CONFIG.COLS.Employee], cMil=map[CONFIG.COLS.Milestone]||0, cWhen=map[CONFIG.COLS.MilestoneDate]||map[CONFIG.COLS.Timestamp]||0;
    var cEff=map[CONFIG.COLS.PointsRollingEffective]||0, cRoll=map[CONFIG.COLS.PointsRolling]||0, cFin=map[CONFIG.COLS.FinalPoints]||0, cCdwn=map[CONFIG.COLS.Cooldown_Window]||0;
    if(!cEmp||!cMil) return null;
    var n=Math.max(0, Math.min(last, beforeRow-1)-1); if (n<=0) return null;
    var vals=s.getRange(2,1,n,Math.max(cEmp,cMil,cWhen,cEff,cRoll,cFin,cCdwn,1)).getValues();
    for (var i=vals.length-1;i>=0;i--){
      var physical=i+2; if (physical>=beforeRow) continue;
      var row=vals[i]; if (String(row[cEmp-1]||'').trim()!==String(employee).trim()) continue;
      var label=String(row[cMil-1]||'').trim(); if (!label) continue;
      var eff=cEff?Number(row[cEff-1]):NaN, roll=cRoll?Number(row[cRoll-1]):NaN, fin=cFin?Number(row[cFin-1]):NaN;
      var points=!isNaN(eff)?eff:(!isNaN(roll)?roll:(!isNaN(fin)?fin:0));
      var when=(cWhen && row[cWhen-1]!=='' && row[cWhen-1]!=null)?toDate_(row[cWhen-1]):null;
      var cdwn=(cCdwn && row[cCdwn-1]!=='' && row[cCdwn-1]!=null)?toDate_(row[cCdwn-1]):null;
      var ctx=rowCtx_(s, physical);
      var mapped=milestoneTextToTier_(label), ptsTier=tierForPoints_(Number(points||0));
      return { row:physical, when:when||null, pointsRolling:Number(points||0), cooldownUntil:cdwn||null, milestone:label, ctx:ctx, active:isMilestoneRowActive_(ctx), mappedTier:(mapped!=null)?Number(mapped):null, pointsTier:(ptsTier!=null)?Number(ptsTier):null, tier:(mapped!=null)?Number(mapped):null, labelAmbiguous: (mapped==null && !!label) };
    }
    return null;
  }catch(e){ logError('findLatestMilestoneForEmployee_', e, {employee:employee, opts:opts}); return null; }
}

// Probation window check (L2 start → suspension → probation)
function isProbationActiveForEmployee_(employee, now){
  try{
    now=now||new Date(); var s=sh_(CONFIG.TABS.EVENTS), data=s.getDataRange().getValues(); if (data.length<2) return false;
    var hdr=data[0], cEmp=hdr.indexOf(CONFIG.COLS.Employee), cMil=hdr.indexOf(CONFIG.COLS.Milestone), cWhen=hdr.indexOf(CONFIG.COLS.MilestoneDate); if(cWhen<0) cWhen=hdr.indexOf(CONFIG.COLS.Timestamp);
    var cProbActive=hdr.indexOf(CONFIG.COLS.Probation_Active); if (cProbActive<0) cProbActive=hdr.indexOf('Probation_Active'); if (cProbActive<0) cProbActive=hdr.indexOf('Probation Active');
    var cEff=hdr.indexOf(CONFIG.COLS.PointsRollingEffective); var cRoll=(cEff>=0)?cEff:hdr.indexOf(CONFIG.COLS.PointsRolling);
    if (cEmp<0||cMil<0) return false;
    var best=null;
    for (var r=1;r<data.length;r++){
      var row=data[r]; if (String(row[cEmp]||'').trim()!==String(employee).trim()) continue;
      var label=String(row[cMil]||'').toLowerCase().trim(); var when=(cWhen>=0 && row[cWhen])?toDate_(row[cWhen]):null; var roll=(cRoll>=0)?Number(row[cRoll]||NaN):NaN;
      var looksL2=/\b(10|10pt|1-?week|final warning|probation|30-?day)\b/i.test(label) || (!isNaN(roll) && roll>=10 && roll<15);
      if (!looksL2||!when) continue;
      var rowProb=false; if (cProbActive>=0){ var pv=row[cProbActive]; var ps=(typeof pv==='string')?pv.trim().toLowerCase():pv; rowProb=(pv===true||pv===1||ps==='true'||ps==='yes'||ps==='1'); }
      if (!(label.indexOf('probation')!==-1 && rowProb)) continue;
      if (!best || when>best.when) best={when:when};
    }
    if (!best) return false;
    var suspensionDays=_policyNumber_('SUSPENSION_L2_DAYS',7), probationDays=_policyNumber_('PROBATION_DAYS',30);
    var returnDate=addDaysLocal_(best.when, suspensionDays), probationEnds=addDaysLocal_(returnDate, probationDays);
    return toDate_(now) <= probationEnds;
  }catch(e){ logError('isProbationActiveForEmployee_', e, {employee:employee}); return false; }
}
var isOnProbationForEmployee_ = isProbationActiveForEmployee_;

// ============================================================================
// On submit orchestration (normalized ingest → append → hooks)
// ============================================================================
function parsePolicyInfraction_(raw){
  var out={policy:'', infraction:'', unused:''}; if(raw==null) return out;
  var s=String(raw).trim(); if(!s) return out; s=s.replace(/\u2014|\u2013/g,'-');
  var rest=s, m;
  m=rest.match(/^\s*(\[[^\]]+\])\s*(.+)$/); if(m){ out.policy=m[1].trim(); rest=m[2].trim(); }
  if(!out.policy){ m=rest.match(/^\s*(\([^\)]+\))\s*(.+)$/); if(m){ out.policy=m[1].trim(); rest=m[2].trim(); } }
  if(!out.policy){ m=rest.match(/^\s*(\<[^\>]+\>)\s*(.+)$/); if(m){ out.policy=m[1].trim(); rest=m[2].trim(); } }
  if(!out.policy){ m=rest.match(/^\s*([A-Za-z0-9&()\/\-\s]{1,80}?)\s*[:\-–—;]\s*(.+)$/); if(m&&m[2]){ var cand=m[1].trim(); if(/[A-Za-z]/.test(cand) && !/\b\d+\s*pts?$/i.test(cand)){ out.policy=cand; rest=m[2].trim(); } } }
  var pts=rest.match(/^(.*?)(\s*[-–—]\s*\d+(?:\.\d+)?\s*pts?\.?)\s*$/i);
  if(pts){ out.infraction=pts[1].trim(); out.unused=pts[2].trim(); } else { out.infraction=rest.trim(); }
  return out;
}

function _onFormSubmit_impl(e){
  return withLocks_('onFormSubmit', function(){
    var events=sh_(CONFIG.TABS.EVENTS), hdrs=headers_(events), map=headerIndexMap_(events);
    var out=new Array(hdrs.length); for (var i=0;i<out.length;i++) out[i]='';

    if (e && e.namedValues){
      for (var q in e.namedValues){ if(!e.namedValues.hasOwnProperty(q)) continue;
        var target=(CONFIG.FORM_TO_EVENTS && CONFIG.FORM_TO_EVENTS[q])||q;
        var col=map[target]; if(col){ out[col-1]=Array.isArray(e.namedValues[q])?e.namedValues[q][0]:e.namedValues[q]; }
      }
    } else if (e && Array.isArray(e.values)) {
      for (var k=0;k<Math.min(e.values.length,out.length);k++) out[k]=e.values[k];
    }

    // Split Policy from Infraction
    var infCol=map[CONFIG.COLS.Infraction]||0, polCol=map[CONFIG.COLS.RelevantPolicy]||0;
    if (infCol){
      var rawInf=String(out[infCol-1]||''); var parsed=parsePolicyInfraction_(rawInf);
      out[infCol-1]=parsed.infraction||'';
      if (polCol && (!out[polCol-1] || String(out[polCol-1]).trim()==='')) out[polCol-1]=(parsed.policy||'').replace(/^[\[\(<]+|[\]\)>]+$/g,'').trim();
    }

    // Ensure Points/Effective headers are good
    try{
      var pCol=map[CONFIG.COLS.Points], headF=(pCol?String(events.getRange(1,pCol).getFormula()||''):'');
      if (pCol && /ARRAYFORMULA/i.test(headF)) updatePointsArrayHeaderIfNeeded_(events);
      ensureRollingEffectiveHeader_(events);
    }catch(upErr){ logError('onSubmit_updatePointsHeaders', upErr); }
    SpreadsheetApp.flush(); Utilities.sleep(200);

    // Append
    var targetRow = appendEventsRow_(out);
    SpreadsheetApp.flush();

    // Ledger: PositivePoints
    try {
      // Resolve the row we just wrote to Events
      var submitRow =
        (typeof appendedRowIndex !== 'undefined' && Number(appendedRowIndex)) ||
        (typeof targetRow !== 'undefined' && Number(targetRow)) ||
        (e && e.range && e.range.getRow && Number(e.range.getRow())) ||
        (function(){ try{
            var eventsTab = (CONFIG && CONFIG.TABS && CONFIG.TABS.EVENTS) ? CONFIG.TABS.EVENTS : 'Events';
            var evSh = SpreadsheetApp.getActive().getSheetByName(eventsTab);
            return evSh && evSh.getLastRow ? Number(evSh.getLastRow()) : 0;
        }catch(_){ return 0; } })() ||
        0;

      var did = false;
      if (submitRow >= 2) {
        did = recordPositiveCreditFromEvent(submitRow);
      } else {
        logError && logError('positive_ledger_from_submit_no_row', null, { resolvedRow: submitRow });
      }

      logInfo_ && logInfo_('positive_ledger_after_submit', { row: submitRow, did: did });
    } catch (e2) {
      logError && logError('positive_ledger_from_submit', e2, { row: submitRow || targetRow });
    }


    // Probation guard for Positive Point Removal (if header not owning points)
    try{
      var headerHasArray=(pCol && /ARRAYFORMULA/i.test(String(events.getRange(1,pCol).getFormula()||'')));
      var ctx=rowCtx_(events, targetRow);
      var evtVal=String(ctx.get(CONFIG.COLS.EventType)||'');
      var isRemoval=/^Positive Point Removal$/i.test(evtVal);
      var emp=String(ctx.get(CONFIG.COLS.Employee)||'');
      if (isRemoval && emp && typeof isOnProbationForEmployee_==='function' && isOnProbationForEmployee_(emp)){
        if (!headerHasArray && pCol){ ctx.set(CONFIG.COLS.Points, 0); logInfo_('onSubmit_probationBlock','Blocked positive points for '+emp+' row='+targetRow); }
        var nCol=map[CONFIG.COLS.NotesReviewer]; if(nCol){
          var prev=String(events.getRange(targetRow,nCol).getValue()||'');
          events.getRange(targetRow,nCol).setValue((prev?prev+' ; ':'')+'Positive points blocked due to active probation.');
        }
      }
    }catch(pbErr){ logError('onSubmit_probationBlock_err', pbErr, {row:targetRow}); }

    // Create Event PDF (idempotent) - BUT SKIP Performance Issues
    try{
      var eventType = String(ctx.get(CONFIG.COLS.EventType) || '').trim();
      var isPerformanceIssue = eventType.toLowerCase() === 'performance issue';
      
      if (!isPerformanceIssue) {
        var pdfHdr=CONFIG.COLS.PdfLink||'Write-Up PDF', pdfIdx=map[pdfHdr]||0,
            existing=(pdfIdx?String(events.getRange(targetRow,pdfIdx).getValue()||''):'');
        if (!existing){
          var pdfId=null;
          try{ if (typeof createEventRecordPdf_==='function') pdfId=createEventRecordPdf_(targetRow); }catch(pErr){ logError('onSubmit_createPdf', pErr, {row:targetRow}); }
          logInfo_('onSubmit_pdf', 'created '+(pdfId||'null')+' row='+targetRow);
        } else {
          logInfo_('onSubmit_pdf','skip existing row='+targetRow);
        }
      } else {
        logInfo_('onSubmit_pdf','skip performance issue - will create specific PDF below');
      }
    }catch(e3){ logError('onSubmit_pdf_top', e3, {row:targetRow}); }

    // Performance Issue handling (CORRECTED)
    try {
      var eventType = String(ctx.get(CONFIG.COLS.EventType) || '').trim();
      if (eventType.toLowerCase() === 'performance issue') {
        // Performance Issues generate 0 points
        if (map[CONFIG.COLS.Points]) {
          events.getRange(targetRow, map[CONFIG.COLS.Points]).setValue(0);
        }
        
        // Performance Issues should use PERF_ISSUE template, not EVENT_RECORD
        // We need to create a separate PDF using the PERF_ISSUE template
        try {
          var perfPdfId = createConsequencePdf_(targetRow, 'Performance Issue');
          if (perfPdfId) {
            // Update the PDF link to point to the PERF_ISSUE PDF instead of EVENT_RECORD
            var pdfHdr = CONFIG.COLS.PdfLink || 'Write-Up PDF';
            var pdfIdx = map[pdfHdr] || 0;
            if (pdfIdx) {
              var perfPdfUrl = 'https://drive.google.com/file/d/' + perfPdfId + '/view';
              setRichLinkSafe(events, targetRow, pdfHdr, 'View PDF', perfPdfUrl);
            }
          }
          logInfo_('onSubmit_perf_pdf', 'created performance issue PDF: ' + (perfPdfId || 'null') + ' row=' + targetRow);
        } catch (perfPdfErr) {
          logError && logError('onSubmit_perf_pdf_err', perfPdfErr, { row: targetRow });
        }
        
        // Update Performance Issue count and check for Growth Plan trigger
        try {
          if (typeof handlePerformanceIssueSubmit === 'function') {
            handlePerformanceIssueSubmit(targetRow);
          }
        } catch (perfErr) {
          logError && logError('onSubmit_perf_handler_err', perfErr, { row: targetRow });
        }
      }
    } catch (perfTopErr) {
      logError && logError('onSubmit_perf_top_err', perfTopErr, { row: targetRow });
    }

    // 🔔 Slack: notify docs channel once PDF exists (disciplinary/milestone only)
    try{
      var evtForNotify = String(rowCtx_(events, targetRow).get(CONFIG.COLS.EventType)||'').trim().toLowerCase();
      if (!/^positive/.test(evtForNotify)) {
        SpreadsheetApp.flush(); Utilities.sleep(200); // ensure link is written
        var pdfIdx2 = map[CONFIG.COLS.PdfLink||'Write-Up PDF']||0;
        // Be strict: only treat as "has PDF" if the cell actually contains a URL
        var hasPdf = false;
        if (pdfIdx2) {
          var cell = events.getRange(targetRow, pdfIdx2);
          var rtv = cell.getRichTextValue && cell.getRichTextValue();
          var link = (rtv && rtv.getLinkUrl && rtv.getLinkUrl()) || '';
          if (!link) {
            var disp = cell.getDisplayValue ? cell.getDisplayValue() : String(cell.getValue()||'');
            var m = String(disp||'').match(/https:\/\/drive\.google\.com\/[^\s)]+/i);
            link = m ? m[0] : '';
          }
          hasPdf = !!link;
        }
        if (hasPdf && typeof notifyDocs_ === 'function'){
          notifyDocs_(targetRow);
        }
      }
    }catch(notifyErr){ logError('notifyDocs_after_pdf', notifyErr, {row:targetRow}); }

    // 🔔 Slack: also notify for Positive Credits
    try {
      var evtVal = String(rowCtx_(events, targetRow).get(CONFIG.COLS.EventType)||'');
      if (/^positive/i.test(evtVal) && typeof notifyPositiveCredit_==='function'){
        notifyPositiveCredit_(targetRow);
      }
    }catch(nPos){ logError && logError('notifyPositive_after_submit',nPos,{row:targetRow}); }

    // Post hooks
    try{ if (typeof afterSubmitPoints_==='function') afterSubmitPoints_(targetRow); }catch(e4){ logError('afterSubmitPoints_', e4, {row:targetRow}); }

    return targetRow;
  }, 30000);
}


// Smooth shim for external caller
function ensureRollingEffectiveHeader(sheetOrName){ try{ return ensureRollingEffectiveHeader_(sheetOrName); }catch(e){ logError('ensureRollingEffectiveHeader_shim', e); return false; } }



// ============================================================================
// On-edit: create milestone PDF when director written
// ============================================================================
// main.js (or wherever it lives)
function onEdit_CreatePdfWhenDirectorSet(e){
  try{
    if (!e || !e.range || !e.range.getSheet) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== CONFIG.TABS.EVENTS) return;
    var row = e.range.getRow(); if (row < 2) return;

    var map    = headerIndexMap_(sh);
    var cdHdr  = CONFIG.COLS.ConsequenceDirector || 'Consequence Director';
    var cdCol  = map[cdHdr] || 0;
    if (!cdCol || e.range.getColumn() !== cdCol) return;

    var evtHdr  = CONFIG.COLS.EventType || 'EventType';
    var pendHdr = CONFIG.COLS.PendingStatus || 'Pending Status';
    var pdfHdr  = CONFIG.COLS.PdfLink || 'Write-Up PDF';

    var evt      = map[evtHdr]  ? String(sh.getRange(row, map[evtHdr]).getDisplayValue()||'').trim()  : '';
    var pending  = map[pendHdr] ? String(sh.getRange(row, map[pendHdr]).getDisplayValue()||'').trim() : '';
    var cd       = cdCol        ? String(sh.getRange(row, cdCol).getDisplayValue()||'').trim()        : '';
    var existing = map[pdfHdr]  ? String(sh.getRange(row, map[pdfHdr]).getDisplayValue()||'').trim()  : '';

    var isMilestone = evt.toLowerCase() === 'milestone';
    var isPending   = pending.toLowerCase() === 'pending';

    logInfo_('claimed_gate_check', {row, evt, pending, cd, hasPdf: !!existing, isMilestone, isPending});
    if (!(isMilestone && isPending && cd) || existing) return;

    var id = maybeCreatePdfForMilestoneRow_(sh, row);
    logInfo_('claimed_after_pdf', {row, pdfId: id});

  }catch(err){
    logError('onEdit_CreatePdfWhenDirectorSet', err);
  }
}

// utils.js
function positiveTierLabelForAction_(actionName){
  try {
    if (!actionName) return 'Positive Credit';
    var s = sh_(CONFIG.TABS.POSITIVE_MAP || 'PositiveMap');
    var hdrs = headers_(s);
    var iAct = hdrs.indexOf('Positive Action');
    var iTyp = hdrs.indexOf('Type');
    if (iAct === -1 || iTyp === -1) return 'Positive Credit';

    var last = s.getLastRow();
    if (last < 2) return 'Positive Credit';

    var vals = s.getRange(2, 1, last-1, Math.max(iAct+1, iTyp+1)).getValues();
    var key = String(actionName).trim();
    for (var r=0; r<vals.length; r++){
      if (String(vals[r][iAct]).trim() === key){
        var t = String(vals[r][iTyp] || '').trim().toLowerCase();
        if (t === 'minor')    return 'Minor Positive Credit';
        if (t === 'moderate') return 'Moderate Positive Credit';
        if (t === 'major')    return 'Major Positive Credit';
        if (t === 'grace')    return 'Grace';
        return 'Positive Credit'; // unknown types default safely
      }
    }
    return 'Positive Credit';
  } catch(e){
    logError && logError('positiveTierLabelForAction_', e, {action: actionName});
    return 'Positive Credit';
  }
}




// utils.js (or where it is)
function maybeCreatePdfForMilestoneRow_(sheetNameOrObj, row){
  return withBackoff_('maybeCreatePdfForMilestoneRow', function(){
    var sh = (typeof sheetNameOrObj==='string') ? sh_(sheetNameOrObj) : sheetNameOrObj;
    if (!sh) return null;

    var map       = headerIndexMap_(sh),
        evtHdr    = CONFIG.COLS.EventType || 'EventType',
        pendingHdr= CONFIG.COLS.PendingStatus || 'Pending Status',
        pdfHdr    = CONFIG.COLS.PdfLink || 'Write-Up PDF',
        cdHdr     = CONFIG.COLS.ConsequenceDirector || 'Consequence Director';

    var evt = map[evtHdr] ? String(sh.getRange(row, map[evtHdr]).getDisplayValue()||'').trim().toLowerCase() : '';
    var pending = map[pendingHdr] ? String(sh.getRange(row, map[pendingHdr]).getDisplayValue()||'').trim().toLowerCase() : '';
    var existing = map[pdfHdr] ? String(sh.getRange(row, map[pdfHdr]).getDisplayValue()||'').trim() : '';
    var cd = map[cdHdr] ? String(sh.getRange(row, map[cdHdr]).getDisplayValue()||'').trim() : '';

    logInfo_('claimed_gate_pdfFn', {row, evt, pending, cd, hasPdf: !!existing});
    if (evt!=='milestone' || pending!=='pending' || existing || !cd) return null;

    var pdfId = null;
    try{
      if (typeof inferMilestoneTemplate_ === 'function'){
        var inferred=null; try{ inferred=inferMilestoneTemplate_(row);}catch(_){}
        var chosen=(inferred&&inferred.templateId)||(CONFIG.TEMPLATES&&CONFIG.TEMPLATES.MILESTONE_5)||null;
        if (chosen && typeof createMilestonePdf_==='function') pdfId=createMilestonePdf_(row, chosen);
        else if (typeof createEventRecordPdf_==='function')     pdfId=createEventRecordPdf_(row);
      } else if (typeof createMilestonePdf_==='function'){ pdfId=createMilestonePdf_(row,null);
      } else if (typeof createEventRecordPdf_==='function'){ pdfId=createEventRecordPdf_(row); }
    }catch(e){ logError('claimed_pdf_mergeErr', e, {row}); }

    if (pdfId){
      try{
        var pdfUrl='https://drive.google.com/file/d/'+pdfId+'/view';
        setRichLinkSafe(sh,row,pdfHdr,'View PDF',pdfUrl);
        if (map[pendingHdr]) sh.getRange(row,map[pendingHdr]).setValue('Claimed');
        logAudit('system','milestone_pdf_created', row, {pdfId});

        logInfo_('claimed_notify_attempt', {row, pdfUrl});
        if (typeof notifyLeadersMilestone_==='function'){
          try { notifyLeadersMilestone_(row, pdfUrl); }
          catch(nerr){ logError('claimed_notify_error', nerr, {row}); }
        } else {
          logInfo_('claimed_notify_missing_fn', {row});
        }
      }catch(_){}
    } else {
      logInfo_('claimed_no_pdf', {row});
    }
    return pdfId;
  },4,250);
}




// ============================================================================
// Doc helpers (token replace + history table builder)
// ============================================================================
function _escRegex_(s){ return String(s).replace(/[-\/\\^$*+?.()|[\]{}]/g,'\\$&'); }

function replaceTokensInBody_(body, tokenMap){
  if(!body||!tokenMap) return;
  Object.keys(tokenMap).forEach(function(key){
    var val=tokenMap[key], pattern='\\{\\{'+_escRegex_(key)+'\\}\\}';
    try{ body.replaceText(pattern, val==null?'':String(val)); }catch(_){}
  });
}

function findHistoryTable_(body, expectedHeaders){
  try{
    var tables=body.getTables?body.getTables():[];
    for (var t=0;t<tables.length;t++){
      var tbl=tables[t]; if (!tbl.getNumRows||tbl.getNumRows()<1) continue;
      var hdr=tbl.getRow(0); if (hdr.getNumCells()<expectedHeaders.length) continue;
      var ok=true;
      for (var c=0;c<expectedHeaders.length;c++){
        var want=String(expectedHeaders[c]).trim().toLowerCase();
        var got=String(hdr.getCell(c).getText()).trim().toLowerCase();
        if (want==='event' && got==='event type') continue;
        if (want==='pts'   && got==='points') continue;
        if (want==='roll'  && got==='rolling') continue;
        if (want==='pdf'   && got==='write-up pdf') continue;
        if (got!==want){ ok=false; break; }
      }
      if (ok) return tbl;
    }
  }catch(_){}
  return null;
}

function fillHistoryTable_(table, tableData, writeMap, rows){
  try{
    if(!table||!Array.isArray(tableData)||tableData.length<2) return;
    var header=table.getRow(0), headerCols=header.getNumCells();
    var sample=(table.getNumRows()>=2)?table.getRow(1):(function(){var r=table.appendTableRow(); for(var c0=0;c0<headerCols;c0++) r.appendTableCell(''); return r;})();
    var tmpl=[]; for (var tc=0; tc<sample.getNumCells(); tc++){ var cell=sample.getCell(tc), attrs={}; try{attrs=cell.getAttributes()||{}; delete attrs[DocumentApp.Attribute.BOLD]; delete attrs[DocumentApp.Attribute.ITALIC]; }catch(_){}
      var p; try{ var paras=cell.getParagraphs(); p=(paras&&paras.length)?paras[0]:cell.appendParagraph(''); }catch(e){ p=cell.appendParagraph(''); }
      var t=p.editAsText(); if(!t) t=p.appendText(''); var added=false; if(t.getText().length===0){ t.insertText(0,'x'); added=true; }
      var textAttrs; try{ textAttrs=t.getAttributes(0)||{}; }catch(_){ textAttrs={}; } if(added){ try{t.deleteText(0,0);}catch(_){ } }
      tmpl.push({cellAttrs:attrs, textAttrs:textAttrs});
    }
    while (table.getNumRows()>1) table.removeRow(1);
    var ncols=tableData[0].length, BASE={}; BASE[DocumentApp.Attribute.FONT_FAMILY]='Arial'; BASE[DocumentApp.Attribute.FONT_SIZE]=6; BASE[DocumentApp.Attribute.BOLD]=false;

    function appendRow(arr){
      var newRow=table.appendTableRow();
      for (var cc=0; cc<ncols; cc++){
        var txt=(cc<arr.length && arr[cc]!=null)?String(arr[cc]):'';
        var cell=newRow.appendTableCell('');
        var t=tmpl[Math.min(cc, tmpl.length-1)]||{};
        try{ if(t.cellAttrs) cell.setAttributes(t.cellAttrs);}catch(_){}
        try{ var paras=cell.getParagraphs(); for (var k=paras.length-1;k>=0;k--) cell.removeParagraph(paras[k]); }catch(_){}
        var p=cell.appendParagraph(txt); try{ p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);}catch(_){}
        var te=p.editAsText(); if (te){ if (t.textAttrs && Object.keys(t.textAttrs).length) te.setAttributes(t.textAttrs); else te.setAttributes(BASE); }
      }
      return newRow;
    }
    for (var r=1;r<tableData.length;r++) appendRow(tableData[r]);
    try{ if (typeof wirePdfLinksIntoTable==='function') wirePdfLinksIntoTable(table, rows, writeMap); }catch(_){}
  }catch(e){ Logger.log('fillHistoryTable_ error: '+String(e)); }
}

function wirePdfLinksIntoTable(tbl, rowsArr, writeMap){
  if(!tbl||!rowsArr||!writeMap) return;
  var pdfCol=Math.max(0, tbl.getRow(0).getNumCells()-1);
  for (var i=0;i<rowsArr.length;i++){
    var drow=i+1; if (drow>=tbl.getNumRows()) break;
    var cell=tbl.getRow(drow).getCell(pdfCol); if (cell.getNumChildren()===0) cell.appendParagraph('');
    var p=cell.getChild(0).asParagraph(), t=p.editAsText(); if(!t) continue;
    var info=writeMap[rowsArr[i].sheetRowIndex]||{}, url=info.url||null;
    t.setText('View PDF'); if (url){ try{ t.setLinkUrl(0, t.getText().length-1, url);}catch(_){ } }
  }
}

/********************** Admin Backfill PDFs **************************
 * Adds an Admin menu and a resumable queue to (re)generate missing
 * Write-Up PDFs for Events rows.
 *
 * Uses existing helpers in your project:
 *   - sh_, headers_, headerIndexMap_, readLinkUrlFromCell_
 *   - createEventRecordPdf_, withBackoff_, setRichLinkSafe (inside create…)
 *   - CONFIG.TABS.EVENTS, CONFIG.TABS.LOGS, CONFIG.COLS.*
 *********************************************************************/

function addAdminBackfillMenu_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Backfill Menu')
    .addSubMenu(
      ui.createMenu('Backfill PDFs')
        .addItem('Queue missing PDFs — last 60 days', 'adminBackfillMissingPdfs_60d')
        .addItem('Queue missing PDFs — all rows', 'adminBackfillMissingPdfs_all')
        .addItem('Queue missing PDFs — from selection', 'adminBackfillMissingPdfs_selection')
        .addSeparator()
        .addItem('Resume queue now', 'processBackfillQueue_')
        .addItem('Show queue status', 'showBackfillStatus_')
        .addItem('Cancel queue', 'cancelBackfillQueue_')
    )
    .addToUi();
}

// Call this from your existing onOpen() or add this standalone onOpen:
// function onOpen() {
//   try { addAdminBackfillMenu_(); } catch(_) {}
//   // if you already have an onOpen elsewhere, just call addAdminBackfillMenu_() in it
// }

/** === Queue storage === **/
const _QKEY_ROWS   = 'pdf_q_rows_csv';   // comma-separated row numbers
const _QKEY_POS    = 'pdf_q_pos';        // zero-based index into rows list
const _QKEY_OPTS   = 'pdf_q_opts';       // JSON: {batchSize,lastDays}
const _QKEY_ACTIVE = 'pdf_q_active';     // '1' while queue should keep running

function _props_() { return PropertiesService.getScriptProperties(); }
function _set(key, val){ _props_().setProperty(key, String(val)); }
function _get(key, def){ var v=_props_().getProperty(key); return (v==null?def:v); }
function _clr(key){ _props_().deleteProperty(key); }

/** Admin entry points **/
function adminBackfillMissingPdfs_60d(){ backfillEnqueueMissingPdfs_({ lastDays: 60, batchSize: 10 }); }
function adminBackfillMissingPdfs_all(){ backfillEnqueueMissingPdfs_({ lastDays: 0,  batchSize: 10 }); }

function adminBackfillMissingPdfs_selection(){
  const s = sh_(CONFIG.TABS.EVENTS);
  const rng = s.getActiveRange();
  if (!rng) return SpreadsheetApp.getUi().alert('Select a block of rows first.');
  const start = rng.getRow(), end = start + rng.getNumRows() - 1;
  backfillEnqueueMissingPdfs_({ fromRow: start, toRow: end, batchSize: 10 });
}

/**
 * Scan Events for rows that look like write-ups but have no PDF link.
 * Enqueue rows and kick off processing.
 *
 * opts: { lastDays?:number, fromRow?:number, toRow?:number, batchSize?:number }
 */
function backfillEnqueueMissingPdfs_(opts) {
  opts = opts || {};
  const s = sh_(CONFIG.TABS.EVENTS);
  const hdrs = headers_(s);
  const map = headerIndexMap_(s);

  const cPdf  = map[CONFIG.COLS.PdfLink] || map['Write-Up PDF'] || 0;
  const cEvt  = map[CONFIG.COLS.EventType] || map['EventType'] || 0;
  const cDate = map[CONFIG.COLS.IncidentDate] || map['IncidentDate'] || map['Timestamp'] || 0;

  const lastRow = s.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No event rows to scan.'); return; }

  const r1 = Math.max(2, opts.fromRow || 2);
  const r2 = Math.min(lastRow, opts.toRow || lastRow);

  const tz = Session.getScriptTimeZone() || 'UTC';
  const cutoff = (opts.lastDays && opts.lastDays > 0)
    ? new Date(Date.now() - opts.lastDays * 24 * 60 * 60 * 1000)
    : null;

  // read display values for range once
  const width = Math.max(cPdf, cEvt, cDate, 1);
  const vals = s.getRange(r1, 1, (r2 - r1 + 1), width).getDisplayValues();

  function looksLikeWriteUp(evt) {
    const e = String(evt || '').toLowerCase();
    if (!e) return false;
    if (/positive/.test(e)) return false; // exclude positive credit rows
    return /write.?up/.test(e) || /disciplinary event/.test(e);
  }

  function hasPdfLink(rowIndex) {
    if (!cPdf) return false;
    const cell = s.getRange(rowIndex, cPdf);
    try {
      const u = readLinkUrlFromCell_(cell);
      if (u) return true;
    } catch(_){}
    const disp = cell.getDisplayValue();
    return !!(disp && /https?:\/\//i.test(disp));
  }

  const targetRows = [];
  for (var i = 0; i < vals.length; i++) {
    const row = r1 + i;
    const evt = cEvt ? vals[i][cEvt - 1] : '';
    if (!looksLikeWriteUp(evt)) continue;

    if (cutoff && cDate) {
      const raw = s.getRange(row, cDate).getValue();
      const d = raw instanceof Date ? raw : new Date(raw);
      if (d.toString() !== 'Invalid Date' && d < cutoff) continue;
    }

    if (!hasPdfLink(row)) {
      targetRows.push(row);
    }
  }

  if (!targetRows.length) {
    SpreadsheetApp.getUi().alert('No missing PDFs found for the chosen scope.');
    return;
  }

  // store queue
  _set(_QKEY_ROWS, targetRows.join(','));
  _set(_QKEY_POS, 0);
  _set(_QKEY_OPTS, JSON.stringify({ batchSize: Number(opts.batchSize || 10) }));
  _set(_QKEY_ACTIVE, '1');

  // log
  try { sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'INFO', 'backfill_queue_created', 'rows=' + targetRows.length]); } catch(_){}

  // start processing now
  processBackfillQueue_();
}

/** Show queue status to user */
function showBackfillStatus_(){
  const rowsCsv = _get(_QKEY_ROWS, '');
  const pos = Number(_get(_QKEY_POS, 0));
  const active = _get(_QKEY_ACTIVE, '') === '1';
  const total = rowsCsv ? rowsCsv.split(',').filter(Boolean).length : 0;
  SpreadsheetApp.getUi().alert(
    (active ? 'Queue ACTIVE' : 'Queue idle') + '\n' +
    'Processed: ' + Math.min(pos, total) + ' / ' + total
  );
}

/** Cancel queue and remove scheduled triggers */
function cancelBackfillQueue_(){
  _clr(_QKEY_ROWS); _clr(_QKEY_POS); _clr(_QKEY_OPTS); _clr(_QKEY_ACTIVE);
  _clearBackfillTriggers_();
  SpreadsheetApp.getUi().alert('Backfill queue cleared.');
}

/** Process up to batchSize rows and reschedule if more remain */
function processBackfillQueue_(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // let an in-flight run finish

  try {
    if (_get(_QKEY_ACTIVE, '') !== '1') { _clearBackfillTriggers_(); return; }

    const rowsCsv = _get(_QKEY_ROWS, '');
    if (!rowsCsv) { _clearBackfillTriggers_(); return; }

    const rows = rowsCsv.split(',').filter(Boolean).map(function(s){ return Number(s); });
    var pos = Number(_get(_QKEY_POS, 0)) || 0;
    const opts = JSON.parse(_get(_QKEY_OPTS, '{"batchSize":10}'));
    const batch = Math.max(1, Number(opts.batchSize || 10));

    const s = sh_(CONFIG.TABS.EVENTS);
    const start = pos;
    const end = Math.min(rows.length, pos + batch);

    for (var i = start; i < end; i++){
      var r = rows[i];
      var ok = false, pdfId = null, err = null;
      try {
        pdfId = withBackoff_('backfillRow', function(){ return createEventRecordPdf_(r); }, 3, 300);
        ok = !!pdfId;
      } catch(e){
        err = String(e);
      }
      try {
        sh_(CONFIG.TABS.LOGS).appendRow([new Date(), ok?'INFO':'ERROR',
          ok?'backfill_row_ok':'backfill_row_err', 'row=' + r, ok ? String(pdfId) : err]);
      } catch(_){}
      Utilities.sleep(150); // be polite to Drive
    }

    pos = end;
    _set(_QKEY_POS, pos);

    if (pos < rows.length && _get(_QKEY_ACTIVE,'')==='1') {
      _ensureBackfillTrigger_(); // schedule next chunk
    } else {
      _clearBackfillTriggers_();
      _clr(_QKEY_ROWS); _clr(_QKEY_POS); _clr(_QKEY_OPTS); _clr(_QKEY_ACTIVE);
      try { sh_(CONFIG.TABS.LOGS).appendRow([new Date(), 'INFO', 'backfill_done', 'rows=' + rows.length]); } catch(_){}
    }

  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

/** Schedule a time-based trigger if one isn't already present */
function _ensureBackfillTrigger_(){
  var exists = ScriptApp.getProjectTriggers().some(function(t){ return t.getHandlerFunction() === 'processBackfillQueue_'; });
  if (!exists) {
    ScriptApp.newTrigger('processBackfillQueue_').timeBased().after(30 * 1000).create(); // run again in ~30s
  }
}
function _clearBackfillTriggers_(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'processBackfillQueue_') ScriptApp.deleteTrigger(t);
  });
}

// ============================================================================
// Probation flags writer (script-owned mode)
// ============================================================================
CONFIG.SCRIPT_OWNS_PROBATION = (typeof CONFIG!=='undefined' && typeof CONFIG.SCRIPT_OWNS_PROBATION!=='undefined')
  ? !!CONFIG.SCRIPT_OWNS_PROBATION : false;
function scriptShouldWriteProbation_(){ return !!(typeof CONFIG!=='undefined' && CONFIG.SCRIPT_OWNS_PROBATION); }

function setProbationFlagsForLabel_(eventsSheetOrName, rowIndex, milestoneLabel){
  try{
    if (typeof scriptShouldWriteProbation_==='function' && !scriptShouldWriteProbation_()) return false;
    var sh=(typeof eventsSheetOrName==='string')?sh_(eventsSheetOrName):eventsSheetOrName; if(!sh) return false;
    var map=headerIndexMap_(sh);
    var cStart=map[CONFIG.COLS.Probation_Start]||map['Probation_Start']||0;
    var cEnd=map[CONFIG.COLS.Probation_End]||map['Probation_End']||0;
    var cActive=map[CONFIG.COLS.Probation_Active]||map['Probation_Active']||map['Probation Active']||0;
    var cMil=map[CONFIG.COLS.Milestone]||0, cMilDate=map[CONFIG.COLS.MilestoneDate]||0, cInc=map[CONFIG.COLS.IncidentDate]||0, cTs=map[CONFIG.COLS.Timestamp]||0;

    var label=(milestoneLabel!=null?String(milestoneLabel):(cMil?String(sh.getRange(rowIndex,cMil).getDisplayValue()||''):'')); label=label.trim(); var lower=label.toLowerCase();

    if (/probation\s*failure/i.test(lower)){ 
      // if (cActive){ sh.getRange(rowIndex,cActive).setValue(true); logAudit('system','probation_flags_pf_active',rowIndex,{label:label}); return true; }  // COMMENTED OUT - using formula now
      return true; 
    }

    var looksProb=/\bprobation\b/i.test(lower);
    try{ var n2=CONFIG&&CONFIG.POLICY&&CONFIG.POLICY.MILESTONE_NAMES&&CONFIG.POLICY.MILESTONE_NAMES.LEVEL_2; if (!looksProb && n2 && /probation/i.test(String(n2))) looksProb=(String(n2).trim().toLowerCase()===lower); }catch(_){}
    if (!looksProb) return false;

    var when=null; try{ if(cMilDate) when=sh.getRange(rowIndex,cMilDate).getValue()||when; if(!when && cInc) when=sh.getRange(rowIndex,cInc).getValue()||when; if(!when && cTs) when=sh.getRange(rowIndex,cTs).getValue()||when; }catch(_){}
    if (!when) when=new Date();
    var probationDays=_policyNumber_('PROBATION_DAYS',30);
    var start=when; // Start probation on the milestone date (no suspension delay)
    var end=addDaysLocal_(start, probationDays);
    var wrote=false;
    if (cStart){ sh.getRange(rowIndex,cStart).setValue(start); wrote=true; }
    if (cEnd){ sh.getRange(rowIndex,cEnd).setValue(end); wrote=true; }
    // if (cActive){ sh.getRange(rowIndex,cActive).setValue(true); wrote=true; }  // COMMENTED OUT - using formula now
    if (wrote) try{ logAudit('system','probation_flags_set',rowIndex,{label:label,start:start,end:end}); }catch(_){}
    return wrote;
  }catch(e){ logError('setProbationFlagsForLabel_', e, {row:rowIndex}); return false; }
}

function _policyNumber_(key, defaultValue) {
  try {
    if (typeof loadPolicyFromSheet_ === 'function') {
      loadPolicyFromSheet_();
    }
    var value = CONFIG.POLICY && CONFIG.POLICY[key];
    return Number(value) || Number(defaultValue) || 0;
  } catch(e) {
    return Number(defaultValue) || 0;
  }
}
