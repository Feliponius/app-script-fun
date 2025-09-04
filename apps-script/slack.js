// ============== Low-level sender (with backoff) ===================
function postSlack_(webhookUrl, payload){
  if (!webhookUrl){
    logError('postSlack_missing_url', new Error('No webhook URL'), {payload:payload});
    return false;
  }
  return withBackoff_('postSlack', function(){
    var res = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      muteHttpExceptions: true,
      payload: JSON.stringify(payload)
    });
    var code = res.getResponseCode();
    var body = res.getContentText();
    logInfo_('postSlack_http', {code:code, body:body.slice(0,500)});
    if (code >= 200 && code < 300) return true;
    throw new Error('Slack HTTP '+code+': '+body);
  }, 4, 250);
}


// ============== Helpers ===========================================
function _readCell_(sh, map, row, colName){
  var c = map[colName] || map[CONFIG.COLS[colName]] || 0;
  return c ? String(sh.getRange(row, c).getDisplayValue() || '') : '';
}
function _extractPdfUrlFromCell_(val){
  // Accept either a rich link or a raw URL
  var s = String(val||'');
  var m = s.match(/https:\/\/drive\.google\.com\/[^\s)]+/i);
  return m ? m[0] : (s.indexOf('http')===0 ? s : '');
}

// ============== Docs-channel notifier (write-up) ==================
// inside slack.js
function notifyDocs_(row){
  try{
    if (!CONFIG.ENABLE_SLACK_DOCS || !CONFIG.DOCS_WEBHOOK){
      logInfo_('notifyDocs_skip_switch_or_url', {enabled:CONFIG.ENABLE_SLACK_DOCS, hasUrl:!!CONFIG.DOCS_WEBHOOK});
      return false;
    }

    var sh  = sh_(CONFIG.TABS.EVENTS), map = headerIndexMap_(sh);

    // alias-tolerant helpers
    function col(){
      var names = Array.prototype.slice.call(arguments).flat().filter(Boolean);
      for (var i=0;i<names.length;i++){
        var key = names[i];
        var c = map[key] || (CONFIG.COLS && map[CONFIG.COLS[key]]) || 0;
        if (c) return c;
      }
      return 0;
    }
    function getCell(h){ var c = col(h); return c ? sh.getRange(row, c) : null; }
    function val(h){ var r = getCell(h); return r ? String(r.getDisplayValue() || '') : ''; }

    // Event type (normalize once)
    var evt = val('EventType');
    var evtLc = String(evt || '').trim().toLowerCase();

    // ---------- POSITIVE BRANCH (no PDF) ----------
    if (/^positive/.test(evtLc)) {
      var emp   = val('Employee');
      var lead  = val('Lead');
      var act   = val('PositiveAction') || val('Positive Action') || '';
      var pts   = val('Points'); // optional, show if you want
      // pull label from PositiveMap -> 'Minor/Moderate/Major Positive Credit' or 'Grace'
      var label = (typeof positiveTierLabelForAction_ === 'function')
        ? positiveTierLabelForAction_(act)
        : 'Positive Credit';

      var parts = [];
      parts.push('*' + (emp || 'Unknown') + '*');
      parts.push('earned a *' + label + '*');
      if (act) parts.push('for *' + act + '*');
      if (pts && !isNaN(Number(pts))) parts.push('(' + pts + ' pt)');
      if (lead) parts.push('(noted by ' + lead + ')');

      // optional deep link to row
      var rowUrl = (typeof sheetRowUrl_ === 'function') ? sheetRowUrl_(sh, row) : '';
      if (rowUrl) parts.push('— <' + rowUrl + '|Open row>');

      var text = parts.join(' ');
      logInfo_('notifyDocs_positive_payload', {row:row, text:text});
      var sent = postSlack_(CONFIG.DOCS_WEBHOOK, { text:text, mrkdwn:true });
      if (sent) logAudit && logAudit('slack','docs_post_positive', row, {employee:emp, eventType:evt});
      return sent;
    }

    // ---------- DISCIPLINARY BRANCH ----------
    // Allow-list enforcement (only for non-positive events)
    if (CONFIG.ONLY_DISCIPLINARY_SLACK){
      var allowed = (CONFIG.DOCS_EVENT_TYPES || []).map(function(t){ return String(t||'').toLowerCase(); });
      if (allowed.indexOf(evtLc) === -1){
        logInfo_('notifyDocs_skip_eventType', {evt:evt, allowed:CONFIG.DOCS_EVENT_TYPES});
        return false;
      }
    }

    var emp   = val('Employee');
    var lead  = val('Lead');
    var infr  = val('Infraction');
    var pts   = val('Points');

    // read PDF link (rich text or plain)
    var pdfCell = getCell('PdfLink');
    var pdfUrl = '';
    if (pdfCell){
      var rtv = pdfCell.getRichTextValue && pdfCell.getRichTextValue();
      if (rtv && rtv.getLinkUrl) pdfUrl = rtv.getLinkUrl() || '';
      if (!pdfUrl){
        var s = pdfCell.getDisplayValue ? pdfCell.getDisplayValue() : String(pdfCell.getValue()||'');
        var m = String(s||'').match(/https:\/\/drive\.google\.com\/[^\s)]+/i);
        if (m) pdfUrl = m[0];
      }
    }

    var text = '*' + emp + '* was written up for *' + (infr || evt || '—') + '* by ' + (lead || 'Unknown');
    if (pts) text += ' — ' + pts + ' pts';
    if (pdfUrl) text += ' — <' + pdfUrl + '|View PDF>';

    logInfo_('notifyDocs_payload', {evt:evt, emp:emp, lead:lead, infr:infr, pts:pts, pdfUrl:pdfUrl});
    var okSend = postSlack_(CONFIG.DOCS_WEBHOOK, { text:text, mrkdwn:true });
    if (okSend) logAudit && logAudit('slack','docs_post', row, {employee:emp, eventType:evt, url:pdfUrl});
    return okSend;

  } catch(err){
    logError && logError('notifyDocs_', err, {row:row});
    return false;
  }
}


function notifyPositiveCredit_(row){
  try{
    // Single entry point to avoid double posting:
    return notifyDocs_(row);
  }catch(err){
    logError && logError('notifyPositiveCredit_', err, {row:row});
    return false;
  }
}



// NEW: notify leaders when a Milestone row is created (Pending, no PDF yet)
function notifyLeadersMilestonePending_(row){
  try{
    if (!CONFIG.ENABLE_SLACK_LEADERS_PENDING || !CONFIG.LEADERS_WEBHOOK) return false;
    var sh  = sh_(CONFIG.TABS.EVENTS), map = headerIndexMap_(sh);
    function get(h){ var c=map[h]||map[CONFIG.COLS[h]]||0; return c?String(sh.getRange(row,c).getDisplayValue()||''):''; }

    var evt = get('EventType').toLowerCase();
    if (evt !== 'milestone') return false;

    var emp = get('Employee');
    var mil = get('Milestone') || 'Milestone';
    var pend= get('PendingStatus') || get('Pending Status') || '';
    var url = sheetRowUrl_(sh, row);

    var text = ':bell: *'+mil+'* pending for *'+emp+'*. A director needs to claim it.'
            + (url ? ' — <'+url+'|Open row>' : '');

    var ok = postSlack_(CONFIG.LEADERS_WEBHOOK, { text:text, mrkdwn:true });
    if (ok) logAudit('slack','leaders_pending_post', row, {employee:emp, milestone:mil, url:url});
    return ok;
  }catch(err){ logError('notifyLeadersMilestonePending_', err, {row:row}); return false; }
}

// EXISTING: claimed/after-PDF notifier — gate it off by config
// slack.js
// slack.js
function notifyLeadersMilestone_(row, pdfUrl){
  try{
    var enabledClaimed = (typeof CONFIG.ENABLE_SLACK_LEADERS_CLAIMED !== 'undefined')
      ? CONFIG.ENABLE_SLACK_LEADERS_CLAIMED
      : CONFIG.ENABLE_SLACK_LEADERS;

    if (!enabledClaimed || !CONFIG.LEADERS_WEBHOOK){
      logInfo_('leaders_claimed_skip_switch', {enabledClaimed, hasWebhook: !!CONFIG.LEADERS_WEBHOOK});
      return false;
    }

    var sh  = sh_(CONFIG.TABS.EVENTS), map = headerIndexMap_(sh);
    function get(h){ var c=map[h]||map[CONFIG.COLS[h]]||0; return c?String(sh.getRange(row,c).getDisplayValue()||''):''; }

    var evt = (get('EventType')||'').trim().toLowerCase();
    if (evt !== 'milestone'){
      logInfo_('leaders_claimed_skip_evt', {row, evt});
      return false;
    }

    var emp = get('Employee');
    var mil = get('Milestone') || 'Milestone';
    var dir = get('ConsequenceDirector') || 'Director';
    var url = pdfUrl || '';

    var text=':white_check_mark: *'+mil+'* claimed for *'+emp+'* by '+dir+(url?' — <'+url+'|View PDF>':'');
    logInfo_('leaders_claimed_payload', {row, emp, mil, dir, url});
    var ok = postSlack_(CONFIG.LEADERS_WEBHOOK, {text, mrkdwn:true});
    logInfo_('leaders_claimed_post_result', {ok});
    if (ok) logAudit('slack','leaders_claimed_post', row, {employee:emp, milestone:mil, url:url});
    return ok;
  }catch(err){
    logError('notifyLeadersMilestone_', err, {row});
    return false;
  }
}

function onEdit_NotifyLeadersOnClaim_(e){
  try{
    if (!e || !e.range || !e.range.getSheet) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== CONFIG.TABS.EVENTS) return;

    var row = e.range.getRow(); if (row < 2) return;
    var map = headerIndexMap_(sh);

    var pendHdr = CONFIG.COLS.PendingStatus || 'Pending Status';
    var pdfHdr  = CONFIG.COLS.PdfLink || 'Write-Up PDF';
    var evtHdr  = CONFIG.COLS.EventType || 'EventType';

    var pendCol = map[pendHdr] || 0;
    if (!pendCol || e.range.getColumn() !== pendCol) return; // only when Pending edited

    var evt = map[evtHdr] ? String(sh.getRange(row, map[evtHdr]).getDisplayValue()||'').trim().toLowerCase() : '';
    if (evt !== 'milestone') return;

    var newPend = String(e.value || sh.getRange(row, pendCol).getDisplayValue() || '').trim().toLowerCase();
    if (['claimed','complete','completed'].indexOf(newPend) === -1) return;

    // need a PDF link to include in Slack
    var pdfCell = map[pdfHdr] ? sh.getRange(row, map[pdfHdr]) : null;
    var pdfUrl = '';
    if (pdfCell){
      var rtv = pdfCell.getRichTextValue && pdfCell.getRichTextValue();
      if (rtv && rtv.getLinkUrl) pdfUrl = rtv.getLinkUrl() || '';
      if (!pdfUrl){
        var s = pdfCell.getDisplayValue ? pdfCell.getDisplayValue() : String(pdfCell.getValue()||'');
        var m = String(s||'').match(/https:\/\/drive\.google\.com\/[^\s)]+/i);
        if (m) pdfUrl = m[0];
      }
    }

    logInfo_ && logInfo_('leaders_claimed_fallback_attempt', {row: row, pending: newPend, pdfUrl: pdfUrl});
    if (typeof notifyLeadersMilestone_ === 'function'){
      notifyLeadersMilestone_(row, pdfUrl);
    }
  }catch(err){
    logError && logError('onEdit_NotifyLeadersOnClaim_', err);
  }
}




