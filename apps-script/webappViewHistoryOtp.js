// /** OTP Web App (read-only) — no Google OAuth needed */
// const WA_ALLOWLIST = {
//   leads: ['philippixler@gmail.com'],     // add allowed addresses here
//   directors: ['you@gmail.com']
// };
// const WA_TABS = { EVENTS: 'Events', DIRECTORY: 'Directory' };
// const WA_COLS = {
//   Employee:'Employee', IncidentDate:'IncidentDate', EventType:'EventType',
//   Infraction:'Infraction', Lead:'Lead', Points:'Points',
//   PointsRolling:'PointsRolling', PointsRollingEffective:'PointsRolling (Effective)',
//   NotesReviewer:'Notes / Reviewer', PdfLink:'Write-Up PDF'
// };
// const OTP_TTL_MIN = 10;       // code expires in 10 minutes
// const SESS_TTL_MIN = 60 * 24; // session token ~24h

// function doGet() {
//   return HtmlService.createHtmlOutput(`
// <!doctype html><meta name="viewport" content="width=device-width,initial-scale=1">
// <title>CLEAR — My History</title>
// <style>
//   body{font-family:Inter,Arial;margin:24px;max-width:720px}
//   .card{border:1px solid #e5e7eb;border-radius:16px;padding:16px;margin:0 0 16px;box-shadow:0 1px 2px rgba(0,0,0,.05)}
//   .muted{color:#6b7280}.small{font-size:12px} button{padding:8px 12px;border-radius:8px}
//   input{padding:8px 10px;border:1px solid #e5e7eb;border-radius:8px;width:100%}
//   table{width:100%;border-collapse:collapse;margin-top:8px}
//   th,td{border-bottom:1px solid #f3f4f6;text-align:left;padding:8px}
//   th{font-weight:600;font-size:12px;color:#4b5563;text-transform:uppercase}
// </style>

// <div id="signin" class="card">
//   <h2>Sign in to view your CLEAR history</h2>
//   <div class="small muted">Use the email we have on file for you.</div>
//   <div style="display:grid;gap:8px;margin-top:8px">
//     <input id="email" placeholder="you@example.com" autocomplete="email">
//     <div><button onclick="sendCode()">Send code</button> <span id="msg" class="small muted"></span></div>
//     <div id="codeRow" style="display:none;gap:8px">
//       <input id="code" placeholder="6-digit code" maxlength="6" style="max-width:140px">
//       <button onclick="verify()">Verify</button>
//     </div>
//   </div>
// </div>

// <div id="app" class="card" style="display:none">
//   <div class="small muted" id="who"></div>
//   <h2>Current Effective Points</h2>
//   <div id="points" style="font-size:28px;font-weight:700;margin-bottom:8px">—</div>
//   <h3>Grace Available</h3>
//   <div id="grace" style="font-size:22px;margin-bottom:8px">—</div>
//   <h3>Work History</h3>
//   <div class="small muted" id="summary"></div>
//   <table><thead><tr><th>Date</th><th>Event</th><th>Infraction/Notes</th><th>Pts</th><th>Roll</th><th>Lead</th><th>PDF</th></tr></thead>
//   <tbody id="rows"><tr><td colspan="7" class="muted">Loading…</td></tr></tbody></table>
// </div>

// <script>
// function esc(s){return String(s||'').replace(/[&<>"']/g,m=>({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;', "'":'&#39;' }[m]));}
// function linkCell(u,t){return u?'<a target="_blank" rel="noopener" href="'+esc(u)+'">'+esc(t||'View')+'</a>':'';}
// function sendCode(){
//   const email = document.getElementById('email').value.trim();
//   document.getElementById('msg').textContent = 'Sending…';
//   google.script.run.withSuccessHandler(function(res){
//     document.getElementById('msg').textContent = res.ok? 'Code sent. Check your email.' : (res.error||'Failed.');
//     if(res.ok) document.getElementById('codeRow').style.display='grid';
//   }).requestSigninCode(email);
// }
// function verify(){
//   const email = document.getElementById('email').value.trim();
//   const code = document.getElementById('code').value.trim();
//   google.script.run.withSuccessHandler(function(res){
//     if(!res.ok){ alert(res.error||'Invalid code'); return; }
//     // load data
//     google.script.run.withSuccessHandler(function(data){
//       document.getElementById('signin').style.display='none';
//       document.getElementById('app').style.display='block';
//       document.getElementById('who').textContent = 'Signed in as '+email+' — records for '+(data.employee||'');
//       document.getElementById('points').textContent = (data.effectivePoints!=null)? String(data.effectivePoints):'—';
//       document.getElementById('grace').textContent = data.graceAvailableText || (data.graceAvailable!=null? String(data.graceAvailable):'—');
//       const body = document.getElementById('rows'); body.innerHTML='';
//       (data.rows||[]).forEach(function(r){
//         const tr = document.createElement('tr');
//         tr.innerHTML = '<td>'+esc(r.date||'')+'</td><td>'+esc(r.event||'')+'</td>'+
//           '<td>'+esc(r.infraction||'')+(r.notes?'<div class="small muted">'+esc(r.notes)+'</div>':'')+'</td>'+
//           '<td>'+esc(r.points==null?'':String(r.points))+'</td><td>'+esc(r.roll==null?'':String(r.roll))+'</td>'+
//           '<td>'+esc(r.lead||'')+'</td><td>'+linkCell(r.pdfUrl,'View PDF')+'</td>';
//         body.appendChild(tr);
//       });
//       document.getElementById('summary').textContent = 'Last updated: '+(data.generatedAt||'');
//     }).getMyOverviewForEmail(res.email);
//   }).verifySigninCode(email, code);
// }
// </script>
//   `).setTitle('CLEAR — My History (OTP)');
// }

// /** ===== Server-side auth (OTP) ===== */
// function requestSigninCode(email){
//   email = String(email||'').trim().toLowerCase();
//   if (!isAllowedEmail_(email)) return {ok:false, error:'Email not allowed'};
//   const code = ('000000' + Math.floor(Math.random()*1e6)).slice(-6);
//   CacheService.getScriptCache().put('otp:'+email, code, OTP_TTL_MIN*60);
//   try { MailApp.sendEmail(email, 'Your CLEAR code', 'Your verification code: '+code+' (expires in '+OTP_TTL_MIN+' minutes)'); }
//   catch(e){ return {ok:false, error:'Mail failed'}; }
//   return {ok:true};
// }
// function verifySigninCode(email, code){
//   email = String(email||'').trim().toLowerCase();
//   const cached = CacheService.getScriptCache().get('otp:'+email);
//   if (!cached || String(cached) !== String(code)) return {ok:false, error:'Invalid or expired code'};
//   const token = Utilities.getUuid();
//   CacheService.getScriptCache().put('sess:'+token, email, SESS_TTL_MIN*60);
//   return {ok:true, token:token, email:email};
// }
// function isAllowedEmail_(email){
//   const list = []
//     .concat(WA_ALLOWLIST.leads||[])
//     .concat(WA_ALLOWLIST.directors||[])
//     .map(x=>String(x||'').toLowerCase().trim());
//   return list.indexOf(email) !== -1;
// }

// /** ===== Data (reuses your existing read-only logic) ===== */
// function getMyOverviewForEmail(email){
//   const employee = resolveEmployeeName(email);
//   const out = { email, employee, effectivePoints:null, graceAvailable:null, graceAvailableText:null, rows:[], generatedAt:formatStamp_(new Date()) };
//   if (!employee) return out;
//   return getMyOverviewForEmployee_(employee);
// }

// /** ===== Helpers (same as before; trimmed) ===== */
// function ss_(){ return SpreadsheetApp.getActive(); }
// function headers_(sheet){ return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h=>String(h||'').trim()); }
// function firstFoundIndex_(hdr, names){ for (var i=0;i<names.length;i++){ var k = hdr.indexOf(names[i]); if (k !== -1) return k; } return -1; }
// function safeString_(v){ return (v==null)?'':String(v); }
// function formatStamp_(d){ var tz = Session.getScriptTimeZone()||'UTC'; return Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm'); }
// function tryFormatDate_(v,tz){ try{ if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v,tz,'yyyy-MM-dd'); var d=new Date(v); if(!isNaN(d)) return Utilities.formatDate(d,tz,'yyyy-MM-dd'); }catch(_){}
//   var s=String(v||''); return s.length>=10?s.slice(0,10):s; }

// function resolveEmployeeName(email){
//   if (!email) return '';
//   var dir = ss_().getSheetByName(WA_TABS.DIRECTORY);
//   if (dir){
//     var data=dir.getDataRange().getValues(), h=(data[0]||[]).map(v=>String(v||'').trim());
//     var iE=h.indexOf('Email'), iN=h.indexOf('Employee');
//     if (iE!==-1 && iN!==-1) for (var r=1;r<data.length;r++){
//       if (String(data[r][iE]||'').trim().toLowerCase()===email) return String(data[r][iN]||'').trim();
//     }
//   }
//   return '';
// }
// function getMyOverviewForEmployee_(employee){
//   var s=ss_().getSheetByName(WA_TABS.EVENTS); if(!s) return {employee, rows:[], generatedAt:formatStamp_(new Date())};
//   var hdr=headers_(s), idx=n=>hdr.indexOf(n), iEmp=idx(WA_COLS.Employee);
//   var iDate=firstFoundIndex_(hdr,[WA_COLS.IncidentDate,'Date','Timestamp']), iEvt=idx(WA_COLS.EventType), iInf=idx(WA_COLS.Infraction),
//       iLead=idx(WA_COLS.Lead), iPts=idx(WA_COLS.Points), iRollE=idx(WA_COLS.PointsRollingEffective), iRoll=idx(WA_COLS.PointsRolling),
//       iNotes=firstFoundIndex_(hdr,[WA_COLS.NotesReviewer,'Notes','Reviewer','Notes/Reviewer']),
//       iPdf=firstFoundIndex_(hdr,[WA_COLS.PdfLink,'PDF Link','WriteUpPDF','Signed_PDF_Link']);
//   var vals=s.getDataRange().getValues(), tz=Session.getScriptTimeZone()||'UTC', rows=[], lastEff=null;
//   for (var r=1;r<vals.length;r++){
//     if (String(vals[r][iEmp]||'').trim().toLowerCase()!==employee.toLowerCase()) continue;
//     var roll=null; if(iRollE!==-1&&isFinite(Number(vals[r][iRollE]))) roll=Number(vals[r][iRollE]); else if(iRoll!==-1&&isFinite(Number(vals[r][iRoll]))) roll=Number(vals[r][iRoll]);
//     if (roll!=null) lastEff=roll;
//     var pdfUrl=''; if (iPdf!==-1) pdfUrl=String(vals[r][iPdf]||'');
//     rows.push({ date:tryFormatDate_(vals[r][iDate],tz), event:safeString_(vals[r][iEvt]), infraction:safeString_(vals[r][iInf]),
//       notes:(iNotes!==-1)?safeString_(vals[r][iNotes]):'', points:(iPts!==-1&&vals[r][iPts]!==''&&vals[r][iPts]!=null)?Number(vals[r][iPts]):null,
//       roll:roll, lead:safeString_(vals[r][iLead]), pdfUrl:pdfUrl });
//   }
//   var out={ employee, rows, effectivePoints:(lastEff!=null)?lastEff:null, graceAvailable:null, graceAvailableText:null, generatedAt:formatStamp_(new Date()) };
//   return out;
// }
