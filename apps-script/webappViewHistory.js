// /** 
//  * WEB APP — Employee Self View (Read-Only)
//  * Purpose: Let a signed-in team member view THEIR OWN history and current effective points.
//  * Security: Deploy Web App “Execute as: User accessing the app” + restrict access to specific emails.
//  * No writes. No side effects. Single-file (no .html needed).
//  *
//  * Optional Directory tab: (recommended)
//  *   Tab name: Directory
//  *   Columns (exact headers): Email | Employee
//  *   Use this to map login emails → Employee names used in Events.
//  */

// // === Add near the top ===
// const WA_ALLOWLIST = (typeof CONFIG !== 'undefined' && CONFIG.ALLOW) ? CONFIG.ALLOW : {
//   // Put the personal Google accounts that may use this POC web app.
//   leads: ['philippixler@gmail.com','lead2@gmail.com'],
//   directors: ['you@gmail.com','director1@gmail.com','director2@gmail.com']
// };

// // Your Google Identity Services Web client ID (create one at console.cloud.google.com -> Credentials -> OAuth 2.0 Client IDs)
// const WA_GOOGLE_CLIENT_ID = '348120683913-63jpj9pnu8j0ibvfh3c3s1296k8k8n5q.apps.googleusercontent.com';


// const WA_CFG = (typeof CONFIG !== 'undefined') ? CONFIG : {};
// const WA_TABS = Object.assign({ EVENTS: 'Events', DIRECTORY: 'Directory' }, WA_CFG.TABS || {});
// const WA_COLS = Object.assign({
//   Employee: 'Employee',
//   IncidentDate: 'IncidentDate',
//   EventType: 'EventType',
//   Infraction: 'Infraction',
//   Lead: 'Lead',
//   Points: 'Points',
//   PointsRolling: 'PointsRolling',
//   PointsRollingEffective: 'PointsRolling (Effective)',
//   NotesReviewer: 'Notes / Reviewer',
//   PdfLink: 'Write-Up PDF',
//   GraceApplied: 'Grace Applied'
// }, WA_CFG.COLS || {});

// function doGet() {
//   var html = HtmlService.createHtmlOutput(`
// <!doctype html>
// <meta name="viewport" content="width=device-width,initial-scale=1" />
// <title>CLEAR — My History</title>
// <style>
//   body{font-family:Inter,Arial,sans-serif;margin:24px;max-width:1100px}
//   .card{border:1px solid #e5e7eb;border-radius:16px;padding:16px;margin:0 0 16px 0;box-shadow:0 1px 2px rgba(0,0,0,.05)}
//   .muted{color:#6b7280}.small{font-size:12px}
//   h1{font-size:22px;margin:0 0 8px 0} h2{font-size:18px;margin:0 0 8px 0}
//   table{width:100%;border-collapse:collapse;margin-top:8px}
//   th,td{border-bottom:1px solid #f3f4f6;text-align:left;padding:8px}
//   th{font-weight:600;font-size:12px;color:#4b5563;text-transform:uppercase;letter-spacing:.02em}
//   td{font-size:14px;color:#111827;vertical-align:top}
//   a{color:#2563eb;text-decoration:none} a:hover{text-decoration:underline}
//   .center{display:flex;justify-content:center;align-items:center;min-height:30vh;flex-direction:column;gap:12px}
//   .warn{color:#b45309}.ok{color:#065f46}
//   #app{display:none}
// </style>

// <div id="signin" class="card center">
//   <h1>Sign in to view your CLEAR history</h1>
//   <div id="g_id_onload"
//        data-client_id="${WA_GOOGLE_CLIENT_ID}"
//        data-context="signin"
//        data-auto_prompt="false"
//        data-callback="onGoogleSignIn"></div>
//   <div class="g_id_signin"
//        data-type="standard"
//        data-shape="pill"
//        data-theme="outline"
//        data-text="continue_with"
//        data-size="large"
//        data-logo_alignment="left"></div>
//   <div class="muted small">Use the Google account our system has on file for you.</div>
// </div>

// <div class="small muted">client: ${WA_GOOGLE_CLIENT_ID}</div>

// <div id="app" class="card">
//   <h1>My CLEAR History</h1>
//   <div class="muted small" id="who">Loading…</div>

//   <div class="card" style="margin-top:12px">
//     <h2>Current Effective Points</h2>
//     <div id="points" style="font-size:28px;font-weight:700">—</div>
//     <div class="muted small">From your most recent eligible event</div>
//   </div>

//   <div class="card">
//     <h2>Grace Available</h2>
//     <div id="grace" style="font-size:28px;font-weight:700">—</div>
//     <div class="muted small">Credits you may qualify for</div>
//   </div>

//   <div class="card">
//     <h2>Work History</h2>
//     <div class="muted small" id="summary"></div>
//     <table id="hist"><thead>
//       <tr><th>Date</th><th>Event</th><th>Infraction / Notes</th><th>Pts</th><th>Roll</th><th>Lead</th><th>PDF</th></tr>
//     </thead><tbody id="rows"><tr><td colspan="7" class="muted">Loading…</td></tr></tbody></table>
//   </div>
// </div>

// <script>
//   // Load GIS library
//   (function(d,s,id){var js=d.createElement(s);js.id=id;js.src="https://accounts.google.com/gsi/client";
//     var f=d.getElementsByTagName(s)[0];f.parentNode.insertBefore(js,f);})(document,"script","gis");

//   function esc(s){return String(s||'').replace(/[&<>"']/g,m=>({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;', "'":'&#39;' }[m]));}
//   function linkCell(url,text){ if(!url) return ''; return '<a target="_blank" rel="noopener" href="'+esc(url)+'">'+esc(text||'View')+'</a>'; }

//   // Called by Google after sign-in
//   function onGoogleSignIn(resp){
//     const token = resp && resp.credential;
//     if(!token){ alert('Sign-in failed.'); return; }
//     google.script.run.withSuccessHandler(function(auth){
//       if(!auth || !auth.ok){ document.getElementById('signin').innerHTML = '<div class="warn">Access denied.</div>'; return; }
//       // Hide sign-in, show app
//       document.getElementById('signin').style.display='none';
//       document.getElementById('app').style.display='block';

//       // Fetch data for this user
//       google.script.run.withSuccessHandler(function(data){
//         var who = document.getElementById('who');
//         if (data && data.employee){
//           who.innerHTML = 'Signed in as <b>'+esc(auth.email||'?')+'</b> — Showing records for <b>'+esc(data.employee)+'</b>';
//         } else {
//           who.innerHTML = '<span class="warn">We couldn\\'t match your login to an employee record. Please contact a Director.</span><br/><span class="small muted">'+esc(auth.email||'')+'</span>';
//         }

//         document.getElementById('points').textContent = (data && data.effectivePoints != null) ? String(data.effectivePoints) : '—';
//         document.getElementById('grace').textContent  = (data && data.graceAvailableText) ? data.graceAvailableText : (data && data.graceAvailable != null ? String(data.graceAvailable) : '—');

//         var body = document.getElementById('rows');
//         body.innerHTML = '';
//         var rows = (data && data.rows) || [];
//         if (!rows.length){
//           body.innerHTML = '<tr><td colspan="7" class="muted">No history yet.</td></tr>';
//         } else {
//           rows.forEach(function(r){
//             var tr = document.createElement('tr');
//             tr.innerHTML =
//               '<td>'+esc(r.date||'')+'</td>'+
//               '<td>'+esc(r.event||'')+'</td>'+
//               '<td>'+esc(r.infraction||'')+(r.notes?'<div class="small muted">'+esc(r.notes)+'</div>':'')+'</td>'+
//               '<td>'+esc(r.points==null ? '' : String(r.points))+'</td>'+
//               '<td>'+esc(r.roll==null ? '' : String(r.roll))+'</td>'+
//               '<td>'+esc(r.lead||'')+'</td>'+
//               '<td>'+ (r.pdfUrl ? linkCell(r.pdfUrl, 'View PDF') : '') +'</td>';
//             body.appendChild(tr);
//           });
//         }
//         var s = document.getElementById('summary');
//         if (data && data.generatedAt){
//           s.textContent = 'Last updated: '+data.generatedAt+(data.rows && data.rows.length? ' — '+data.rows.length+' record(s)':'');
//         }
//       }).getMyOverviewForEmail(auth.email);

//     }).authorize(token);
//   }
// </script>
//   `).setTitle('CLEAR — My History');
//   return html;
// }

// // ---------- Server: data builder (read-only) ----------

// function getMyOverview() {
//   var email = getCurrentUserEmail();
//   var employee = resolveEmployeeName(email); // uses Directory tab if present; safe fallback
//   var out = { email: email, employee: employee, effectivePoints: null, graceAvailable: null, graceAvailableText: null, rows: [], generatedAt: formatStamp_(new Date()) };

//   if (!employee) {
//     return out; // not matched; UI shows a helpful message
//   }

//   var s = ss_().getSheetByName(WA_TABS.EVENTS);
//   if (!s) return out;

//   var hdr = headers_(s);
//   function idx(name){ var i = hdr.indexOf(name); return i >= 0 ? i : -1; }

//   var iEmp   = idx(WA_COLS.Employee);
//   var iDate  = firstFoundIndex_(hdr, [WA_COLS.IncidentDate, 'Date', 'Timestamp']);
//   var iEvt   = idx(WA_COLS.EventType);
//   var iInf   = idx(WA_COLS.Infraction);
//   var iLead  = idx(WA_COLS.Lead);
//   var iPts   = idx(WA_COLS.Points);
//   var iRollE = idx(WA_COLS.PointsRollingEffective);
//   var iRoll  = idx(WA_COLS.PointsRolling);
//   var iNotes = firstFoundIndex_(hdr, [WA_COLS.NotesReviewer, 'Notes', 'Reviewer', 'Notes/Reviewer']);
//   var iPdf   = firstFoundIndex_(hdr, [WA_COLS.PdfLink, 'PDF Link', 'WriteUpPDF', 'Signed_PDF_Link']);

//   var values = s.getDataRange().getValues();
//   var tz = Session.getScriptTimeZone() || 'UTC';

//   var rows = [];
//   var lastEffective = null;

//   for (var r = 1; r < values.length; r++){
//     var row = values[r];
//     if (iEmp === -1) break;
//     if (String(row[iEmp]||'').trim().toLowerCase() !== String(employee).trim().toLowerCase()) continue;

//     // read basics
//     var date = tryFormatDate_(row[iDate], tz);
//     var evt  = safeString_(row[iEvt]);
//     var inf  = safeString_(row[iInf]);
//     var lead = safeString_(row[iLead]);
//     var pts  = (iPts !== -1 && row[iPts] !== '' && row[iPts] != null) ? Number(row[iPts]) : null;

//     // effective rolling preference
//     var rollHere = null;
//     if (iRollE !== -1 && isFinite(Number(values[r][iRollE]))) rollHere = Number(values[r][iRollE]);
//     else if (iRoll !== -1 && isFinite(Number(values[r][iRoll]))) rollHere = Number(values[r][iRoll]);
//     if (rollHere != null) lastEffective = rollHere;

//     // notes
//     var notes = (iNotes !== -1) ? safeString_(row[iNotes]) : '';

//     // pdf link (attempt to read rich link if helper exists)
//     var pdfUrl = '';
//     try {
//       if (typeof readLinkUrlFromCell_ === 'function' && iPdf !== -1) {
//         var url = readLinkUrlFromCell_(s.getRange(r+1, iPdf+1));
//         pdfUrl = url || String(s.getRange(r+1, iPdf+1).getDisplayValue() || '');
//       } else if (iPdf !== -1) {
//         pdfUrl = String(row[iPdf] || '');
//       }
//     } catch (_){}

//     rows.push({ date: date, event: evt, infraction: inf, notes: notes, points: pts, roll: rollHere, lead: lead, pdfUrl: pdfUrl });
//   }

//   out.rows = rows;
//   out.effectivePoints = (lastEffective != null) ? lastEffective : null;

//   // Grace available (best-effort)
//   try {
//     if (typeof countUniversalCredits_ === 'function') {
//       var g = countUniversalCredits_(employee);
//       out.graceAvailable = g && typeof g.count === 'number' ? g.count : null;
//       out.graceAvailableText = (g && typeof g.count === 'number') ? (g.count + ' universal') : null;
//     } else {
//       // simple fallback: look for a "Positive Points" / credits ledger if present
//       var guess = tryCountGraceFallback_(employee);
//       if (guess != null) {
//         out.graceAvailable = guess;
//         out.graceAvailableText = String(guess);
//       }
//     }
//   } catch (_){}

//   return out;
// }

// // ---------- Helpers (isolated; read-only) ----------

// // Verify the Google ID token with Google and return the email (consumer-safe).
// function authorize(idToken){
//   if (!idToken) return { ok:false, error:'no_token' };
//   var info = verifyIdToken_(idToken);
//   if (!info || !info.email || !info.email_verified) return { ok:false, error:'invalid_token' };

//   var email = String(info.email).toLowerCase();
//   var ok = isAllowedEmail_(email);
//   return ok ? { ok:true, email: email } : { ok:false, email: email, error:'not_allowed' };
// }

// function isAllowedEmail_(email){
//   var list = []
//     .concat(WA_ALLOWLIST.leads || [])
//     .concat(WA_ALLOWLIST.directors || [])
//     .map(function(e){ return String(e||'').toLowerCase().trim(); });
//   return list.indexOf(String(email||'').toLowerCase().trim()) !== -1;
// }

// function getMyOverviewForEmail(email){
//   var employee = resolveEmployeeName(email);
//   var out = { email: email, employee: employee, effectivePoints: null, graceAvailable: null, graceAvailableText: null, rows: [], generatedAt: formatStamp_(new Date()) };
//   if (!employee) return out;
//   // reuse your existing read-only builder (same as getMyOverview but without Session.getActiveUser)
//   return getMyOverviewForEmployee_(employee);
// }

// function getMyOverviewForEmployee_(employee){
//   var s = ss_().getSheetByName(WA_TABS.EVENTS);
//   if (!s) return { employee: employee, rows: [], generatedAt: formatStamp_(new Date()) };

//   var hdr = headers_(s);
//   function idx(name){ var i = hdr.indexOf(name); return i >= 0 ? i : -1; }
//   var iEmp   = idx(WA_COLS.Employee);
//   var iDate  = firstFoundIndex_(hdr, [WA_COLS.IncidentDate, 'Date', 'Timestamp']);
//   var iEvt   = idx(WA_COLS.EventType);
//   var iInf   = idx(WA_COLS.Infraction);
//   var iLead  = idx(WA_COLS.Lead);
//   var iPts   = idx(WA_COLS.Points);
//   var iRollE = idx(WA_COLS.PointsRollingEffective);
//   var iRoll  = idx(WA_COLS.PointsRolling);
//   var iNotes = firstFoundIndex_(hdr, [WA_COLS.NotesReviewer, 'Notes', 'Reviewer', 'Notes/Reviewer']);
//   var iPdf   = firstFoundIndex_(hdr, [WA_COLS.PdfLink, 'PDF Link', 'WriteUpPDF', 'Signed_PDF_Link']);

//   var values = s.getDataRange().getValues();
//   var tz = Session.getScriptTimeZone() || 'UTC';

//   var rows = [], lastEffective = null;
//   for (var r = 1; r < values.length; r++){
//     if (String(values[r][iEmp]||'').trim().toLowerCase() !== String(employee).trim().toLowerCase()) continue;

//     var rollHere = null;
//     if (iRollE !== -1 && isFinite(Number(values[r][iRollE]))) rollHere = Number(values[r][iRollE]);
//     else if (iRoll !== -1 && isFinite(Number(values[r][iRoll]))) rollHere = Number(values[r][iRoll]);

//     if (rollHere != null) lastEffective = rollHere;

//     var pdfUrl = '';
//     try {
//       if (typeof readLinkUrlFromCell_ === 'function' && iPdf !== -1) {
//         var url = readLinkUrlFromCell_(s.getRange(r+1, iPdf+1));
//         pdfUrl = url || String(s.getRange(r+1, iPdf+1).getDisplayValue() || '');
//       } else if (iPdf !== -1) {
//         pdfUrl = String(values[r][iPdf] || '');
//       }
//     } catch (_){}

//     rows.push({
//       date: tryFormatDate_(values[r][iDate], tz),
//       event: safeString_(values[r][iEvt]),
//       infraction: safeString_(values[r][iInf]),
//       notes: (iNotes !== -1) ? safeString_(values[r][iNotes]) : '',
//       points: (iPts !== -1 && values[r][iPts] !== '' && values[r][iPts] != null) ? Number(values[r][iPts]) : null,
//       roll: rollHere,
//       lead: safeString_(values[r][iLead]),
//       pdfUrl: pdfUrl
//     });
//   }

//   var out = {
//     employee: employee,
//     rows: rows,
//     effectivePoints: (lastEffective != null) ? lastEffective : null,
//     graceAvailable: null,
//     graceAvailableText: null,
//     generatedAt: formatStamp_(new Date())
//   };

//   // Grace count (reuse your helper if present)
//   try {
//     if (typeof countUniversalCredits_ === 'function') {
//       var g = countUniversalCredits_(employee);
//       out.graceAvailable = (g && typeof g.count === 'number') ? g.count : null;
//       out.graceAvailableText = (g && typeof g.count === 'number') ? (g.count + ' universal') : null;
//     } else {
//       var guess = tryCountGraceFallback_(employee);
//       if (guess != null) { out.graceAvailable = guess; out.graceAvailableText = String(guess); }
//     }
//   } catch (_){}

//   return out;
// }

// // Validate ID token against Google and basic claims.
// function verifyIdToken_(idToken){
//   try{
//     var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
//     if (res.getResponseCode() !== 200) return null;
//     var obj = JSON.parse(res.getContentText() || '{}');

//     // Minimal checks: audience and expiration
//     if (obj.aud !== WA_GOOGLE_CLIENT_ID) return null;
//     if (Number(obj.exp || 0) * 1000 < Date.now()) return null;
//     return obj; // contains email, email_verified, sub, etc.
//   }catch(e){ return null; }
// }

// function getCurrentUserEmail(){
//   return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
// }

// // Prefer Directory.Email -> Employee mapping; fallback tries to infer by Events exact matches
// function resolveEmployeeName(email){
//   if (!email) return '';
//   // 1) Directory tab mapping
//   var dir = ss_().getSheetByName(WA_TABS.DIRECTORY);
//   if (dir){
//     var data = dir.getDataRange().getValues();
//     var h = (data[0] || []).map(function(v){ return String(v||'').trim(); });
//     var iEmail = h.indexOf('Email');
//     var iEmp   = h.indexOf('Employee');
//     if (iEmail !== -1 && iEmp !== -1){
//       for (var r=1; r < data.length; r++){
//         var e = String(data[r][iEmail] || '').trim().toLowerCase();
//         if (e && e === email) return String(data[r][iEmp] || '').trim();
//       }
//     }
//   }
//   // 2) Fallback: if someone typed their email as Employee anywhere (rare), use first non-empty
//   var s = ss_().getSheetByName(WA_TABS.EVENTS);
//   if (s){
//     var vals = s.getDataRange().getValues();
//     var hdr = headers_(s);
//     var iEmp = hdr.indexOf(WA_COLS.Employee);
//     if (iEmp !== -1){
//       // naive: look for a row whose Lead or Notes contains the email and grab that Employee
//       var iLead = hdr.indexOf(WA_COLS.Lead);
//       var iNotes = hdr.indexOf(WA_COLS.NotesReviewer);
//       for (var r=1; r<vals.length; r++){
//         var hay = [vals[r][iLead], vals[r][iNotes]].map(function(v){return String(v||'').toLowerCase();}).join(' ');
//         if (hay.indexOf(email) !== -1){
//           var emp = String(vals[r][iEmp] || '').trim();
//           if (emp) return emp;
//         }
//       }
//     }
//   }
//   return '';
// }

// // Grace fallback: counts rows in a "Positive Points" style ledger that aren’t consumed.
// // You can replace this with your real ledger logic, or leave it as null.
// function tryCountGraceFallback_(employee){
//   try{
//     var tab = (WA_CFG.TABS && WA_CFG.TABS.POSITIVE_POINTS) ? WA_CFG.TABS.POSITIVE_POINTS : 'Positive Points';
//     var s = ss_().getSheetByName(tab);
//     if (!s) return null;
//     var vals = s.getDataRange().getValues();
//     if (vals.length < 2) return 0;
//     var h = vals[0].map(function(v){ return String(v||'').trim(); });
//     function find(names){ for (var i=0;i<names.length;i++){ var idx = h.indexOf(names[i]); if (idx !== -1) return idx; } return -1; }
//     var iEmp = find(['Employee','Name']);
//     var iConsumed = find(['Consumed?','Consumed','Used?']);
//     if (iEmp === -1 || iConsumed === -1) return null;
//     var count = 0;
//     for (var r=1; r<vals.length; r++){
//       if (String(vals[r][iEmp]||'').trim().toLowerCase() !== String(employee).trim().toLowerCase()) continue;
//       var used = String(vals[r][iConsumed]||'').toLowerCase();
//       var isUsed = (vals[r][iConsumed] === true) || used === 'true' || used === 'yes' || used === 'y' || used === '1';
//       if (!isUsed) count++;
//     }
//     return count;
//   }catch(e){ return null; }
// }

// // Utilities borrowed (read-only) — safe fallbacks if your project doesn’t expose these
// function ss_(){ return SpreadsheetApp.getActive(); }
// function headers_(sheet){ return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h||'').trim(); }); }
// function firstFoundIndex_(hdr, names){ for (var i=0;i<names.length;i++){ var idx = hdr.indexOf(names[i]); if (idx !== -1) return idx; } return -1; }
// function safeString_(v){ return (v == null) ? '' : String(v); }
// function tryFormatDate_(v, tz){
//   try {
//     if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
//     var d = new Date(v);
//     if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
//   } catch (_){}
//   var s = String(v||''); return s.length >= 10 ? s.slice(0,10) : s;
// }
// function formatStamp_(d){ var tz = Session.getScriptTimeZone() || 'UTC'; return Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm'); }
