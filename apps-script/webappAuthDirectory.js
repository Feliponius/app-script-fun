/** CLEAR Web App (Auth + Directory) — single-file drop-in
 * Features: Email OTP, Password login, Reset password, Invite flow, Pending self-claim
 * Storage: Directory sheet and CacheService for short-lived tokens
 * Security: Passwords hashed = base64(HMAC-SHA256( (salt + password), PEPPER ))
 */

/* ========= CONFIG (EDIT) ========= */
const AUTH = {
  DIRECTORS: ['you@gmail.com','director1@gmail.com'],             // TODO: your director emails
  PEPPER: 'change-this-long-random-secret-string-please',          // TODO: a long random secret
  OTP_TTL_MIN: 10,
  RESET_TTL_MIN: 15,
  SESSION_TTL_MIN: 60*24, // 24h
};

const TABS = { DIRECTORY: 'Directory', REQUESTS: 'AccessRequests', EVENTS: 'Events' };
const DIR_COLS = ['Email','Employee','Role','Verified','Salt','PassHash','CreatedAt','UpdatedAt','LastLogin'];
const REQ_COLS = ['RequestedAt','Email','Employee','Status'];

/* ========= UI ========= */
function doGet(e){
  return HtmlService.createHtmlOutput(`
<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1">
<title>CLEAR — Sign in</title>
<style>
  body{font-family:Inter,Arial;margin:24px;max-width:820px}
  .card{border:1px solid #e5e7eb;border-radius:16px;padding:16px;margin:0 0 16px;box-shadow:0 1px 2px rgba(0,0,0,.05)}
  .muted{color:#6b7280}.small{font-size:12px}
  input,button{padding:10px;border-radius:10px;border:1px solid #e5e7eb}
  button{cursor:pointer}
  .row{display:flex;gap:8px;flex-wrap:wrap}
  table{width:100%;border-collapse:collapse;margin-top:10px}
  th,td{border-bottom:1px solid #f3f4f6;text-align:left;padding:8px}
  th{font-weight:600;font-size:12px;color:#4b5563;text-transform:uppercase}
  .right{float:right}
  a{color:#2563eb;text-decoration:none}
</style>

<div id="signin" class="card">
  <h2>CLEAR — Sign in</h2>
  <div class="small muted">Use the email we have on file for you.</div>
  <div class="row" style="margin-top:8px">
    <input id="email" placeholder="you@example.com" style="min-width:260px" autocomplete="email">
    <input id="password" type="password" placeholder="Password (if set)" style="min-width:260px">
  </div>
  <div class="row">
    <button onclick="login()">Sign in</button>
    <button onclick="sendOtp()">Send code instead</button>
    <button onclick="showCreate()">Create password</button>
    <button onclick="showReset()">Forgot password</button>
    <span id="msg" class="small muted"></span>
  </div>
</div>

<div id="create" class="card" style="display:none">
  <h3>Create password</h3>
  <div class="small muted">We’ll verify your email first.</div>
  <div class="row" style="margin-top:8px">
    <input id="cEmail" placeholder="you@example.com" style="min-width:260px">
    <input id="cEmployee" placeholder="Your name as in system (for first-time claim)" style="min-width:260px">
    <button onclick="sendOtpCreate()">Send verify code</button>
  </div>
  <div class="row" id="cCodeRow" style="display:none">
    <input id="cCode" placeholder="6-digit code" maxlength="6" style="max-width:140px">
    <input id="cPass1" type="password" placeholder="New password" style="min-width:220px">
    <input id="cPass2" type="password" placeholder="Repeat password" style="min-width:220px">
    <button id="btnSetPwd" onclick="finishCreate()">Set password</button>
  </div>
  <div id="cMsg" class="small muted"></div>
</div>

<div id="reset" class="card" style="display:none">
  <h3>Reset password</h3>
  <div class="row">
    <input id="rEmail" placeholder="you@example.com" style="min-width:260px">
    <button onclick="startReset()">Send reset code</button>
  </div>
  <div class="row" id="rCodeRow" style="display:none">
    <input id="rCode" placeholder="6-digit code" maxlength="6" style="max-width:140px">
    <input id="rPass1" type="password" placeholder="New password" style="min-width:220px">
    <input id="rPass2" type="password" placeholder="Repeat password" style="min-width:220px">
    <button onclick="finishReset()">Reset</button>
  </div>
  <div id="rMsg" class="small muted"></div>
</div>

<div id="app" class="card" style="display:none">
  <div class="small muted" id="who"></div>
  <h3>Current Effective Points</h3>
  <div id="points" style="font-size:28px;font-weight:700;margin-bottom:8px">—</div>
  <h3>Grace Available</h3>
  <div id="grace" style="font-size:22px;margin-bottom:8px">—</div>
  <h3>Work History</h3>
  <div class="small muted" id="summary"></div>
  <table><thead><tr><th>Date</th><th>Event</th><th>Infraction/Notes</th><th>Pts</th><th>Roll</th><th>Lead</th><th>PDF</th></tr></thead>
  <tbody id="rows"><tr><td colspan="7" class="muted">Loading…</td></tr></tbody></table>
</div>

<script>
// ---------- helpers ----------
function $(id){ return document.getElementById(id); }
function safeShow(id){ var el=$(id); if (el) el.style.display='block'; }
function safeHide(id){ var el=$(id); if (el) el.style.display='none'; }
function msg(id, text){ var el=$(id); if (el) el.textContent = text || ''; }
function esc(s){
  return String(s||'').replace(/[&<>"']/g, function(m){
    return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;', "'":'&#39;'}[m]);
  });
}
function link(u,t){ return u ? '<a target="_blank" rel="noopener" href="'+esc(u)+'">'+esc(t||'View')+'</a>' : ''; }

function showCreate(){ safeShow('create'); safeHide('reset'); }
function showReset(){ safeShow('reset'); safeHide('create'); }

// ---------- auth actions ----------
function login(){
  const email = $('email')?.value.trim() || '';
  const pass  = $('password')?.value || '';
  msg('msg','Signing in…');

  google.script.run
    .withSuccessHandler(function(res){
      if(!res || !res.ok){ msg('msg', (res && res.error) || 'Login failed'); return; }
      msg('msg','Loading your data…');
      loadData(res.email);
    })
    .withFailureHandler(function(err){
      msg('msg','Login error: ' + (err && err.message ? err.message : String(err)));
    })
    .loginWithPassword(email, pass);
}

function sendOtp(){
  const email = $('email')?.value.trim() || '';
  msg('msg','Sending code…');
  google.script.run
    .withSuccessHandler(function(r){ msg('msg', r && r.ok ? 'Code sent — check email' : (r && r.error || 'Unable to send code')); })
    .withFailureHandler(function(err){ msg('msg','Error: ' + (err && err.message ? err.message : String(err))); })
    .requestSigninCode(email);
}

function sendOtpCreate(){
  const email = $('cEmail')?.value.trim() || '';
  const emp   = $('cEmployee')?.value.trim() || '';
  msg('cMsg','Sending verify code…');
  google.script.run
    .withSuccessHandler(function(r){
      if(r && r.ok){ safeShow('cCodeRow'); msg('cMsg','Code sent.'); }
      else { msg('cMsg', r && r.error || 'Failed'); }
    })
    .withFailureHandler(function(err){ msg('cMsg','Error: ' + (err && err.message ? err.message : String(err))); })
    .requestCreateCode(email, emp);
}

function finishCreate(){
  const email = $('cEmail')?.value.trim() || '';
  const code  = $('cCode')?.value.trim() || '';
  const p1    = $('cPass1')?.value || '';
  const p2    = $('cPass2')?.value || '';

  if (!email || !code){ msg('cMsg','Enter your email and the 6-digit code'); return; }
  if (p1 !== p2){ msg('cMsg','Passwords do not match'); return; }
  if (p1.length < 8){ msg('cMsg','Password must be at least 8 characters'); return; }

  const btn = $('btnSetPwd'); if (btn) btn.disabled = true;
  msg('cMsg','Setting password…');

  google.script.run
    .withSuccessHandler(function(r){
      if (!r || !r.ok){
        msg('cMsg', (r && r.error) ? String(r.error) : 'Error');
        if (btn) btn.disabled = false;
        return;
      }
      msg('cMsg','✅ Password set. You can sign in above.');
      safeHide('cCodeRow');
      if ($('cPass1')) $('cPass1').value = '';
      if ($('cPass2')) $('cPass2').value = '';
      if (btn) btn.disabled = false;
    })
    .withFailureHandler(function(err){
      msg('cMsg','Error: ' + (err && err.message ? err.message : String(err)));
      if (btn) btn.disabled = false;
    })
    .completeCreate(email, code, p1);
}

function startReset(){
  const email = $('rEmail')?.value.trim() || '';
  msg('rMsg','Sending reset code…');
  google.script.run
    .withSuccessHandler(function(r){
      if(r && r.ok){ safeShow('rCodeRow'); msg('rMsg','Code sent.'); }
      else { msg('rMsg', r && r.error || 'Error'); }
    })
    .withFailureHandler(function(err){ msg('rMsg','Error: ' + (err && err.message ? err.message : String(err))); })
    .requestResetCode(email);
}

function finishReset(){
  const email = $('rEmail')?.value.trim() || '';
  const code  = $('rCode')?.value.trim() || '';
  const p1    = $('rPass1')?.value || '';
  const p2    = $('rPass2')?.value || '';
  if (p1 !== p2){ msg('rMsg','Passwords do not match'); return; }

  google.script.run
    .withSuccessHandler(function(r){
      msg('rMsg', r && r.ok ? 'Password reset — sign in above.' : (r && r.error || 'Error'));
      if (r && r.ok) safeHide('rCodeRow');
    })
    .withFailureHandler(function(err){ msg('rMsg','Error: ' + (err && err.message ? err.message : String(err))); })
    .completeReset(email, code, p1);
}

// ---------- data load ----------
function loadData(email){
  google.script.run
    .withSuccessHandler(function(data){
      safeHide('signin');
      safeShow('app');

      if ($('who')) $('who').textContent =
        'Signed in as ' + email + ' — records for ' + (data && data.employee || '(not found)');
      if ($('points')) $('points').textContent =
        (data && data.effectivePoints != null) ? String(data.effectivePoints) : '—';
      if ($('grace')) $('grace').textContent =
        (data && data.graceAvailableText) ? data.graceAvailableText :
        (data && data.graceAvailable != null ? String(data.graceAvailable) : '—');

      const body = $('rows');
      if (body){
        body.innerHTML = '';
        const rows = (data && data.rows) || [];
        if (!rows.length){
          body.innerHTML = '<tr><td colspan="7" class="muted">No history yet.</td></tr>';
        } else {
          rows.forEach(function(r){
            const tr = document.createElement('tr');
            tr.innerHTML =
              '<td>'+ (r.date||'') +'</td>'+
              '<td>'+ (r.event||'') +'</td>'+
              '<td>'+ (r.infraction||'') + (r.notes?'<div class="small muted">'+esc(r.notes)+'</div>':'') +'</td>'+
              '<td>'+ (r.points==null?'':String(r.points)) +'</td>'+
              '<td>'+ (r.roll==null?'':String(r.roll)) +'</td>'+
              '<td>'+ (r.lead||'') +'</td>'+
              '<td>'+ (r.pdfUrl?('<a target="_blank" rel="noopener" href="'+esc(r.pdfUrl)+'">View PDF</a>'):'') +'</td>';
            body.appendChild(tr);
          });
        }
      }
      if ($('summary')) $('summary').textContent = 'Last updated: ' + (data && data.generatedAt || '');
      msg('msg','');
    })
    .withFailureHandler(function(err){
      msg('msg','Data error: ' + (err && err.message ? err.message : String(err)));
    })
    .getMyOverviewForEmail(email);
}
</script>
  `).setTitle('CLEAR — Sign in');
}


/* ========= Auth: Passwords + OTP ========= */
function loginWithPassword(email, password){
  try{
    email = norm_(email);
    if (!email || !password) return { ok:false, error:'Missing credentials' };
    const row = getOrCreateDirRow_(email);
    if (!row || !row.Email) return { ok:false, error:'Directory unavailable' };
    if (!row.PassHash) return { ok:false, error:'No password set. Use "Create password" or code sign-in.' };
    if (!verifyHash_(password, row.Salt, row.PassHash)) return { ok:false, error:'Incorrect password' };
    markLogin_(email);
    return { ok:true, email };
  }catch(e){
    return { ok:false, error:'Server error: ' + (e && e.message ? e.message : String(e)) };
  }
}

// Map a login email -> the exact Employee name used in Events
function resolveEmployeeName(email){
  email = norm_(email);
  if (!email) return '';

  // Prefer the Directory tab (authoritative mapping)
  var sh  = dir_();                 // ensures the "Directory" sheet exists
  var map = readDirMap_(sh);        // { email -> { Employee, ... } }
  var rec = map[email];
  if (rec && rec.Employee) return String(rec.Employee).trim();

  // Fallback: no mapping yet → no employee name (dashboard can still open)
  return '';
}


function markLogin_(email){
  try{
    email = norm_(email);
    const sh  = dir_();
    const map = readDirMap_(sh);
    const r   = map[email];
    if (!r) return;
    writeDirFields_(sh, r._row, { LastLogin: now_() });
  }catch(_){ /* no-op */ }
}


function requestSigninCode(email){
  email = norm_(email); if (!email) return {ok:false};
  const row = getOrCreateDirRow_(email);
  if (!row || !row.Email) return {ok:false, error:'Not in Directory. Ask a director to invite you or use Create password.'};
  const code=('000000'+Math.floor(Math.random()*1e6)).slice(-6);
  CacheService.getScriptCache().put('otp:'+email, code, AUTH.OTP_TTL_MIN*60);
  try { MailApp.sendEmail(email, 'Your CLEAR code', 'Code: '+code+' (expires in '+AUTH.OTP_TTL_MIN+' minutes)'); } catch(_){ return {ok:false,error:'Mail failed'} }
  return {ok:true};
}

/* Create password (verify email first). If not in Directory, create Pending row. */
function requestCreateCode(email, employee){
  email=norm_(email); if(!email) return {ok:false,error:'Enter email'};
  const exists = !!findDirRowIndex_(email);
  if (!exists && !employee) return {ok:false,error:'Enter your name for Pending record'};
  const code=('000000'+Math.floor(Math.random()*1e6)).slice(-6);
  CacheService.getScriptCache().put('create:'+email, JSON.stringify({code:code, employee:employee||''}), AUTH.OTP_TTL_MIN*60);
  try { MailApp.sendEmail(email,'Verify your email','Your verification code: '+code+' (expires in '+AUTH.OTP_TTL_MIN+' minutes)'); } catch(_){ return {ok:false,error:'Mail failed'} }
  return {ok:true};
}
function completeCreate(email, code, newPass){
  try{
    email = String(email||'').trim().toLowerCase();
    if (!newPass || newPass.length < 8) return { ok:false, error:'Password must be at least 8 characters' };

    const payload = CacheService.getScriptCache().get('create:'+email);
    if (!payload) return { ok:false, error:'Code expired or not requested' };
    const obj = JSON.parse(payload || '{}');
    if (String(obj.code) !== String(code)) return { ok:false, error:'Invalid code' };

    const sh = dir_(); if (!sh) return { ok:false, error:'Directory sheet missing' };
    const map = readDirMap_(sh);
    let r = map[email];
    if (!r){
      sh.appendRow([email, obj.employee||'', 'Employee', true, '', '', now_(), now_(), '' ]);
      r = readDirMap_(sh)[email];
    }

    const salt = Utilities.getUuid();
    const hash = makeHash_(newPass, salt);
    writeDirFields_(sh, r._row, { Salt:salt, PassHash:hash, Verified:true, UpdatedAt:now_() });

    return { ok:true };
  }catch(e){
    return { ok:false, error:'Server error: ' + (e && e.message ? e.message : String(e)) };
  }
}


/* Reset password — email code */
function requestResetCode(email){
  email=norm_(email); if(!email) return {ok:false,error:'Enter email'};
  const r = getOrCreateDirRow_(email); if (!r || !r.Email) return {ok:false,error:'Not in Directory'};
  const code=('000000'+Math.floor(Math.random()*1e6)).slice(-6);
  CacheService.getScriptCache().put('reset:'+email, code, AUTH.RESET_TTL_MIN*60);
  try { MailApp.sendEmail(email,'CLEAR reset code','Reset code: '+code+' (expires in '+AUTH.RESET_TTL_MIN+' minutes)'); } catch(_){ return {ok:false,error:'Mail failed'} }
  return {ok:true};
}
function completeReset(email, code, newPass){
  email=norm_(email); const cached=CacheService.getScriptCache().get('reset:'+email);
  if (!cached || String(cached)!==String(code||'')) return {ok:false,error:'Invalid/expired code'};
  const sh=dir_(); const map=readDirMap_(sh); const r=map[email]; if(!r) return {ok:false,error:'Not found'};
  const salt=Utilities.getUuid(); const hash=makeHash_(newPass, salt);
  writeDirFields_(sh, r._row, { Salt:salt, PassHash:hash, UpdatedAt:now_() });
  return {ok:true};
}

/* ===== Invites (directors only) ===== */
function sendInvite(email, employee, role){
  email=norm_(email); if(!isDirector_(getCaller_())) return {ok:false,error:'Directors only'};
  role = role||'Employee';
  const sh=dir_(); const map=readDirMap_(sh);
  if (!map[email]){
    sh.appendRow([email, employee||'', role, false, '', '', now_(), now_(), '']);
  } else {
    writeDirFields_(sh, map[email]._row, { Employee:(employee||map[email].Employee), Role:role, Verified:false, UpdatedAt:now_() });
  }
  // notify user to create password
  try{ MailApp.sendEmail(email, 'You have been invited to CLEAR', 'Visit your CLEAR link and choose "Create password".'); } catch(_){}
  return {ok:true};
}

/* ===== Self-claim request (optional) ===== */
function requestAccess(email, employee){
  email=norm_(email); if(!email||!employee) return {ok:false,error:'Enter email & name'};
  const sh=req_(); if (sh.getLastRow()===0) sh.appendRow(REQ_COLS);
  sh.appendRow([now_(), email, employee, 'NEW']);
  // notify directors
  try{ MailApp.sendEmail(AUTH.DIRECTORS.join(','), 'CLEAR access request', email+' requests access as '+employee); }catch(_){}
  return {ok:true};
}

/* ========= Data (history) — same as earlier POC, trimmed ========= */
function getMyOverviewForEmail(email){
  const employee = resolveEmployeeName(email);
  const out = { email, employee, effectivePoints:null, graceAvailable:null, graceAvailableText:null, rows:[], generatedAt:now_() };
  if (!employee) return out;
  return getMyOverviewForEmployee_(employee);
}

function getMyOverviewForEmployee_(employee){
  const s=SpreadsheetApp.getActive().getSheetByName(TABS.EVENTS); if(!s) return {employee,rows:[],generatedAt:now_()};
  const hdr = s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const idx = n => hdr.indexOf(n);
  const iEmp=idx('Employee'), iDate=firstIdx_(hdr,['IncidentDate','Date','Timestamp']);
  const iEvt=idx('EventType'), iInf=idx('Infraction'), iLead=idx('Lead');
  const iPts=idx('Points'), iRollE=idx('PointsRolling (Effective)'), iRoll=idx('PointsRolling');
  const iNotes=firstIdx_(hdr,['Notes / Reviewer','Notes','Reviewer','Notes/Reviewer']);
  const iPdf=firstIdx_(hdr,['Write-Up PDF','PDF Link','WriteUpPDF','Signed_PDF_Link']);
  const tz=Session.getScriptTimeZone()||'UTC';

  const vals=s.getDataRange().getValues(); const rows=[]; let lastEff=null;
  for (let r=1;r<vals.length;r++){
    if (String(vals[r][iEmp]||'').trim().toLowerCase()!==String(employee).trim().toLowerCase()) continue;
    let roll=null; if(iRollE!==-1&&isFinite(Number(vals[r][iRollE]))) roll=Number(vals[r][iRollE]); else if(iRoll!==-1&&isFinite(Number(vals[r][iRoll]))) roll=Number(vals[r][iRoll]);
    if (roll!=null) lastEff=roll;
    rows.push({
      date:fmtDate_(vals[r][iDate],tz),
      event:String(vals[r][iEvt]||''),
      infraction:String(vals[r][iInf]||''),
      notes:(iNotes!==-1)?String(vals[r][iNotes]||''):'',
      points:(iPts!==-1 && vals[r][iPts]!=='' && vals[r][iPts]!=null)?Number(vals[r][iPts]):null,
      roll:roll,
      lead:String(vals[r][iLead]||''),
      pdfUrl:(iPdf!==-1)?String(vals[r][iPdf]||''):''
    });
  }
  return { employee, rows, effectivePoints:(lastEff!=null)?lastEff:null, graceAvailable:null, graceAvailableText:null, generatedAt:now_() };
}

/* ========= Directory helpers ========= */
function dir_(){ const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName(TABS.DIRECTORY); if(!sh){ sh=ss.insertSheet(TABS.DIRECTORY); sh.appendRow(DIR_COLS); } return sh; }
function req_(){ const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName(TABS.REQUESTS); if(!sh){ sh=ss.insertSheet(TABS.REQUESTS); sh.appendRow(REQ_COLS); } return sh; }
function readDirMap_(sh){
  const vals=sh.getDataRange().getValues(); const h=vals[0]; const m={};
  const idx={}; DIR_COLS.forEach((n,i)=>idx[n]=h.indexOf(n));
  for(let r=1;r<vals.length;r++){
    const row=vals[r]; const email=norm_(row[idx.Email]);
    if (!email) continue;
    m[email]={ _row:r+1, Email:email, Employee:row[idx.Employee]||'', Role:row[idx.Role]||'Employee',
      Verified:Boolean(row[idx.Verified]===true || String(row[idx.Verified]).toLowerCase()==='true'),
      Salt:String(row[idx.Salt]||''), PassHash:String(row[idx.PassHash]||'')
    };
  }
  return m;
}
function writeDirFields_(sh,row,patch){
  const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; const map={}; h.forEach((n,i)=>map[n]=i+1);
  Object.keys(patch).forEach(k=>{ if(map[k]) sh.getRange(row,map[k]).setValue(patch[k]); });
}
function findDirRowIndex_(email){ const sh=dir_(); const vals=sh.getRange(1,1,sh.getLastRow(),1).getValues(); for(let r=2;r<=vals.length;r++){ if(String(vals[r-1][0]||'').trim().toLowerCase()===email) return r; } return 0; }
function getOrCreateDirRow_(email){
  const sh=dir_(); const map=readDirMap_(sh); let r=map[email]; if(r) return r;
  // create a bare pending row so OTP can work; directors can fill Employee later
  sh.appendRow([email,'','Employee',false,'','',now_(),now_(),'']);
  return readDirMap_(sh)[email];
}

/* ========= Crypto / utils ========= */
function makeHash_(password, salt){
  const value = String(salt || '') + String(password || '');
  const sigBytes = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256, // algorithm
    value,                                // message
    AUTH.PEPPER                           // secret key (pepper)
  );
  return Utilities.base64Encode(sigBytes); // store this
}

function verifyHash_(password, salt, hash){
  return makeHash_(password, salt) === String(hash || '');
}
function isDirector_(email){ return (AUTH.DIRECTORS||[]).map(norm_).indexOf(norm_(email))!==-1; }
function getCaller_(){ return String(Session.getActiveUser().getEmail()||'').toLowerCase(); }
function now_(){ const tz=Session.getScriptTimeZone()||'UTC'; return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'); }
function norm_(s){ return String(s||'').trim().toLowerCase(); }
function firstIdx_(hdr, names){ for (var i=0;i<names.length;i++){ const k=hdr.indexOf(names[i]); if(k!==-1) return k; } return -1; }
function fmtDate_(v,tz){ try{ if(v instanceof Date && !isNaN(v)) return Utilities.formatDate(v,tz,'yyyy-MM-dd'); const d=new Date(v); if(!isNaN(d)) return Utilities.formatDate(d,tz,'yyyy-MM-dd'); }catch(_){}
  const s=String(v||''); return s.length>=10? s.slice(0,10) : s; }

/* ========= OPTIONAL: directors-only helpers you can call from the Script editor ========= */
// Send an invite (pre-seed/overwrite Directory row, email the user to create password)
function adminInvite(email, employee, role){ return sendInvite(email, employee, role); }
// Approve a pending self-claim row (set Verified=TRUE, set Employee/Role)
function adminApprove(email, employee, role){
  email=norm_(email); if(!isDirector_(getCaller_())) return {ok:false,error:'Directors only'};
  const sh=dir_(); const map=readDirMap_(sh); const r=map[email]; if(!r) return {ok:false,error:'Not found'};
  writeDirFields_(sh, r._row, { Verified:true, Employee:employee||r.Employee, Role:role||r.Role, UpdatedAt:now_() });
  try{ MailApp.sendEmail(email,'Your CLEAR access is approved','You can sign in with your password or request a code.'); }catch(_){}
  return {ok:true};
}
