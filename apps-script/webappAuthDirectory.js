/** CLEAR Web App (Auth + Directory) — single-file drop-in
 * Features: Email OTP, Password login, Reset password, Invite flow, Pending self-claim
 * Storage: Directory sheet and CacheService for short-lived tokens
 * Security: Passwords hashed = base64(HMAC-SHA256( (salt + password), PEPPER ))
 */

/* ========= CONFIG (EDIT) ========= */
const AUTH = {
  PEPPER: '8f2h9k4m7p3q5w6e8r9t2y4u8i3o5p7a9s2d5f8g3h7j4k6l8z9x2c5v8b3n7m', // ⚠️ CHANGE THIS TO A SECURE RANDOM STRING
  OTP_TTL_MIN: 10,
  RESET_TTL_MIN: 15,
  SESSION_TTL_MIN: 60*24, // 24 hours
};

const TABS = { DIRECTORY: 'Directory', REQUESTS: 'AccessRequests', EVENTS: 'Events' };
const DIR_COLS = ['Email','Employee','Role','Verified','Salt','PassHash','CreatedAt','UpdatedAt','LastLogin'];
const REQ_COLS = ['RequestedAt','Email','Employee','Status'];

// Add this missing function
function now_(){
  const tz = Session.getScriptTimeZone() || 'UTC';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
}

// Add these other missing utility functions
function norm_(s){ 
  return String(s || '').trim().toLowerCase(); 
}

function firstIdx_(hdr, names){ 
  for (var i=0;i<names.length;i++){ 
    const k=hdr.indexOf(names[i]); 
    if(k!==-1) return k; 
  } 
  return -1; 
}

function getCaller_(){
  return String(Session.getActiveUser().getEmail() || '').toLowerCase();
}

/* ========= UI ========= */
function doGet(e){
  try {
    console.log('🚀 doGet called with parameters:', e);

    const params = (e && e.parameter) || {};
    const page = params.page;
    const sessionId = params.session;
    console.log('🎫 Session ID:', sessionId);

    const user = sessionId ? getUserFromSession(sessionId) : null;
    console.log('👤 User from session:', user);

    const pageMap = {
      login: renderLoginPage,
      director: () => renderDirectorDashboard(user, sessionId),
      lead: () => renderLeadDashboard(user, sessionId),
      employee: () => renderEmployeeDashboard(user, sessionId),
      employeeSearch: () => renderEmployeeSearchPage(user, sessionId)
    };

    if (!user && page !== 'login') {
      console.log('❌ No user found, rendering login page');
      return renderLoginPage();
    }

    if (page && pageMap[page]) {
      console.log('📄 Rendering page:', page);
      return pageMap[page]();
    }

    if (!user) {
      console.log('❌ No user after page check, rendering login');
      return renderLoginPage();
    }

    console.log('🎯 User role:', user.role);

    // Route based on user role
    switch(user.role) {
      case 'director':
        console.log('🎬 Rendering director dashboard');
        return renderDirectorDashboard(user, sessionId);
      case 'lead':
        console.log('👥 Rendering lead dashboard');
        return renderLeadDashboard(user, sessionId);
      case 'employee':
      default:
        console.log('👷 Rendering employee dashboard');
        return renderEmployeeDashboard(user, sessionId);
    }
  } catch (error) {
    console.log('💥 Error in doGet:', error);
    Logger.log('Error in doGet: ' + error);
    return HtmlService.createHtmlOutput('<h1>Error</h1><p>' + error + '</p>');
  }
}

function renderLoginPage() {
  return HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <title>CLEAR — Login</title>
  <meta charset="UTF-8">
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; }
    .container { max-width: 400px; margin: 0 auto; }
    h1 { text-align: center; color: #333; }
    .form-group { margin: 15px 0; }
    label { display: block; margin-bottom: 5px; }
    input { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; }
    button { width: 100%; padding: 10px; background: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer; }
    button:hover { background: #0056b3; }
    button:disabled { background: #ccc; cursor: not-allowed; }
    .msg { margin: 10px 0; padding: 10px; border-radius: 4px; }
    .error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
    .success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
  </style>
</head>
<body>
  <div class="container">
    <h1>CLEAR Login</h1>
    
    <div class="form-group">
      <label for="email">Email:</label>
      <input type="email" id="email" required>
    </div>
    
    <div class="form-group">
      <label for="password">Password:</label>
      <input type="password" id="password" required>
    </div>
    
    <button id="loginBtn" onclick="login()">Sign In</button>
    
    <div id="msg"></div>
  </div>

  <script>
    function $(id) { return document.getElementById(id); }
    
    function login() {
      var email = $('email').value;
      var password = $('password').value;
      
      if (!email || !password) {
        showMsg('Please enter email and password', 'error');
        return;
      }
      
      var btn = $('loginBtn');
      btn.disabled = true;
      btn.textContent = 'Signing in...';
      
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.ok) {
            showMsg('Login successful! Redirecting...', 'success');
            const base = window.location.href.split('?')[0];
            google.script.run
              .withSuccessHandler(() => {
                google.script.host.close();
                window.open(base + '?session=' + result.sessionId, '_top');
              })
              .closeAndReopenWithSession(result.sessionId);
          } else {
            showMsg(result.error || 'Login failed', 'error');
            btn.disabled = false;
            btn.textContent = 'Sign In';
          }
        })
        .withFailureHandler(function(error) {
          showMsg('Login error: ' + error.message, 'error');
          btn.disabled = false;
          btn.textContent = 'Sign In';
        })
        .loginWithPassword(email, password);
    }
    
    function showMsg(text, type) {
      var msgDiv = $('msg');
      msgDiv.textContent = text;
      msgDiv.className = 'msg ' + (type || 'error');
    }
  </script>
</body>
</html>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function closeAndReopenWithSession(sessionId) {
  // This function doesn't need to do anything - it's just called to ensure
  // the google.script.run pipeline is properly closed before we close the host
  console.log('📤 closeAndReopenWithSession called with:', sessionId);
  return true;
}

// Add this function to handle different dashboard types
function renderDirectorDashboard(user, sessionId) {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>CLEAR — Director Dashboard</title>
      <style>
        .director-nav { display: flex; gap: 10px; margin-bottom: 20px; }
        .director-nav button { padding: 10px 15px; background: #2563eb; color: white; border: none; border-radius: 5px; cursor: pointer; }
        .director-nav button:hover { background: #1d4ed8; }
        .content-area { padding: 20px; border: 1px solid #e5e7eb; border-radius: 8px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 20px; }
        .stat-card { background: #f8fafc; padding: 15px; border-radius: 8px; text-align: center; }
        .stat-number { font-size: 24px; font-weight: bold; color: #dc2626; }
        .pending-milestones { background: #fef3c7; border-left: 4px solid #f59e0b; padding: 10px; margin: 10px 0; }
        .logout-btn { background: #dc2626; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; float: right; margin-bottom: 10px; }
        .logout-btn:hover { background: #b91c1c; }
      </style>
    </head>
    <body>
      <button class="logout-btn" id="logoutBtn">Logout</button>
      <h1>Director Dashboard</h1>
      
      <nav class="director-nav">
        <button id="employeeSearchBtn">Employee Search</button>
        <button id="pendingMilestonesBtn">Pending Milestones</button>
        <button id="graceRequestsBtn">Grace Requests</button>
        <button id="reportsBtn">Reports</button>
        <button id="bulkOperationsBtn">Bulk Operations</button>
      </nav>
      
      <div id="content-area">
        <div class="stats-grid">
          <div class="stat-card">
            <h3>Pending Milestones</h3>
            <div class="stat-number" id="pendingCount">—</div>
          </div>
          <div class="stat-card">
            <h3>Active Probation</h3>
            <div class="stat-number" id="probationCount">—</div>
          </div>
          <div class="stat-card">
            <h3>Grace Requests</h3>
            <div class="stat-number" id="graceCount">—</div>
          </div>
          <div class="stat-card">
            <h3>This Month's Events</h3>
            <div class="stat-number" id="eventsCount">—</div>
          </div>
        </div>
        
        <div id="pending-milestones-list">
          <!-- Dynamic content loaded here -->
        </div>
      </div>
      
      <script>
        const sessionId = ${JSON.stringify(sessionId)};

        // ---------- Helper Functions ----------
        function $(id) { return document.getElementById(id); }

        // ---------- Event Listeners Setup ----------
        document.addEventListener('DOMContentLoaded', function() {
          console.log('🎯 DIRECTOR DASHBOARD LOADED');
          
          // Attach button event listeners
          $('logoutBtn').addEventListener('click', logout);
          $('employeeSearchBtn').addEventListener('click', showEmployeeSearch);
          $('pendingMilestonesBtn').addEventListener('click', showPendingMilestones);
          $('graceRequestsBtn').addEventListener('click', showGraceRequests);
          $('reportsBtn').addEventListener('click', showReports);
          $('bulkOperationsBtn').addEventListener('click', showBulkOperations);
          
          // Load dashboard data
          loadDashboardData();
        });
        
        function logout() {
          console.log('Director logout clicked');
          window.open('?page=login', '_top');
        }
        
        // ---------- Dashboard Functions ----------
        function loadDashboardData() {
          console.log('Loading director dashboard data...');
          
          google.script.run
            .withSuccessHandler(function(data) {
              console.log('✅ Director dashboard data received:', data);
              
              $('pendingCount').textContent = data.pendingMilestones || 0;
              $('probationCount').textContent = data.activeProbation || 0;
              $('graceCount').textContent = data.graceRequests || 0;
              $('eventsCount').textContent = data.monthlyEvents || 0;
              
              // Load pending milestones
              loadPendingMilestones();
            })
            .withFailureHandler(function(err) {
              console.error('❌ Error loading director dashboard data:', err);
              $('pendingCount').textContent = 'Error';
              $('probationCount').textContent = 'Error';
              $('graceCount').textContent = 'Error';
              $('eventsCount').textContent = 'Error';
            })
            .getDirectorDashboardData();
        }
          
        function loadPendingMilestones() {
          console.log('Loading pending milestones...');
          
          google.script.run
            .withSuccessHandler(function(milestones) {
              const container = $('pending-milestones-list');
              if (milestones.length === 0) {
                container.innerHTML = '<p>No pending milestones.</p>';
                return;
              }
              
              let html = '<h3>Pending Milestones Requiring Attention</h3>';
              milestones.forEach(function(milestone) {
                html += '<div class="pending-milestones">' +
                  '<strong>' + milestone.employee + '</strong> - ' + milestone.milestone + 
                  ' (Row: ' + milestone.row + ')' +
                  '<button onclick="assignDirector(' + milestone.row + ')">Assign Director</button>' +
                  '<button onclick="viewDetails(' + milestone.row + ')">View Details</button>' +
                  '</div>';
              });
              container.innerHTML = html;
            })
            .withFailureHandler(function(err) {
              console.error('❌ Error loading pending milestones:', err);
              $('pending-milestones-list').innerHTML = '<p>Error loading milestones</p>';
            })
            .getPendingMilestones();
        }
        
        function assignDirector(row) {
          const director = prompt('Enter director name:');
          if (director) {
            google.script.run
              .withSuccessHandler(function() { loadPendingMilestones(); })
              .assignMilestoneDirector(row, director);
          }
        }
        
        function showEmployeeSearch() {
          window.open('?session=' + encodeURIComponent(sessionId) + '&page=employeeSearch', '_top');
        }
        
        function showPendingMilestones() {
          console.log('Showing pending milestones');
          $('content-area').innerHTML = '<h3>Pending Milestones</h3><div id="milestones-container">Loading...</div>';
          loadPendingMilestones();
        }
        
        function showGraceRequests() {
          console.log('Showing grace requests');
          $('content-area').innerHTML = '<h3>Grace Requests</h3><div id="grace-container">Coming soon...</div>';
        }
        
        function showReports() {
          console.log('Showing reports');
          $('content-area').innerHTML = '<h3>Reports</h3><button onclick="generateMonthlyReport()">Generate Monthly Report</button><div id="report-container"></div>';
        }
        
        function showBulkOperations() {
          console.log('Showing bulk operations');
          $('content-area').innerHTML = '<h3>Bulk Operations</h3><div id="bulk-container">Coming soon...</div>';
        }
      </script>
    </body>
    </html>
  `)
    .setTitle('CLEAR — Director Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// For directors
function renderLeadDashboard(user, sessionId) {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>CLEAR — Lead Dashboard</title>
      <style>
        .lead-nav { display: flex; gap: 10px; margin-bottom: 20px; }
        .lead-nav button { padding: 10px 15px; background: #2563eb; color: white; border: none; border-radius: 5px; cursor: pointer; }
        .lead-nav button:hover { background: #1d4ed8; }
        .content-area { padding: 20px; border: 1px solid #e5e7eb; border-radius: 8px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 20px; }
        .stat-card { background: #f8fafc; padding: 15px; border-radius: 8px; text-align: center; }
        .stat-number { font-size: 24px; font-weight: bold; color: #dc2626; }
        .pending-milestones { background: #fef3c7; border-left: 4px solid #f59e0b; padding: 10px; margin: 10px 0; }
      </style>
    </head>
    <body>
      <h1>Lead Dashboard</h1>
      
      <nav class="lead-nav">
        <button onclick="showEmployeeSearch()">Employee Search</button>
        <button onclick="showPendingMilestones()">Pending Milestones</button>
        <button onclick="showReports()">Reports</button>
        <button onclick="showGraceRequests()">Grace Requests</button>
      </nav>
      
      <div id="content-area">
        <div class="stats-grid">
          <div class="stat-card">
            <h3>Pending Milestones</h3>
            <div class="stat-number" id="pendingCount">—</div>
          </div>
          <div class="stat-card">
            <h3>Active Probation</h3>
            <div class="stat-number" id="probationCount">—</div>
          </div>
          <div class="stat-card">
            <h3>Grace Requests</h3>
            <div class="stat-number" id="graceCount">—</div>
          </div>
          <div class="stat-card">
            <h3>This Month's Events</h3>
            <div class="stat-number" id="eventsCount">—</div>
          </div>
        </div>
        
        <div id="pending-milestones-list">
          <!-- Dynamic content loaded here -->
        </div>
      </div>
      
      <script>
        const sessionId = ${JSON.stringify(sessionId)};

        // Load dashboard data
        google.script.run
          .withSuccessHandler(function(data) {
            document.getElementById('pendingCount').textContent = data.pendingMilestones || 0;
            document.getElementById('probationCount').textContent = data.activeProbation || 0;
            document.getElementById('graceCount').textContent = data.graceRequests || 0;
            document.getElementById('eventsCount').textContent = data.monthlyEvents || 0;
            
            // Load pending milestones
            loadPendingMilestones();
          })
          .getDirectorDashboardData();
          
        function loadPendingMilestones() {
          google.script.run
            .withSuccessHandler(function(milestones) {
              const container = document.getElementById('pending-milestones-list');
              if (milestones.length === 0) {
                container.innerHTML = '<p>No pending milestones.</p>';
                return;
              }
              
              let html = '<h3>Pending Milestones Requiring Attention</h3>';
              milestones.forEach(function(milestone) {
                html += '<div class="pending-milestones">' +
                  '<strong>' + milestone.employee + '</strong> - ' + milestone.milestone + 
                  ' (Row: ' + milestone.row + ')' +
                  '<button onclick="assignDirector(' + milestone.row + ')">Assign Director</button>' +
                  '<button onclick="viewDetails(' + milestone.row + ')">View Details</button>' +
                  '</div>';
              });
              container.innerHTML = html;
            })
            .getPendingMilestones();
        }
        
        function assignDirector(row) {
          const director = prompt('Enter director name:');
          if (director) {
            google.script.run
              .withSuccessHandler(function() { loadPendingMilestones(); })
              .assignMilestoneDirector(row, director);
          }
        }
        
        function showEmployeeSearch() {
          window.open('?session=' + encodeURIComponent(sessionId) + '&page=employeeSearch', '_top');
        }
        
        function showPendingMilestones() {
          document.getElementById('content-area').innerHTML = '<h3>Pending Milestones</h3><div id="milestones-container">Loading...</div>';
          loadPendingMilestones();
        }
        
        function showReports() {
          document.getElementById('content-area').innerHTML = '<h3>Reports</h3><button onclick="generateMonthlyReport()">Generate Monthly Report</button><div id="report-container"></div>';
        }
      </script>
    </body>
    </html>
  `)
    .setTitle('CLEAR — Lead Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function renderEmployeeDashboard(user, sessionId) {
  return HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <title>CLEAR — Dashboard</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; }
    .card { border: 1px solid #ddd; padding: 20px; margin: 10px 0; border-radius: 8px; }
    .logout-btn { background: #dc2626; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; float: right; }
  </style>
</head>
<body>
  <div class="card">
    <button class="logout-btn" onclick="logout()">Logout</button>
    <h1>My CLEAR Dashboard</h1>
    <p>Welcome, ` + (user.employee || user.email) + `!</p>
    <p>Your role: ` + (user.role || 'employee') + `</p>
  </div>

  <div class="card">
    <h2>Current Status</h2>
    <div id="points">Loading points...</div>
    <div id="grace">Loading grace status...</div>
  </div>

  <div class="card">
    <h2>Your History</h2>
    <div id="history">Loading history...</div>
  </div>

  <script>
    const sessionId = ${JSON.stringify(sessionId)};
    function logout() {
      window.open('?page=login', '_top');
    }
    
    // Load dashboard data
    google.script.run
      .withSuccessHandler(function(data) {
        document.getElementById('points').textContent = 'Points: ' + (data.effectivePoints || '—');
        document.getElementById('grace').textContent = 'Grace: ' + (data.graceAvailableText || '—');
        
        if (data.rows && data.rows.length > 0) {
          var html = '<table border="1"><tr><th>Date</th><th>Event</th><th>Points</th></tr>';
          data.rows.slice(0, 5).forEach(function(row) {
            html += '<tr><td>' + (row.date || '') + '</td><td>' + (row.event || '') + '</td><td>' + (row.points || '') + '</td></tr>';
          });
          html += '</table>';
          document.getElementById('history').innerHTML = html;
        } else {
          document.getElementById('history').textContent = 'No history found.';
        }
      })
      .withFailureHandler(function(err) {
        document.getElementById('points').textContent = 'Error loading data';
        document.getElementById('history').textContent = 'Error: ' + err.message;
      })
      .getMyOverviewForEmail('` + (user.email || '') + `');
  </script>
</body>
</html>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function renderEmployeeSearchPage(user, sessionId) {
  return HtmlService.createHtmlOutputFromFile('employeeLookup')
    .setTitle('CLEAR — Employee Search')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}


/* ========= Auth: Passwords + OTP ========= */
function loginWithPassword(email, password){
  try{
    console.log('🔐 loginWithPassword called for:', email);
    email = norm_(email);
    console.log('📧 Normalized email:', email);
    
    if (!email || !password) {
      console.log('❌ Missing credentials');
      return { ok:false, error:'Missing credentials' };
    }
    
    const row = getOrCreateDirRow_(email);
    console.log('📋 Directory row found:', row);
    
    if (!row || !row.Email) {
      console.log('❌ Directory unavailable');
      return { ok:false, error:'Directory unavailable' };
    }
    if (!row.PassHash) {
      console.log('❌ No password set');
      return { ok:false, error:'No password set. Use "Create password" or code sign-in.' };
    }
    if (!verifyHash_(password, row.Salt, row.PassHash)) {
      console.log('❌ Incorrect password');
      return { ok:false, error:'Incorrect password' };
    }
    
    markLogin_(email);
    console.log('✅ Authentication successful');
    
    // After successful authentication
    console.log('🎯 Calling getUserRole...');
    const role = getUserRole(email);
    console.log('🎭 Role from getUserRole:', role);
    
    console.log('📝 Calling createUserSession...');
    const sessionId = createUserSession(email); // Create session
    console.log('🆔 Session created:', sessionId);
    
    const employee = resolveEmployeeName(email);
    console.log('👤 Employee name:', employee);
    
    const result = { 
      ok: true, 
      email: email,
      role: role,
      employee: employee,
      sessionId: sessionId // Return session ID
    };
    
    console.log('✅ loginWithPassword result:', result);
    return result;
  }catch(e){
    console.log('💥 loginWithPassword error:', e);
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
  email=norm_(email); if(!isCallerDirector()) return {ok:false,error:'Directors only'};
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

function isCallerDirector() {
  const callerEmail = getCaller_();
  return getUserRole(callerEmail) === 'director';
}

function getDirectorEmails() {
  try {
    const sh = dir_();
    const map = readDirMap_(sh);
    const directors = [];
    
    for (const email in map) {
      const user = map[email];
      if (user && String(user.Role).toLowerCase().trim() === 'director') {
        directors.push(email);
      }
    }
    
    return directors;
  } catch (e) {
    Logger.log('getDirectorEmails error: ' + e);
    return [];
  }
}

/* ===== Self-claim request (optional) ===== */
function requestAccess(email, employee){
  email=norm_(email); if(!email||!employee) return {ok:false,error:'Enter email & name'};
  const sh=req_(); if (sh.getLastRow()===0) sh.appendRow(REQ_COLS);
  sh.appendRow([now_(), email, employee, 'NEW']);
  // notify directors
  const directorEmails = getDirectorEmails();
  if (directorEmails.length > 0) {
    try{ MailApp.sendEmail(directorEmails.join(','), 'CLEAR access request', email+' requests access as '+employee); }catch(_){}
  }
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
function norm_(s){ return String(s||'').trim().toLowerCase(); }
function firstIdx_(hdr, names){ for (var i=0;i<names.length;i++){ const k=hdr.indexOf(names[i]); if(k!==-1) return k; } return -1; }
function fmtDate_(v,tz){ try{ if(v instanceof Date && !isNaN(v)) return Utilities.formatDate(v,tz,'yyyy-MM-dd'); const d=new Date(v); if(!isNaN(d)) return Utilities.formatDate(d,tz,'yyyy-MM-dd'); }catch(_){}
  const s=String(v||''); return s.length>=10? s.slice(0,10) : s; }

/* ========= OPTIONAL: directors-only helpers you can call from the Script editor ========= */
// Send an invite (pre-seed/overwrite Directory row, email the user to create password)
function adminInvite(email, employee, role){ return sendInvite(email, employee, role); }
// Approve a pending self-claim row (set Verified=TRUE, set Employee/Role)
function adminApprove(email, employee, role){
  email=norm_(email); if(!isCallerDirector()) return {ok:false,error:'Directors only'};
  const sh=dir_(); const map=readDirMap_(sh); const r=map[email]; if(!r) return {ok:false,error:'Not found'};
  writeDirFields_(sh, r._row, { Verified:true, Employee:employee||r.Employee, Role:role||r.Role, UpdatedAt:now_() });
  try{ MailApp.sendEmail(email,'Your CLEAR access is approved','You can sign in with your password or request a code.'); }catch(_){}
  return {ok:true};
}

function getUserRole(email) {
  email = norm_(email);
  console.log('🔍 getUserRole called for email:', email);
  
  // Only check Directory sheet roles
  try {
    const sh = dir_();
    const map = readDirMap_(sh);
    const user = map[email];
    
    console.log('📋 User from Directory:', user);
    
    if (user && user.Role) {
      const rawRole = String(user.Role);
      const role = rawRole.toLowerCase().trim();
      console.log('🎭 Raw role from sheet:', rawRole, '| Normalized role:', role);
      
      if (role === 'director' || role === 'admin') {
        console.log('✅ Returning director role');
        return 'director';
      }
      if (role === 'lead' || role === 'supervisor') {
        console.log('✅ Returning lead role');
        return 'lead';
      }
      console.log('✅ Returning employee role (default)');
    } else {
      console.log('❌ No role found for user, using default employee');
    }
  } catch (e) {
    console.log('❌ getUserRole error:', e);
    Logger.log('getUserRole error: ' + e);
  }
  
  return 'employee'; // default role
}

function getCaller_(){
  return String(Session.getActiveUser().getEmail() || '').toLowerCase();
}

function getDirectorDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventsSheet = ss.getSheetByName('Events');
    
    // Get pending milestones count
    const pendingMilestones = getPendingMilestonesCount();
    
    // Get active probation count
    const activeProbation = getActiveProbationCount();
    
    // Get grace requests
    const graceRequests = getPendingGraceRequestsCount();
    
    // Get monthly events
    const monthlyEvents = getMonthlyEventsCount();
    
    return {
      pendingMilestones: pendingMilestones,
      activeProbation: activeProbation,
      graceRequests: graceRequests,
      monthlyEvents: monthlyEvents
    };
  } catch (e) {
    Logger.log('getDirectorDashboardData error: ' + e);
    return { error: 'Failed to load dashboard data' };
  }
}

function getPendingMilestones() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventsSheet = ss.getSheetByName('Events');
    const data = eventsSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column indices
    const employeeCol = headers.indexOf('Employee');
    const milestoneCol = headers.indexOf('Milestone');
    const pendingCol = headers.indexOf('Pending Status');
    const directorCol = headers.indexOf('Consequence Director');
    
    const milestones = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[milestoneCol] && 
          String(row[pendingCol]).toLowerCase() === 'pending' && 
          !row[directorCol]) {
        milestones.push({
          row: i + 1,
          employee: row[employeeCol],
          milestone: row[milestoneCol],
          pendingStatus: row[pendingCol]
        });
      }
    }
    
    return milestones;
  } catch (e) {
    Logger.log('getPendingMilestones error: ' + e);
    return [];
  }
}

function assignMilestoneDirector(row, directorName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventsSheet = ss.getSheetByName('Events');
    const headers = eventsSheet.getRange(1, 1, 1, eventsSheet.getLastColumn()).getValues()[0];
    
    const directorCol = headers.indexOf('Consequence Director');
    if (directorCol === -1) {
      throw new Error('Consequence Director column not found');
    }
    
    eventsSheet.getRange(row, directorCol + 1).setValue(directorName);
    
    // Log the assignment
    Logger.log('Assigned director ' + directorName + ' to milestone row ' + row);
    
    return { success: true };
  } catch (e) {
    Logger.log('assignMilestoneDirector error: ' + e);
    return { success: false, error: e.toString() };
  }
}

function getPendingMilestonesCount() {
  return getPendingMilestones().length;
}

function getActiveProbationCount() {
  // Implement based on your probation logic
  return 0; // Placeholder
}

function getPendingGraceRequestsCount() {
  // Implement based on your grace system
  return 0; // Placeholder
}

function getMonthlyEventsCount() {
  // Implement monthly event counting
  return 0; // Placeholder
}

function createUserSession(email) {
  const sessionId = Utilities.getUuid();
  const userData = {
    email: email,
    role: getUserRole(email),
    employee: resolveEmployeeName(email),
    created: new Date().getTime()
  };
  
  // Store in cache with 6-hour expiration
  CacheService.getScriptCache().put(
    'session:' + sessionId,
    JSON.stringify(userData),
    6 * 60 * 60 // 6 hours
  );
  
  return sessionId;
}

function getUserFromSession(sessionId) {
  try {
    const cached = CacheService.getScriptCache().get('session:' + sessionId);
    if (!cached) return null;
    
    return JSON.parse(cached);
  } catch (e) {
    return null;
  }
}

// ---------- DEBUG FUNCTIONS ----------
function debugUserDirectory(email) {
  try {
    console.log('🔍 DEBUG: Checking Directory for email:', email);
    
    email = norm_(email);
    console.log('📧 Normalized email:', email);
    
    const sh = dir_();
    if (!sh) {
      console.log('❌ Directory sheet not found');
      return { error: 'Directory sheet not found' };
    }
    
    const map = readDirMap_(sh);
    console.log('📋 Full Directory map keys:', Object.keys(map));
    
    const user = map[email];
    console.log('👤 User data from Directory:', user);
    
    if (user) {
      console.log('📝 User details:');
      console.log('  - Email:', user.Email);
      console.log('  - Employee:', user.Employee);
      console.log('  - Role:', user.Role);
      console.log('  - Verified:', user.Verified);
      console.log('  - Row number:', user._row);
      
      // Test getUserRole
      console.log('🎯 Testing getUserRole...');
      const role = getUserRole(email);
      console.log('🎭 getUserRole result:', role);
      
      return {
        found: true,
        email: user.Email,
        employee: user.Employee,
        role: user.Role,
        verified: user.Verified,
        row: user._row,
        getUserRoleResult: role
      };
    } else {
      console.log('❌ User not found in Directory');
      
      // Check if Directory has any data
      const data = sh.getDataRange().getValues();
      console.log('📊 Directory sheet has', data.length - 1, 'rows of data');
      
      if (data.length > 1) {
        console.log('📋 First few Directory entries:');
        for (let i = 1; i < Math.min(6, data.length); i++) {
          console.log('  Row', i + 1, ':', data[i][0], '|', data[i][1], '|', data[i][2]);
        }
      }
      
      return { 
        found: false, 
        directorySize: data.length - 1,
        sampleData: data.slice(1, Math.min(6, data.length)).map(row => ({
          email: row[0],
          employee: row[1], 
          role: row[2]
        }))
      };
    }
  } catch (e) {
    console.log('💥 debugUserDirectory error:', e);
    return { error: e.toString() };
  }
}

// Make this function available to the web app
function testUserDirectory() {
  const email = Session.getActiveUser().getEmail();
  return debugUserDirectory(email);
}
