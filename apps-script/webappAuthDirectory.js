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
  // Single-page shell: client handles all routing via google.script.history
  return HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>CLEAR - Performance Management Portal</title>
  <style>
    /* === CLEAR Design System === */
    :root {
      --cfa-red: #D73527;
      --cfa-red-dark: #B02A1F;
      --cfa-blue: #1E3A8A;
      --cfa-blue-light: #3B82F6;
      --cfa-gray-50: #F9FAFB;
      --cfa-gray-100: #F3F4F6;
      --cfa-gray-200: #E5E7EB;
      --cfa-gray-300: #D1D5DB;
      --cfa-gray-400: #9CA3AF;
      --cfa-gray-500: #6B7280;
      --cfa-gray-600: #4B5563;
      --cfa-gray-700: #374151;
      --cfa-gray-800: #1F2937;
      --cfa-gray-900: #111827;
      --success: #10B981;
      --warning: #F59E0B;
      --error: #EF4444;
      --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
      --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
      --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
      --radius: 8px;
      --radius-lg: 12px;
    }

    * { box-sizing: border-box; }
    
    body { 
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      margin: 0; 
      background: var(--cfa-gray-50);
      color: var(--cfa-gray-800);
      line-height: 1.6;
    }

    /* === Loading States === */
    .loading-spinner {
      display: inline-block;
      width: 20px;
      height: 20px;
      border: 2px solid var(--cfa-gray-300);
      border-radius: 50%;
      border-top-color: var(--cfa-red);
      animation: spin 1s ease-in-out infinite;
    }

    @keyframes spin {
      to { transform: rotate(360deg); }
    }

    .loading-overlay {
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(255, 255, 255, 0.9);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 1000;
    }

    /* === Login Screen === */
    .login-container {
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      background: linear-gradient(135deg, var(--cfa-gray-50) 0%, var(--cfa-gray-100) 100%);
    }

    .login-card {
      background: white;
      padding: 3rem;
      border-radius: var(--radius-lg);
      box-shadow: var(--shadow-lg);
      width: 100%;
      max-width: 400px;
      text-align: center;
    }

    .login-logo {
      width: 80px;
      height: 80px;
      background: var(--cfa-red);
      border-radius: 50%;
      margin: 0 auto 2rem;
      display: flex;
      align-items: center;
      justify-content: center;
      color: white;
      font-size: 2rem;
      font-weight: bold;
    }

    .login-title {
      font-size: 1.875rem;
      font-weight: 700;
      color: var(--cfa-gray-900);
      margin-bottom: 0.5rem;
    }

    .login-subtitle {
      color: var(--cfa-gray-600);
      margin-bottom: 2rem;
    }

    .form-group {
      margin-bottom: 1.5rem;
      text-align: left;
    }

    .form-label {
      display: block;
      font-weight: 600;
      color: var(--cfa-gray-700);
      margin-bottom: 0.5rem;
    }

    .form-input {
      width: 100%;
      padding: 0.75rem 1rem;
      border: 2px solid var(--cfa-gray-200);
      border-radius: var(--radius);
      font-size: 1rem;
      transition: border-color 0.2s, box-shadow 0.2s;
    }

    .form-input:focus {
      outline: none;
      border-color: var(--cfa-blue);
      box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
    }

    .btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 0.75rem 1.5rem;
      border: none;
      border-radius: var(--radius);
      font-size: 1rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
      text-decoration: none;
      min-height: 44px;
    }

    .btn:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }

    .btn-primary {
      background: var(--cfa-red);
      color: white;
      width: 100%;
    }

    .btn-primary:hover:not(:disabled) {
      background: var(--cfa-red-dark);
      transform: translateY(-1px);
      box-shadow: var(--shadow-md);
    }

    .btn-secondary {
      background: var(--cfa-gray-200);
      color: var(--cfa-gray-700);
    }

    .btn-secondary:hover {
      background: var(--cfa-gray-300);
    }

    .btn-danger {
      background: var(--error);
      color: white;
    }

    .btn-danger:hover {
      background: #DC2626;
    }

    /* === Dashboard Layout === */
    .dashboard {
      min-height: 100vh;
      display: flex;
      flex-direction: column;
    }

    .topbar {
      background: white;
      border-bottom: 1px solid var(--cfa-gray-200);
      padding: 1rem 2rem;
      display: flex;
      justify-content: space-between;
      align-items: center;
      box-shadow: var(--shadow-sm);
    }

    .brand {
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }

    .brand-logo {
      width: 32px;
      height: 32px;
      background: var(--cfa-red);
      border-radius: 6px;
      display: flex;
      align-items: center;
      justify-content: center;
      color: white;
      font-weight: bold;
      font-size: 0.875rem;
    }

    .brand-text {
      font-size: 1.25rem;
      font-weight: 700;
      color: var(--cfa-gray-900);
    }

    .topbar-actions {
      display: flex;
      align-items: center;
      gap: 1rem;
    }

    .user-info {
      display: flex;
      align-items: center;
      gap: 0.5rem;
      color: var(--cfa-gray-600);
      font-size: 0.875rem;
    }

    .user-avatar {
      width: 32px;
      height: 32px;
      background: var(--cfa-gray-300);
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-weight: 600;
      color: var(--cfa-gray-700);
    }

    /* === Main Content === */
    .main-content {
      flex: 1;
      padding: 2rem;
      max-width: 1200px;
      margin: 0 auto;
      width: 100%;
    }

    .page-header {
      margin-bottom: 2rem;
    }

    .page-title {
      font-size: 2rem;
      font-weight: 700;
      color: var(--cfa-gray-900);
      margin-bottom: 0.5rem;
    }

    .page-subtitle {
      color: var(--cfa-gray-600);
      font-size: 1.125rem;
    }

    /* === Cards === */
    .cards-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 1.5rem;
      margin-bottom: 2rem;
    }

    .card {
      background: white;
      border-radius: var(--radius-lg);
      padding: 1.5rem;
      box-shadow: var(--shadow-sm);
      border: 1px solid var(--cfa-gray-200);
      transition: all 0.2s;
    }

    .card:hover {
      box-shadow: var(--shadow-md);
      transform: translateY(-2px);
    }

    .card-header {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      margin-bottom: 1rem;
    }

    .card-title {
      font-size: 0.875rem;
      font-weight: 600;
      color: var(--cfa-gray-600);
      text-transform: uppercase;
      letter-spacing: 0.05em;
    }

    .card-icon {
      width: 40px;
      height: 40px;
      border-radius: var(--radius);
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 1.25rem;
    }

    .card-icon.primary { background: rgba(59, 130, 246, 0.1); color: var(--cfa-blue); }
    .card-icon.success { background: rgba(16, 185, 129, 0.1); color: var(--success); }
    .card-icon.warning { background: rgba(245, 158, 11, 0.1); color: var(--warning); }
    .card-icon.danger { background: rgba(239, 68, 68, 0.1); color: var(--error); }

    .card-value {
      font-size: 2.5rem;
      font-weight: 700;
      color: var(--cfa-gray-900);
      line-height: 1;
      margin-bottom: 0.5rem;
    }

    .card-description {
      color: var(--cfa-gray-600);
      font-size: 0.875rem;
    }

    /* === Status Chips === */
    .status-chip {
      display: inline-flex;
      align-items: center;
      padding: 0.25rem 0.75rem;
      border-radius: 9999px;
      font-size: 0.75rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.05em;
    }

    .status-chip.success {
      background: rgba(16, 185, 129, 0.1);
      color: var(--success);
    }

    .status-chip.warning {
      background: rgba(245, 158, 11, 0.1);
      color: var(--warning);
    }

    .status-chip.danger {
      background: rgba(239, 68, 68, 0.1);
      color: var(--error);
    }

    .status-chip.neutral {
      background: var(--cfa-gray-100);
      color: var(--cfa-gray-600);
    }

    /* === Tables === */
    .table-container {
      background: white;
      border-radius: var(--radius-lg);
      box-shadow: var(--shadow-sm);
      border: 1px solid var(--cfa-gray-200);
      overflow: hidden;
    }

    .table {
      width: 100%;
      border-collapse: collapse;
    }

    .table th {
      background: var(--cfa-gray-50);
      padding: 1rem;
      text-align: left;
      font-weight: 600;
      color: var(--cfa-gray-700);
      border-bottom: 1px solid var(--cfa-gray-200);
      font-size: 0.875rem;
      text-transform: uppercase;
      letter-spacing: 0.05em;
    }

    .table td {
      padding: 1rem;
      border-bottom: 1px solid var(--cfa-gray-100);
      color: var(--cfa-gray-800);
    }

    .table tbody tr:hover {
      background: var(--cfa-gray-50);
    }

    /* === Navigation === */
    .nav-tabs {
      display: flex;
      border-bottom: 1px solid var(--cfa-gray-200);
      margin-bottom: 2rem;
    }

    .nav-tab {
      padding: 0.75rem 1.5rem;
      border: none;
      background: none;
      color: var(--cfa-gray-600);
      font-weight: 600;
      cursor: pointer;
      border-bottom: 2px solid transparent;
      transition: all 0.2s;
    }

    .nav-tab.active {
      color: var(--cfa-red);
      border-bottom-color: var(--cfa-red);
    }

    .nav-tab:hover:not(.active) {
      color: var(--cfa-gray-800);
    }

    /* === Messages === */
    .message {
      padding: 1rem;
      border-radius: var(--radius);
      margin-bottom: 1rem;
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }

    .message.error {
      background: rgba(239, 68, 68, 0.1);
      color: var(--error);
      border: 1px solid rgba(239, 68, 68, 0.2);
    }

    .message.success {
      background: rgba(16, 185, 129, 0.1);
      color: var(--success);
      border: 1px solid rgba(16, 185, 129, 0.2);
    }

    .message.warning {
      background: rgba(245, 158, 11, 0.1);
      color: var(--warning);
      border: 1px solid rgba(245, 158, 11, 0.2);
    }

    /* === Responsive === */
    @media (max-width: 768px) {
      .main-content {
        padding: 1rem;
      }
      
      .topbar {
        padding: 1rem;
      }
      
      .cards-grid {
        grid-template-columns: 1fr;
      }
      
      .login-card {
        margin: 1rem;
        padding: 2rem;
      }
    }

    /* === Animations === */
    .fade-in {
      animation: fadeIn 0.3s ease-in-out;
    }

    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(10px); }
      to { opacity: 1; transform: translateY(0); }
    }
  </style>
</head>
<body>
  <div id="app"></div>
  <script>
    (function(){
      const state = { session: null, user: null };
      const $ = (id) => document.getElementById(id);

      function setQuerySession(session) {
        try {
          const base = window.location.href.split('?')[0];
          const hash = window.location.hash || '';
          const qs = session ? ('?session=' + encodeURIComponent(session)) : '';
          window.top.history.replaceState(null, '', base + qs + hash);
        } catch (_) {}
      }

      function showMessage(text, type = 'error') {
        const app = $('app');
        const messageEl = document.createElement('div');
        messageEl.className = \`message \${type} fade-in\`;
        messageEl.innerHTML = \`
          <div class="loading-spinner"></div>
          <span>\${text}</span>
        \`;
        
        // Remove existing messages
        const existing = app.querySelector('.message');
        if (existing) existing.remove();
        
        app.insertBefore(messageEl, app.firstChild);
        
        // Auto-remove after 5 seconds
        setTimeout(() => {
          if (messageEl.parentNode) {
            messageEl.remove();
          }
        }, 5000);
      }

      function showLoading() {
        const overlay = document.createElement('div');
        overlay.className = 'loading-overlay';
        overlay.innerHTML = \`
          <div style="text-align: center;">
            <div class="loading-spinner" style="width: 40px; height: 40px; border-width: 4px;"></div>
            <p style="margin-top: 1rem; color: var(--cfa-gray-600);">Loading...</p>
          </div>
        \`;
        document.body.appendChild(overlay);
        return overlay;
      }

      function hideLoading(overlay) {
        if (overlay && overlay.parentNode) {
          overlay.remove();
        }
      }

      function render(token) {
        const route = token || '';
        if (!state.user) {
          return renderLogin();
        }
        if (route.startsWith('dashboard:')) {
          const role = route.split(':')[1] || state.user.role || 'employee';
          return renderDashboard(role);
        }
        if (route === 'employeeSearch') {
          return renderEmployeeSearch();
        }
        return renderDashboard(state.user.role || 'employee');
      }

      function renderLogin() {
        const app = $('app');
        app.innerHTML = \`
          <div class="login-container">
            <div class="login-card fade-in">
              <div class="login-logo">C</div>
              <h1 class="login-title">CLEAR</h1>
              <p class="login-subtitle">Performance Management Portal</p>
              
              <form id="loginForm">
                <div class="form-group">
                  <label for="email" class="form-label">Email Address</label>
                  <input type="email" id="email" class="form-input" required 
                         placeholder="Enter your email" autocomplete="email" />
                </div>
                
                <div class="form-group">
                  <label for="password" class="form-label">Password</label>
                  <input type="password" id="password" class="form-input" required 
                         placeholder="Enter your password" autocomplete="current-password" />
                </div>
                
                <button type="submit" id="loginBtn" class="btn btn-primary">
                  <span id="loginBtnText">Sign In</span>
                  <div id="loginSpinner" class="loading-spinner" style="display: none; margin-left: 0.5rem;"></div>
                </button>
              </form>
              
              <div id="msg" style="margin-top: 1rem;"></div>
            </div>
          </div>
        \`;

        $('loginForm').addEventListener('submit', (e) => {
          e.preventDefault();
          const email = $('email').value;
          const password = $('password').value;
          const btn = $('loginBtn');
          const btnText = $('loginBtnText');
          const spinner = $('loginSpinner');
          
          if (!email || !password) { 
            showMessage('Please enter both email and password', 'error'); 
            return; 
          }
          
          btn.disabled = true;
          btnText.textContent = 'Signing in...';
          spinner.style.display = 'inline-block';
          
          google.script.run
            .withSuccessHandler((result) => {
              if (result && result.ok) {
                state.session = result.sessionId;
                state.user = { email: result.email, role: result.role, employee: result.employee };
                setQuerySession(state.session);
                const token = 'dashboard:' + (result.role || 'employee');
                google.script.history.replace(token);
                render(token);
              } else {
                showMessage((result && result.error) || 'Login failed', 'error');
                btn.disabled = false;
                btnText.textContent = 'Sign In';
                spinner.style.display = 'none';
              }
            })
            .withFailureHandler((err) => {
              showMessage('Login error: ' + (err && err.message ? err.message : err), 'error');
              btn.disabled = false;
              btnText.textContent = 'Sign In';
              spinner.style.display = 'none';
            })
            .loginWithPassword(email, password);
        });
      }

      function renderDashboard(role) {
        const app = $('app');
        const roleTitle = role.charAt(0).toUpperCase() + role.slice(1);
        
        app.innerHTML = \`
          <div class="dashboard">
            <div class="topbar">
              <div class="brand">
                <div class="brand-logo">C</div>
                <div class="brand-text">CLEAR</div>
              </div>
              <div class="topbar-actions">
                <div class="user-info">
                  <div class="user-avatar">\${(state.user?.employee || state.user?.email || 'U').charAt(0).toUpperCase()}</div>
                  <span>\${state.user?.employee || state.user?.email || 'User'}</span>
                </div>
                \${role === 'director' || role === 'lead' ? '<button id="toSearch" class="btn btn-secondary">Employee Search</button>' : ''}
                <button id="logoutBtn" class="btn btn-danger">Logout</button>
              </div>
            </div>
            
            <div class="main-content">
              <div class="page-header">
                <h1 class="page-title">\${roleTitle} Dashboard</h1>
                <p class="page-subtitle">Welcome back, \${state.user?.employee || state.user?.email || 'User'}</p>
              </div>
              
              <div id="dashboardContent">
                <div class="loading-overlay">
                  <div style="text-align: center;">
                    <div class="loading-spinner" style="width: 40px; height: 40px; border-width: 4px;"></div>
                    <p style="margin-top: 1rem; color: var(--cfa-gray-600);">Loading dashboard...</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        \`;

        // Event listeners
        $('logoutBtn').addEventListener('click', () => {
          const s = state.session;
          state.session = null; 
          state.user = null;
          setQuerySession('');
          try { google.script.history.replace(''); } catch(_){}
          if (s) { try { google.script.run.logoutSession(s); } catch(_){} }
          render('');
        });

        if ($('toSearch')) {
          $('toSearch').addEventListener('click', () => {
            google.script.history.push('employeeSearch');
          });
        }

        // Load role-specific content
        loadDashboardContent(role);
      }

      function loadDashboardContent(role) {
        const contentEl = $('dashboardContent');
        
        if (role === 'director') {
          loadDirectorDashboard(contentEl);
        } else if (role === 'lead') {
          loadLeadDashboard(contentEl);
        } else {
          loadEmployeeDashboard(contentEl);
        }
      }

      function loadDirectorDashboard(container) {
        google.script.run
          .withSuccessHandler((data) => {
            if (data && !data.error) {
              container.innerHTML = \`
                <div class="cards-grid">
                  <div class="card fade-in">
                    <div class="card-header">
                      <div class="card-title">Pending Milestones</div>
                      <div class="card-icon warning">⚠️</div>
                    </div>
                    <div class="card-value">\${data.pendingMilestones || 0}</div>
                    <div class="card-description">Requiring director assignment</div>
                  </div>
                  
                  <div class="card fade-in">
                    <div class="card-header">
                      <div class="card-title">Active Probation</div>
                      <div class="card-icon danger">🚨</div>
                    </div>
                    <div class="card-value">\${data.activeProbation || 0}</div>
                    <div class="card-description">Employees on probation</div>
                  </div>
                  
                  <div class="card fade-in">
                    <div class="card-header">
                      <div class="card-title">Grace Requests</div>
                      <div class="card-icon primary">🤝</div>
                    </div>
                    <div class="card-value">\${data.graceRequests || 0}</div>
                    <div class="card-description">Pending approval</div>
                  </div>
                  
                  <div class="card fade-in">
                    <div class="card-header">
                      <div class="card-title">Monthly Events</div>
                      <div class="card-icon success">📊</div>
                    </div>
                    <div class="card-value">\${data.monthlyEvents || 0}</div>
                    <div class="card-description">This month's incidents</div>
                  </div>
                </div>
                
                <div class="nav-tabs">
                  <button class="nav-tab active" data-tab="pending">Pending Milestones</button>
                  <button class="nav-tab" data-tab="reports">Reports</button>
                  <button class="nav-tab" data-tab="bulk">Bulk Operations</button>
                </div>
                
                <div id="tabContent">
                  <div class="table-container">
                    <table class="table">
                      <thead>
                        <tr>
                          <th>Employee</th>
                          <th>Milestone</th>
                          <th>Status</th>
                          <th>Actions</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td colspan="4" style="text-align: center; padding: 2rem; color: var(--cfa-gray-500);">
                            Loading pending milestones...
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>
              \`;
              
              // Load pending milestones
              loadPendingMilestones();
            } else {
              container.innerHTML = \`
                <div class="message error">
                  <span>Failed to load dashboard data. Please try again.</span>
                </div>
              \`;
            }
          })
          .withFailureHandler((err) => {
            container.innerHTML = \`
              <div class="message error">
                <span>Network error: \${err.message || err}</span>
              </div>
            \`;
          })
          .getDirectorDashboardData();
      }

      function loadLeadDashboard(container) {
        container.innerHTML = \`
          <div class="cards-grid">
            <div class="card fade-in">
              <div class="card-header">
                <div class="card-title">Team Overview</div>
                <div class="card-icon primary">👥</div>
              </div>
              <div class="card-value">—</div>
              <div class="card-description">Team performance metrics</div>
            </div>
            
            <div class="card fade-in">
              <div class="card-header">
                <div class="card-title">Pending Reviews</div>
                <div class="card-icon warning">📋</div>
              </div>
              <div class="card-value">—</div>
              <div class="card-description">Items requiring attention</div>
            </div>
          </div>
          
          <div class="message warning">
            <span>Lead dashboard features coming soon. Contact your director for team-specific data.</span>
          </div>
        \`;
      }

      function loadEmployeeDashboard(container) {
        container.innerHTML = \`
          <div class="cards-grid">
            <div class="card fade-in">
              <div class="card-header">
                <div class="card-title">Current Status</div>
                <div class="card-icon primary">📊</div>
              </div>
              <div class="card-value" id="statusValue">Loading...</div>
              <div class="card-description" id="statusDescription">Your performance metrics</div>
            </div>
            
            <div class="card fade-in">
              <div class="card-header">
                <div class="card-title">Active Milestone</div>
                <div class="card-icon warning">🎯</div>
              </div>
              <div class="card-value" id="milestoneValue">Loading...</div>
              <div class="card-description" id="milestoneDescription">Current active milestone</div>
            </div>
            
            <div class="card fade-in">
              <div class="card-header">
                <div class="card-title">Grace Available</div>
                <div class="card-icon success">🤝</div>
              </div>
              <div class="card-value" id="graceValue">Loading...</div>
              <div class="card-description" id="graceDescription">Available grace credits</div>
            </div>
          </div>
          
          <div class="table-container" style="margin-top: 2rem;">
            <h3 style="padding: 1rem 1rem 0; margin: 0; color: var(--cfa-gray-800);">Recent History</h3>
            <table class="table">
              <thead>
                <tr>
                  <th>Date</th>
                  <th>Event</th>
                  <th>Points</th>
                  <th>Status</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td colspan="5" style="text-align: center; padding: 2rem; color: var(--cfa-gray-500);">
                    Loading your history...
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        \`;
        
        // Load employee data
        loadEmployeeData();
      }

      function loadEmployeeData() {
        if (!state.user?.email) return;
        
        // Load main overview data
        google.script.run
          .withSuccessHandler((data) => {
            if (data && !data.error) {
              // Update status card
              const statusCard = document.getElementById('statusValue');
              const statusDesc = document.getElementById('statusDescription');
              if (statusCard) {
                statusCard.textContent = data.effectivePoints || '0';
                statusDesc.textContent = 'Rolling points (effective)';
              }
              
              // Update history table
              const tbody = document.querySelector('.table tbody');
              if (tbody && data.rows && data.rows.length > 0) {
                tbody.innerHTML = data.rows.slice(0, 10).map(row => {
                  const isAppealable = row.eventType && 
                    row.eventType.toLowerCase() === 'disciplinary event' && 
                    row.active;
                  
                  return \`
                    <tr>
                      <td>\${row.date || '—'}</td>
                      <td>\${row.event || '—'}</td>
                      <td>\${row.points || '—'}</td>
                      <td><span class="status-chip \${row.active ? 'warning' : 'neutral'}">\${row.active ? 'Active' : 'Completed'}</span></td>
                      <td>
                        \${isAppealable ? \`
                          <button class="btn btn-secondary" onclick="showAppealModal(\${row.row})" style="padding: 0.25rem 0.5rem; font-size: 0.75rem;">
                            Appeal
                          </button>
                        \` : '—'}
                      </td>
                    </tr>
                  \`;
                }).join('');
              } else if (tbody) {
                tbody.innerHTML = \`
                  <tr>
                    <td colspan="5" style="text-align: center; padding: 2rem; color: var(--cfa-gray-500);">
                      No history found.
                    </td>
                  </tr>
                \`;
              }
            }
          })
          .withFailureHandler((err) => {
            showMessage('Failed to load your data: ' + (err.message || err), 'error');
          })
          .getMyOverviewForEmail(state.user.email);
        
        // Load active milestone
        google.script.run
          .withSuccessHandler((milestone) => {
            const milestoneValue = document.getElementById('milestoneValue');
            const milestoneDesc = document.getElementById('milestoneDescription');
            if (milestoneValue) {
              if (milestone) {
                milestoneValue.textContent = milestone.milestone || 'Active';
                milestoneDesc.textContent = milestone.consequence || 'Current milestone';
              } else {
                milestoneValue.textContent = 'None';
                milestoneDesc.textContent = 'No active milestones';
              }
            }
          })
          .withFailureHandler((err) => {
            const milestoneValue = document.getElementById('milestoneValue');
            if (milestoneValue) {
              milestoneValue.textContent = 'Error';
            }
          })
          .getActiveMilestone(state.user.email);
        
        // Load grace credits
        google.script.run
          .withSuccessHandler((credits) => {
            const graceValue = document.getElementById('graceValue');
            const graceDesc = document.getElementById('graceDescription');
            if (graceValue) {
              if (credits && credits.length > 0) {
                graceValue.textContent = credits.length;
                graceDesc.innerHTML = credits.slice(0, 3).map(credit => 
                  \`<div style="font-size: 0.75rem; margin: 0.25rem 0;">• \${credit.reason || 'Grace credit'}</div>\`
                ).join('');
              } else {
                graceValue.textContent = '0';
                graceDesc.textContent = 'No grace credits available';
              }
            }
          })
          .withFailureHandler((err) => {
            const graceValue = document.getElementById('graceValue');
            if (graceValue) {
              graceValue.textContent = 'Error';
            }
          })
          .getGraceCredits(state.user.email);
      }

      // Global function for appeal modal
      window.showAppealModal = function(row) {
        const reason = prompt('Please provide a reason for your appeal:');
        if (reason && reason.trim()) {
          const overlay = showLoading();
          
          google.script.run
            .withSuccessHandler((result) => {
              hideLoading(overlay);
              if (result.success) {
                showMessage('Appeal submitted successfully! Your milestone has been canceled.', 'success');
                // Reload the data to reflect changes
                setTimeout(() => {
                  loadEmployeeData();
                }, 1000);
              } else {
                showMessage('Failed to submit appeal: ' + (result.error || 'Unknown error'), 'error');
              }
            })
            .withFailureHandler((err) => {
              hideLoading(overlay);
              showMessage('Error submitting appeal: ' + (err.message || err), 'error');
            })
            .submitAppeal(state.user.email, row, reason.trim());
        }
      };

      function loadPendingMilestones() {
        google.script.run
          .withSuccessHandler((milestones) => {
            const tbody = document.querySelector('.table tbody');
            if (tbody) {
              if (milestones && milestones.length > 0) {
                tbody.innerHTML = milestones.map(milestone => \`
                  <tr>
                    <td>\${milestone.employee || '—'}</td>
                    <td>\${milestone.milestone || '—'}</td>
                    <td><span class="status-chip warning">Pending</span></td>
                    <td>
                      <button class="btn btn-secondary" onclick="assignDirector(\${milestone.row})" style="padding: 0.5rem 1rem; font-size: 0.875rem;">
                        Assign Director
                      </button>
                    </td>
                  </tr>
                \`).join('');
              } else {
                tbody.innerHTML = \`
                  <tr>
                    <td colspan="4" style="text-align: center; padding: 2rem; color: var(--cfa-gray-500);">
                      No pending milestones.
                    </td>
                  </tr>
                \`;
              }
            }
          })
          .withFailureHandler((err) => {
            const tbody = document.querySelector('.table tbody');
            if (tbody) {
              tbody.innerHTML = \`
                <tr>
                  <td colspan="4" style="text-align: center; padding: 2rem; color: var(--error);">
                    Error loading milestones: \${err.message || err}
                  </td>
                </tr>
              \`;
            }
          })
          .getPendingMilestones();
      }

      function renderEmployeeSearch() {
        const app = $('app');
        app.innerHTML = \`
          <div class="dashboard">
            <div class="topbar">
              <div class="brand">
                <div class="brand-logo">C</div>
                <div class="brand-text">CLEAR</div>
              </div>
              <div class="topbar-actions">
                <button id="backBtn" class="btn btn-secondary">← Back to Dashboard</button>
                <button id="logoutBtn" class="btn btn-danger">Logout</button>
              </div>
            </div>
            
            <div class="main-content">
              <div class="page-header">
                <h1 class="page-title">Employee Search</h1>
                <p class="page-subtitle">Search and view employee performance data</p>
              </div>
              
              <div class="message warning">
                <span>Employee search functionality is coming soon. This feature will allow you to search for employees and view their performance history.</span>
              </div>
            </div>
          </div>
        \`;

        $('backBtn').addEventListener('click', () => {
          const token = 'dashboard:' + (state.user && state.user.role || 'employee');
          google.script.history.replace(token);
          render(token);
        });

        $('logoutBtn').addEventListener('click', () => {
          const s = state.session;
          state.session = null; 
          state.user = null;
          setQuerySession('');
          try { google.script.history.replace(''); } catch(_){}
          if (s) { try { google.script.run.logoutSession(s); } catch(_){} }
          render('');
        });
      }

      function bootstrap(){
        google.script.url.getLocation((loc) => {
          const session = (loc && loc.parameter && loc.parameter.session) ? loc.parameter.session : null;
          const token = (loc && typeof loc.hash === 'string') ? loc.hash : '';
          if (session) {
            google.script.run
              .withSuccessHandler((user) => {
                if (user) {
                  state.session = session; 
                  state.user = user;
                  setQuerySession(session);
                  const initial = token || ('dashboard:' + (user.role || 'employee'));
                  render(initial);
                } else {
                  render('');
                }
              })
              .withFailureHandler(() => render(''))
              .getUserFromSession(session);
          } else {
            render(token);
          }
          google.script.history.setChangeHandler((e) => render((e && e.token) || ''));
        });
      }

      // Global functions for milestone assignment
      window.assignDirector = function(row) {
        const director = prompt('Enter director name:');
        if (director) {
          google.script.run
            .withSuccessHandler(() => {
              showMessage('Director assigned successfully', 'success');
              loadPendingMilestones();
            })
            .withFailureHandler((err) => {
              showMessage('Failed to assign director: ' + (err.message || err), 'error');
            })
            .assignMilestoneDirector(row, director);
        }
      };

      bootstrap();
    })();
  </script>
</body>
</html>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
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
    
    console.log(' Calling createUserSession...');
    const sessionId = createUserSession(email); // Create session
    console.log(' Session created:', sessionId);
    
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
  return getEmployeeHistoryWithAppeals(email);
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
  console.log(' getUserRole called for email:', email);
  
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

// SPA: allow logout to invalidate a session server-side
function logoutSession(sessionId) {
  try {
    if (!sessionId) return false;
    CacheService.getScriptCache().remove('session:' + sessionId);
    return true;
  } catch (e) {
    return false;
  }
}

// Add these new functions after the existing backend functions

function getActiveMilestone(email) {
  try {
    const employee = resolveEmployeeName(email);
    if (!employee) return null;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventsSheet = ss.getSheetByName('Events');
    if (!eventsSheet) return null;
    
    const data = eventsSheet.getDataRange().getValues();
    const headers = data[0];
    
    const employeeCol = headers.indexOf('Employee');
    const activeCol = headers.indexOf('Active');
    const milestoneCol = headers.indexOf('Milestone');
    const consequenceCol = headers.indexOf('Consequence');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[employeeCol] || '').trim().toLowerCase() === employee.toLowerCase() &&
          String(row[activeCol] || '').toLowerCase() === 'true') {
        return {
          milestone: row[milestoneCol] || '',
          consequence: row[consequenceCol] || '',
          row: i + 1
        };
      }
    }
    
    return null;
  } catch (e) {
    Logger.log('getActiveMilestone error: ' + e);
    return null;
  }
}

function getGraceCredits(email) {
  try {
    const employee = resolveEmployeeName(email);
    if (!employee) return [];
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const positivePointsSheet = ss.getSheetByName('PositivePoints');
    if (!positivePointsSheet) return [];
    
    const data = positivePointsSheet.getDataRange().getValues();
    const headers = data[0];
    
    const employeeCol = headers.indexOf('Employee');
    const creditReasonCol = headers.indexOf('Credit Reason');
    const dateCol = headers.indexOf('Date');
    const pointsCol = headers.indexOf('Points');
    
    const credits = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[employeeCol] || '').trim().toLowerCase() === employee.toLowerCase()) {
        credits.push({
          reason: row[creditReasonCol] || '',
          date: row[dateCol] || '',
          points: row[pointsCol] || 0,
          row: i + 1
        });
      }
    }
    
    return credits;
  } catch (e) {
    Logger.log('getGraceCredits error: ' + e);
    return [];
  }
}

function getEmployeeHistoryWithAppeals(email) {
  try {
    const employee = resolveEmployeeName(email);
    if (!employee) return { rows: [], generatedAt: now_() };
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventsSheet = ss.getSheetByName('Events');
    if (!eventsSheet) return { rows: [], generatedAt: now_() };
    
    const hdr = eventsSheet.getRange(1,1,1,eventsSheet.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
    const idx = n => hdr.indexOf(n);
    const iEmp=idx('Employee'), iDate=firstIdx_(hdr,['IncidentDate','Date','Timestamp']);
    const iEvt=idx('EventType'), iInf=idx('Infraction'), iLead=idx('Lead');
    const iPts=idx('Points'), iRollE=idx('PointsRolling (Effective)'), iRoll=idx('PointsRolling');
    const iNotes=firstIdx_(hdr,['Notes / Reviewer','Notes','Reviewer','Notes/Reviewer']);
    const iPdf=firstIdx_(hdr,['Write-Up PDF','PDF Link','WriteUpPDF','Signed_PDF_Link']);
    const iActive=idx('Active');
    const iConsequence=idx('Consequence');
    const iGraceReason=idx('Grace Reason');
    const tz=Session.getScriptTimeZone()||'UTC';

    const vals=eventsSheet.getDataRange().getValues(); const rows=[]; let lastEff=null;
    for (let r=1;r<vals.length;r++){
      if (String(vals[r][iEmp]||'').trim().toLowerCase()!==String(employee).trim().toLowerCase()) continue;
      let roll=null; if(iRollE!==-1&&isFinite(Number(vals[r][iRollE]))) roll=Number(vals[r][iRollE]); else if(iRoll!==-1&&isFinite(Number(vals[r][iRoll]))) roll=Number(vals[r][iRoll]);
      if (roll!=null) lastEff=roll;
      
      const eventType = String(vals[r][iEvt]||'');
      let eventDisplay = '';
      
      // Determine what to show in the Event column
      if (eventType.toLowerCase() === 'disciplinary event') {
        eventDisplay = String(vals[r][iInf]||'');
      } else if (eventType.toLowerCase() === 'milestone') {
        eventDisplay = String(vals[r][iConsequence]||'');
      } else if (eventType.toLowerCase() === 'grace') {
        eventDisplay = String(vals[r][iGraceReason]||'');
      } else {
        eventDisplay = eventType; // fallback to event type
      }
      
      rows.push({
        date:fmtDate_(vals[r][iDate],tz),
        event: eventDisplay,
        eventType: eventType,
        infraction:String(vals[r][iInf]||''),
        notes:(iNotes!==-1)?String(vals[r][iNotes]||''):'',
        points:(iPts!==-1 && vals[r][iPts]!=='' && vals[r][iPts]!=null)?Number(vals[r][iPts]):null,
        roll:roll,
        lead:String(vals[r][iLead]||''),
        pdfUrl:(iPdf!==-1)?String(vals[r][iPdf]||''):'',
        active: String(vals[r][iActive]||'').toLowerCase() === 'true',
        row: r + 1
      });
    }
    return { employee, rows, effectivePoints:(lastEff!=null)?lastEff:null, graceAvailable:null, graceAvailableText:null, generatedAt:now_() };
  } catch (e) {
    Logger.log('getEmployeeHistoryWithAppeals error: ' + e);
    return { rows: [], generatedAt: now_() };
  }
}

function submitAppeal(email, row, appealReason) {
  try {
    const employee = resolveEmployeeName(email);
    if (!employee) return { success: false, error: 'Employee not found' };
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventsSheet = ss.getSheetByName('Events');
    if (!eventsSheet) return { success: false, error: 'Events sheet not found' };
    
    const data = eventsSheet.getDataRange().getValues();
    const headers = data[0];
    
    const employeeCol = headers.indexOf('Employee');
    const eventTypeCol = headers.indexOf('EventType');
    const activeCol = headers.indexOf('Active');
    
    // Verify this is an active disciplinary event for this employee
    const rowData = data[row - 1];
    if (String(rowData[employeeCol] || '').trim().toLowerCase() !== employee.toLowerCase()) {
      return { success: false, error: 'Event not found for this employee' };
    }
    
    if (String(rowData[eventTypeCol] || '').toLowerCase() !== 'disciplinary event') {
      return { success: false, error: 'Only disciplinary events can be appealed' };
    }
    
    if (String(rowData[activeCol] || '').toLowerCase() !== 'true') {
      return { success: false, error: 'Only active events can be appealed' };
    }
    
    // Set Active to FALSE (effectively canceling the milestone)
    eventsSheet.getRange(row, activeCol + 1).setValue('FALSE');
    
    // Add appeal reason to notes or create an appeal log
    const notesCol = headers.indexOf('Notes / Reviewer');
    if (notesCol !== -1) {
      const currentNotes = String(rowData[notesCol] || '');
      const newNotes = currentNotes + (currentNotes ? '\n' : '') + `APPEAL: ${appealReason}`;
      eventsSheet.getRange(row, notesCol + 1).setValue(newNotes);
    }
    
    // Notify directors
    const directorEmails = getDirectorEmails();
    if (directorEmails.length > 0) {
      try {
        MailApp.sendEmail(
          directorEmails.join(','), 
          'CLEAR Appeal Submitted', 
          `${employee} has appealed a disciplinary event.\n\nAppeal Reason: ${appealReason}\n\nEvent Row: ${row}`
        );
      } catch (e) {
        Logger.log('Failed to send appeal notification: ' + e);
      }
    }
    
    return { success: true, message: 'Appeal submitted successfully' };
  } catch (e) {
    Logger.log('submitAppeal error: ' + e);
    return { success: false, error: 'Failed to submit appeal: ' + e.toString() };
  }
}
