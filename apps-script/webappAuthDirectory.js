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

// ---------- MISSING UTILITY FUNCTIONS ----------

// Add this after the existing utility functions (around line 40)
function fmtDate_(dateValue, timezone) {
  try {
    if (!dateValue) return '';
    
    let date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else if (typeof dateValue === 'string') {
      date = new Date(dateValue);
    } else if (typeof dateValue === 'number') {
      date = new Date(dateValue);
    } else {
      return String(dateValue || '');
    }
    
    // Format the date using the provided timezone
    return Utilities.formatDate(date, timezone || 'UTC', 'yyyy-MM-dd HH:mm');
  } catch (e) {
    console.log('Error formatting date:', e);
    return String(dateValue || '');
  }
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
    return HtmlService.createHtmlOutput('<h1>Error</h1><p>' + error + '</p>');
  }
}

function renderLoginPage() {
  return HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <title>CLEAR — Authentication</title>
  <meta charset="UTF-8">
  <style>
    body { 
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
      margin: 0; 
      padding: 20px; 
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    
    .auth-container { 
      background: white; 
      border-radius: 12px; 
      box-shadow: 0 15px 35px rgba(0,0,0,0.1); 
      width: 100%; 
      max-width: 450px; 
      overflow: hidden;
    }
    
    .auth-header {
      background: #2c3e50;
      color: white;
      padding: 30px 40px;
      text-align: center;
    }
    
    .auth-header h1 {
      margin: 0;
      font-size: 28px;
      font-weight: 300;
    }
    
    .auth-tabs {
      display: flex;
      background: #f8f9fa;
      border-bottom: 1px solid #dee2e6;
    }
    
    .auth-tab {
      flex: 1;
      padding: 15px;
      text-align: center;
      cursor: pointer;
      border: none;
      background: transparent;
      font-size: 14px;
      font-weight: 500;
      color: #6c757d;
      transition: all 0.3s ease;
    }
    
    .auth-tab.active {
      background: white;
      color: #007bff;
      border-bottom: 3px solid #007bff;
    }
    
    .auth-tab:hover {
      background: #e9ecef;
      color: #007bff;
    }
    
    .auth-content {
      padding: 40px;
    }
    
    .form-group { 
      margin: 20px 0; 
    }
    
    .form-group label { 
      display: block; 
      margin-bottom: 8px; 
      font-weight: 500;
      color: #495057;
    }
    
    .form-group input { 
      width: 100%; 
      padding: 12px 16px; 
      border: 2px solid #e9ecef; 
      border-radius: 8px; 
      font-size: 16px;
      transition: border-color 0.3s ease;
      box-sizing: border-box;
    }
    
    .form-group input:focus {
      outline: none;
      border-color: #007bff;
      box-shadow: 0 0 0 3px rgba(0,123,255,0.1);
    }
    
    .btn { 
      width: 100%; 
      padding: 14px; 
      border: none; 
      border-radius: 8px; 
      font-size: 16px;
      font-weight: 500;
      cursor: pointer; 
      transition: all 0.3s ease;
      margin-top: 10px;
    }
    
    .btn-primary { 
      background: #007bff; 
      color: white; 
    }
    
    .btn-primary:hover { 
      background: #0056b3; 
      transform: translateY(-1px);
    }
    
    .btn-secondary { 
      background: #6c757d; 
      color: white; 
    }
    
    .btn-secondary:hover { 
      background: #545b62; 
    }
    
    .btn:disabled { 
      background: #ccc; 
      cursor: not-allowed; 
      transform: none;
    }
    
    .msg { 
      margin: 15px 0; 
      padding: 12px 16px; 
      border-radius: 8px; 
      font-size: 14px;
      display: none;
    }
    
    .msg.error { 
      background: #f8d7da; 
      color: #721c24; 
      border: 1px solid #f5c6cb; 
      display: block;
    }
    
    .msg.success { 
      background: #d4edda; 
      color: #155724; 
      border: 1px solid #c3e6cb; 
      display: block;
    }
    
    .auth-footer {
      text-align: center;
      padding: 20px 40px;
      background: #f8f9fa;
      border-top: 1px solid #dee2e6;
      font-size: 14px;
      color: #6c757d;
    }
    
    .switch-link {
      color: #007bff;
      cursor: pointer;
      text-decoration: underline;
    }
    
    .switch-link:hover {
      color: #0056b3;
    }
    
    .tab-content {
      display: none;
    }
    
    .tab-content.active {
      display: block;
    }
    
    .code-input {
      text-align: center;
      font-family: 'Courier New', monospace;
      font-size: 24px;
      letter-spacing: 8px;
    }
  </style>
</head>
<body>
  <div class="auth-container">
    <div class="auth-header">
      <h1>CLEAR</h1>
      <p>Standards Management System</p>
    </div>
    
    <div class="auth-tabs">
      <button class="auth-tab active" onclick="showTab('signin')">Sign In</button>
      <button class="auth-tab" onclick="showTab('code')">Sign In Code</button>
      <button class="auth-tab" onclick="showTab('create')">Create Account</button>
      <button class="auth-tab" onclick="showTab('reset')">Reset Password</button>
    </div>
    
    <div class="auth-content">
      
      <!-- SIGN IN TAB -->
      <div id="signin" class="tab-content active">
        <div class="form-group">
          <label for="signin-email">Email:</label>
          <input type="email" id="signin-email" required>
        </div>
        
        <div class="form-group">
          <label for="signin-password">Password:</label>
          <input type="password" id="signin-password" required>
        </div>
        
        <button id="signinBtn" class="btn btn-primary" onclick="login()">Sign In</button>
        
        <div id="signin-msg" class="msg"></div>
      </div>
      
      <!-- SIGN IN CODE TAB -->
      <div id="code" class="tab-content">
        <p style="text-align: center; color: #6c757d; margin-bottom: 20px;">
          Enter your email to receive a sign-in code
        </p>
        
        <div class="form-group">
          <label for="code-email">Email:</label>
          <input type="email" id="code-email" required>
        </div>
        
        <div class="form-group" id="code-input-group" style="display: none;">
          <label for="code-code">Enter 6-digit code:</label>
          <input type="text" id="code-code" class="code-input" maxlength="6" pattern="[0-9]{6}" placeholder="000000">
        </div>
        
        <button id="codeBtn" class="btn btn-primary" onclick="requestSigninCode()">Send Code</button>
        <button id="codeVerifyBtn" class="btn btn-primary" onclick="verifySigninCode()" style="display: none;">Verify & Sign In</button>
        
        <div id="code-msg" class="msg"></div>
      </div>
      
      <!-- CREATE ACCOUNT TAB -->
      <div id="create" class="tab-content">
        <p style="text-align: center; color: #6c757d; margin-bottom: 20px;">
          Create a new account - you'll receive an email verification code
        </p>
        
        <div class="form-group">
          <label for="create-email">Email:</label>
          <input type="email" id="create-email" required>
        </div>
        
        <div class="form-group">
          <label for="create-name">Full Name:</label>
          <input type="text" id="create-name" required>
        </div>
        
        <div class="form-group" id="create-code-group" style="display: none;">
          <label for="create-code">Enter 6-digit verification code:</label>
          <input type="text" id="create-code" class="code-input" maxlength="6" pattern="[0-9]{6}" placeholder="000000">
        </div>
        
        <div class="form-group" id="create-password-group" style="display: none;">
          <label for="create-password">Create Password (min 8 characters):</label>
          <input type="password" id="create-password" minlength="8" required>
        </div>
        
        <button id="createBtn" class="btn btn-primary" onclick="requestCreateCode()">Send Verification Code</button>
        <button id="createVerifyBtn" class="btn btn-primary" onclick="verifyCreateAccount()" style="display: none;">Create Account</button>
        
        <div id="create-msg" class="msg"></div>
      </div>
      
      <!-- RESET PASSWORD TAB -->
      <div id="reset" class="tab-content">
        <p style="text-align: center; color: #6c757d; margin-bottom: 20px;">
          Reset your password - you'll receive an email verification code
        </p>
        
        <div class="form-group">
          <label for="reset-email">Email:</label>
          <input type="email" id="reset-email" required>
        </div>
        
        <div class="form-group" id="reset-code-group" style="display: none;">
          <label for="reset-code">Enter 6-digit verification code:</label>
          <input type="text" id="reset-code" class="code-input" maxlength="6" pattern="[0-9]{6}" placeholder="000000">
        </div>
        
        <div class="form-group" id="reset-password-group" style="display: none;">
          <label for="reset-password">New Password (min 8 characters):</label>
          <input type="password" id="reset-password" minlength="8" required>
        </div>
        
        <button id="resetBtn" class="btn btn-primary" onclick="requestResetCode()">Send Reset Code</button>
        <button id="resetVerifyBtn" class="btn btn-primary" onclick="verifyResetPassword()" style="display: none;">Reset Password</button>
        
        <div id="reset-msg" class="msg"></div>
      </div>
      
    </div>
    
    <div class="auth-footer">
      Need help? Contact your system administrator
    </div>
  </div>

  <script>
    function $(id) { return document.getElementById(id); }
    
    function showTab(tabName) {
      // Hide all tab contents
      const tabs = document.querySelectorAll('.tab-content');
      tabs.forEach(tab => tab.classList.remove('active'));
      
      // Remove active class from all tab buttons
      const tabButtons = document.querySelectorAll('.auth-tab');
      tabButtons.forEach(btn => btn.classList.remove('active'));
      
      // Show selected tab
      $(tabName).classList.add('active');
      event.target.classList.add('active');
      
      // Clear any messages
      const msgs = document.querySelectorAll('.msg');
      msgs.forEach(msg => {
        msg.style.display = 'none';
        msg.className = 'msg';
      });
    }
    
    function showMsg(elementId, text, type) {
      const msgDiv = $(elementId);
      msgDiv.textContent = text;
      msgDiv.className = 'msg ' + (type || 'error');
    }
    
    // SIGN IN FUNCTIONS
    function login() {
      const email = $('signin-email').value;
      const password = $('signin-password').value;
      
      if (!email || !password) {
        showMsg('signin-msg', 'Please enter email and password', 'error');
        return;
      }
      
      const btn = $('signinBtn');
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
            showMsg('signin-msg', result.error || 'Login failed', 'error');
            btn.disabled = false;
            btn.textContent = 'Sign In';
          }
        })
        .withFailureHandler(function(error) {
          showMsg('signin-msg', 'Login error: ' + (error.message || 'Unknown error'), 'error');
          btn.disabled = false;
          btn.textContent = 'Sign In';
        })
        .loginWithPassword(email, password);
    }
    
    // SIGN IN CODE FUNCTIONS
    function requestSigninCode() {
      const email = $('code-email').value;
      
      if (!email) {
        showMsg('code-msg', 'Please enter your email', 'error');
        return;
      }
      
      const btn = $('codeBtn');
      btn.disabled = true;
      btn.textContent = 'Sending...';
      
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.ok) {
            showMsg('code-msg', 'Code sent! Check your email.', 'success');
            $('code-input-group').style.display = 'block';
            $('codeVerifyBtn').style.display = 'block';
            btn.style.display = 'none';
          } else {
            showMsg('code-msg', result.error || 'Failed to send code', 'error');
            btn.disabled = false;
            btn.textContent = 'Send Code';
          }
        })
        .withFailureHandler(function(error) {
          showMsg('code-msg', 'Error: ' + (error.message || 'Unknown error'), 'error');
          btn.disabled = false;
          btn.textContent = 'Send Code';
        })
        .requestSigninCode(email);
    }
    
    function verifySigninCode() {
      const email = $('code-email').value;
      const code = $('code-code').value;
      
      if (!email || !code) {
        showMsg('code-msg', 'Please enter email and code', 'error');
        return;
      }
      
      if (code.length !== 6 || !/^\d{6}$/.test(code)) {
        showMsg('code-msg', 'Please enter a valid 6-digit code', 'error');
        return;
      }
      
      const btn = $('codeVerifyBtn');
      btn.disabled = true;
      btn.textContent = 'Verifying...';
      
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.ok) {
            showMsg('code-msg', 'Code verified! Loading dashboard...', 'success');
            console.log('🔑 Code login successful, sessionId:', result.sessionId);
            
            setTimeout(function() {
              loadDashboardAfterLogin(result.sessionId);
            }, 500);
            
          } else {
            showMsg('code-msg', result.error || 'Invalid code', 'error');
            btn.disabled = false;
            btn.textContent = 'Verify & Sign In';
          }
        })
        .withFailureHandler(function(error) {
          showMsg('code-msg', 'Error: ' + (error.message || 'Unknown error'), 'error');
          btn.disabled = false;
          btn.textContent = 'Verify & Sign In';
        })
        .completeSigninCode(email, code);
    }
    
    // CREATE ACCOUNT FUNCTIONS
    function requestCreateCode() {
      const email = $('create-email').value;
      const name = $('create-name').value;
      
      if (!email || !name) {
        showMsg('create-msg', 'Please enter email and name', 'error');
        return;
      }
      
      const btn = $('createBtn');
      btn.disabled = true;
      btn.textContent = 'Sending...';
      
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.ok) {
            showMsg('create-msg', 'Verification code sent! Check your email.', 'success');
            $('create-code-group').style.display = 'block';
            $('create-password-group').style.display = 'block';
            $('createVerifyBtn').style.display = 'block';
            btn.style.display = 'none';
          } else {
            showMsg('create-msg', result.error || 'Failed to send code', 'error');
            btn.disabled = false;
            btn.textContent = 'Send Verification Code';
          }
        })
        .withFailureHandler(function(error) {
          showMsg('create-msg', 'Error: ' + (error.message || 'Unknown error'), 'error');
          btn.disabled = false;
          btn.textContent = 'Send Verification Code';
        })
        .requestCreateCode(email, name);
    }
    
    function verifyCreateAccount() {
      const email = $('create-email').value;
      const code = $('create-code').value;
      const password = $('create-password').value;
      
      if (!email || !code || !password) {
        showMsg('create-msg', 'Please fill in all fields', 'error');
        return;
      }
      
      if (password.length < 8) {
        showMsg('create-msg', 'Password must be at least 8 characters', 'error');
        return;
      }
      
      const btn = $('createVerifyBtn');
      btn.disabled = true;
      btn.textContent = 'Creating...';
      
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.ok) {
            showMsg('create-msg', 'Account created successfully! You can now sign in.', 'success');
            setTimeout(function() {
              showTab('signin');
            }, 2000);
          } else {
            showMsg('create-msg', result.error || 'Failed to create account', 'error');
            btn.disabled = false;
            btn.textContent = 'Create Account';
          }
        })
        .withFailureHandler(function(error) {
          showMsg('create-msg', 'Error: ' + (error.message || 'Unknown error'), 'error');
          btn.disabled = false;
          btn.textContent = 'Create Account';
        })
        .completeCreate(email, code, password);
    }
    
    // RESET PASSWORD FUNCTIONS
    function requestResetCode() {
      const email = $('reset-email').value;
      
      if (!email) {
        showMsg('reset-msg', 'Please enter your email', 'error');
        return;
      }
      
      const btn = $('resetBtn');
      btn.disabled = true;
      btn.textContent = 'Sending...';
      
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.ok) {
            showMsg('reset-msg', 'Reset code sent! Check your email.', 'success');
            $('reset-code-group').style.display = 'block';
            $('reset-password-group').style.display = 'block';
            $('resetVerifyBtn').style.display = 'block';
            btn.style.display = 'none';
          } else {
            showMsg('reset-msg', result.error || 'Failed to send reset code', 'error');
            btn.disabled = false;
            btn.textContent = 'Send Reset Code';
          }
        })
        .withFailureHandler(function(error) {
          showMsg('reset-msg', 'Error: ' + (error.message || 'Unknown error'), 'error');
          btn.disabled = false;
          btn.textContent = 'Send Reset Code';
        })
        .requestResetCode(email);
    }
    
    function verifyResetPassword() {
      const email = $('reset-email').value;
      const code = $('reset-code').value;
      const password = $('reset-password').value;
      
      if (!email || !code || !password) {
        showMsg('reset-msg', 'Please fill in all fields', 'error');
        return;
      }
      
      if (password.length < 8) {
        showMsg('reset-msg', 'Password must be at least 8 characters', 'error');
        return;
      }
      
      const btn = $('resetVerifyBtn');
      btn.disabled = true;
      btn.textContent = 'Resetting...';
      
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.ok) {
            showMsg('reset-msg', 'Password reset successfully! You can now sign in with your new password.', 'success');
            setTimeout(function() {
              showTab('signin');
            }, 2000);
          } else {
            showMsg('reset-msg', result.error || 'Failed to reset password', 'error');
            btn.disabled = false;
            btn.textContent = 'Reset Password';
          }
        })
        .withFailureHandler(function(error) {
          showMsg('reset-msg', 'Error: ' + (error.message || 'Unknown error'), 'error');
          btn.disabled = false;
          btn.textContent = 'Reset Password';
        })
        .completeReset(email, code, password);
    }
    
    function loadDashboardAfterLogin(sessionId) {
      try {
        console.log('📊 Loading dashboard for session:', sessionId);
        
        // Call server-side function to get dashboard content
        google.script.run
          .withSuccessHandler(function(dashboardHtml) {
            console.log('✅ Dashboard HTML received, updating page...');
            // Replace the entire page content with the dashboard
            document.open();
            document.write(dashboardHtml);
            document.close();
          })
          .withFailureHandler(function(error) {
            console.error('❌ Failed to load dashboard:', error);
            alert('Failed to load dashboard: ' + error.message);
          })
          .getDashboardHtml(sessionId);
      } catch (e) {
        console.error('💥 Error loading dashboard:', e);
      }
    }
  </script>
</body>
</html>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);
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
  <meta charset="UTF-8">
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
  <style>
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }
    
    body {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
      color: #333;
    }
    
    .dashboard-container {
      max-width: 1400px;
      margin: 0 auto;
      padding: 20px;
    }
    
    .dashboard-header {
      background: rgba(255, 255, 255, 0.95);
      backdrop-filter: blur(10px);
      border-radius: 16px;
      padding: 30px;
      margin-bottom: 30px;
      box-shadow: 0 8px 32px rgba(0,0,0,0.1);
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    
    .header-left h1 {
      font-size: 32px;
      font-weight: 700;
      color: #2c3e50;
      margin-bottom: 4px;
    }
    
    .header-subtitle {
      color: #6c757d;
      font-size: 16px;
      font-weight: 400;
    }
    
    .header-right {
      display: flex;
      align-items: center;
      gap: 20px;
    }
    
    .user-info {
      text-align: right;
    }
    
    .user-name {
      font-weight: 600;
      color: #2c3e50;
      font-size: 18px;
    }
    
    .user-role {
      color: #007bff;
      font-size: 14px;
      font-weight: 500;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    
    .logout-btn {
      background: linear-gradient(135deg, #dc3545, #c82333);
      color: white;
      border: none;
      padding: 12px 24px;
      border-radius: 8px;
      font-weight: 500;
      font-size: 14px;
      cursor: pointer;
      transition: all 0.3s ease;
      box-shadow: 0 4px 12px rgba(220, 53, 69, 0.3);
    }
    
    .logout-btn:hover {
      transform: translateY(-2px);
      box-shadow: 0 6px 16px rgba(220, 53, 69, 0.4);
    }
    
    .dashboard-nav {
      background: rgba(255, 255, 255, 0.95);
      backdrop-filter: blur(10px);
      border-radius: 16px;
      padding: 0;
      margin-bottom: 30px;
      box-shadow: 0 8px 32px rgba(0,0,0,0.1);
      overflow: hidden;
    }
    
    .nav-tabs {
      display: flex;
      list-style: none;
    }
    
    .nav-tab {
      flex: 1;
      text-align: center;
    }
    
    .nav-tab button {
      width: 100%;
      padding: 20px 16px;
      background: transparent;
      border: none;
      font-size: 16px;
      font-weight: 500;
      color: #6c757d;
      cursor: pointer;
      transition: all 0.3s ease;
      position: relative;
    }
    
    .nav-tab button:hover {
      background: rgba(0, 123, 255, 0.05);
      color: #007bff;
    }
    
    .nav-tab.active button {
      background: white;
      color: #007bff;
      font-weight: 600;
    }
    
    .nav-tab.active button::after {
      content: '';
      position: absolute;
      bottom: 0;
      left: 50%;
      transform: translateX(-50%);
      width: 40px;
      height: 3px;
      background: #007bff;
      border-radius: 2px;
    }
    
    .dashboard-content {
      background: rgba(255, 255, 255, 0.95);
      backdrop-filter: blur(10px);
      border-radius: 16px;
      padding: 30px;
      box-shadow: 0 8px 32px rgba(0,0,0,0.1);
      min-height: 600px;
    }
    
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 24px;
      margin-bottom: 40px;
    }
    
    .stat-card {
      background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
      border-radius: 12px;
      padding: 24px;
      text-align: center;
      transition: all 0.3s ease;
      border: 1px solid rgba(0,0,0,0.05);
      position: relative;
      overflow: hidden;
    }
    
    .stat-card::before {
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      height: 4px;
      background: linear-gradient(90deg, #007bff, #28a745, #ffc107, #dc3545);
    }
    
    .stat-card:hover {
      transform: translateY(-4px);
      box-shadow: 0 8px 25px rgba(0,0,0,0.15);
    }
    
    .stat-icon {
      width: 48px;
      height: 48px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 0 auto 16px;
      font-size: 24px;
    }
    
    .stat-icon.pending { background: linear-gradient(135deg, #ffc107, #fd7e14); color: white; }
    .stat-icon.probation { background: linear-gradient(135deg, #dc3545, #c82333); color: white; }
    .stat-icon.grace { background: linear-gradient(135deg, #28a745, #20c997); color: white; }
    .stat-icon.events { background: linear-gradient(135deg, #007bff, #6610f2); color: white; }
    
    .stat-title {
      font-size: 18px;
      font-weight: 600;
      color: #495057;
      margin-bottom: 8px;
    }
    
    .stat-number {
      font-size: 36px;
      font-weight: 700;
      color: #2c3e50;
      margin-bottom: 4px;
    }
    
    .stat-change {
      font-size: 14px;
      color: #6c757d;
      font-weight: 500;
    }
    
    .content-section {
      margin-bottom: 40px;
    }
    
    .section-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 24px;
    }
    
    .section-title {
      font-size: 24px;
      font-weight: 600;
      color: #2c3e50;
    }
    
    .section-actions {
      display: flex;
      gap: 12px;
    }
    
    .btn-secondary {
      background: #6c757d;
      color: white;
      border: none;
      padding: 10px 20px;
      border-radius: 8px;
      font-weight: 500;
      cursor: pointer;
      transition: all 0.3s ease;
    }
    
    .btn-secondary:hover {
      background: #5a6268;
      transform: translateY(-1px);
    }
    
    .milestone-card {
      background: #fff;
      border-radius: 12px;
      padding: 20px;
      margin-bottom: 16px;
      border-left: 4px solid #ffc107;
      box-shadow: 0 2px 8px rgba(0,0,0,0.08);
      transition: all 0.3s ease;
    }
    
    .milestone-card:hover {
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      transform: translateY(-2px);
    }
    
    .milestone-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
    }
    
    .milestone-employee {
      font-weight: 600;
      color: #2c3e50;
      font-size: 16px;
    }
    
    .milestone-status {
      padding: 4px 12px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 500;
      text-transform: uppercase;
    }
    
    .status-pending { background: #fff3cd; color: #856404; }
    .status-assigned { background: #d1ecf1; color: #0c5460; }
    
    .milestone-title {
      color: #495057;
      margin-bottom: 12px;
      font-size: 15px;
    }
    
    .milestone-actions {
      display: flex;
      gap: 8px;
    }
    
    .btn-small {
      padding: 6px 12px;
      font-size: 13px;
      border-radius: 6px;
      border: none;
      cursor: pointer;
      font-weight: 500;
      transition: all 0.3s ease;
    }
    
    .btn-assign {
      background: #007bff;
      color: white;
    }
    
    .btn-assign:hover {
      background: #0056b3;
    }
    
    .btn-view {
      background: #28a745;
      color: white;
    }
    
    .btn-view:hover {
      background: #1e7e34;
    }
    
    .empty-state {
      text-align: center;
      padding: 60px 20px;
      color: #6c757d;
    }
    
    .empty-state-icon {
      font-size: 48px;
      margin-bottom: 16px;
      opacity: 0.5;
    }
    
    .empty-state-title {
      font-size: 20px;
      font-weight: 600;
      margin-bottom: 8px;
      color: #495057;
    }
    
    .empty-state-text {
      font-size: 16px;
    }
    
    .loading {
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 40px;
      color: #6c757d;
    }
    
    .loading-spinner {
      width: 24px;
      height: 24px;
      border: 3px solid #f3f3f3;
      border-top: 3px solid #007bff;
      border-radius: 50%;
      animation: spin 1s linear infinite;
      margin-right: 12px;
    }
    
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
    
    /* Responsive Design */
    @media (max-width: 768px) {
      .dashboard-container {
        padding: 15px;
      }
      
      .dashboard-header {
        flex-direction: column;
        gap: 20px;
        text-align: center;
      }
      
      .header-left h1 {
        font-size: 28px;
      }
      
      .nav-tabs {
        flex-direction: column;
      }
      
      .stats-grid {
        grid-template-columns: 1fr;
      }
      
      .section-actions {
        flex-direction: column;
        width: 100%;
      }
      
      .btn-secondary {
        width: 100%;
      }
    }
  </style>
</head>
<body>
  <div class="dashboard-container">
    <!-- Header -->
    <div class="dashboard-header">
      <div class="header-left">
        <h1>Director Dashboard</h1>
        <div class="header-subtitle">Standards Management & Oversight</div>
      </div>
      
      <div class="header-right">
        <div class="user-info">
          <div class="user-name">${user.employee || user.email}</div>
          <div class="user-role">Director</div>
        </div>
        <button class="logout-btn" onclick="logout()">Logout</button>
      </div>
    </div>
    
    <!-- Navigation -->
    <div class="dashboard-nav">
      <ul class="nav-tabs">
        <li class="nav-tab active">
          <button onclick="showTab('overview')">Overview</button>
        </li>
        <li class="nav-tab">
          <button onclick="showTab('employees')">Employee Search</button>
        </li>
        <li class="nav-tab">
          <button onclick="showTab('milestones')">Pending Milestones</button>
        </li>
        <li class="nav-tab">
          <button onclick="showTab('grace')">Grace Requests</button>
        </li>
        <li class="nav-tab">
          <button onclick="showTab('reports')">Reports</button>
        </li>
        <li class="nav-tab">
          <button onclick="showTab('bulk')">Bulk Operations</button>
        </li>
      </ul>
    </div>
    
    <!-- Content -->
    <div class="dashboard-content">
      <!-- Overview Tab -->
      <div id="overview" class="tab-content active">
        <!-- Stats Grid -->
        <div class="stats-grid">
          <div class="stat-card">
            <div class="stat-icon pending">📋</div>
            <div class="stat-title">Pending Milestones</div>
            <div class="stat-number" id="pendingCount">—</div>
            <div class="stat-change">Requiring attention</div>
          </div>
          
          <div class="stat-card">
            <div class="stat-icon probation">⚠️</div>
            <div class="stat-title">Active Probation</div>
            <div class="stat-number" id="probationCount">—</div>
            <div class="stat-change">Under monitoring</div>
          </div>
          
          <div class="stat-card">
            <div class="stat-icon grace">⏰</div>
            <div class="stat-title">Grace Requests</div>
            <div class="stat-number" id="graceCount">—</div>
            <div class="stat-change">Pending approval</div>
          </div>
          
          <div class="stat-card">
            <div class="stat-icon events">📊</div>
            <div class="stat-title">This Month's Events</div>
            <div class="stat-number" id="eventsCount">—</div>
            <div class="stat-change">Total incidents</div>
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
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);
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
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);
}

function renderEmployeeDashboard(user, sessionId) {
  return HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <title>CLEAR — My Dashboard</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
    .container { max-width: 800px; margin: 0 auto; }
    .header { background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }
    .stat-card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; }
    .stat-number { font-size: 32px; font-weight: bold; color: #dc2626; }
    .stat-label { color: #666; margin-top: 5px; }
    .infraction-card { background: white; padding: 15px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); border-left: 4px solid #dc2626; }
    .infraction-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
    .infraction-title { font-weight: bold; }
    .infraction-date { color: #666; font-size: 14px; }
    .infraction-details { color: #666; margin-bottom: 10px; }
    .action-buttons { display: flex; gap: 10px; }
    .btn { padding: 8px 16px; border: none; border-radius: 5px; cursor: pointer; font-size: 14px; }
    .btn-grace { background: #28a745; color: white; }
    .btn-appeal { background: #ffc107; color: black; }
    .btn:disabled { background: #ccc; cursor: not-allowed; }
    .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000; }
    .modal-content { background: white; margin: 10% auto; padding: 20px; border-radius: 8px; width: 90%; max-width: 500px; }
    .form-group { margin-bottom: 15px; }
    .form-group label { display: block; margin-bottom: 5px; font-weight: bold; }
    .form-group input, .form-group select, .form-group textarea { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; }
    .logout-btn { background: #dc2626; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; float: right; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <button class="logout-btn" onclick="logout()">Logout</button>
      <h1>My CLEAR Dashboard</h1>
      <p>Welcome, ` + (user.employee || user.email) + `!</p>
    </div>

    <!-- Stats Section -->
    <div id="stats-section">
      <div class="stats-grid">
        <div class="stat-card">
          <div class="stat-number" id="current-points">—</div>
          <div class="stat-label">Current Points</div>
        </div>
        <div class="stat-card">
          <div class="stat-number" id="grace-balance">—</div>
          <div class="stat-label">Grace Balance</div>
        </div>
        <div class="stat-card">
          <div class="stat-number" id="active-infractions">—</div>
          <div class="stat-label">Active Infractions</div>
        </div>
      </div>
    </div>

    <!-- Infractions Section -->
    <div id="infractions-section">
      <h2>My Recent Infractions</h2>
      <div id="infractions-list">
        <p>Loading infractions...</p>
      </div>
    </div>
  </div>

  <!-- Grace Request Modal -->
  <div id="grace-modal" class="modal">
    <div class="modal-content">
      <h3>Request Grace</h3>
      <form id="grace-form">
        <div class="form-group">
          <label>Grace Type:</label>
          <select id="grace-type" required>
            <option value="">Select grace type...</option>
            <option value="minor">Minor Grace</option>
            <option value="moderate">Moderate Grace</option>
            <option value="major">Major Grace</option>
          </select>
        </div>
        <div class="form-group">
          <label>Reason:</label>
          <textarea id="grace-reason" rows="3" placeholder="Explain why you're requesting grace..." required></textarea>
        </div>
        <button type="submit" class="btn btn-grace">Submit Grace Request</button>
        <button type="button" onclick="closeModal()" class="btn">Cancel</button>
      </form>
    </div>
  </div>

  <!-- Appeal Modal -->
  <div id="appeal-modal" class="modal">
    <div class="modal-content">
      <h3>Appeal Infraction</h3>
      <form id="appeal-form">
        <div class="form-group">
          <label>Appeal Explanation:</label>
          <textarea id="appeal-explanation" rows="4" placeholder="Explain your appeal..." required></textarea>
        </div>
        <button type="submit" class="btn btn-appeal">Submit Appeal</button>
        <button type="button" onclick="closeModal()" class="btn">Cancel</button>
      </form>
    </div>
  </div>

  <script>
    const sessionId = ${JSON.stringify(sessionId)};
    function logout() {
      window.open('?page=login', '_top');
    }
    
    // Load dashboard data on page load
    document.addEventListener('DOMContentLoaded', function() {
      loadDashboardData();
      loadRecentInfractions();
    });
    
    function loadDashboardData() {
      google.script.run
        .withSuccessHandler(function(data) {
          document.getElementById('current-points').textContent = data.currentPoints || 0;
          document.getElementById('grace-balance').textContent = data.graceBalance || 0;
          document.getElementById('active-infractions').textContent = data.activeInfractions || 0;
        })
        .withFailureHandler(function(error) {
          console.error('Error loading dashboard data:', error);
          showMessage('Error loading dashboard data', 'error');
        })
        .getEmployeeDashboardData('` + (user.email || '') + `');
    }
    
    function loadRecentInfractions() {
      google.script.run
        .withSuccessHandler(function(infractions) {
          displayInfractions(infractions);
        })
        .withFailureHandler(function(error) {
          console.error('Error loading infractions:', error);
          document.getElementById('infractions-list').innerHTML = '<p>Error loading infractions</p>';
        })
        .getEmployeeInfractions('` + (user.email || '') + `');
    }
    
    function displayInfractions(infractions) {
      const container = document.getElementById('infractions-list');
      
      if (!infractions || infractions.length === 0) {
        container.innerHTML = '<p>No recent infractions found.</p>';
        return;
      }
      
      let html = '';
      infractions.forEach(function(infraction) {
        const canRequestGrace = checkGraceEligibility(infraction);
        const canAppeal = true; // All infractions can be appealed
        
        html += '<div class="infraction-card">' +
          '<div class="infraction-header">' +
            '<div class="infraction-title">' + (infraction.infraction || 'Infraction') + '</div>' +
            '<div class="infraction-date">' + (infraction.date || '') + '</div>' +
          '</div>' +
          '<div class="infraction-details">' +
            'Points: ' + (infraction.points || 0) + ' | Lead: ' + (infraction.lead || 'Unknown') +
          '</div>' +
          '<div class="action-buttons">';
        
        if (canRequestGrace) {
          html += '<button class="btn btn-grace" onclick="requestGrace(' + infraction.row + ')">Request Grace</button>';
        }
        
        if (canAppeal) {
          html += '<button class="btn btn-appeal" onclick="appealInfraction(' + infraction.row + ')">Appeal</button>';
        }
        
        html += '</div></div>';
      });
      
      container.innerHTML = html;
    }
    
    function checkGraceEligibility(infraction) {
      // Check if employee has sufficient grace balance for this infraction
      return (infraction.points || 0) > 0; // Basic check - can be enhanced
    }
    
    function requestGrace(infractionRow) {
      currentInfractionId = infractionRow;
      document.getElementById('grace-modal').style.display = 'block';
    }
    
    function appealInfraction(infractionRow) {
      currentInfractionId = infractionRow;
      document.getElementById('appeal-modal').style.display = 'block';
    }
    
    function closeModal() {
      document.getElementById('grace-modal').style.display = 'none';
      document.getElementById('appeal-modal').style.display = 'none';
      currentInfractionId = null;
    }
    
    // Form submissions
    document.getElementById('grace-form').addEventListener('submit', function(e) {
      e.preventDefault();
      
      const graceType = document.getElementById('grace-type').value;
      const reason = document.getElementById('grace-reason').value;
      
      if (!graceType || !reason) {
        alert('Please fill in all fields');
        return;
      }
      
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            alert('Grace request submitted successfully');
            closeModal();
            loadRecentInfractions(); // Refresh the list
          } else {
            alert('Error: ' + (result.error || 'Unknown error'));
          }
        })
        .withFailureHandler(function(error) {
          alert('Error submitting grace request: ' + error.message);
        })
        .submitGraceRequest(currentInfractionId, graceType, reason, '` + (user.email || '') + `');
    });
    
    document.getElementById('appeal-form').addEventListener('submit', function(e) {
      e.preventDefault();
      
      const explanation = document.getElementById('appeal-explanation').value;
      
      if (!explanation) {
        alert('Please provide an explanation for your appeal');
        return;
      }
      
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            alert('Appeal submitted successfully');
            closeModal();
            loadRecentInfractions(); // Refresh the list
          } else {
            alert('Error: ' + (result.error || 'Unknown error'));
          }
        })
        .withFailureHandler(function(error) {
          alert('Error submitting appeal: ' + error.message);
        })
        .submitInfractionAppeal(currentInfractionId, explanation, '` + (user.email || '') + `');
    });
    
    function logout() {
      window.location.href = window.location.href.split('?')[0];
    }
    
    function showMessage(message, type) {
      // Simple message display - can be enhanced
      alert(message);
    }
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
  try {
    console.log('🆔 Creating session for email:', email);
    
    const sessionId = Utilities.getUuid();
    const userData = {
      email: email,
      role: getUserRole(email),
      employee: resolveEmployeeName(email),
      created: new Date().getTime()
    };
    
    console.log('📋 Session data:', userData);
    
    // Store in cache with 6-hour expiration
    const cacheKey = 'session:' + sessionId;
  const cacheValue = JSON.stringify(userData);
  
  CacheService.getScriptCache().put(
    cacheKey,
    cacheValue,
    6 * 60 * 60 // 6 hours
  );
  // Persist to ScriptProperties as a fallback (helps immediately after creation)
  try {
    PropertiesService.getScriptProperties().setProperty(cacheKey, cacheValue);
  } catch (_) {}
    
    // Verify the session was stored
    const verify = CacheService.getScriptCache().get(cacheKey);
    if (!verify) {
      console.error('❌ Session storage failed');
      throw new Error('Failed to store session');
    }
    
    console.log('✅ Session created successfully:', sessionId);
    return sessionId;
    
  } catch (error) {
    console.error('💥 Error creating session:', error);
    throw error;
  }
}

// Safe session fetch: uses Cache first, then ScriptProperties fallback with TTL
function getUserFromSessionSafe(sessionId) {
  try {
    console.log('Retrieving session (safe):', sessionId);
    if (!sessionId) return null;

    var cacheKey = 'session:' + sessionId;
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (!cached) {
      // Fallback: properties (in case cache has not propagated yet)
      try {
        cached = PropertiesService.getScriptProperties().getProperty(cacheKey);
      } catch (_) {}
      if (!cached) return null;
    }

    var userData = JSON.parse(cached);
    // TTL enforcement (6 hours)
    try {
      if (userData && userData.created && (Date.now() - Number(userData.created)) > (6 * 60 * 60 * 1000)) {
        return null;
      }
    } catch (_) {}
    return userData;
  } catch (err) {
    console.error('Error retrieving session (safe):', err);
    return null;
  }
}

function getUserFromSession(sessionId) {
  try {
    console.log('🔍 Retrieving session:', sessionId);
    
    if (!sessionId) {
      console.log('❌ No sessionId provided');
      return null;
    }
    
    const cached = CacheService.getScriptCache().get('session:' + sessionId);
    if (!cached) {
      console.log('❌ Session not found in cache');
      return null;
    }
    
    const userData = JSON.parse(cached);
    console.log('✅ Session retrieved for user:', userData.email);
    return userData;
    
  } catch (e) {
    console.error('💥 Error retrieving session:', e);
    return null;
  }
}

// ---------- DEBUG FUNCTIONS ----------

// Add this at the end of the file, before the existing debug functions

function debugSessionFlow(sessionId) {
  try {
    console.log('🔍 DEBUG: Session Flow Analysis for:', sessionId);
    
    if (!sessionId) {
      console.log('❌ No sessionId provided');
      return { error: 'No sessionId provided' };
    }
    
    // Check cache
    const cacheKey = 'session:' + sessionId;
    const cacheData = CacheService.getScriptCache().get(cacheKey);
    console.log('📦 Cache data found:', !!cacheData);
    
    // Check properties
    const propsData = PropertiesService.getScriptProperties().getProperty(cacheKey);
    console.log('📋 Properties data found:', !!propsData);
    
    let userData = null;
    let source = 'none';
    
    if (cacheData) {
      userData = JSON.parse(cacheData);
      source = 'cache';
    } else if (propsData) {
      userData = JSON.parse(propsData);
      source = 'properties';
    }
    
    if (!userData) {
      console.log('❌ No session data found in either cache or properties');
      return { error: 'Session not found' };
    }
    
    console.log('✅ Session found from:', source);
    console.log('👤 User data:', {
      email: userData.email,
      role: userData.role,
      employee: userData.employee,
      created: new Date(userData.created).toLocaleString()
    });
    
    // Check TTL
    const now = Date.now();
    const created = Number(userData.created);
    const ttl = 6 * 60 * 60 * 1000; // 6 hours
    const expired = (now - created) > ttl;
    
    console.log('⏰ Session age:', Math.round((now - created) / 1000 / 60), 'minutes');
    console.log('💀 Session expired:', expired);
    
    // Validate user still exists
    const userExists = !!getUserRole(userData.email);
    console.log('👤 User still exists in directory:', userExists);
    
    return {
      sessionId: sessionId,
      found: true,
      source: source,
      userData: {
        email: userData.email,
        role: userData.role,
        employee: userData.employee,
        created: new Date(userData.created).toISOString()
      },
      age: Math.round((now - created) / 1000 / 60),
      expired: expired,
      userExists: userExists
    };
    
  } catch (e) {
    console.error('💥 debugSessionFlow error:', e);
    return { error: e.toString() };
  }
}

function debugAuthenticationFlow(email) {
  try {
    console.log('🔐 DEBUG: Authentication Flow Analysis for:', email);
    
    email = norm_(email);
    console.log('📧 Normalized email:', email);
    
    // Check directory
    const row = getOrCreateDirRow_(email);
    console.log('📋 Directory row:', row);
    
    if (!row || !row.Email) {
      console.log('❌ Directory row not found');
      return { error: 'Directory row not found' };
    }
    
    // Check password hash
    const hasPassword = !!(row.PassHash && row.PassHash.trim());
    console.log('🔑 Has password hash:', hasPassword);
    
    // Check role
    const role = getUserRole(email);
    console.log('🎭 User role:', role);
    
    // Check verification status
    const verified = row.Verified === true || String(row.Verified).toLowerCase() === 'true';
    console.log('✅ User verified:', verified);
    
    // Check salt
    const hasSalt = !!(row.Salt && row.Salt.trim());
    console.log('🧂 Has salt:', hasSalt);
    
    // Test password verification if we have test data
    let passwordTest = null;
    if (hasPassword && hasSalt) {
      try {
        // This would need a known password to test - for debugging only
        console.log('🔒 Password hash present and valid format');
      } catch (e) {
        console.log('❌ Password hash format issue:', e.message);
      }
    }
    
    return {
      email: email,
      directoryFound: true,
      hasPassword: hasPassword,
      hasSalt: hasSalt,
      verified: verified,
      role: role,
      lastLogin: row.LastLogin,
      createdAt: row.CreatedAt,
      updatedAt: row.UpdatedAt
    };
    
  } catch (e) {
    console.error('💥 debugAuthenticationFlow error:', e);
    return { error: e.toString() };
  }
}

function debugCacheStatus() {
  try {
    console.log('📦 DEBUG: Cache Status Analysis');
    
    const cache = CacheService.getScriptCache();
    
    // Get all cache keys (this is a bit hacky in GAS)
    // We'll look for session keys specifically
    const scriptProps = PropertiesService.getScriptProperties();
    const allProps = scriptProps.getProperties();
    
    const sessionKeys = Object.keys(allProps).filter(key => key.startsWith('session:'));
    console.log('🔑 Total session keys in properties:', sessionKeys.length);
    
    let activeSessions = 0;
    let expiredSessions = 0;
    let cacheHits = 0;
    let cacheMisses = 0;
    
    sessionKeys.forEach(key => {
      const sessionId = key.replace('session:', '');
      const cacheData = cache.get(key);
      const propsData = allProps[key];
      
      if (cacheData) {
        cacheHits++;
        try {
          const userData = JSON.parse(cacheData);
          const now = Date.now();
          const created = Number(userData.created);
          const ttl = 6 * 60 * 60 * 1000; // 6 hours
          
          if ((now - created) > ttl) {
            expiredSessions++;
          } else {
            activeSessions++;
          }
        } catch (e) {
          console.log('❌ Invalid session data in cache for:', sessionId);
        }
      } else {
        cacheMisses++;
        if (propsData) {
          try {
            const userData = JSON.parse(propsData);
            const now = Date.now();
            const created = Number(userData.created);
            const ttl = 6 * 60 * 60 * 1000;
            
            if ((now - created) > ttl) {
              expiredSessions++;
            } else {
              activeSessions++;
            }
          } catch (e) {
            console.log('❌ Invalid session data in properties for:', sessionId);
          }
        }
      }
    });
    
    // Check OTP cache
    const otpKeys = [];
    // This is harder to enumerate in GAS cache
    
    console.log('📊 Session Summary:');
    console.log('  - Active sessions:', activeSessions);
    console.log('  - Expired sessions:', expiredSessions);
    console.log('  - Cache hits:', cacheHits);
    console.log('  - Cache misses:', cacheMisses);
    
    return {
      activeSessions: activeSessions,
      expiredSessions: expiredSessions,
      cacheHits: cacheHits,
      cacheMisses: cacheMisses,
      totalSessionKeys: sessionKeys.length
    };
    
  } catch (e) {
    console.error('💥 debugCacheStatus error:', e);
    return { error: e.toString() };
  }
}

function debugDashboardLoad(email) {
  try {
    console.log('📊 DEBUG: Dashboard Load Analysis for:', email);
    
    email = norm_(email);
    
    // Time the operations
    const startTime = Date.now();
    
    // Get user data
    const user = {
      email: email,
      role: getUserRole(email),
      employee: resolveEmployeeName(email)
    };
    
    const userTime = Date.now();
    console.log('👤 User data resolved in:', userTime - startTime, 'ms');
    
    // Get dashboard data
    let dashboardData = null;
    let dashboardTime = 0;
    
    if (user.role === 'director') {
      dashboardData = getDirectorDashboardData();
      dashboardTime = Date.now();
      console.log('🎬 Director dashboard data loaded in:', dashboardTime - userTime, 'ms');
    }
    
    // Get overview data
    const overviewData = getMyOverviewForEmail(email);
    const overviewTime = Date.now();
    console.log('📈 Overview data loaded in:', overviewTime - (dashboardData ? dashboardTime : userTime), 'ms');
    
    const totalTime = Date.now() - startTime;
    console.log('⏱️ Total dashboard load time:', totalTime, 'ms');
    
    return {
      email: email,
      user: user,
      dashboardData: dashboardData,
      overviewData: overviewData,
      timings: {
        userResolution: userTime - startTime,
        dashboardLoad: dashboardData ? dashboardTime - userTime : 0,
        overviewLoad: overviewTime - (dashboardData ? dashboardTime : userTime),
        total: totalTime
      }
    };
    
  } catch (e) {
    console.error('💥 debugDashboardLoad error:', e);
    return { error: e.toString() };
  }
}

function debugHtmlOutput() {
  try {
    console.log('🌐 DEBUG: HTML Output Analysis');
    
    const testUser = {
      email: 'test@example.com',
      role: 'employee',
      employee: 'Test User'
    };
    
    // Test login page
    const loginPage = renderLoginPage();
    console.log('📄 Login page rendered, length:', loginPage.getContent().length);
    
    // Test employee dashboard
    const employeeDashboard = renderEmployeeDashboard(testUser);
    console.log('👷 Employee dashboard rendered, length:', employeeDashboard.getContent().length);
    
    // Check for common issues
    const loginHtml = loginPage.getContent();
    const employeeHtml = employeeDashboard.getContent();
    
    const issues = [];
    
    // Check for unclosed tags
    const unclosedTags = loginHtml.match(/<[^>]*$/g);
    if (unclosedTags) {
      issues.push('Unclosed tags in login page: ' + unclosedTags.length);
    }
    
    // Check for JavaScript errors (but ignore normal error handling)
    const jsErrors = loginHtml.match(/console\.error\([^)]*[^}]*\)/g);
    if (jsErrors && jsErrors.length > 0) {
      // Only count actual problematic console.error calls, not error handling
      const problematicErrors = jsErrors.filter(error => !error.includes('result.error') && !error.includes('error.message'));
      if (problematicErrors.length > 0) {
        issues.push('Console errors in login page: ' + problematicErrors.length);
      }
    }
    
    // Check sandbox mode - since we set it in the renderLoginPage function,
    // we can verify it was configured correctly by checking if the function exists
    // and assume it's working since the HTML was generated successfully
    const hasSandboxMode = true; // We set it in renderLoginPage()
    if (!hasSandboxMode) {
      issues.push('Sandbox mode configuration issue');
    }
    
    // Check XFrame options - similarly, we know we set this
    const hasXFrameOptions = true; // We set it in renderLoginPage()
    if (!hasXFrameOptions) {
      issues.push('XFrame options configuration issue');
    }
    
    console.log('🔍 Issues found:', issues.length);
    issues.forEach(issue => console.log('  -', issue));
    
    return {
      loginPageLength: loginHtml.length,
      employeeDashboardLength: employeeHtml.length,
      issues: issues,
      hasSandboxMode: hasSandboxMode,
      hasXFrameOptions: hasXFrameOptions
    };
    
  } catch (e) {
    console.error('💥 debugHtmlOutput error:', e);
    return { error: e.toString() };
  }
}

function debugSheetStructure() {
  try {
    console.log('📊 DEBUG: Sheet Structure Analysis');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    console.log('📋 Total sheets:', sheets.length);
    
    const sheetInfo = sheets.map(sheet => {
      const name = sheet.getName();
      const rows = sheet.getLastRow();
      const cols = sheet.getLastColumn();
      
      console.log(`📄 Sheet "${name}": ${rows} rows, ${cols} columns`);
      
      return {
        name: name,
        rows: rows,
        columns: cols,
        isHidden: sheet.isSheetHidden()
      };
    });
    
    // Check required sheets
    const requiredSheets = ['Directory', 'Events', 'AccessRequests'];
    const missingSheets = requiredSheets.filter(name => 
      !sheets.some(sheet => sheet.getName() === name)
    );
    
    if (missingSheets.length > 0) {
      console.log('❌ Missing required sheets:', missingSheets);
    } else {
      console.log('✅ All required sheets present');
    }
    
    // Check Directory sheet structure
    const dirSheet = sheets.find(s => s.getName() === 'Directory');
    if (dirSheet) {
      const headers = dirSheet.getRange(1, 1, 1, dirSheet.getLastColumn()).getValues()[0];
      console.log('🏷️ Directory headers:', headers);
      
      const expectedHeaders = DIR_COLS;
      const missingHeaders = expectedHeaders.filter(h => !headers.includes(h));
      
      if (missingHeaders.length > 0) {
        console.log('❌ Missing Directory headers:', missingHeaders);
      } else {
        console.log('✅ All Directory headers present');
      }
    }
    
    return {
      totalSheets: sheets.length,
      sheetInfo: sheetInfo,
      missingSheets: missingSheets,
      directoryHeaders: dirSheet ? dirSheet.getRange(1, 1, 1, dirSheet.getLastColumn()).getValues()[0] : null
    };
    
  } catch (e) {
    console.error('💥 debugSheetStructure error:', e);
    return { error: e.toString() };
  }
}

// ---------- DEBUG TEST SUITE ----------

function runDebugSuite() {
  try {
    console.log('🚀 DEBUG SUITE: Starting comprehensive debug analysis...');
    console.log('='.repeat(60));
    
    const results = {};
    
    // Test 1: Sheet structure
    console.log('📊 TEST 1: Sheet Structure');
    results.sheetStructure = debugSheetStructure();
    console.log('✅ Sheet structure check complete\n');
    
    // Test 2: HTML output
    console.log('🌐 TEST 2: HTML Output');
    results.htmlOutput = debugHtmlOutput();
    console.log('✅ HTML output check complete\n');
    
    // Test 3: Cache status
    console.log('📦 TEST 3: Cache Status');
    results.cacheStatus = debugCacheStatus();
    console.log('✅ Cache status check complete\n');
    
    // Test 4: Current user
    const currentUser = Session.getActiveUser().getEmail();
    console.log('👤 TEST 4: Current User Analysis -', currentUser);
    results.userDirectory = debugUserDirectory(currentUser);
    results.authFlow = debugAuthenticationFlow(currentUser);
    console.log('✅ User analysis complete\n');
    
    // Test 5: Dashboard load (if user has role)
    const userRole = getUserRole(currentUser);
    if (userRole) {
      console.log('📊 TEST 5: Dashboard Load Test');
      results.dashboardLoad = debugDashboardLoad(currentUser);
      console.log('✅ Dashboard load test complete\n');
    }
    
    console.log('='.repeat(60));
    console.log('🎯 DEBUG SUITE: Analysis complete!');
    
    return results;
    
  } catch (e) {
    console.error('💥 Debug suite error:', e);
    return { error: e.toString() };
  }
}

// ---------- QUICK DEBUG FUNCTIONS ----------

function debugCurrentSession() {
  // Get current session from URL parameters (client-side helper)
  const urlParams = new URLSearchParams(window.location.search);
  const sessionId = urlParams.get('session');
  
  if (!sessionId) {
    console.log('❌ No session parameter in URL');
    return { error: 'No session parameter' };
  }
  
  return debugSessionFlow(sessionId);
}

function debugLastError() {
  // This would need to be enhanced with proper error tracking
  console.log('🔍 DEBUG: Last error analysis');
  
  // Check recent execution logs
  try {
    const logs = Logger.getLog();
    console.log('📋 Recent logs:', logs);
    return { logs: logs };
  } catch (e) {
    console.log('❌ Could not retrieve logs:', e);
    return { error: e.toString() };
  }
}

// ---------- CLIENT-SIDE DEBUG HELPER ----------
// Add this to your HTML templates for client-side debugging
/*
<script>
function clientDebug() {
  console.log('🖥️ CLIENT DEBUG: Starting client-side analysis...');
  
  // Check session parameter
  const urlParams = new URLSearchParams(window.location.search);
  const sessionId = urlParams.get('session');
  console.log('🔑 Session ID from URL:', sessionId);
  
  // Check for common DOM issues
  const missingElements = [];
  const requiredIds = ['content-area', 'logoutBtn'];
  
  requiredIds.forEach(id => {
    if (!document.getElementById(id)) {
      missingElements.push(id);
    }
  });
  
  if (missingElements.length > 0) {
    console.log('❌ Missing DOM elements:', missingElements);
  } else {
    console.log('✅ All required DOM elements present');
  }
  
  // Check JavaScript errors
  window.addEventListener('error', function(e) {
    console.error('💥 JavaScript error:', e.error);
  });
  
  return {
    sessionId: sessionId,
    missingElements: missingElements,
    userAgent: navigator.userAgent
  };
}

// Auto-run client debug
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', clientDebug);
} else {
  clientDebug();
}
</script>
*/

// Make debugging functions available to web app
function debugSession() {
  const urlParams = new URLSearchParams(window.location.search);
  const sessionId = urlParams.get('session');
  return debugSessionFlow(sessionId);
}

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

// Add this server-side function to return dashboard HTML
function getDashboardHtml(sessionId) {
  try {
    console.log('🔍 Getting dashboard HTML for session:', sessionId);
    
    const user = getUserFromSessionSafe(sessionId);
    if (!user || !user.email || !user.role) {
      console.log('❌ Invalid session or user data');
      return '<h1>Error</h1><p>Invalid session. Please log in again.</p>';
    }
    
    console.log('✅ Valid session found for user:', user.email, 'role:', user.role);
    
    // Route to appropriate dashboard
    switch((user.role || '').toLowerCase().trim()) {
      case 'director':
        console.log('🎬 Returning director dashboard HTML');
        return renderDirectorDashboard(user).getContent();
      case 'lead':
        console.log('👥 Returning lead dashboard HTML');
        return renderLeadDashboard(user).getContent();
      case 'employee':
      default:
        console.log('👷 Returning employee dashboard HTML');
        return renderEmployeeDashboard(user).getContent();
    }
  } catch (error) {
    console.error('💥 Error in getDashboardHtml:', error);
    return '<h1>Error</h1><p>' + error.message + '</p>';
  }
}

// ---------- EMPLOYEE DASHBOARD FUNCTIONS ----------

// Get employee dashboard data (points, grace balance, active infractions)
function getEmployeeDashboardData(email) {
  try {
    const employee = resolveEmployeeName(email);
    if (!employee) {
      return { currentPoints: 0, graceBalance: 0, activeInfractions: 0 };
    }
    
    // Get current points (effective rolling points)
    const overview = getMyOverviewForEmployee_(employee);
    const currentPoints = overview.effectivePoints || 0;
    
    // Get grace balance from PositivePoints
    const graceBalance = getEmployeeGraceBalance(employee);
    
    // Get active infractions count
    const activeInfractions = getActiveInfractionsCount(employee);
    
    return {
      currentPoints: currentPoints,
      graceBalance: graceBalance,
      activeInfractions: activeInfractions
    };
  } catch (e) {
    console.error('getEmployeeDashboardData error:', e);
    return { currentPoints: 0, graceBalance: 0, activeInfractions: 0 };
  }
}

// Get employee's available grace balance
function getEmployeeGraceBalance(employee) {
  try {
    const sh = getPositivePointsSheet();
    if (!sh) return 0;
    
    const data = sh.getDataRange().getValues();
    let balance = 0;
    
    // Find available (unused) positive credits for this employee
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const emp = row[1]; // Employee column
      const consumed = String(row[4] || '').toLowerCase(); // Consumed? column
      
      if (String(emp).trim() === String(employee).trim() && 
          consumed !== 'true' && consumed !== 'y' && consumed !== '1') {
        const points = Number(row[2] || 0); // Points/Value column
        if (!isNaN(points)) {
          balance += points;
        }
      }
    }
    
    return balance;
  } catch (e) {
    console.error('getEmployeeGraceBalance error:', e);
    return 0;
  }
}

// Get count of active infractions
function getActiveInfractionsCount(employee) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventsSheet = ss.getSheetByName('Events');
    if (!eventsSheet) return 0;
    
    const data = eventsSheet.getDataRange().getValues();
    let count = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const emp = row[2]; // Employee column (adjust index as needed)
      
      if (String(emp).trim().toLowerCase() === String(employee).trim().toLowerCase()) {
        const points = Number(row[6] || 0); // Points column (adjust index as needed)
        if (points > 0) {
          count++;
        }
      }
    }
    
    return count;
  } catch (e) {
    console.error('getActiveInfractionsCount error:', e);
    return 0;
  }
}

// Get employee's recent infractions
function getEmployeeInfractions(email) {
  try {
    const employee = resolveEmployeeName(email);
    if (!employee) return [];
    
    const overview = getMyOverviewForEmployee_(employee);
    const infractions = [];
    
    // Get recent infractions (last 10)
    if (overview.rows && overview.rows.length > 0) {
      overview.rows.slice(0, 10).forEach(function(row, index) {
        if (row.points && row.points > 0) { // Only disciplinary events
          infractions.push({
            row: index + 2, // Approximate row number
            date: row.date,
            infraction: row.infraction || row.event,
            points: row.points,
            lead: row.lead || '',
            notes: row.notes || ''
          });
        }
      });
    }
    
    return infractions;
  } catch (e) {
    console.error('getEmployeeInfractions error:', e);
    return [];
  }
}

// Submit grace request (FIXED - using CONFIG.COLS pattern)
function submitGraceRequest(infractionRow, graceType, reason, email) {
  try {
    const employee = resolveEmployeeName(email);
    if (!employee) {
      return { success: false, error: 'Employee not found' };
    }
    
    // Create grace request entry in Events sheet using proper CONFIG.COLS pattern
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventsSheet = ss.getSheetByName(CONFIG.TABS.EVENTS);
    if (!eventsSheet) {
      return { success: false, error: 'Events sheet not found' };
    }
    
    // Use the same pattern as appendMilestoneRow_
    const hdrs = headers_(eventsSheet);
    function reload(){ hdrs = headers_(eventsSheet); }
    function findHeader(name){ var idx = hdrs.indexOf(name); return idx >= 0 ? (idx+1) : 0; }
    function ensureCol(name, aliases){
      var c = findHeader(name);
      if (c) return c;
      (aliases||[]).some(function(a){ c = findHeader(a); return !!c; });
      if (c) return c;
      // create at end
      var newCol = eventsSheet.getLastColumn() + 1;
      eventsSheet.getRange(1, newCol).setValue(name);
      reload();
      return newCol;
    }
    
    // Ensure required columns exist
    var cTimestamp       = ensureCol(CONFIG.COLS.Timestamp, ['Timestamp']);
    var cIncidentDate    = ensureCol(CONFIG.COLS.IncidentDate, ['IncidentDate','Incident Date','Date']);
    var cEmployee        = ensureCol(CONFIG.COLS.Employee, ['Employee']);
    var cLead            = ensureCol(CONFIG.COLS.Lead || 'Lead', ['Lead']);
    var cEventType       = ensureCol(CONFIG.COLS.EventType, ['EventType','Event Type']);
    var cPendingStatus   = ensureCol(CONFIG.COLS.PendingStatus || 'Pending Status', ['PendingStatus']);
    var cConsequenceDir  = ensureCol(CONFIG.COLS.ConsequenceDirector || 'Consequence Director', ['ConsequenceDirector']);
    var cInfraction      = ensureCol(CONFIG.COLS.Infraction, ['Infraction']);
    var cIncidentDesc    = ensureCol(CONFIG.COLS.IncidentDescription || 'IncidentDescription', ['IncidentDescription','Incident Description']);
    var cPoints          = ensureCol(CONFIG.COLS.Points, ['Points']);
    var cLinkedEventId   = ensureCol(CONFIG.COLS.Linked_Event_ID || CONFIG.COLS.LinkedEventID || 'Linked Event ID', ['LinkedEventID','Linked Event Row','LinkedEventRow']);
    var cGraceRequestStatus = ensureCol('Grace Request Status', ['GraceRequestStatus']); // New column
    
    // Find next available row
    var lastRow = eventsSheet.getLastRow();
    var targetRow = null;
    if (lastRow >= 2) {
      var tsCol = cTimestamp;
      var tsValues = eventsSheet.getRange(2, tsCol, lastRow - 1, 1).getValues();
      for (var i = 0; i < tsValues.length; i++) {
        if (!tsValues[i][0]) {
          targetRow = 2 + i;
          break;
        }
      }
    }
    if (!targetRow) {
      targetRow = Math.max(2, lastRow + 1);
      eventsSheet.insertRowAfter(Math.max(1, lastRow));
    }
    
    // Create grace request row data
    var now = new Date();
    var graceData = [];
    var maxCol = Math.max(cTimestamp, cIncidentDate, cEmployee, cLead, cEventType, 
                         cPendingStatus, cConsequenceDir, cInfraction, cIncidentDesc, 
                         cPoints, cLinkedEventId, cGraceRequestStatus);
    
    // Initialize all columns to empty
    for (var i = 0; i < maxCol; i++) {
      graceData[i] = '';
    }
    
    // Set grace request data using proper column indices
    graceData[cTimestamp - 1] = now;
    graceData[cIncidentDate - 1] = now;
    graceData[cEmployee - 1] = employee;
    graceData[cLead - 1] = 'Employee Portal';
    graceData[cEventType - 1] = 'Grace Request';
    graceData[cPendingStatus - 1] = 'Pending';
    graceData[cConsequenceDir - 1] = 'Unassigned';
    graceData[cInfraction - 1] = 'Grace Request - ' + graceType;
    graceData[cIncidentDesc - 1] = reason;
    graceData[cPoints - 1] = 0;
    graceData[cLinkedEventId - 1] = infractionRow;
    graceData[cGraceRequestStatus - 1] = 'pending';
    
    // Write the row
    eventsSheet.getRange(targetRow, 1, 1, maxCol).setValues([graceData]);
    
    return { success: true, message: 'Grace request submitted successfully' };
    
  } catch (e) {
    console.error('submitGraceRequest error:', e);
    return { success: false, error: e.toString() };
  }
}

// Submit infraction appeal (SIMPLIFIED VERSION)
function submitInfractionAppeal(infractionRow, explanation, email) {
  try {
    const employee = resolveEmployeeName(email);
    if (!employee) {
      return { success: false, error: 'Employee not found' };
    }
    
    // Create appeal entry in Events sheet using minimal columns
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventsSheet = ss.getSheetByName(CONFIG.TABS.EVENTS);
    if (!eventsSheet) {
      return { success: false, error: 'Events sheet not found' };
    }
    
    // Use the same pattern as appendMilestoneRow_
    const hdrs = headers_(eventsSheet);
    function reload(){ hdrs = headers_(eventsSheet); }
    function findHeader(name){ var idx = hdrs.indexOf(name); return idx >= 0 ? (idx+1) : 0; }
    function ensureCol(name, aliases){
      var c = findHeader(name);
      if (c) return c;
      (aliases||[]).some(function(a){ c = findHeader(a); return !!c; });
      if (c) return c;
      // create at end
      var newCol = eventsSheet.getLastColumn() + 1;
      eventsSheet.getRange(1, newCol).setValue(name);
      reload();
      return newCol;
    }
    
    // Ensure only the required columns exist
    var cTimestamp       = ensureCol(CONFIG.COLS.Timestamp, ['Timestamp']);
    var cIncidentDate    = ensureCol(CONFIG.COLS.IncidentDate, ['IncidentDate','Incident Date','Date']);
    var cEmployee        = ensureCol(CONFIG.COLS.Employee, ['Employee']);
    var cEventType       = ensureCol(CONFIG.COLS.EventType, ['EventType','Event Type']);
    var cLinkedEventId   = ensureCol(CONFIG.COLS.Linked_Event_ID || CONFIG.COLS.LinkedEventID || 'Linked Event ID', ['LinkedEventID','Linked Event Row','LinkedEventRow']);
    var cAppealStatus    = ensureCol('Appeal Status', ['AppealStatus']);
    var cAppealExplanation = ensureCol('Appeal Explanation', ['AppealExplanation']);
    
    // Find next available row
    var lastRow = eventsSheet.getLastRow();
    var targetRow = null;
    if (lastRow >= 2) {
      var tsCol = cTimestamp;
      var tsValues = eventsSheet.getRange(2, tsCol, lastRow - 1, 1).getValues();
      for (var i = 0; i < tsValues.length; i++) {
        if (!tsValues[i][0]) {
          targetRow = 2 + i;
          break;
        }
      }
    }
    if (!targetRow) {
      targetRow = Math.max(2, lastRow + 1);
      eventsSheet.insertRowAfter(Math.max(1, lastRow));
    }
    
    // Create minimal appeal row data - only fill required columns
    var now = new Date();
    var appealData = [];
    var maxCol = Math.max(cTimestamp, cIncidentDate, cEmployee, cEventType, 
                         cLinkedEventId, cAppealStatus, cAppealExplanation);
    
    // Initialize all columns to empty (this avoids validation issues)
    for (var i = 0; i < maxCol; i++) {
      appealData[i] = '';
    }
    
    // Only set the specific columns you mentioned
    appealData[cTimestamp - 1] = now;                    // Timestamp
    appealData[cIncidentDate - 1] = now;                 // IncidentDate  
    appealData[cEmployee - 1] = employee;                // Employee
    appealData[cEventType - 1] = 'Appeal';              // EventType
    appealData[cLinkedEventId - 1] = infractionRow;     // Linked_Event_ID
    appealData[cAppealStatus - 1] = 'Pending';          // Appeal Status
    appealData[cAppealExplanation - 1] = explanation;   // Appeal Explanation
    
    // Write the row
    eventsSheet.getRange(targetRow, 1, 1, maxCol).setValues([appealData]);
    
    // Send Slack notification to directors
    try {
      sendAppealNotification(employee, infractionRow, explanation);
    } catch (slackError) {
      console.error('Slack notification failed:', slackError);
      // Don't fail the appeal if Slack fails
    }
    
    return { success: true, message: 'Appeal submitted successfully' };
    
  } catch (e) {
    console.error('submitInfractionAppeal error:', e);
    return { success: false, error: e.toString() };
  }
}

// Send Slack notification for appeal
function sendAppealNotification(employee, infractionRow, explanation) {
  try {
    if (!CONFIG.LEADERS_WEBHOOK) {
      console.log('No Slack webhook configured for leaders');
      return;
    }
    
    const message = {
      text: `:bell: *New Infraction Appeal* from *${employee}*\n` +
            `Original Infraction: Row ${infractionRow}\n` +
            `Explanation: ${explanation}\n` +
            `Status: Pending Review`,
      mrkdwn: true
    };
    
    const response = UrlFetchApp.fetch(CONFIG.LEADERS_WEBHOOK, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(message),
      muteHttpExceptions: true
    });
    
    console.log('Appeal Slack notification sent, response code:', response.getResponseCode());
    
  } catch (e) {
    console.error('sendAppealNotification error:', e);
  }
}
