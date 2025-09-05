// config.gs — compatible with CLEAR v2.1 (Apps Script V8; uses var for max compatibility)
var CONFIG = {
  // IDs
  SHEET_ID: '1AsHXKRPCPhcjFGBPJHQ9kqrbKaddIoWucSkApWG9C5U',
  FORM_ID: '1w6z9DHZQJjBTML866EIWmwCvrmlNCGcRRDZamh8FO9I',
  DEST_FOLDER_ID: '1CyGYTg3ajPErjI8bjxwSEtk9p4soUKX2',

  // Tabs (table names == tab names)
  TABS: {
    EVENTS: 'Events',
    RUBRIC: 'Rubric',
    POSITIVE: 'PositivePoints',      // legacy name you used
    AUDIT: 'Audit',
    LOGS: 'Logs',
    CONFIG: 'Config',
    DATA_VALIDATION: 'Data-Validation'
  },

  POSITIVE_TIER_LABELS: {
    minor: 'Minor Positive Credit',
    moderate: 'Moderate Positive Credit',
    major: 'Major Poistive Credit',
    grace: 'Grace',
  },

  // Events columns — EXACT headers
  COLS: {
    Timestamp: 'Timestamp',
    IncidentDate: 'IncidentDate',

    // Grace model (user-facing)
    GraceFlag: 'Grace',                 // checkbox (director sets)
    GracedBy: 'Graced By',              // director name
    GraceReason: 'Grace Reason',        // short reason
    GraceApplied: 'Grace Applied',      // TRUE/YES once applied
    GraceTier: 'Grace Tier',            // Minor|Moderate|Major (system)
    GraceLedgerRow: 'Grace Ledger Row', // row number in PositivePoints
    LinkedEventRow: 'Linked Event Row', // on the synthetic "Grace" row

    // Existing columns
    Employee: 'Employee',
    Lead: 'Lead',
    EventType: 'EventType',
    ConsequenceDirector: 'Consequence Director',
    PendingStatus: 'Pending Status',
    PositiveActionType: 'PositiveAction',
    RelevantPolicy: 'Policy',
    Infraction: 'Infraction',
    NotesReviewer: 'IncidentDescription',
    Points: 'Points',
    CorrectiveActions: 'CorrectiveActions',
    TeamMemberStatement: 'TeamMemberStatement',
    RedLine: 'RedLine',
    FinalPoints: 'FinalPoints',
    Active: 'Active',
    PointsRolling: 'PointsRolling',
    PointsRollingEffective: 'PointsRolling (Effective)',
    Milestone: 'Milestone',
    MilestoneDate: 'MilestoneDate',
    Probation_Active: 'Probation_Active',
    ProbationStart: 'Probation_Start',
    ProbationEnd: 'Probation_End',
    Perf_NoPickup: 'Perf_NoPickup',
    Perf_NoPickup_End: 'Perf_NoPickup_End',
    Per_NoPickup_Active: 'Per_NoPickup_Active',
    
    // Performance Issue Tracking Columns
    PerfIssueCount: 'Perf Issue Count',                    // Count of Performance Issues for employee
    PerfGrowthPlanDate: 'Perf Growth Plan Date',          // When growth plan was assigned
    PerfGrowthPlanDeadline: 'Perf Growth Plan Deadline',  // Growth plan deadline
    PerfGrowthPlanDecision: 'Perf Growth Plan Decision',  // Success/Failure decision
    PerfReductionStatus: 'Perf Reduction Status',         // None/Indefinite/Greater
    PerfReturnToGoodStanding: 'Perf Return to Good Standing', // Date returned to good standing
    
    Linked_Event_ID: 'Linked_Event_ID',
    PdfLink: 'Write-Up PDF',
    Signed: 'Signed_PDF_Link',
    Audit: 'Audit',

    // Legacy nullify kept for back-compat in formulas
    Nullify: 'Nullify',
    NullifiedBy: 'Nullified By',

    ProbationFailure: 'Probation_Failure',
    ProbationActive: 'Probation_Active',      // alias w/o underscore
    IncidentDescription: 'IncidentDescription',// alias so both names work
    PositiveAction: 'PositiveAction',          // alias in addition to PositiveActionType
    LinkedEventID: 'Linked_Event_ID',          // common alias used in helpers
    WriteUpPDF: 'Write-Up PDF',                // some code uses WriteUpPDF
    SignedPDFLink: 'Signed_PDF_Link'           // some code uses SignedPDFLink
  },

  // Director-only actions (not used by v2.1 but left for compatibility)
  DIRECTOR_ACTIONS: [
    'Milestone — 5 pts (2-Day Suspension + Written Plan)',
    'Milestone — 10 pts (1-Week + Final Warning + 30-Day Probation)',
    'Termination',
    'Policy-Protected Notice',
    'Probation — Start',
    'Probation — Complete',
    'Probation — Failure'
  ],

  // People & permissions (not used when running without Workspace)
  DIRECTORS: [
    'Philip Pixler',
    'Kristie Wright',
    'Quana Slaughter',
    'Cristian Zavala'
  ],

  // Slack webhooks (optional)
  DOCS_WEBHOOK: 'https://hooks.slack.com/services/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',
  LEADERS_WEBHOOK: 'https://hooks.slack.com/services/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',

  ONLY_DISCIPLINARY_SLACK: true,        // post to docs channel only for disciplinary events
  ENABLE_SLACK_DOCS:       true,        // master switch
  ENABLE_SLACK_LEADERS:    true,        // master switch
  ENABLE_SLACK_LEADERS_PENDING: true,   // ping when Milestone row is created
  ENABLE_SLACK_LEADERS_CLAIMED: true,  // ping after PDF is created/claimed (turn OFF per your request)


  DOCS_EVENT_TYPES: [
    'Disciplinary Event',
    'Disciplinary Event (Write-Up)',
    'Positive Credit',
  ],

  // Template Doc IDs (replace EVENT_RECORD at minimum)
  TEMPLATES: {
    EVENT_RECORD: '17AbFEgK_SyBTqE8frFHZ-BY2oDhi8CtjYGg6py0aEtM',
    MILESTONE_5:  '1kfcIYeLajNQHeSptYLHe1w65C2MtR6d7jA2WEQJkzU4',
    MILESTONE_10: '1eO5YJzAp1rPQm8nEvkh0GiBJchMNgI82nTFJyTIpL1U',
    MILESTONE_15: '1M5ps3uCwUFcGNBfRqURPY2A7C5JmpS16YiNoKoBjVbA',
    POLICY_PROTECTED: '',
    PROBATION_FAILURE: '1ejCM7pVhzgal6STKECGop9QeRcuQ8bjzAP1y0QIk3jI',
    EMP_HISTORY: '1aHEEchp3HDbuzlKEQ57iDr_F3L9Dm5PKG1qm6KceGs8',
    
    // Performance Issue Templates
    PERF_ISSUE: '1wpGcqxTMxaL1P4m1kuj8KCpw05ZFL4o9J4Un6F0UTEs',                    // Performance Issue documentation
    PERF_GROWTH_PLAN: '1s_wH0JCGl6f9tlUsf8QI6SNSG7nWS2BprjCehjZtQSM',              // Growth Plan consequence (after 2 Performance Issues)
    PERF_GREATER_REDUCTION: '1fJegqntStZU25dgzb5cTIsTxsmefaw-mEbsr3T-_Q2I',        // Greater Reduction consequence
    PERF_FAILURE_TERMINATION: '1heDEgCl6EShhQuIeY7xPv7MTuflCY3h5lHh_96HyFl4',       // Performance Failure Termination
    PERF_GA_SUCCESS: '1r9b-36yxz_IYiAWanWY_9gRwrJ5J6OntTeR3CjbC9aM',                                                             // Return to Good Standing (MISSING - needs template ID)
    PERF_INDEFINITE_REDUCTION: '1kpvI5ANwn_ehyyzXzOcCqr6J_F2tUKo1ENf4yAqnoiM'                                                    // Indefinite Reduction (MISSING - needs template ID)
  },

  // Optional: map template token names to source keys (used by buildPdfDataFromRow_)
  PDF_ALIASES: {
    "RelevantPolicy": ["Policy", "Relevant Policy"],
    "EventID": ["LinkedEventID", "Linked_Event_ID", "Row"],
    "WriteUpTitle": ["WriteUpTitle", "Title"]
  },

  // Runtime
  NIGHTLY_HOUR: 3
};

// temporary: if false, script will NOT write Probation_* columns (sheet owns them)
CONFIG.SCRIPT_OWNS_PROBATION = true;
function scriptShouldWriteProbation_(){ return !!(CONFIG && CONFIG.SCRIPT_OWNS_PROBATION); }

// --- Compatibility aliases (safe): map older names other code may expect
CONFIG.TABS.POSITIVE_MAP = CONFIG.TABS.POSITIVE_MAP || 'PositiveMap'; // tab with PositiveAction -> Credit Type
CONFIG.TABS.POSITIVE_POINTS = CONFIG.TABS.POSITIVE_POINTS || CONFIG.TABS.POSITIVE || 'PositivePoints';
CONFIG.COLS.PositiveActionType = CONFIG.COLS.PositiveActionType || CONFIG.COLS.PositiveAction || 'PositiveAction';

// Form → Events header mapping
CONFIG.FORM_TO_EVENTS = {
  "Timestamp": "Timestamp",
  "Date of Incident": "IncidentDate",
  "Employee": "Employee",
  "Event Type": "EventType",
  "Lead": "Lead",
  "Positive Action": "PositiveAction",
  "Infraction": "Infraction",
  "Incident Description": "IncidentDescription",           // For disciplinary events
  "Performance Incident Description": "IncidentDescription", // For performance issues
  "Corrective Actions": "CorrectiveActions",
  "Team Member Statement": "TeamMemberStatement"
};

// Form → PositivePoints header mapping (feeds tier directly)
CONFIG.FORM_TO_POSITIVE = {
  "Timestamp": "Timestamp",
  "Date of Incident": "IncidentDate",
  "Employee": "Employee",
  "Lead": "Lead",
  "Positive Action": "Credit Type",   // <-- tier: Minor | Moderate | Major
  "Incident Description": "Notes"
};

// Load policy (reads Config tab; provides safe defaults if rows missing)
function loadPolicyFromSheet_() {
  var sh = sh_(CONFIG.TABS.CONFIG);
  var lastRow = Math.max(0, sh.getLastRow());
  var numRows = Math.max(0, lastRow - 1); // number of config data rows (exclude header)

  var cfg = {};
  if (numRows > 0) {
    var values = sh.getRange(2, 1, numRows, 2).getValues();
    values.forEach(function(row){
      var k = row[0];
      var v = row[1];
      if (k) cfg[k] = v;
    });
  }

  // Normalize CONFIG.POLICY from raw cfg (loaded from Config sheet)
  CONFIG.POLICY = {
    // numeric milestone thresholds (sheet keys: MILESTONE_LEVEL_1, etc.)
    MILESTONES: {
      LEVEL_1: Number(cfg.MILESTONE_LEVEL_1 || cfg.MILESTONE_1 || cfg.MILESTONELEVEL1 || cfg.MILESTONELEVEL_1 || 5),
      LEVEL_2: Number(cfg.MILESTONE_LEVEL_2 || cfg.MILESTONE_2 || cfg.MILESTONELEVEL2 || cfg.MILESTONELEVEL_2 || 10),
      LEVEL_3: Number(cfg.MILESTONE_LEVEL_3 || cfg.MILESTONE_3 || cfg.MILESTONELEVEL3 || cfg.MILESTONELEVEL_3 || 15)
    },

    // human-readable milestone names (sheet keys: MILESTONE_NAME_1, etc.)
    MILESTONE_NAMES: {
      LEVEL_1: String(cfg.MILESTONE_NAME_1 || cfg.MILESTONE_1_NAME || cfg.MILESTONE_NAME1 || cfg.MILESTONE_NAME_1 || '') || '',
      LEVEL_2: String(cfg.MILESTONE_NAME_2 || cfg.MILESTONE_2_NAME || cfg.MILESTONE_NAME2 || cfg.MILESTONE_NAME_2 || '') || '',
      LEVEL_3: String(cfg.MILESTONE_NAME_3 || cfg.MILESTONE_3_NAME || cfg.MILESTONE_NAME3 || cfg.MILESTONE_NAME_3 || '') || ''
    },

    // other numeric policy fields
    COOLDOWN_WEEKS: Number(cfg.COOLDOWN_WEEKS || cfg.COOLDOWN || cfg.COOLDOWNWEEKS || 4),
    DOUBLE_TRIGGER_DAYS: Number(cfg.DOUBLE_TRIGGER_DAYS || cfg.DOUBLE_TRIGGER || 0),
    PROBATION_DAYS: Number(cfg.PROBATION_DAYS || cfg.PROBATION || 30),
    POSITIVE_POINTS_CAP: Number(cfg.POSITIVE_POINTS_CAP || cfg.POSITIVE_POINTS || 0),
    GRACE_COMMITTEE_CAP: Number(cfg.GRACE_COMMITTEE_CAP || cfg.GRACE_COMMIT || 0),
    ROLLING_DAYS: Number(cfg.ROLLING_DAYS || 180),
    // probation failure threshold (sheet key: PROBATION_FAILURE) — default 14 per your config screenshot
    PROBATION_FAILURE: Number(cfg.PROBATION_FAILURE || cfg.PROBATION_FAIL || cfg.PROBATION_FAIL_THRESHOLD || 14)
  };

  // --- Grace & Credits (read from Config with safe defaults)
  CONFIG.POLICY.CREDITS = [
    { name: 'Minor',    maxPoints: Number(cfg.CREDIT_MINOR_MAX    || 2),   cooldownDays: Number(cfg.CREDITS_COOLDOWN_MINOR_DAYS    || 0) },
    { name: 'Moderate', maxPoints: Number(cfg.CREDIT_MODERATE_MAX || 4),   cooldownDays: Number(cfg.CREDITS_COOLDOWN_MODERATE_DAYS || 0) },
    { name: 'Major',    maxPoints: Number(cfg.CREDIT_MAJOR_MAX    || 999), cooldownDays: Number(cfg.CREDITS_COOLDOWN_MAJOR_DAYS    || 0) }
  ];

  CONFIG.POLICY.GRACE = {
    enabled: String(cfg.GRACE_ENABLED || 'TRUE').toUpperCase() !== 'FALSE',
    eventTypeName: String(cfg.GRACE_EVENTTYPE_NAME || 'Grace'),
    eligibleWindowDays: Number(cfg.GRACE_ELIGIBLE_WINDOW_DAYS || CONFIG.POLICY.ROLLING_DAYS || 180),
    excludeRedLines: String(cfg.GRACE_EXCLUDE_REDLINES || 'TRUE').toUpperCase() !== 'FALSE',
    requireApproval: String(cfg.CREDITS_REQUIRE_APPROVAL || 'TRUE').toUpperCase() !== 'FALSE'
  };

  // --- Convenience aliases for older code paths (do not remove)
  CONFIG.POLICY.MILESTONES_LEVEL_1 = CONFIG.POLICY.MILESTONES.LEVEL_1;
  CONFIG.POLICY.MILESTONES_LEVEL_2 = CONFIG.POLICY.MILESTONES.LEVEL_2;
  CONFIG.POLICY.MILESTONES_LEVEL_3 = CONFIG.POLICY.MILESTONES.LEVEL_3;
  CONFIG.POLICY.MILESTONE_NAME_1 = CONFIG.POLICY.MILESTONE_NAMES.LEVEL_1;
  CONFIG.POLICY.MILESTONE_NAME_2 = CONFIG.POLICY.MILESTONE_NAMES.LEVEL_2;
  CONFIG.POLICY.MILESTONE_NAME_3 = CONFIG.POLICY.MILESTONE_NAMES.LEVEL_3;
  CONFIG.POLICY.PROBATION_FAIL = CONFIG.POLICY.PROBATION_FAILURE;

  return CONFIG.POLICY;
}

// Lightweight config validator (name with underscore to match earlier code)
function validateConfig(){
  var issues = [];
  if (typeof CONFIG === 'undefined') { issues.push('CONFIG is not defined — paste your config.gs first.'); Logger.log(issues.join('\n')); return issues; }
  if (!CONFIG.SHEET_ID) issues.push('CONFIG.SHEET_ID is missing.');
  if (!CONFIG.TABS || typeof CONFIG.TABS !== 'object') issues.push('CONFIG.TABS is missing or not an object.');
  if (!CONFIG.COLS || typeof CONFIG.COLS !== 'object') issues.push('CONFIG.COLS is missing or not an object.');

  // required tabs
  var requiredTabs = [
    CONFIG.TABS.EVENTS || 'Events',
    CONFIG.TABS.RUBRIC || 'Rubric',
    CONFIG.TABS.POSITIVE_POINTS || CONFIG.TABS.POSITIVE || 'PositivePoints',
    CONFIG.TABS.AUDIT || 'Audit',
    CONFIG.TABS.LOGS || 'Logs'
  ];

  var ss;
  try {
    ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  } catch(e) {
    issues.push('Unable to open spreadsheet by CONFIG.SHEET_ID: ' + (e.message || e));
    Logger.log('CONFIG issues:\n' + issues.join('\n'));
    return issues;
  }

  var existingTabs = ss.getSheets().map(function(s){ return s.getName(); });
  requiredTabs.forEach(function(t){ if (t && existingTabs.indexOf(t) === -1) issues.push('Missing tab: ' + t); });

  // check Events headers
  var eventsName = CONFIG.TABS.EVENTS || 'Events';
  if (existingTabs.indexOf(eventsName) !== -1){
    var eventsSheet = ss.getSheetByName(eventsName);
    var hdrs = eventsSheet.getRange(1,1,1,Math.max(1,eventsSheet.getLastColumn())).getValues()[0];
    Object.keys(CONFIG.COLS).forEach(function(k){
      var h = CONFIG.COLS[k];
      if (h && hdrs.indexOf(h)===-1) issues.push('Events missing header: "'+h+'"');
    });
  } else {
    issues.push('Events sheet not present, skipping header checks.');
  }

  if (!CONFIG.TEMPLATES || !CONFIG.TEMPLATES.EVENT_RECORD) issues.push('CONFIG.TEMPLATES.EVENT_RECORD (Doc ID) is missing.');

  if (issues.length) {
    Logger.log('CONFIG issues:\n' + issues.join('\n'));
  } else {
    Logger.log('Config looks good ✅');
  }
  return issues;
}
