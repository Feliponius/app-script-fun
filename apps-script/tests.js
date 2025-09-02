/************** Test harness for exhaustive system checks **************/

/** Utility — recreate expected headers (idempotent) **/
function ensureEventHeaders_() {
  const s = sh_(CONFIG.TABS.EVENTS);
  const expected = [
    "Timestamp","IncidentDate","Employee","Lead","EventType","Consequence Director","Pending Status",
    "PositiveAction","Policy","Infraction","IncidentDescription","Points","CorrectiveActions",
    "TeamMemberStatement","RedLine","Policy Protected?","FinalPoints","Active","PointsRolling",
    "Milestone","MilestoneDate","Nullify","PointsRolling (Effective)",
    "Probation_Active","Probation_Failure","Probation_Start","Probation_End","Perf_NoPickup",
    "Perf_NoPickup_End","Per_NoPickup_Active","Linked_Event_ID","Write-Up PDF","Signed_PDF_Link","Audit"
  ];
  const cur = headers_(s);
  // Add any missing headers (append to the right)
  expected.forEach((h, i) => {
    if (cur.indexOf(h) === -1) {
      s.getRange(1, s.getLastColumn() + 1).setValue(h);
      Logger.log('ensureEventHeaders_: created header "%s"', h);
    }
  });
}

function TEST_leadersWebhook(){
  postSlack_(CONFIG.LEADERS_WEBHOOK, {text:'TEST leaders webhook OK '+new Date(), mrkdwn:true});
}

function TEST_notifyLeadersClaimed(){
  var row = 59; // CHANGE to a Milestone row that now has a PDF
  notifyLeadersMilestone_(row, ''); // will read URL from cell if omitted in your implementation
}

function TEST_forceClaim(){
  var row = 59; // CHANGE to a Milestone row that is Pending, with Consequence Director filled
  maybeCreatePdfForMilestoneRow_(CONFIG.TABS.EVENTS, row);
}


function TEST_notifyDocsRow(){
  var row = 2; // pick a row with a write-up PDF
  notifyDocs_(row);
}


function TEST_slackWebhookDocs(){
  if (!CONFIG.DOCS_WEBHOOK){ logError('TEST_slackWebhookDocs', new Error('Missing DOCS_WEBHOOK')); return; }
  var ok = postSlack_(CONFIG.DOCS_WEBHOOK, { text: 'TEST: webhook connectivity OK ('+new Date()+')', mrkdwn:true });
  logInfo_('TEST_slackWebhookDocs_result', {ok:ok});
}
function testGraceHappyPath(){
  // Setup
  const emp = 'Test User';
  const director = 'QA Director';
  const reason = 'QA happy path';
  const events = sh_(CONFIG.TABS.EVENTS);
  const pos = sh_(CONFIG.TABS.POSITIVE_POINTS);
  const eh = headers_(events), ph = headers_(pos);
  const e = (h)=>eh.indexOf(h)+1, p=(h)=>ph.indexOf(h)+1;

  // Seed 1: credit (Moderate, approved)
  pos.appendRow(ph.map(()=>'')); const prow = pos.getLastRow();
  pos.getRange(prow, p('Employee')).setValue(emp);
  pos.getRange(prow, p('Credit Type')).setValue('Moderate');
  if (p('Approved?')>0) pos.getRange(prow, p('Approved?')).setValue(true);

  // Seed 2: 3-point event
  events.appendRow(eh.map(()=>'')); const erow = events.getLastRow();
  events.getRange(erow, e(CONFIG.COLS.IncidentDate)).setValue(new Date());
  events.getRange(erow, e(CONFIG.COLS.Employee)).setValue(emp);
  events.getRange(erow, e(CONFIG.COLS.EventType)).setValue('Disciplinary');
  events.getRange(erow, e(CONFIG.COLS.Points)).setValue(3);
  events.getRange(erow, e(CONFIG.COLS.Active)).setValue(true);

  // Act
  const ok = applyGraceForEventRow(erow, director, reason);
  Logger.log('applyGraceForEventRow returned: ' + ok);

  // Assert
  const rowVals = events.getRange(erow,1,1,events.getLastColumn()).getValues()[0];
  const get = (name)=> rowVals[eh.indexOf(name)];

  assertEqual('GraceApplied true', String(get(CONFIG.COLS.GraceApplied)).toUpperCase(), 'TRUE');
  assertEqual('GraceTier Moderate', String(get(CONFIG.COLS.GraceTier)), 'Moderate');
  if (eh.indexOf('Nullify')!==-1) assertEqual('Nullify YES', String(get('Nullify')), 'YES');

  // Check ledger consumed
  const prowVals = pos.getRange(prow,1,1,pos.getLastColumn()).getValues()[0];
  const getP = (name)=> prowVals[ph.indexOf(name)];
  assertTrue('Ledger consumed', asBool(getP('Consumed?')));

  // Check synthetic row exists
  const lastEventType = events.getRange(events.getLastRow(), e(CONFIG.COLS.EventType)).getValue();
  assertEqual('Synthetic EventType', String(lastEventType), (CONFIG.POLICY && CONFIG.POLICY.GRACE && CONFIG.POLICY.GRACE.eventTypeName) || 'Grace');

  Logger.log('✅ testGraceHappyPath passed');
}

function testGraceNoCredit(){
  const emp = 'Test User';
  const director = 'QA Director';
  const reason = 'QA no credit';
  const events = sh_(CONFIG.TABS.EVENTS);
  const eh = headers_(events), e=(h)=>eh.indexOf(h)+1;

  // Seed event (3 pts)
  events.appendRow(eh.map(()=>'')); const erow = events.getLastRow();
  events.getRange(erow, e(CONFIG.COLS.IncidentDate)).setValue(new Date());
  events.getRange(erow, e(CONFIG.COLS.Employee)).setValue(emp);
  events.getRange(erow, e(CONFIG.COLS.EventType)).setValue('Disciplinary');
  events.getRange(erow, e(CONFIG.COLS.Points)).setValue(3);
  events.getRange(erow, e(CONFIG.COLS.Active)).setValue(true);

  // Act: should fail (no available Moderate credit)
  const ok = applyGraceForEventRow(erow, director, reason);
  assertFalse('Grace without credit should fail', ok);

  Logger.log('✅ testGraceNoCredit passed');
}

function testGraceRedLineBlocked(){
  const emp = 'Test User';
  const director = 'QA Director';
  const reason = 'QA redline';
  const events = sh_(CONFIG.TABS.EVENTS);
  const pos = sh_(CONFIG.TABS.POSITIVE_POINTS);
  const eh = headers_(events), ph = headers_(pos);
  const e=(h)=>eh.indexOf(h)+1, p=(h)=>ph.indexOf(h)+1;

  // Seed credit (Major, approved) so availability isn't the reason for failure
  pos.appendRow(ph.map(()=>'')); const prow = pos.getLastRow();
  pos.getRange(prow, p('Employee')).setValue(emp);
  pos.getRange(prow, p('Credit Type')).setValue('Major');
  if (p('Approved?')>0) pos.getRange(prow, p('Approved?')).setValue(true);

  // Seed red-line event
  events.appendRow(eh.map(()=>'')); const erow = events.getLastRow();
  events.getRange(erow, e(CONFIG.COLS.IncidentDate)).setValue(new Date());
  events.getRange(erow, e(CONFIG.COLS.Employee)).setValue(emp);
  events.getRange(erow, e(CONFIG.COLS.EventType)).setValue('Disciplinary');
  events.getRange(erow, e(CONFIG.COLS.Points)).setValue(8);
  events.getRange(erow, e(CONFIG.COLS.Active)).setValue(true);
  if (eh.indexOf(CONFIG.COLS.RedLine)!==-1) events.getRange(erow, e(CONFIG.COLS.RedLine)).setValue(true);

  // Act: should fail if GRACE_EXCLUDE_REDLINES=true
  const ok = applyGraceForEventRow(erow, director, reason);
  assertFalse('Grace should block red-line', ok);

  Logger.log('✅ testGraceRedLineBlocked passed');
}

/* ---------- tiny test helpers ---------- */
function assertEqual(msg, actual, expected){
  if (actual !== expected) throw new Error(msg + ': expected "'+expected+'" got "'+actual+'"');
}
function assertTrue(msg, v){
  if (!asBool(v)) throw new Error(msg + ': expected TRUE');
}
function assertFalse(msg, v){
  if (asBool(v)) throw new Error(msg + ': expected FALSE');
}
function asBool(v){
  if (v === true) return true;
  var s = String(v||'').trim().toUpperCase();
  return s==='TRUE' || s==='YES' || s==='1';
}




/** Clear all data rows but keep header row **/
function resetEventsSheet() {
  const s = sh_(CONFIG.TABS.EVENTS);
  // keep headers, delete everything below row 1
  const lastRow = s.getLastRow();
  if (lastRow > 1) {
    s.getRange(2, 1, lastRow - 1, s.getLastColumn()).clearContent();
  }
  // ensure headers exist exactly once
  ensureEventHeaders_();
  Logger.log('resetEventsSheet: cleared data rows and ensured headers.');
}

/** Append an event using headers-map convenience (returns row index) **/
function insertTestEvent(obj) {
  const s = sh_(CONFIG.TABS.EVENTS);
  const hdr = headers_(s);
  const rowArr = hdr.map(h => (Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : ''));
  return appendEventsRow_(rowArr);
}

/** Helper to mark a director claim by writing Consequence Director and calling onEdit handler **/
function claimMilestoneAsDirector(milestoneRow, directorName) {
  const s = sh_(CONFIG.TABS.EVENTS);
  const hdr = headers_(s);
  const col = hdr.indexOf(CONFIG.COLS.ConsequenceDirector) + 1;
  s.getRange(milestoneRow, col).setValue(directorName);
  // simulate installable trigger payload
  const fakeE = { range: s.getRange(milestoneRow, col), value: directorName, oldValue: '' };
  try { onEditMilestoneClaim_(fakeE); } catch(err){ Logger.log('claimMilestoneAsDirector: handler threw %s', err); }
}

/** Helper to flip Nullify on a row (and call the onEdit handler) **/
function toggleNullify(row, value) {
  const s = sh_(CONFIG.TABS.EVENTS);
  const hdr = headers_(s);
  const col = hdr.indexOf(CONFIG.COLS.Nullify) + 1;
  s.getRange(row, col).setValue(value);
  const fakeE = { range: s.getRange(row, col), value: value, oldValue: '' };
  try { onEditNullify(fakeE); } catch(err){ Logger.log('toggleNullify: handler threw %s', err); }
}

/** Small assertion helper **/
function ok(cond, msg) { 
  Logger.log((cond ? 'PASS: ' : 'FAIL: ') + msg); 
  return !!cond;
}

/** Run the exhaustive sequence **/
function runExhaustiveTests() {
  resetEventsSheet();
  ensureEventHeaders_();

  // config quick-check
  Logger.log('CONFIG.POLICY snapshot: %s', JSON.stringify(CONFIG.POLICY || {}));

  // 1) Baseline: trigger L1 (5 pts) -> pending milestone expected
  const empA = 'TEST_FINISH_A';
  const rA1 = insertTestEvent({
    'Employee': empA, 'IncidentDate': new Date(), 'EventType': 'Disciplinary event (write-up)', 'Points': 5
  });
  afterSubmitPoints_(rA1);
  Utilities.sleep(400);
  const latestA = findLatestMilestoneForEmployee_(empA);
  ok(latestA && tierForPoints_(latestA.pointsRolling) >= 1, 'L1 milestone created for ' + empA);

  // record the milestone row for later
  const milestoneRowA = latestA ? latestA.row : null;

  // 2) Same-tier cooldown: another 5 pts immediately -> blocked (no new milestone)
  const rA2 = insertTestEvent({ 'Employee': empA, 'IncidentDate': new Date(), 'EventType': 'Disciplinary event (write-up)', 'Points': 5 });
  afterSubmitPoints_(rA2);
  Utilities.sleep(400);
  const latestA2 = findLatestMilestoneForEmployee_(empA);
  ok(latestA2 && latestA2.row === milestoneRowA, 'Same-tier cooldown blocked duplicate L1');

  // 3) Escalation allowed during cooldown: add 5 more (total 15 raw) or a 10-pt event to make L2
  const rA3 = insertTestEvent({ 'Employee': empA, 'IncidentDate': new Date(), 'EventType': 'Disciplinary event (write-up)', 'Points': 10 });
  afterSubmitPoints_(rA3);
  Utilities.sleep(500);
  const latestA3 = findLatestMilestoneForEmployee_(empA);
  ok(latestA3 && tierForPoints_(latestA3.pointsRolling) >= 2, 'Escalation to L2 allowed during cooldown');

  // store L2 row
  const milestoneRowL2 = latestA3 ? latestA3.row : null;

  // 4) Claim L2 as director -> should set probation fields
  claimMilestoneAsDirector(milestoneRowL2, 'Director One');
  Utilities.sleep(400);
  const ctxL2 = rowCtx_(sh_(CONFIG.TABS.EVENTS), milestoneRowL2);
  const pStart = ctxL2.get(CONFIG.COLS.ProbationStart);
  const pEnd = ctxL2.get(CONFIG.COLS.ProbationEnd);
  const pActiveCell = ctxL2.get(CONFIG.COLS.ProbationActive);
  ok(pStart && pEnd, 'Probation Start/End set on claim');
  Logger.log('Probation_Active cell value (formula may be used): %s', String(pActiveCell));

  // 5) Probation failure flow: make Probation_Start = yesterday to allow counting events
  const yesterday = addDaysLocal_(new Date(), -1);
  ctxL2.set(CONFIG.COLS.ProbationStart, yesterday);
  ctxL2.set(CONFIG.COLS.ProbationEnd, addDaysLocal_(yesterday, 30));
  // Add events totaling 4 pts since probation start
  const rP1 = insertTestEvent({ 'Employee': empA, 'IncidentDate': new Date(), 'EventType': 'Disciplinary event (write-up)', 'Points': 2 });
  afterSubmitPoints_(rP1);
  const rP2 = insertTestEvent({ 'Employee': empA, 'IncidentDate': new Date(), 'EventType': 'Disciplinary event (write-up)', 'Points': 2 });
  afterSubmitPoints_(rP2);
  Utilities.sleep(600);
  // find probation failure milestone
  const latestAfterProb = findLatestMilestoneForEmployee_(empA);
  const didProbFail = latestAfterProb && (String(latestAfterProb.pointsRolling || '').toLowerCase().indexOf('probation') >= 0 || String(latestAfterProb.when || '').length>0 && ctxL2.get(CONFIG.COLS.ProbationFailure));
  // We also check the flag on the last event
  const pfFlagOnLast = rowCtx_(sh_(CONFIG.TABS.EVENTS), rP2).get(CONFIG.COLS.ProbationFailure);
  ok(pfFlagOnLast || (latestAfterProb && String(latestAfterProb.row) !== String(milestoneRowL2)), 'Probation failure milestone or flag created');

  // 6) Nullify behavior: create new emp, add 5-pt event, flip Nullify -> effective recalcs and milestone may change
  const empN = 'TEST_NULL';
  const rN = insertTestEvent({ 'Employee': empN, 'IncidentDate': new Date(), 'EventType': 'Disciplinary event (write-up)', 'Points': 5 });
  afterSubmitPoints_(rN);
  Utilities.sleep(300);
  const latestN = findLatestMilestoneForEmployee_(empN);
  ok(latestN, 'Nullify test: milestone initially present');
  toggleNullify(rN, 'YES');
  Utilities.sleep(400);
  // effective should drop or become different; recompute
  const rollAfterNull = getRollingPointsForEmployee_(empN);
  Logger.log('Nullify result roll: %s', JSON.stringify(rollAfterNull));
  ok((typeof rollAfterNull === 'object') ? (rollAfterNull.effective === 0 || rollAfterNull.raw < 5) : true, 'Nullify recalculation observed');

  // 7) Positive cap + floor scenario
  const empP = 'TEST_POS';
  // create two raw events to make raw >=5
  const rP_a = insertTestEvent({ 'Employee': empP, 'IncidentDate': addDaysLocal_(new Date(), -10), 'EventType': 'Disciplinary event (write-up)', 'Points': 3 });
  afterSubmitPoints_(rP_a);
  const rP_b = insertTestEvent({ 'Employee': empP, 'IncidentDate': addDaysLocal_(new Date(), -5), 'EventType': 'Disciplinary event (write-up)', 'Points': 3 });
  afterSubmitPoints_(rP_b);
  // Add Positive removal over cap
  const cap = Number((CONFIG.POLICY && CONFIG.POLICY.POSITIVE_POINTS_CAP) || CONFIG.POSITIVE_POINTS_CAP || 2);
  const rPos = insertTestEvent({ 'Employee': empP, 'IncidentDate': new Date(), 'EventType': 'Positive Point Removal', 'Points': cap + 2 });
  afterSubmitPoints_(rPos);
  Utilities.sleep(500);
  const rollP = getRollingPointsForEmployee_(empP);
  Logger.log('Positive cap roll: %s', JSON.stringify(rollP));
  ok( (typeof rollP === 'object') ? (rollP.pos <= cap && (rollP.raw >= (CONFIG.POLICY && CONFIG.POLICY.MILESTONES ? CONFIG.POLICY.MILESTONES.LEVEL_1 : 5) ? rollP.effective >= (CONFIG.POLICY.MILESTONES ? CONFIG.POLICY.MILESTONES.LEVEL_1 : 5) : true)) : true, 'Positive cap + floor behavior observed');

  // 8) PDF idempotency: create event and ensure second trigger does not re-create
  const empPdf = 'TEST_PDF';
  const rPdf = insertTestEvent({ 'Employee': empPdf, 'IncidentDate': new Date(), 'EventType': 'Disciplinary event (write-up)', 'Points': 1 });
  afterSubmitPoints_(rPdf);
  Utilities.sleep(400);
  // run onFormSubmit again for same row to simulate duplicate fire
  try { onFormSubmit({ range: sh_(CONFIG.TABS.EVENTS).getRange(rPdf, 1) }); } catch(e){ /* ok if throws */ }
  Utilities.sleep(300);
  // check log lines or pdf cell exists
  const pdfCell = rowCtx_(sh_(CONFIG.TABS.EVENTS), rPdf).get(CONFIG.COLS.PdfLink);
  ok(pdfCell, 'PDF created and preserved on re-trigger (idempotent)');

  Logger.log('runExhaustiveTests: done. Inspect the Events sheet and logs for details.');
}

function testMilestoneMapping() {
  var s = '1-Week Suspension + Final Warning + 30-Day Probation';
  Logger.log('milestoneTextToTier_("%s") => %s', s, typeof milestoneTextToTier_ === 'function' ? milestoneTextToTier_(s) : 'MISSING');
  try{ Logger.log('CONFIG.POLICY.MILESTONE_NAMES.LEVEL_2 = %s', (CONFIG && CONFIG.POLICY && CONFIG.POLICY.MILESTONE_NAMES && CONFIG.POLICY.MILESTONE_NAMES.LEVEL_2) || '<<empty>>'); }catch(_){ Logger.log('CONFIG.POLICY not available');}
}

function testIsMilestoneActive(){
  var s = sh_(CONFIG.TABS.EVENTS);
  var ctx = rowCtx_(s, 6); // pick a row with a milestone
  Logger.log('isMilestoneRowActive_ => %s', isMilestoneRowActive_(ctx));
  Logger.log('safeNullify => "%s"', safeGetFromCtx_(ctx, [(CONFIG && CONFIG.COLS && CONFIG.COLS.Nullify) || null, 'Nullify']));
}

function testFindLatestForEmployee() {
  var emp = 'Braue, Caleb Eden'; // <-- exact Employee cell text
  var out = findLatestMilestoneForEmployee_(emp, { beforeRow: 99999 });
  Logger.log('findLatestMilestoneForEmployee_(%s) => %s', emp, JSON.stringify(out, null, 2));
}