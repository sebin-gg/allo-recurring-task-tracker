/*************************************************************
 * Allo Innoware — Recurring Task Tracker
 * Single-file Apps Script.
 * Setup: paste into Extensions → Apps Script, save, then run
 *   setupSheets() once, then installTriggers().
 *************************************************************/

var SHEETS = {
  CONFIG: 'Config',
  TASKS: 'TaskMaster',
  CYCLES: 'Cycles',
  MYTASKS: 'MyTasks',
  DASH: 'Dashboard',
  LOG: 'Log'
};

var STATUS = { SCHEDULED: 'Scheduled', OPEN: 'Open', COMPLETED: 'Completed' };
var DAY_MS = 86400000;

var CFG_DEFAULTS = {
  MANAGER_NAME: 'Manager',
  MANAGER_EMAIL: 'manager-test@example.com',       // <- replace with YOUR test email
  LEADERSHIP_EMAIL: 'leadership-test@example.com', // <- replace with YOUR test email
  ALERT_HOUR: '18',          // daily check runs in this hour; same-day alert for items due today
  ESCALATE_AFTER_DAYS: '1',  // escalate when today >= due + N days and still incomplete
  MGR_ACK_SLA_HOURS: '24'    // escalate if manager alert unacknowledged after N hours
};

var TASK_HEADERS = ['TaskID','Task Name','Assignee','Assignee Email','Freq Days','Requires Ticket (Y/N)','Anchor Date','Active (Y/N)','Notes'];
var CYCLE_HEADERS = ['CycleKey','TaskID','Task Name','Assignee','Assignee Email','Freq Days','Cycle Start','Due Date','Status','Completed By','Completed At','Ticket / Proof','Mgr Alert Sent At','Escalated At','Mgr Ack By','Mgr Ack At'];
var LOG_HEADERS = ['LogKey','Type','CycleKey','At','Detail'];

/* ---------------- setup ---------------- */

function setupSheets() {
  var ss = SpreadsheetApp.getActive();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ensureSheet_(ss, SHEETS.TASKS, TASK_HEADERS);
    ensureSheet_(ss, SHEETS.CYCLES, CYCLE_HEADERS);
    ensureSheet_(ss, SHEETS.LOG, LOG_HEADERS);
    var cfg = ensureSheet_(ss, SHEETS.CONFIG, ['Key', 'Value']);
    var my = ensureSheet_(ss, SHEETS.MYTASKS, []);
    var dash = ensureSheet_(ss, SHEETS.DASH, []);

    // Config defaults (do not overwrite user edits)
    var existing = cfg.getDataRange().getValues();
    var keys = {};
    for (var i = 1; i < existing.length; i++) keys[existing[i][0]] = true;
    var row = cfg.getLastRow() + 1;
    Object.keys(CFG_DEFAULTS).forEach(function (k) {
      if (!keys[k]) { cfg.getRange(row++, 1, 1, 2).setValues([[k, CFG_DEFAULTS[k]]]); }
    });

    // Seed demo tasks only if TaskMaster is empty
    var tasks = sheet_(SHEETS.TASKS);
    if (tasks.getLastRow() < 2) {
      var today = startOfDay_(new Date());
      tasks.getRange(2, 1, 4, 9).setValues([
        ['T-001', 'Daily Operations Update', 'Employee A', 'emp-a-test@example.com', 1, 'N', today, 'Y', 'WhatsApp group update'],
        ['T-002', 'Portal Task A', 'Employee B', 'emp-b-test@example.com', 7, 'Y', today, 'Y', 'Portal ticket required'],
        ['T-003', 'Portal Task B', 'Employee C', 'emp-c-test@example.com', 15, 'Y', today, 'Y', 'Portal ticket required'],
        ['T-004', 'Portal Task C', 'Employee D', 'emp-d-test@example.com', 30, 'Y', today, 'Y', 'Portal ticket required']
      ]);
    }

    // MyTasks employee selector
    my.getRange('A1').setValue('MY PENDING TASKS — update Ticket ID / proof here, then tick Done. Nothing else to edit anywhere.');
    my.getRange('A2').setValue('Employee:');
    var assigneeRange = tasks.getRange(2, 3, Math.max(tasks.getLastRow() - 1, 1), 1);
    my.getRange('B2').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(assigneeRange, true).build());
    if (!my.getRange('B2').getValue()) my.getRange('B2').setValue(assigneeRange.getCell(1, 1).getValue());

    dash.getRange('A1').setValue('MANAGER DASHBOARD');
    dash.getRange('A2').setValue('Manager name (used for acknowledgements):');
    if (!dash.getRange('B2').getValue()) dash.getRange('B2').setValue(cfg_('MANAGER_NAME'));

    ensureCycles_();
    renderMyTasks_();
    renderDashboard_();
  } finally { lock.releaseLock(); }
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runDailyChecks') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDailyChecks').timeBased()
    .everyDays(1).atHour(parseInt(cfg_('ALERT_HOUR'), 10)).create();
  SpreadsheetApp.getActive().toast('Daily trigger installed for hour ' + cfg_('ALERT_HOUR') + ':00.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Recurring Tasks')
    .addItem('Run daily checks now', 'runDailyChecks')
    .addItem('Refresh My Tasks', 'renderMyTasks')
    .addItem('Refresh Dashboard', 'renderDashboard')
    .addSeparator()
    .addItem('Setup sheets', 'setupSheets')
    .addItem('Install daily trigger', 'installTriggers')
    .addSeparator()
    .addItem('DEMO: run all tests (one click)', 'demoRunAll')
    .addItem('DEMO: apply test emails', 'demoApplyTestEmails')
    .addItem('DEMO: set first open task due today', 'demoDueToday')
    .addItem('DEMO: age first open task for escalation', 'demoForceEscalation')
    .addToUi();
}

/* ---------------- cycle engine ---------------- */

function ensureCycles() { // menu-safe wrapper
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try { ensureCycles_(); } finally { lock.releaseLock(); }
}

function ensureCycles_() {
  var today = startOfDay_(new Date());
  var cyc = sheet_(SHEETS.CYCLES);
  var existing = cycleKeySet_();

  readTasks_().forEach(function (t) {
    if (t.active !== 'Y') return;
    var k0 = Math.max(0, Math.floor((today - t.anchor) / DAY_MS / t.freq));
    [k0, k0 + 1].forEach(function (k) {
      var start = new Date(t.anchor.getTime() + k * t.freq * DAY_MS);
      var key = t.id + '|' + iso_(start);
      if (existing[key]) return;
      var status = start > today ? STATUS.SCHEDULED : STATUS.OPEN;
      cyc.appendRow([key, t.id, t.name, t.assignee, t.email, t.freq, start,
        new Date(start.getTime() + (t.freq - 1) * DAY_MS), status, '', '', '', '', '', '', '']);
      existing[key] = true;
      log_('CYCLE|' + key, 'CYCLE_OPEN', key, 'status=' + status);
    });
  });

  // flip Scheduled -> Open when start date reached
  var vals = cyc.getDataRange().getValues();
  for (var r = 1; r < vals.length; r++) {
    if (vals[r][8] === STATUS.SCHEDULED && vals[r][6] && vals[r][6] <= today) {
      cyc.getRange(r + 1, 9).setValue(STATUS.OPEN);
    }
  }
}

/* ---------------- daily checks: alerts + escalation ---------------- */

function runDailyChecks() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    ensureCycles_();
    var now = new Date();
    var today = startOfDay_(now);
    var escDays = parseInt(cfg_('ESCALATE_AFTER_DAYS'), 10);
    var ackSlaMs = parseFloat(cfg_('MGR_ACK_SLA_HOURS')) * 3600000;
    var cyc = sheet_(SHEETS.CYCLES);
    var vals = cyc.getDataRange().getValues();
    var logKeys = logKeySet_();
    var alertItems = [], escItems = [];

    for (var i = 1; i < vals.length; i++) {
      var v = vals[i], row = i + 1;
      if (!v[0] || v[8] === STATUS.COMPLETED || v[8] === STATUS.SCHEDULED) continue;
      var key = v[0], due = v[7];

      // same-day manager alert: due <= today, not completed, not yet alerted
      if (due <= today && !v[12] && !logKeys['MGR_ALERT|' + key]) {
        alertItems.push({ row: row, key: key, task: v[2], who: v[3], due: due });
      }

      // escalation: overdue beyond SLA, OR alert sent but unacknowledged beyond ack SLA
      var overdueEsc = today > new Date(due.getTime() + escDays * DAY_MS);
      var ackEsc = v[12] && !v[15] && (now - v[12]) >= ackSlaMs;
      if (!v[13] && (overdueEsc || ackEsc) && !logKeys['ESC|' + key]) {
        escItems.push({ row: row, key: key, task: v[2], who: v[3], due: due,
          reason: overdueEsc ? 'unresolved past escalation SLA' : 'manager alert unacknowledged' });
      }
    }

    if (alertItems.length) {
      var body = 'The following recurring tasks are still incomplete (due today or overdue):\n\n' +
        alertItems.map(function (a) {
          return '• ' + a.task + ' — ' + a.who + ' (due ' + iso_(a.due) + ')';
        }).join('\n') +
        '\n\nPlease review the Dashboard tab and acknowledge each item.';
      if (sendMail_(cfg_('MANAGER_EMAIL'), '', '[OVERDUE] Recurring tasks pending — ' + alertItems.length + ' item(s)', body, 'MGR_ALERT')) {
        alertItems.forEach(function (a) {
          cyc.getRange(a.row, 13).setValue(now);
          log_('MGR_ALERT|' + a.key, 'MGR_ALERT', a.key, 'manager emailed');
        });
      }
    }

    if (escItems.length) {
      var ebody = 'ESCALATION: the following recurring tasks remain unresolved:\n\n' +
        escItems.map(function (a) {
          return '• ' + a.task + ' — ' + a.who + ' (due ' + iso_(a.due) + ', ' + a.reason + ')';
        }).join('\n') + '\n\nManager was alerted earlier; see Dashboard tab.';
      if (sendMail_(cfg_('LEADERSHIP_EMAIL'), cfg_('MANAGER_EMAIL'), '[ESCALATION] Recurring tasks unresolved — ' + escItems.length + ' item(s)', ebody, 'ESCALATION')) {
        escItems.forEach(function (a) {
          cyc.getRange(a.row, 14).setValue(now);
          log_('ESC|' + a.key, 'ESCALATION', a.key, a.reason);
        });
      }
    }

    renderMyTasks_();
    renderDashboard_();
    SpreadsheetApp.getActive().toast('Checks done: ' + alertItems.length + ' alert(s), ' + escItems.length + ' escalation(s).');
  } finally { lock.releaseLock(); }
}

/* ---------------- completion + acknowledgement ---------------- */

function completeCycle(cycleKey, proof, by) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var cyc = sheet_(SHEETS.CYCLES);
    var vals = cyc.getDataRange().getValues();
    var row = -1, v = null;
    for (var i = 1; i < vals.length; i++) if (vals[i][0] === cycleKey) { row = i + 1; v = vals[i]; break; }
    var ss = SpreadsheetApp.getActive();
    if (row < 0) { log_('ERR|' + cycleKey + '|' + Date.now(), 'ERROR', cycleKey, 'cycle not found'); return false; }
    if (v[8] === STATUS.COMPLETED) { ss.toast('Already completed — no duplicate entry.'); return false; }
    var t = taskById_(v[1]);
    if (t && t.reqTicket === 'Y' && !String(proof).trim()) {
      ss.toast('Ticket ID is required for "' + v[2] + '". Enter it, then tick Done.');
      return false;
    }
    cyc.getRange(row, 9, 1, 4).setValues([[STATUS.COMPLETED, by, new Date(),
      String(proof).trim() || 'Self-attested in MyTasks sheet']]);
    log_('COMPLETE|' + cycleKey, 'COMPLETE', cycleKey, 'by=' + by + ' proof=' + (proof || 'self-attest'));
    ensureCycles_(); // open/track next cycle automatically
    ss.toast('Completed: ' + v[2]);
    return true;
  } finally { lock.releaseLock(); }
}

function acknowledge(cycleKey, by) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var cyc = sheet_(SHEETS.CYCLES);
    var vals = cyc.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (vals[i][0] === cycleKey) {
        if (vals[i][15]) { SpreadsheetApp.getActive().toast('Already acknowledged.'); return; }
        cyc.getRange(i + 1, 15, 1, 2).setValues([[by, new Date()]]);
        log_('ACK|' + cycleKey, 'MGR_ACK', cycleKey, 'by=' + by);
        return;
      }
    }
  } finally { lock.releaseLock(); }
}

/* ---------------- onEdit router ---------------- */

function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet(), name = sh.getName(), r = e.range.getRow(), c = e.range.getColumn();
  var a1 = sh.getRange(r, 1).getValue();

  if (name === SHEETS.MYTASKS) {
    if (e.range.getA1Notation() === 'B2') { renderMyTasks(); return; }
    if (c === 7 && e.value === 'TRUE' && typeof a1 === 'string' && a1.indexOf('|') > 0) {
      var ok = completeCycle(a1, sh.getRange(r, 6).getValue(), sh.getRange('B2').getValue());
      if (!ok) sh.getRange(r, 7).setValue(false);
      renderMyTasks(); renderDashboard();
    }
  }
  if (name === SHEETS.DASH && c === 7 && e.value === 'TRUE' && typeof a1 === 'string' && a1.indexOf('|') > 0) {
    acknowledge(a1, sh.getRange('B2').getValue() || 'Manager');
    renderDashboard();
  }
}

/* ---------------- rendering ---------------- */

function renderMyTasks() { renderMyTasks_(); }
function renderMyTasks_() {
  var my = sheet_(SHEETS.MYTASKS);
  var tasks = sheet_(SHEETS.TASKS);
  if (tasks.getLastRow() > 1) { // keep B2 picker in sync with new assignees
    my.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(tasks.getRange(2, 3, tasks.getLastRow() - 1, 1), true).build());
  }
  var emp = my.getRange('B2').getValue();
  var today = startOfDay_(new Date());
  my.getRange('A5:H200').clear({ contentsOnly: true }).clearDataValidations();
  my.getRange('A5:G5').setValues([['CycleKey', 'Task', 'Due', 'Status', 'Proof needed', 'Ticket ID / proof →', 'Done?']])
    .setFontWeight('bold');

  var rows = sheet_(SHEETS.CYCLES).getDataRange().getValues();
  var tasks = {}; readTasks_().forEach(function (t) { tasks[t.id] = t; });
  var out = [], keys = [];
  for (var i = 1; i < rows.length; i++) {
    var v = rows[i];
    if (v[3] !== emp || v[8] === STATUS.COMPLETED || v[8] === STATUS.SCHEDULED) continue;
    if (v[6] > today) continue;
    var d = Math.round((v[7] - today) / DAY_MS);
    var status = d < 0 ? 'OVERDUE by ' + (-d) + 'd' : (d === 0 ? 'Due today' : 'Due in ' + d + 'd');
    var t = tasks[v[1]];
    out.push([v[0], v[2], v[7], status, t && t.reqTicket === 'Y' ? 'Ticket ID (required)' : 'Note (optional)', '', false]);
  }
  out.sort(function (a, b) { return a[2] - b[2]; });
  if (!out.length) { my.getRange('A6').setValue('(nothing pending — all caught up)'); return; }
  var rng = my.getRange(6, 1, out.length, 7);
  rng.setValues(out);
  my.getRange(6, 7, out.length, 1).insertCheckboxes();
  my.getRange(6, 3, out.length, 1).setNumberFormat('yyyy-mm-dd');
}

function renderDashboard() { renderDashboard_(); }
function renderDashboard_() {
  var dash = sheet_(SHEETS.DASH);
  var today = startOfDay_(new Date());
  var rows = sheet_(SHEETS.CYCLES).getDataRange().getValues();
  dash.getRange('A4:H400').clear({ contentsOnly: true }).clearDataValidations();

  // Section 1: per-employee summary (live formulas)
  dash.getRange('A4').setValue('Per-employee summary').setFontWeight('bold');
  dash.getRange('A5:D5').setValues([['Employee', 'Pending (not yet due)', 'Overdue', 'Completed last 30d']]).setFontWeight('bold');
  var emps = [], seen = {};
  readTasks_().forEach(function (t) { if (!seen[t.assignee]) { seen[t.assignee] = 1; emps.push(t.assignee); } });
  emps.forEach(function (e, i) {
    var r = 6 + i;
    dash.getRange(r, 1).setValue(e);
    dash.getRange(r, 2).setFormula('=COUNTIFS(Cycles!$D:$D,$A' + r + ',Cycles!$I:$I,"Open",Cycles!$H:$H,">="&TODAY())');
    dash.getRange(r, 3).setFormula('=COUNTIFS(Cycles!$D:$D,$A' + r + ',Cycles!$I:$I,"Open",Cycles!$H:$H,"<"&TODAY())');
    dash.getRange(r, 4).setFormula('=COUNTIFS(Cycles!$D:$D,$A' + r + ',Cycles!$I:$I,"Completed",Cycles!$K:$K,">="&TODAY()-30)');
  });

  // Section 2: items needing acknowledgement
  var r0 = 6 + emps.length + 2;
  dash.getRange(r0, 1).setValue('Needs acknowledgement (alerted, still open)').setFontWeight('bold');
  dash.getRange(r0 + 1, 1, 1, 7).setValues([['CycleKey', 'Task', 'Assignee', 'Due', 'Days overdue', 'Alert sent', 'Ack?']]).setFontWeight('bold');
  var ackRows = [];
  for (var i = 1; i < rows.length; i++) {
    var v = rows[i];
    if (v[8] === STATUS.COMPLETED || v[8] === STATUS.SCHEDULED || !v[12] || v[15]) continue;
    ackRows.push([v[0], v[2], v[3], v[7], Math.max(0, Math.round((today - v[7]) / DAY_MS)), v[12], false]);
  }
  var r1;
  if (ackRows.length) {
    dash.getRange(r0 + 2, 1, ackRows.length, 7).setValues(ackRows);
    dash.getRange(r0 + 2, 7, ackRows.length, 1).insertCheckboxes();
    r1 = r0 + 2 + ackRows.length + 2;
  } else {
    dash.getRange(r0 + 2, 1).setValue('(none)');
    r1 = r0 + 4;
  }

  // Section 3: all open / overdue items
  dash.getRange(r1, 1).setValue('All open & overdue items').setFontWeight('bold');
  dash.getRange(r1 + 1, 1, 1, 8).setValues([['CycleKey', 'Task', 'Assignee', 'Due', 'Status', 'Alert sent', 'Escalated', 'Acked']]).setFontWeight('bold');
  var all = [];
  for (var j = 1; j < rows.length; j++) {
    var w = rows[j];
    if (w[8] === STATUS.COMPLETED || w[8] === STATUS.SCHEDULED) continue;
    var dd = Math.round((w[7] - today) / DAY_MS);
    all.push([w[0], w[2], w[3], w[7],
      dd < 0 ? 'OVERDUE ' + (-dd) + 'd' : (dd === 0 ? 'Due today' : 'Open'),
      w[12] ? 'Y' : '', w[13] ? 'Y' : '', w[15] ? 'Y' : '']);
  }
  all.sort(function (a, b) { return a[3] - b[3]; });
  if (all.length) dash.getRange(r1 + 2, 1, all.length, 8).setValues(all);
  else dash.getRange(r1 + 2, 1).setValue('(none)');
}

/* ---------------- demo helpers ---------------- */

// Sets the first Open cycle's due date to TODAY so a manual runDailyChecks()
// simulates the 18:00 end-of-day run -> same-day manager alert only.
function demoDueToday() {
  var cyc = sheet_(SHEETS.CYCLES);
  var vals = cyc.getDataRange().getValues();
  var today = startOfDay_(new Date());
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][8] === STATUS.OPEN) {
      cyc.getRange(i + 1, 8).setValue(today);
      SpreadsheetApp.getActive().toast('"' + vals[i][2] + '" due date set to today. Now run "Run daily checks now".');
      return;
    }
  }
  SpreadsheetApp.getActive().toast('No open cycle found.');
}

// Simulates elapsed time: due 2 days ago + manager alert stamped 25h ago with
// no acknowledgement -> next runDailyChecks() escalates to leadership.
function demoForceEscalation() {
  var cyc = sheet_(SHEETS.CYCLES);
  var vals = cyc.getDataRange().getValues();
  var twoDaysAgo = new Date(startOfDay_(new Date()).getTime() - 2 * DAY_MS);
  var longAgo = new Date(Date.now() - 25 * 3600000);
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][8] === STATUS.OPEN && !vals[i][13]) {
      cyc.getRange(i + 1, 8).setValue(twoDaysAgo);
      if (!vals[i][12]) {
        cyc.getRange(i + 1, 13).setValue(longAgo);
        log_('MGR_ALERT|' + vals[i][0], 'MGR_ALERT', vals[i][0], 'simulated alert (demo)');
      }
      SpreadsheetApp.getActive().toast('"' + vals[i][2] + '" aged: due 2 days ago, alert sent 25h ago. Now run "Run daily checks now".');
      return;
    }
  }
  SpreadsheetApp.getActive().toast('No suitable open cycle found.');
}

// One-click: fill Config emails + employee emails with the owner's test
// addresses (all owned by the submitter; foces@ceconline.edu excluded).
function demoApplyTestEmails() {
  var emails = {
    MANAGER_EMAIL: 'sebinmathew543@gmail.com',
    LEADERSHIP_EMAIL: 'sebin.dev.yt@gmail.com'
  };
  Object.keys(emails).forEach(function (k) { adminSet(k, emails[k]); });
  var empEmails = {
    'T-001': 'sebinmathew543@gmail.com',
    'T-002': 'chn24bt402@ceconline.edu',
    'T-003': 'sebinmathew@ieee.org',
    'T-004': 'emp-d-test@example.com'
  };
  var tasks = sheet_(SHEETS.TASKS);
  var vals = tasks.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (empEmails[vals[i][0]]) tasks.getRange(i + 1, 4).setValue(empEmails[vals[i][0]]);
  }
  SpreadsheetApp.getActive().toast('Test emails applied: manager + leadership set; 4 employees assigned.');
}

function firstCycleKey_(taskId) {
  var vals = sheet_(SHEETS.CYCLES).getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).indexOf(taskId + '|') === 0 && vals[i][8] === STATUS.OPEN) return vals[i][0];
  }
  return null;
}

function alertedCycleKey_() {
  var vals = sheet_(SHEETS.CYCLES).getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][12] && !vals[i][15] && vals[i][8] !== STATUS.COMPLETED) return vals[i][0];
  }
  return null;
}

// One-click full demo: runs the equivalent of all 8 test cases in order and
// leaves a complete, screenshot-ready state in Cycles / MyTasks / Dashboard / Log.
function demoRunAll() {
  demoApplyTestEmails();                                   // emails -> Config + TaskMaster
  installTriggers();                                       // daily 18:00 trigger installed

  // T1 + T2: complete a 7-day ticket task from a single point of entry
  var k1 = firstCycleKey_('T-002');
  if (k1) completeCycle(k1, 'TKT-1001', 'Employee B');

  // T3: daily task without a ticket (self-attestation + note)
  var k2 = firstCycleKey_('T-001');
  if (k2) completeCycle(k2, 'Posted in WA group 9:15am', 'Employee A');

  // T4: same-day manager alert on a due-today item
  demoDueToday();
  runDailyChecks();

  // T5: manager acknowledges the alerted item
  var k3 = alertedCycleKey_();
  if (k3) acknowledge(k3, 'Manager');

  // T6: unresolved past SLA -> leadership escalation (both paths satisfied)
  demoForceEscalation();
  runDailyChecks();

  // T8: re-run twice — must create zero new cycles/emails/entries
  runDailyChecks();
  runDailyChecks();

  renderMyTasks_();
  renderDashboard_();
  SpreadsheetApp.getActive().toast('Demo complete: see Cycles, MyTasks, Dashboard and Log.');
}

/* ---------------- admin / API helpers ---------------- */

// Set a Config value remotely (clasp run adminSet -p '["KEY","value"]').
function adminSet(key, value) {
  var cfg = sheet_(SHEETS.CONFIG);
  var vals = cfg.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] === key) { cfg.getRange(i + 1, 2).setValue(value); return 'set ' + key + '=' + value; }
  }
  cfg.appendRow([key, value]);
  return 'added ' + key + '=' + value;
}

// Dump a tab's values as JSON for remote verification/screenshots.
function adminDump(sheetName, maxRows) {
  var sh = sheet_(sheetName);
  if (!sh) return 'no sheet ' + sheetName;
  return JSON.stringify(sh.getDataRange().getValues().slice(0, maxRows || 60));
}

/* ---------------- utilities ---------------- */

function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (headers.length && sh.getLastRow() === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  return sh;
}

function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

function iso_(d) { return Utilities.formatDate(d, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd'); }

function cfg_(key) {
  var vals = sheet_(SHEETS.CONFIG).getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) if (vals[i][0] === key) return vals[i][1];
  return CFG_DEFAULTS[key] || '';
}

function readTasks_() {
  var vals = sheet_(SHEETS.TASKS).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    var v = vals[i];
    if (!v[0]) continue;
    out.push({ id: v[0], name: v[1], assignee: v[2], email: v[3], freq: parseInt(v[4], 10),
      reqTicket: String(v[5]).toUpperCase(), anchor: startOfDay_(new Date(v[6])), active: String(v[7]).toUpperCase() });
  }
  return out;
}

function taskById_(id) {
  var ts = readTasks_();
  for (var i = 0; i < ts.length; i++) if (ts[i].id === id) return ts[i];
  return null;
}

function cycleKeySet_() {
  var vals = sheet_(SHEETS.CYCLES).getRange('A:A').getValues();
  var s = {};
  for (var i = 0; i < vals.length; i++) if (vals[i][0]) s[vals[i][0]] = true;
  return s;
}

function logKeySet_() {
  var vals = sheet_(SHEETS.LOG).getRange('A:A').getValues();
  var s = {};
  for (var i = 0; i < vals.length; i++) if (vals[i][0]) s[vals[i][0]] = true;
  return s;
}

function log_(key, type, cycleKey, detail) {
  sheet_(SHEETS.LOG).appendRow([key, type, cycleKey, new Date(), detail]);
}

// Guarded send: placeholder example.com addresses are never emailed;
// the intended email is recorded in Log as EMAIL_SKIPPED instead.
function sendMail_(to, cc, subject, body, type) {
  var skipped = /@example\.com/i.test(to) || !to;
  if (cc && /@example\.com/i.test(cc)) cc = ''; // never cc a placeholder
  if (skipped) {
    log_('EMAIL_SKIPPED|' + type + '|' + Date.now(), 'EMAIL_SKIPPED', '', 'to=' + to + ' subject=' + subject);
    return true; // treated as sent so stamps/keys still record for dedupe demos
  }
  try {
    MailApp.sendEmail({ to: to, cc: cc || '', subject: subject, body: body });
    return true;
  } catch (err) {
    log_('ERR|' + type + '|' + Date.now(), 'ERROR', '', 'sendMail failed: ' + err);
    return false; // unstamped -> retried next run
  }
}
