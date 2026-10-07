/* ==========================================================================
   UKRO Rope Rescue Challenge — scoring backend (Google Apps Script)
   --------------------------------------------------------------------------
   Lives inside the event's Google Sheet. The assessor scoresheet and the
   organiser standings page (both on GitHub Pages) talk to it.

   ONE-OFF SETUP (about 10 minutes):
     1. Create a blank Google Sheet. Extensions > Apps Script.
     2. Delete the sample code, paste in this whole file, click Save.
     3. Choose the function "setup" in the toolbar and click Run.
        Approve the permission prompt (it only touches this one sheet).
        This creates the tabs, pre-filled with the 2026 teams and timetable.
     4. In the sheet, open the Settings tab and change both PINs.
     5. Deploy > New deployment > type "Web app".
          Execute as: Me      Who has access: Anyone
        Copy the Web app URL (ends in /exec) into the two HTML pages.
     6. After editing this script later: Deploy > Manage deployments >
        edit (pencil) > Version: New version > Deploy. The URL stays the same.

   HOW THE DATA WORKS
     - Submissions is an append-only log. Every scoresheet sent is one row;
       nothing is ever overwritten.
     - If a sheet is re-sent for the same team and scenario, the LATEST one
       counts; earlier versions stay visible and are flagged on the
       organiser page.
     - To exclude a submission, type VOID in its Status column.
     - Scores are recalculated here from the individual deductions, so the
       totals can't be mistyped.
   ========================================================================== */

/* ---------------- Scoring rules (edit here if the rules change) ---------- */
var RULES = {
  technicalMax: 100,
  commandMax: 50,
  medicalMax: 50,
  nonSafetyCritical: 2,   // technical & command, per occurrence
  safetyCritical: 25,     // technical & command, per occurrence
  medicalMinor: 2,        // medical, highest deduction once per topic
  medicalCritical: 10,
  completionBonus: 25,
  maxCountPerBox: 20
};

/* ---------------- Assessment criteria (match the paper sheets) ----------- */
var CRITERIA = {
  technical: [
    ['anchors', 'Anchors & Belays', 'Strong anchors, appropriate belays & redundancy'],
    ['rigging', 'Rigging & Systems', 'Tidy, efficient rigging; systems monitored'],
    ['ropeProtection', 'Rope Protection', 'Effective protection & rope management'],
    ['edge', 'Edge Transitions', 'Controlled, protected movement at edges'],
    ['hauling', 'Hauling & Lowering', 'Smooth, controlled & efficient operation'],
    ['equipment', 'Equipment Selection', 'Appropriate kit, proficiently used'],
    ['teamSafety', 'Team Safety', 'Safe PPE, workspace & working systems'],
    ['systemCheck', 'System Safety Check', 'Systems checked, monitored & critical points controlled'],
    ['casualty', 'Casualty & Stretcher Management', 'Protected casualty; smooth, controlled movement'],
    ['attendant', 'Attendant Skills', 'Secure attachment, positioning & casualty care']
  ],
  command: [
    ['hazardId', 'Hazard ID', 'Identify, control, monitor & review hazards'],
    ['planning', 'Planning', 'Clear priorities, tasking & workable plan'],
    ['planAB', 'Plan A & B and Emergency Plan', 'Viable contingencies; reassess & adapt'],
    ['communication', 'Communication', 'Clear instructions; understood & confirmed'],
    ['positioning', 'Positioning', 'Maintains overview of team & operations'],
    ['taskFocus', 'Task Focus', 'Maintains command; avoids task fixation'],
    ['coaching', 'Coaching', 'Appropriate coaching without losing command'],
    ['decisions', 'Decision Making', 'Consult, decide & communicate promptly'],
    ['teamMgmt', 'Team Management', 'Logical deployment & coordinated activity'],
    ['tempo', 'Team Tempo', 'Controlled pace, motivation & welfare'],
    ['sa', 'Situational Awareness', 'Constantly monitors & reviews operations'],
    ['assessorCheck', 'Assessor Safety Check', 'Safety check before critical movement']
  ],
  medical: [
    ['scene', 'Scene assessment and safety', 'Consideration of mechanism of injury (request for additional resources)'],
    ['communication', 'Communication', 'Initial indirect and direct contact and reassurance to casualty; communication between team leader, wider team and medic'],
    ['catHaem', 'Cat Haem / Massive Haem', 'Rapid, effective, systematic check, consideration, justification and appropriate treatment and kit/equipment use of life-threatening bleed'],
    ['airway', 'Airway', 'Rapid, effective, systematic check, consideration, justification and appropriate treatment and kit/equipment of airway and consideration of C-spine'],
    ['breathing', 'Breathing / Respiration', 'Rapid, effective, systematic check, consideration, justification and appropriate treatment and kit/equipment of chest trauma or illness'],
    ['circulation', 'Circulation', 'Rapid, effective, systematic check, consideration, justification and appropriate treatment and kit/equipment of shock, lethal triad and suspected internal bleeding'],
    ['disability', 'Disability / Head', 'Rapid, effective, systematic check, consideration, justification and appropriate treatment and kit/equipment of neurological system'],
    ['environment', 'Environment / Handling / Heat', 'Rapid, effective, systematic check, consideration, justification and appropriate treatment and kit/equipment of protection from the environment, heat management, casualty comfort and packaging for extrication and environment'],
    ['casualtyPpe', 'Casualty PPE, comfort and safety', 'Throughout'],
    ['handover', 'Hand over', ''],
    ['casualtyComms', 'Casualty communication and reassurance', 'Throughout'],
    ['kitSelection', 'Medical kit selection', 'Appropriate selection and justification of medical kit/equipment to leave with team, take to casualty and utilise']
  ]
};

/* ---------------- Tab layout ---------------- */
var TAB = {
  settings: 'Settings',
  teams: 'Teams',
  locations: 'Locations',
  schedule: 'Schedule',
  opening: 'Opening Day',
  submissions: 'Submissions'
};

var SUB_HEADERS = ['Received', 'Submission ID', 'Sheet', 'Date', 'Rotation', 'Location', 'Team',
  'Assessor', 'Team leader', 'Technical', 'Command', 'Completed', 'Medical',
  'Time allocation', 'Completion time', 'Casualty contact', 'Safety critical detail',
  'Good practice', 'Key learning', 'Comments', 'Detail (JSON)', 'Status'];

/* ==========================================================================
   SETUP — run once from the Apps Script editor
   ========================================================================== */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function tab(name, headers, rows, widths) {
    var sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) return sh;           // never overwrite live data
    if (!sh) sh = ss.insertSheet(name);
    sh.clear();
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#2B3990').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    if (rows && rows.length) {
      sh.getRange(2, 1, rows.length, headers.length).setNumberFormat('@').setValues(rows);
    }
    (widths || []).forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    return sh;
  }

  tab(TAB.settings, ['Setting', 'Value', 'Notes'], [
    ['Event name', 'UKRO National Rope Rescue Challenge 2026', 'Shown at the top of both pages'],
    ['Assessor PIN', 'change-me-assessor', 'Shared with assessors. Use 6+ characters.'],
    ['Organiser PIN', 'change-me-organiser', 'Organisers only — unlocks the standings.'],
    ['Opening day target minutes', '90', 'Points are earned for each minute under this'],
    ['Opening day points per minute', '3', '']
  ], [220, 320, 360]);

  tab(TAB.teams, ['Team'], SAMPLE_TEAMS.map(function (t) { return [t]; }), [240]);

  tab(TAB.locations, ['Location', 'Medical scored?'],
    SAMPLE_LOCATIONS.map(function (l) { return [l[0], l[1] ? 'Yes' : 'No']; }), [320, 140]);

  tab(TAB.schedule, ['Date (YYYY-MM-DD)', 'Rotation', 'Time', 'Location', 'Team'],
    SAMPLE_SCHEDULE.map(function (r) { return [r[0], String(r[1]), r[2], r[3], r[4]]; }),
    [150, 80, 120, 320, 200]);

  tab(TAB.opening, ['Team', 'Time (minutes)', 'Outcome (Finished / DNF)'],
    SAMPLE_TEAMS.map(function (t) { return [t, '', '']; }), [240, 130, 200]);

  var sub = tab(TAB.submissions, SUB_HEADERS, [], []);
  sub.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');

  var s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1 && s1.getLastRow() === 0) ss.deleteSheet(s1);
  ss.setActiveSheet(ss.getSheetByName(TAB.settings));
}

/* ==========================================================================
   WEB APP ENTRY POINTS
   Pages POST JSON as text/plain (avoids a CORS preflight).
   ========================================================================== */
function doGet() {
  return json_({ ok: true, service: 'UKRO scoring', hint: 'POST only' });
}

function doPost(e) {
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var data = readAll_();
    var role = roleFor_(req.pin, data.settings);
    if (!role) return json_({ ok: false, error: 'bad_pin', message: 'PIN not recognised.' });

    if (req.action === 'config') {
      return json_({ ok: true, role: role, config: publicConfig_(data) });
    }
    if (req.action === 'submit') {
      var lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        var fresh = readAll_();                        // re-read inside the lock
        var res = acceptSubmission_(req.submission, fresh, new Date());
        if (res.ok && !res.duplicate) {
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.submissions).appendRow(res.row);
        }
        delete res.row;
        return json_(res);
      } finally {
        lock.releaseLock();
      }
    }
    if (req.action === 'standings') {
      if (role !== 'organiser') return json_({ ok: false, error: 'forbidden', message: 'Organiser PIN required.' });
      return json_({ ok: true, result: buildStandings(data) });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ==========================================================================
   READING THE SHEET
   ========================================================================== */
function readAll_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  function rows(name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  }
  return parseData(
    rows(TAB.settings), rows(TAB.teams), rows(TAB.locations),
    rows(TAB.schedule), rows(TAB.opening), rows(TAB.submissions)
  );
}

/* ==========================================================================
   PURE LOGIC — no Google services below this line (tested outside Sheets)
   ========================================================================== */
function str_(v) { return v === null || v === undefined ? '' : String(v).trim(); }

function normDate_(v) {
  if (v instanceof Date) {
    var m = v.getMonth() + 1, d = v.getDate();
    return v.getFullYear() + '-' + (m < 10 ? '0' : '') + m + '-' + (d < 10 ? '0' : '') + d;
  }
  return str_(v);
}

function key_(date, rotation, location, team) {
  return [normDate_(date), str_(rotation), str_(location), str_(team)].join('|');
}

function parseData(settingsRows, teamRows, locRows, schedRows, openRows, subRows) {
  var settings = {};
  settingsRows.forEach(function (r) { if (str_(r[0])) settings[str_(r[0])] = str_(r[1]); });

  var teams = teamRows.map(function (r) { return str_(r[0]); }).filter(Boolean);

  var locations = locRows.filter(function (r) { return str_(r[0]); }).map(function (r) {
    return { name: str_(r[0]), medical: /^y/i.test(str_(r[1])) };
  });

  var schedule = schedRows.filter(function (r) { return str_(r[0]) && str_(r[4]); }).map(function (r) {
    return { date: normDate_(r[0]), rotation: str_(r[1]), time: str_(r[2]), location: str_(r[3]), team: str_(r[4]) };
  });

  var opening = openRows.filter(function (r) { return str_(r[0]); }).map(function (r) {
    return { team: str_(r[0]), minutes: str_(r[1]), outcome: str_(r[2]) };
  });

  var submissions = subRows.filter(function (r) { return str_(r[1]); }).map(function (r, i) {
    var o = {};
    SUB_HEADERS.forEach(function (h, j) { o[h] = r[j]; });
    return {
      order: i,
      received: o['Received'] instanceof Date ? o['Received'].toISOString() : str_(o['Received']),
      id: str_(o['Submission ID']),
      sheet: str_(o['Sheet']),
      date: normDate_(o['Date']),
      rotation: str_(o['Rotation']),
      location: str_(o['Location']),
      team: str_(o['Team']),
      assessor: str_(o['Assessor']),
      teamLeader: str_(o['Team leader']),
      technical: num_(o['Technical']),
      command: num_(o['Command']),
      completed: str_(o['Completed']),
      medical: num_(o['Medical']),
      void: /^void/i.test(str_(o['Status']))
    };
  });

  return { settings: settings, teams: teams, locations: locations, schedule: schedule, opening: opening, submissions: submissions };
}

function num_(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(v);
  return isFinite(n) ? n : null;
}

function roleFor_(pin, settings) {
  pin = str_(pin);
  if (!pin) return null;
  if (pin === str_(settings['Organiser PIN'])) return 'organiser';
  if (pin === str_(settings['Assessor PIN'])) return 'assessor';
  return null;
}

function publicConfig_(data) {
  return {
    event: data.settings['Event name'] || 'UKRO Rope Rescue Challenge',
    teams: data.teams,
    locations: data.locations,
    schedule: data.schedule,
    criteria: CRITERIA,
    rules: RULES
  };
}

/* ---------- Scoring a single sheet ---------- */
function scoreTC(deductions) {
  var out = { technical: RULES.technicalMax, command: RULES.commandMax, errors: [] };
  ['technical', 'command'].forEach(function (section) {
    var given = (deductions && deductions[section]) || {};
    CRITERIA[section].forEach(function (c) {
      var d = given[c[0]] || {};
      var nsc = Number(d.nsc || 0), sc = Number(d.sc || 0);
      if (!isInt_(nsc) || !isInt_(sc)) { out.errors.push(c[1] + ': counts must be whole numbers 0–' + RULES.maxCountPerBox); return; }
      out[section] -= nsc * RULES.nonSafetyCritical + sc * RULES.safetyCritical;
    });
  });
  return out;
}

function scoreMedical(deductions) {
  var out = { medical: RULES.medicalMax, errors: [] };
  var given = (deductions && deductions.medical) || {};
  CRITERIA.medical.forEach(function (c) {
    var lvl = str_((given[c[0]] || {}).level || 'none');
    if (lvl === 'minor') out.medical -= RULES.medicalMinor;
    else if (lvl === 'critical') out.medical -= RULES.medicalCritical;
    else if (lvl !== 'none') out.errors.push(c[1] + ': unknown deduction level');
  });
  return out;
}

function isInt_(n) { return isFinite(n) && Math.floor(n) === n && n >= 0 && n <= RULES.maxCountPerBox; }

/* ---------- Validating and logging a submission ---------- */
function acceptSubmission_(s, data, now) {
  if (!s || typeof s !== 'object') return { ok: false, error: 'invalid', message: 'No scoresheet received.' };
  var id = str_(s.id);
  if (!id) return { ok: false, error: 'invalid', message: 'Missing submission ID.' };

  // Same ID already logged = a retry from a patchy connection. Accept quietly.
  for (var i = 0; i < data.submissions.length; i++) {
    if (data.submissions[i].id === id) return { ok: true, duplicate: true, id: id };
  }

  var sheet = s.sheet === 'medical' ? 'medical' : (s.sheet === 'tc' ? 'tc' : '');
  if (!sheet) return { ok: false, error: 'invalid', message: 'Unknown sheet type.' };

  var k = key_(s.date, s.rotation, s.location, s.team);
  var slot = null;
  data.schedule.forEach(function (r) { if (key_(r.date, r.rotation, r.location, r.team) === k) slot = r; });
  if (!slot) return { ok: false, error: 'not_scheduled', message: 'That team, location and rotation is not on the Schedule tab.' };

  var loc = null;
  data.locations.forEach(function (l) { if (l.name === slot.location) loc = l; });
  if (sheet === 'medical' && loc && !loc.medical) {
    return { ok: false, error: 'invalid', message: 'Medical is not scored at ' + slot.location + '.' };
  }
  if (!str_(s.assessor)) return { ok: false, error: 'invalid', message: 'Assessor name is required.' };

  var tech = '', cmd = '', med = '', completed = '';
  var errors = [];
  if (sheet === 'tc') {
    var t = scoreTC(s.deductions);
    errors = t.errors; tech = t.technical; cmd = t.command;
    completed = s.completed === 'Yes' ? 'Yes' : (s.completed === 'No' ? 'No' : '');
    if (!completed) errors.push('Scenario completion (Yes / No) is required.');
    if ((tech < 0 || cmd < 0) && !s.confirmNegative) {
      return { ok: false, error: 'confirm_negative', message: 'A total is below zero and needs confirming.' };
    }
  } else {
    var m = scoreMedical(s.deductions);
    errors = m.errors; med = m.medical;
    if (med < 0 && !s.confirmNegative) {
      return { ok: false, error: 'confirm_negative', message: 'The medical total is below zero and needs confirming.' };
    }
  }
  if (errors.length) return { ok: false, error: 'invalid', message: errors.join(' ') };

  var detail = { deductions: s.deductions || {}, deviceTime: str_(s.deviceTime), app: str_(s.app) };
  var row = [
    now, id, sheet === 'tc' ? 'Technical & Command' : 'Medical',
    slot.date, slot.rotation, slot.location, slot.team,
    str_(s.assessor), str_(s.teamLeader),
    tech, cmd, completed, med,
    str_(s.timeAllocation), str_(s.completionTime), str_(s.contactTime),
    str_(s.safetyCritical), str_(s.goodPractice), str_(s.keyLearning), str_(s.comments),
    JSON.stringify(detail), ''
  ];
  // Text-format the free-text cells so Sheets doesn't reinterpret "12:30" etc.
  // (also stops a name or note starting with "=" being treated as a formula)
  [3, 7, 8, 13, 14, 15, 16, 17, 18, 19].forEach(function (j) { if (row[j]) row[j] = "'" + row[j]; });
  return { ok: true, id: id, row: row, scores: { technical: tech, command: cmd, completed: completed, medical: med } };
}

/* ---------- Standings ---------- */
function buildStandings(data) {
  var target = Number(data.settings['Opening day target minutes'] || 90);
  var ppm = Number(data.settings['Opening day points per minute'] || 3);

  var medicalAt = {};
  data.locations.forEach(function (l) { medicalAt[l.name] = l.medical; });

  var days = [];
  data.schedule.forEach(function (r) { if (days.indexOf(r.date) < 0) days.push(r.date); });
  days.sort();

  // Group live submissions by scenario, sheet type. Latest (by log order) wins.
  var bySlot = {}, unscheduled = [], voided = 0;
  var slotKeys = {};
  data.schedule.forEach(function (r) { slotKeys[key_(r.date, r.rotation, r.location, r.team)] = true; });
  data.submissions.forEach(function (s) {
    if (s.void) { voided++; return; }
    var k = key_(s.date, s.rotation, s.location, s.team);
    if (!slotKeys[k]) { unscheduled.push(s); return; }
    var type = /^med/i.test(s.sheet) ? 'med' : 'tc';
    bySlot[k] = bySlot[k] || { tc: [], med: [] };
    bySlot[k][type].push(s);
  });

  var teamsMap = {};
  function team(name) {
    if (!teamsMap[name]) {
      teamsMap[name] = { team: name, opening: 0, openingNote: '', days: {}, total: 0,
        scheduled: 0, scored: 0, partial: 0, command: 0, commandCount: 0, commandOf: 0,
        medical: 0, medicalCount: 0, medicalOf: 0 };
      days.forEach(function (d) { teamsMap[name].days[d] = 0; });
    }
    return teamsMap[name];
  }
  data.teams.forEach(team);

  data.opening.forEach(function (o) {
    var t = team(o.team);
    var mins = Number(o.minutes);
    if (/^dnf/i.test(o.outcome)) { t.opening = 0; t.openingNote = 'DNF'; }
    else if (o.minutes !== '' && isFinite(mins)) { t.opening = Math.max(0, Math.round((target - mins) * ppm)); t.openingNote = mins + ' min'; }
    else { t.openingNote = 'Not entered'; }
  });

  var scenarios = [], checks = [];
  data.schedule.forEach(function (r) {
    var k = key_(r.date, r.rotation, r.location, r.team);
    var g = bySlot[k] || { tc: [], med: [] };
    var tc = g.tc.length ? g.tc[g.tc.length - 1] : null;
    var med = g.med.length ? g.med[g.med.length - 1] : null;
    var needsMed = !!medicalAt[r.location];
    var t = team(r.team);
    t.scheduled++;
    t.commandOf++;
    if (needsMed) t.medicalOf++;

    var status, points = null;
    if (tc && (med || !needsMed)) {
      status = 'scored';
      points = tc.technical + tc.command + (tc.completed === 'Yes' ? RULES.completionBonus : 0) + (needsMed ? med.medical : 0);
      t.scored++;
      t.days[r.date] = (t.days[r.date] || 0) + points;
      t.total += points;
    } else if (tc || med) {
      status = 'partial';
      t.partial++;
      checks.push({ type: 'missing', slot: r, message: 'Awaiting ' + (tc ? 'medical' : 'technical & command') + ' sheet' });
    } else {
      status = 'awaiting';
    }
    if (tc) { t.command += tc.command; t.commandCount++; }
    if (med && needsMed) { t.medical += med.medical; t.medicalCount++; }

    [['tc', g.tc, 'Technical & command'], ['med', g.med, 'Medical']].forEach(function (x) {
      if (x[1].length > 1) {
        var vals = x[1].map(function (s) { return x[0] === 'tc' ? (s.technical + '/' + s.command + '/' + s.completed) : String(s.medical); });
        var differ = vals.some(function (v) { return v !== vals[0]; });
        checks.push({ type: differ ? 'conflict' : 'resent', slot: r,
          message: x[2] + ' sent ' + x[1].length + ' times' + (differ ? ' with different scores — latest counts' : ' (same scores)'),
          versions: x[1].map(function (s) { return { received: s.received, assessor: s.assessor, technical: s.technical, command: s.command, completed: s.completed, medical: s.medical }; }) });
      }
    });
    if ((tc && (tc.technical < 0 || tc.command < 0)) || (med && med.medical < 0)) {
      checks.push({ type: 'negative', slot: r, message: 'Negative score confirmed by assessor' });
    }

    scenarios.push({
      date: r.date, rotation: r.rotation, time: r.time, location: r.location, team: r.team,
      medicalScored: needsMed, status: status, points: points,
      tc: tc ? { technical: tc.technical, command: tc.command, completed: tc.completed, assessor: tc.assessor, teamLeader: tc.teamLeader, versions: g.tc.length } : null,
      med: med ? { medical: med.medical, assessor: med.assessor, versions: g.med.length } : null
    });
  });

  unscheduled.forEach(function (s) {
    checks.push({ type: 'unscheduled', slot: { date: s.date, rotation: s.rotation, location: s.location, team: s.team },
      message: 'Submission does not match the Schedule tab — not counted' });
  });

  var list = Object.keys(teamsMap).map(function (n) {
    var t = teamsMap[n];
    t.total += t.opening;
    return t;
  });

  function rank(arr, field) {
    arr.sort(function (a, b) { return b[field] - a[field] || a.team.localeCompare(b.team); });
    var prev = null, prevRank = 0;
    arr.forEach(function (x, i) {
      x.rank = (prev !== null && x[field] === prev) ? prevRank : i + 1;
      prev = x[field]; prevRank = x.rank;
    });
    return arr;
  }

  var standings = rank(list.slice(), 'total').map(function (t) {
    return { rank: t.rank, team: t.team, opening: t.opening, openingNote: t.openingNote,
      days: days.map(function (d) { return t.days[d] || 0; }), total: t.total,
      scored: t.scored, scheduled: t.scheduled, partial: t.partial };
  });

  var leader = rank(list.filter(function (t) { return t.commandCount; }).map(function (t) {
    return { team: t.team, points: t.command, count: t.commandCount, of: t.commandOf };
  }), 'points');
  var medic = rank(list.filter(function (t) { return t.medicalCount; }).map(function (t) {
    return { team: t.team, points: t.medical, count: t.medicalCount, of: t.medicalOf };
  }), 'points');

  var counts = { scored: 0, partial: 0, awaiting: 0 };
  scenarios.forEach(function (s) { counts[s.status]++; });

  return {
    event: data.settings['Event name'] || 'UKRO Rope Rescue Challenge',
    generated: new Date().toISOString(),
    days: days,
    standings: standings,
    awards: { teamLeader: leader, medic: medic,
      commandMax: RULES.commandMax, medicalMax: RULES.medicalMax },
    scenarios: scenarios,
    checks: checks,
    counts: counts,
    voided: voided,
    submissions: data.submissions.length,
    rules: { openingTarget: target, openingPerMinute: ppm, completionBonus: RULES.completionBonus }
  };
}

/* ==========================================================================
   2026 TEMPLATE DATA — used by setup() only. Edit the tabs, not this.
   ========================================================================== */
var SAMPLE_TEAMS = ['Avon', 'Devon & Somerset', 'Greater Manchester', 'Hereford & Worcester', 'Jersey',
  'LFB', 'Lancashire', 'Leicestershire', 'Nottinghamshire', 'Portugal (ANSD)', 'West Midlands', 'West Yorkshire'];

var SAMPLE_LOCATIONS = [
  ['County Hall', true],
  ['Crowne Plaza – Scenario 1', true],
  ['Crowne Plaza – Scenario 2', false],
  ['Meadow Lane – Notts County FC', false],
  ['National Justice Museum', true],
  ['Nottingham Trent University – City Campus', true],
  ['Theatre Royal', true],
  ['Wilford Suspension Bridge', false]
];

var SAMPLE_SCHEDULE = [
  ["2026-10-02",1,"08:00–09:10","County Hall","LFB"],
  ["2026-10-02",1,"08:00–09:10","County Hall","Nottinghamshire"],
  ["2026-10-02",1,"08:00–09:10","Crowne Plaza – Scenario 1","Avon"],
  ["2026-10-02",1,"08:00–09:10","Crowne Plaza – Scenario 1","Devon & Somerset"],
  ["2026-10-02",1,"08:00–09:10","Crowne Plaza – Scenario 2","Hereford & Worcester"],
  ["2026-10-02",1,"08:00–09:10","Crowne Plaza – Scenario 2","Jersey"],
  ["2026-10-02",1,"08:00–09:10","National Justice Museum","West Midlands"],
  ["2026-10-02",1,"08:00–09:10","National Justice Museum","West Yorkshire"],
  ["2026-10-02",1,"08:00–09:10","Theatre Royal","Lancashire"],
  ["2026-10-02",1,"08:00–09:10","Theatre Royal","Leicestershire"],
  ["2026-10-02",1,"08:00–09:10","Wilford Suspension Bridge","Greater Manchester"],
  ["2026-10-02",1,"08:00–09:10","Wilford Suspension Bridge","Portugal (ANSD)"],
  ["2026-10-02",2,"10:30–11:40","County Hall","Greater Manchester"],
  ["2026-10-02",2,"10:30–11:40","County Hall","Portugal (ANSD)"],
  ["2026-10-02",2,"10:30–11:40","Crowne Plaza – Scenario 1","Hereford & Worcester"],
  ["2026-10-02",2,"10:30–11:40","Crowne Plaza – Scenario 1","Jersey"],
  ["2026-10-02",2,"10:30–11:40","Crowne Plaza – Scenario 2","Avon"],
  ["2026-10-02",2,"10:30–11:40","Crowne Plaza – Scenario 2","Devon & Somerset"],
  ["2026-10-02",2,"10:30–11:40","Meadow Lane – Notts County FC","West Midlands"],
  ["2026-10-02",2,"10:30–11:40","Meadow Lane – Notts County FC","West Yorkshire"],
  ["2026-10-02",2,"10:30–11:40","Nottingham Trent University – City Campus","Lancashire"],
  ["2026-10-02",2,"10:30–11:40","Nottingham Trent University – City Campus","Leicestershire"],
  ["2026-10-02",2,"10:30–11:40","Wilford Suspension Bridge","LFB"],
  ["2026-10-02",2,"10:30–11:40","Wilford Suspension Bridge","Nottinghamshire"],
  ["2026-10-02",3,"13:00–14:10","Crowne Plaza – Scenario 1","Lancashire"],
  ["2026-10-02",3,"13:00–14:10","Crowne Plaza – Scenario 1","Leicestershire"],
  ["2026-10-02",3,"13:00–14:10","Meadow Lane – Notts County FC","Greater Manchester"],
  ["2026-10-02",3,"13:00–14:10","Meadow Lane – Notts County FC","LFB"],
  ["2026-10-02",3,"13:00–14:10","National Justice Museum","Nottinghamshire"],
  ["2026-10-02",3,"13:00–14:10","National Justice Museum","Portugal (ANSD)"],
  ["2026-10-02",3,"13:00–14:10","Nottingham Trent University – City Campus","Devon & Somerset"],
  ["2026-10-02",3,"13:00–14:10","Nottingham Trent University – City Campus","Hereford & Worcester"],
  ["2026-10-02",3,"13:00–14:10","Theatre Royal","Avon"],
  ["2026-10-02",3,"13:00–14:10","Theatre Royal","Jersey"],
  ["2026-10-02",3,"13:00–14:10","Wilford Suspension Bridge","West Midlands"],
  ["2026-10-02",3,"13:00–14:10","Wilford Suspension Bridge","West Yorkshire"],
  ["2026-10-02",4,"15:30–16:40","County Hall","West Midlands"],
  ["2026-10-02",4,"15:30–16:40","County Hall","West Yorkshire"],
  ["2026-10-02",4,"15:30–16:40","Crowne Plaza – Scenario 2","Lancashire"],
  ["2026-10-02",4,"15:30–16:40","Crowne Plaza – Scenario 2","Leicestershire"],
  ["2026-10-02",4,"15:30–16:40","Meadow Lane – Notts County FC","Nottinghamshire"],
  ["2026-10-02",4,"15:30–16:40","Meadow Lane – Notts County FC","Portugal (ANSD)"],
  ["2026-10-02",4,"15:30–16:40","National Justice Museum","Greater Manchester"],
  ["2026-10-02",4,"15:30–16:40","National Justice Museum","LFB"],
  ["2026-10-02",4,"15:30–16:40","Nottingham Trent University – City Campus","Avon"],
  ["2026-10-02",4,"15:30–16:40","Nottingham Trent University – City Campus","Jersey"],
  ["2026-10-02",4,"15:30–16:40","Theatre Royal","Devon & Somerset"],
  ["2026-10-02",4,"15:30–16:40","Theatre Royal","Hereford & Worcester"],
  ["2026-10-03",1,"08:00–09:10","County Hall","Devon & Somerset"],
  ["2026-10-03",1,"08:00–09:10","County Hall","Lancashire"],
  ["2026-10-03",1,"08:00–09:10","Crowne Plaza – Scenario 1","Nottinghamshire"],
  ["2026-10-03",1,"08:00–09:10","Crowne Plaza – Scenario 1","West Midlands"],
  ["2026-10-03",1,"08:00–09:10","Crowne Plaza – Scenario 2","Greater Manchester"],
  ["2026-10-03",1,"08:00–09:10","Crowne Plaza – Scenario 2","West Yorkshire"],
  ["2026-10-03",1,"08:00–09:10","National Justice Museum","Avon"],
  ["2026-10-03",1,"08:00–09:10","National Justice Museum","Hereford & Worcester"],
  ["2026-10-03",1,"08:00–09:10","Theatre Royal","LFB"],
  ["2026-10-03",1,"08:00–09:10","Theatre Royal","Portugal (ANSD)"],
  ["2026-10-03",1,"08:00–09:10","Wilford Suspension Bridge","Jersey"],
  ["2026-10-03",1,"08:00–09:10","Wilford Suspension Bridge","Leicestershire"],
  ["2026-10-03",2,"10:30–11:40","County Hall","Jersey"],
  ["2026-10-03",2,"10:30–11:40","County Hall","Leicestershire"],
  ["2026-10-03",2,"10:30–11:40","Crowne Plaza – Scenario 1","Greater Manchester"],
  ["2026-10-03",2,"10:30–11:40","Crowne Plaza – Scenario 1","West Yorkshire"],
  ["2026-10-03",2,"10:30–11:40","Crowne Plaza – Scenario 2","Nottinghamshire"],
  ["2026-10-03",2,"10:30–11:40","Crowne Plaza – Scenario 2","West Midlands"],
  ["2026-10-03",2,"10:30–11:40","Meadow Lane – Notts County FC","Avon"],
  ["2026-10-03",2,"10:30–11:40","Meadow Lane – Notts County FC","Hereford & Worcester"],
  ["2026-10-03",2,"10:30–11:40","Nottingham Trent University – City Campus","LFB"],
  ["2026-10-03",2,"10:30–11:40","Nottingham Trent University – City Campus","Portugal (ANSD)"],
  ["2026-10-03",2,"10:30–11:40","Wilford Suspension Bridge","Devon & Somerset"],
  ["2026-10-03",2,"10:30–11:40","Wilford Suspension Bridge","Lancashire"],
  ["2026-10-03",3,"13:00–14:10","Crowne Plaza – Scenario 1","LFB"],
  ["2026-10-03",3,"13:00–14:10","Crowne Plaza – Scenario 1","Portugal (ANSD)"],
  ["2026-10-03",3,"13:00–14:10","Meadow Lane – Notts County FC","Jersey"],
  ["2026-10-03",3,"13:00–14:10","Meadow Lane – Notts County FC","Lancashire"],
  ["2026-10-03",3,"13:00–14:10","National Justice Museum","Devon & Somerset"],
  ["2026-10-03",3,"13:00–14:10","National Justice Museum","Leicestershire"],
  ["2026-10-03",3,"13:00–14:10","Nottingham Trent University – City Campus","Greater Manchester"],
  ["2026-10-03",3,"13:00–14:10","Nottingham Trent University – City Campus","West Midlands"],
  ["2026-10-03",3,"13:00–14:10","Theatre Royal","Nottinghamshire"],
  ["2026-10-03",3,"13:00–14:10","Theatre Royal","West Yorkshire"],
  ["2026-10-03",3,"13:00–14:10","Wilford Suspension Bridge","Avon"],
  ["2026-10-03",3,"13:00–14:10","Wilford Suspension Bridge","Hereford & Worcester"],
  ["2026-10-03",4,"15:30–16:40","County Hall","Avon"],
  ["2026-10-03",4,"15:30–16:40","County Hall","Hereford & Worcester"],
  ["2026-10-03",4,"15:30–16:40","Crowne Plaza – Scenario 2","LFB"],
  ["2026-10-03",4,"15:30–16:40","Crowne Plaza – Scenario 2","Portugal (ANSD)"],
  ["2026-10-03",4,"15:30–16:40","Meadow Lane – Notts County FC","Devon & Somerset"],
  ["2026-10-03",4,"15:30–16:40","Meadow Lane – Notts County FC","Leicestershire"],
  ["2026-10-03",4,"15:30–16:40","National Justice Museum","Jersey"],
  ["2026-10-03",4,"15:30–16:40","National Justice Museum","Lancashire"],
  ["2026-10-03",4,"15:30–16:40","Nottingham Trent University – City Campus","Nottinghamshire"],
  ["2026-10-03",4,"15:30–16:40","Nottingham Trent University – City Campus","West Yorkshire"],
  ["2026-10-03",4,"15:30–16:40","Theatre Royal","Greater Manchester"],
  ["2026-10-03",4,"15:30–16:40","Theatre Royal","West Midlands"]
];

/* For testing outside Google (Node). Ignored by Apps Script. */
if (typeof module !== 'undefined') {
  module.exports = { RULES: RULES, CRITERIA: CRITERIA, SUB_HEADERS: SUB_HEADERS, parseData: parseData,
    scoreTC: scoreTC, scoreMedical: scoreMedical, acceptSubmission_: acceptSubmission_,
    buildStandings: buildStandings, roleFor_: roleFor_, publicConfig_: publicConfig_,
    SAMPLE_TEAMS: SAMPLE_TEAMS, SAMPLE_LOCATIONS: SAMPLE_LOCATIONS, SAMPLE_SCHEDULE: SAMPLE_SCHEDULE };
}
