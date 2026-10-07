/* ==========================================================================
   UKRO practice mode
   --------------------------------------------------------------------------
   Lets the scoresheets and standings be tried without the Google Sheet.
   Uses the real scoring code from ukro-backend.js, but keeps everything in
   this browser's storage on this device. Nothing is sent anywhere.
   ========================================================================== */
(function () {
  'use strict';
  var KEY = 'ukro-demo-state';

  function blank() {
    return {
      settings: [['Event name', 'UKRO Rope Rescue Challenge — PRACTICE'], ['Assessor PIN', 'demo'], ['Organiser PIN', 'demo'],
                 ['Opening day target minutes', '90'], ['Opening day points per minute', '3']],
      teams: SAMPLE_TEAMS.map(function (t) { return [t]; }),
      locations: SAMPLE_LOCATIONS.map(function (l) { return [l[0], l[1] ? 'Yes' : 'No']; }),
      schedule: SAMPLE_SCHEDULE.map(function (r) { return [r[0], String(r[1]), r[2], r[3], r[4]]; }),
      opening: SAMPLE_TEAMS.map(function (t) { return [t, '', '']; }),
      submissions: []
    };
  }
  function load() { try { return JSON.parse(localStorage.getItem(KEY)) || blank(); } catch (e) { return blank(); } }
  function save(st) { try { localStorage.setItem(KEY, JSON.stringify(st)); } catch (e) {} }
  function data(st) { return parseData(st.settings, st.teams, st.locations, st.schedule, st.opening, st.submissions); }

  function handle(req) {
    var st = load(), d = data(st);
    var role = roleFor_(req.pin, d.settings);
    if (!role) return { ok: false, error: 'bad_pin', message: 'PIN not recognised.' };
    if (req.action === 'config') return { ok: true, role: role, config: publicConfig_(d) };
    if (req.action === 'submit') {
      var res = acceptSubmission_(req.submission, d, new Date());
      if (res.ok && !res.duplicate) {
        st.submissions.push(res.row.map(function (v) {
          if (v instanceof Date) return v.toISOString();
          return typeof v === 'string' && v.charAt(0) === "'" ? v.slice(1) : v;
        }));
        save(st);
      }
      delete res.row; return res;
    }
    if (req.action === 'standings') return { ok: true, result: buildStandings(d) };
    return { ok: false, error: 'unknown_action' };
  }

  /* ---- 2026 results as a sample (aligned with SAMPLE_SCHEDULE) ----
     [technical, command, completed 1/0, medical or null]. Rebuilt from the
     final paper totals, so deductions are spread across criteria to match. */
  var RESULTS_2026 = [[75,48,1,48],[92,44,1,48],[71,48,1,28],[92,46,1,28],[94,42,1,null],[63,34,1,null],[84,40,0,34],[48,14,0,38],[94,44,1,40],[86,44,1,28],[90,44,0,null],[73,48,1,null],[84,48,1,50],[15,19,1,44],[92,46,1,44],[44,46,1,40],[92,46,1,null],[100,50,1,null],[66,11,0,null],[74,26,0,null],[68,46,0,24],[92,46,0,27],[90,50,1,null],[98,44,0,null],[92,48,1,36],[92,46,1,44],[66,18,0,null],[60,30,0,null],[73,46,1,48],[86,48,1,48],[94,50,0,24],[88,38,0,18],[84,42,1,46],[84,42,1,30],[92,42,1,null],[78,42,0,null],[73,19,1,48],[65,30,0,44],[73,44,1,null],[92,48,1,null],[80,40,0,null],[49,13,0,null],[61,3,0,38],[98,48,0,44],[82,46,1,38],[90,42,0,24],[94,48,1,36],[92,44,1,28],[98,46,1,48],[65,19,1,48],[63,40,0,48],[76,46,0,26],[66,34,1,null],[61,34,1,null],[92,48,1,36],[63,21,1,38],[90,46,1,46],[94,38,1,48],[82,36,0,null],[96,48,1,null],[98,48,1,48],[98,48,1,48],[80,40,0,30],[96,49,1,22],[63,34,1,null],[65,40,1,null],[62,36,0,null],[72,31,0,null],[88,50,1,32],[51,36,1,30],[96,40,1,null],[96,48,1,null],[96,48,1,46],[90,50,1,46],[72,31,0,null],[86,48,0,null],[98,48,1,44],[94,46,1,44],[92,48,0,24],[82,38,0,24],[44,42,0,30],[88,42,0,30],[94,44,1,null],[98,48,1,null],[94,48,1,48],[72,44,1,38],[94,50,1,null],[61,30,1,null],[88,48,0,null],[92,40,0,null],[86,40,1,34],[90,40,1,46],[67,42,0,38],[99,38,0,28],[92,42,0,30],[69,44,1,20]];
  var OPENING_2026 = [["Avon",43,"Finished"],["Devon & Somerset",43,"Finished"],["Greater Manchester",64,"Finished"],["Hereford & Worcester",51,"Finished"],["Jersey",76,"Finished"],["LFB",29,"Finished"],["Lancashire",44,"Finished"],["Leicestershire",41,"Finished"],["Nottinghamshire",47,"Finished"],["Portugal (ANSD)",49,"Finished"],["West Midlands",46,"Finished"],["West Yorkshire",null,"DNF"]];

  function spreadTC(section, lost) {
    lost = Math.max(0, lost);
    var sc = 0;
    while ((lost - 25 * sc) % 2 !== 0 && 25 * (sc + 1) <= lost) sc++;
    if ((lost - 25 * sc) % 2 !== 0) lost -= 1;           // a few 2026 paper totals aren't reachable
    while (lost - 25 * sc >= 50 && sc < 3) sc += 2;      // keep it realistic: some -25s for big losses
    var nsc = (lost - 25 * sc) / 2, keys = CRITERIA[section].map(function (c) { return c[0]; }), out = {}, i = 0;
    for (var j = 0; j < sc; j++) { var k = keys[(j * 3 + 1) % keys.length]; out[k] = out[k] || {}; out[k].sc = (out[k].sc || 0) + 1; }
    while (nsc > 0) { var k2 = keys[i++ % keys.length]; out[k2] = out[k2] || {}; out[k2].nsc = (out[k2].nsc || 0) + 1; nsc--; }
    return out;
  }
  function spreadMed(lost) {
    lost = Math.max(0, lost); if (lost % 2) lost -= 1;
    var keys = CRITERIA.medical.map(function (c) { return c[0]; });
    for (var c = 0; c <= 12; c++) {
      var rest = lost - 10 * c;
      if (rest >= 0 && c + rest / 2 <= 12) {
        var o = {}, i = 0;
        for (; i < c; i++) o[keys[i]] = { level: 'critical' };
        for (var j = 0; j < rest / 2; j++, i++) o[keys[i]] = { level: 'minor' };
        return o;
      }
    }
    return {};
  }
  var NAMES = ['Sample Assessor A', 'Sample Assessor B', 'Sample Assessor C', 'Sample Assessor D'];

  function loadSample() {
    var st = blank();
    OPENING_2026.forEach(function (o) {
      st.opening.forEach(function (r) { if (r[0] === o[0]) { r[1] = o[1] == null ? '' : String(o[1]); r[2] = o[2]; } });
    });
    save(st);
    SAMPLE_SCHEDULE.forEach(function (s, i) {
      var r = RESULTS_2026[i], base = { date: s[0], rotation: String(s[1]), location: s[3], team: s[4], assessor: NAMES[i % 4] };
      handle({ action: 'submit', pin: 'demo', submission: Object.assign({}, base, { id: 'sample-tc-' + i, sheet: 'tc', completed: r[2] ? 'Yes' : 'No',
        teamLeader: 'Sample team leader', confirmNegative: true,
        deductions: { technical: spreadTC('technical', 100 - r[0]), command: spreadTC('command', 50 - r[1]) } }) });
      if (r[3] !== null) handle({ action: 'submit', pin: 'demo', submission: Object.assign({}, base, { id: 'sample-med-' + i, sheet: 'medical',
        deductions: { medical: spreadMed(50 - r[3]) } }) });
    });
  }

  window.UKRODemo = {
    call: function (body) {
      return new Promise(function (resolve) { setTimeout(function () { resolve(handle(body)); }, 150); });
    },
    reset: function () { save(blank()); },
    loadSample: loadSample,
    count: function () { return load().submissions.length; }
  };
})();
