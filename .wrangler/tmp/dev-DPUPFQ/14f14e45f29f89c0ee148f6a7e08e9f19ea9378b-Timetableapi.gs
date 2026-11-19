/**
 * ============================================================================
 * Timetableapi.gs — School Timetable (Academic Master = settings/editor,
 * Teacher & ClassTeacher = read-only after "Apply").
 * ----------------------------------------------------------------------------
 * Sheets (created automatically):
 *   "TT Settings"        key | json   (rows: draft, published, load:<class>)
 *   "TT Slots"           Day | Period | Class | Subject | SubjectCode | Teacher | Username | Label | Room   (DRAFT)
 *   "TT Published Slots" same columns                                                                     (what teachers see)
 *
 * Teachers see NOTHING until the Academic Master presses "Apply Timetable".
 *
 * Multi-Level Support: getClasses() inaangalia education_level ya shule
 * (Pre-Primary, Primary, O-Level, A-Level, Combined).
 * ============================================================================
 */
var TT_HEAD = ['Day', 'Period', 'Class', 'Subject', 'SubjectCode', 'Teacher', 'Username', 'Label', 'Room'];

function _ttDefaults_() {
  var school = null;
  try { school = rbacSchool_(); } catch (e) { school = null; }
  if (!school) {
    var st = {};
    try { st = psSettings_(); } catch (e) {}
    school = {
      council: String(st.SchoolCouncil || '').trim().toUpperCase(),
      nameUpper: String(st.SchoolName || '').trim().toUpperCase()
    };
  }
  return {
    councilName: school.council, schoolName: school.nameUpper,
    academicYear: String(new Date().getFullYear()), title: 'SCHOOL TEACHING TIMETABLE',
    breakTime: '10:40 - 11:00', lunchTime: '13:00 - 13:30', assemblyTime: '07:40 - 08:00',
    extraCurriculumTime: '15:00 - 16:00',
    periods: ['08:00 - 08:40', '08:40 - 09:20', '09:20 - 10:00', '10:00 - 10:40', '11:00 - 11:40',
              '11:40 - 12:20', '12:20 - 13:00', '13:30 - 14:10', '14:10 - 14:50'],
    extras: ['Sport & Game', 'Subject Clubs', 'Debate', 'Self Study', 'General Cleanliness'],
    notes: ''
  };
}

function _ttSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#0d1b36').setFontColor('#fff').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function _ttSlotSheet_(pub) { return _ttSheet_(pub ? 'TT Published Slots' : 'TT Slots', TT_HEAD); }
function _ttSetSheet_() { return _ttSheet_('TT Settings', ['Key', 'JSON']); }

function _ttGetSet_(key) {
  var sh = _ttSetSheet_();
  if (sh.getLastRow() < 2) return null;
  var d = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < d.length; i++) if (d[i][0] === key) { try { return JSON.parse(d[i][1]); } catch (e) { return null; } }
  return null;
}
function _ttPutSet_(key, obj) {
  var sh = _ttSetSheet_(), json = JSON.stringify(obj);
  var d = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
  for (var i = 0; i < d.length; i++) if (d[i][0] === key) { sh.getRange(i + 2, 2).setValue(json); return; }
  sh.appendRow([key, json]);
}

function _ttReadSlots_(pub) {
  var sh = _ttSlotSheet_(pub);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().filter(function (r) { return r[0] && r[1]; }).map(function (r) {
    return { day: +r[0], period: +r[1], cls: String(r[2]), subject: String(r[3] || ''), code: String(r[4] || ''),
             teacher: String(r[5] || ''), username: String(r[6] || ''), label: String(r[7] || ''), room: String(r[8] || '') };
  });
}

function _ttAssertAcademic_() {
  var s = getTeacherSession();
  if (!s) throw new Error('Session expired. Please log in again.');
  var pos = String(s.position || '').toLowerCase();
  var role = String(s.role || '').toLowerCase();
  if (!/academic/.test(pos) && !/academic/.test(role) && role !== 'headofschool' && role !== 'schoolmanager' && role !== 'secondmaster') {
    throw new Error('Only the Academic Master (or school leadership) can change the timetable.');
  }
  return s;
}

/** Respects the Admin "restricted / read-only" switch for the Academics dashboard */
function _ttBlocked_() {
  try {
    var chk = checkDashboardWriteAccess(['Timetable']);
    if (chk && chk.ok === false) return chk.message || 'Changes are currently disabled by the Admin.';
  } catch (e) {}
  return '';
}

function _ttTeachers_() {
  return getTeacherList().filter(function (t) {
    return ['accountant', 'librarian', 'lab technician'].indexOf(String(t.position || '').trim().toLowerCase()) < 0;
  }).map(function (t) {
    return { name: t.name, username: t.username, position: t.position || 'Teacher', subjects: t.subjects || [], classes: t.classes || [] };
  });
}

/* ────────────── Academic Master ────────────── */
function getTTAdminData() {
  _ttAssertAcademic_();
  var draft = _ttGetSet_('draft') || _ttDefaults_();
  var def = _ttDefaults_();
  Object.keys(def).forEach(function (k) { if (draft[k] == null) draft[k] = def[k]; });
  var pub = _ttGetSet_('published');
  return {
    settings: draft, slots: _ttReadSlots_(false), classes: getClasses(),
    subjects: getAcademicSubjectsRaw().map(function (s) { return { name: s.name, code: s.code || '' }; }),
    teachers: _ttTeachers_(),
    published: pub ? { appliedAt: pub.appliedAt, appliedBy: pub.appliedBy } : null,
    draftUpdatedAt: draft.updatedAt || ''
  };
}

function saveTTSettings(s) {
  var me = _ttAssertAcademic_();
  var blk = _ttBlocked_(); if (blk) return { success: false, message: blk };
  var def = _ttDefaults_(), out = {};
  Object.keys(def).forEach(function (k) { out[k] = (s && s[k] != null) ? s[k] : def[k]; });
  if (!Array.isArray(out.periods) || out.periods.length !== 9) out.periods = def.periods;
  if (!Array.isArray(out.extras) || out.extras.length !== 5) out.extras = def.extras;
  out.updatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  _ttPutSet_('draft', out);
  return { success: true, message: 'Timetable settings saved (draft).', updatedAt: out.updatedAt };
}

function saveTTSlot(p) {
  var me = _ttAssertAcademic_();
  var blk = _ttBlocked_(); if (blk) return { success: false, message: blk };
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var day = +p.day, per = +p.period, cls = String(p.cls || '');
    if (!(day >= 1 && day <= 5) || !(per >= 1 && per <= 9) || !cls) return { success: false, message: 'Invalid slot.' };
    var sh = _ttSlotSheet_(false), data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues() : [];
    var uname = String(p.username || '');
    if (uname) {                                   // teacher clash check
      for (var i = 0; i < data.length; i++) {
        if (+data[i][0] === day && +data[i][1] === per && String(data[i][6]) === uname && String(data[i][2]) !== cls) {
          return { success: false, clash: true, message: String(data[i][5]) + ' is already teaching ' + data[i][2] + ' at this time (Period ' + per + ').' };
        }
      }
    }
    var row = [day, per, cls, p.subject || '', p.code || '', p.teacher || '', uname, p.label || '', p.room || ''];
    var found = -1;
    for (var j = 0; j < data.length; j++) if (+data[j][0] === day && +data[j][1] === per && String(data[j][2]) === cls) { found = j; break; }
    if (!row[3] && !row[7] && !row[5]) {            // empty -> clear
      if (found >= 0) sh.deleteRow(found + 2);
    } else if (found >= 0) sh.getRange(found + 2, 1, 1, 9).setValues([row]);
    else sh.appendRow(row);
    _ttTouch_();
    return { success: true };
  } finally { lock.releaseLock(); }
}

function _ttTouch_() {
  var d = _ttGetSet_('draft') || _ttDefaults_();
  d.updatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  _ttPutSet_('draft', d);
}

function clearTTClass(cls) {
  _ttAssertAcademic_();
  var sh = _ttSlotSheet_(false);
  if (sh.getLastRow() < 2) return { success: true };
  var d = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  for (var i = d.length - 1; i >= 0; i--) if (String(d[i][2]) === String(cls)) sh.deleteRow(i + 2);
  _ttTouch_();
  return { success: true, message: 'Cleared draft for ' + cls };
}

/** APPLY: copy the draft to the published copy -> visible to Teacher / ClassTeacher accounts */
function applyTimetable() {
  var me = _ttAssertAcademic_();
  var blk = _ttBlocked_(); if (blk) return { success: false, message: blk };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var slots = _ttReadSlots_(false);
    if (!slots.length) return { success: false, message: 'The timetable is empty. Add some periods first.' };
    var pub = _ttSlotSheet_(true);
    if (pub.getLastRow() > 1) pub.getRange(2, 1, pub.getLastRow() - 1, 9).clearContent();
    var rows = slots.map(function (s) { return [s.day, s.period, s.cls, s.subject, s.code, s.teacher, s.username, s.label, s.room]; });
    pub.getRange(2, 1, rows.length, 9).setValues(rows);
    var draft = _ttGetSet_('draft') || _ttDefaults_();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    draft.appliedAt = now;
    _ttPutSet_('draft', draft);
    var snap = JSON.parse(JSON.stringify(draft)); snap.appliedAt = now; snap.appliedBy = me.fullName || me.username;
    _ttPutSet_('published', snap);
    try { logActivity(me.username, 'Timetable applied (' + rows.length + ' periods)'); } catch (e) {}
    return { success: true, appliedAt: now, count: rows.length, message: 'Timetable applied. Teachers and Class Teachers can now see it.' };
  } finally { lock.releaseLock(); }
}

/* ────────────── Teacher / ClassTeacher / Academic (read only) ────────────── */
function getPublishedTimetable() {
  var s = getTeacherSession();
  if (!s) throw new Error('Session expired. Please log in again.');
  var pub = _ttGetSet_('published');
  if (!pub) return { published: false };
  var me = null;
  _ttTeachers_().forEach(function (t) { if (t.username === s.username) me = t; });
  return {
    published: true, appliedAt: pub.appliedAt, appliedBy: pub.appliedBy, settings: pub,
    slots: _ttReadSlots_(true), classes: getClasses(),
    me: { name: (me && me.name) || s.fullName || s.username, username: s.username, classes: (me && me.classes) || [] }
  };
}

/* ═════════════════════════════════════════════════════════════════════════
 * SUBJECTS & PERIODS PER WEEK  +  AUTOMATIC TIMETABLE
 *   1) Academic sets, per class, how many periods each subject has per week ("TT Settings" key load:<class>)
 *   2) generateTimetable() fills the draft grid and puts the assigned teacher's name in each period
 *   3) Everything can still be edited by hand in the Editor, then "Apply Timetable"
 * ═════════════════════════════════════════════════════════════════════════ */
var TT_DAYS_ = 5, TT_PERIODS_ = 9;

function _ttNorm_(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ''); }

/** All subjects that belong to a class (from Manage Subjects > Assigned Classes). */
function _ttSubjectsForClass_(cls) {
  var codes = {};
  getAcademicSubjectsRaw().forEach(function (s) { codes[String(s.name).trim().toLowerCase()] = s.code || ''; });
  var seen = {}, out = [];
  getSubjectsForClassAndTeacher(cls, []).forEach(function (s) {
    var n = String(s.name || '').trim(); if (!n || seen[n.toLowerCase()]) return; seen[n.toLowerCase()] = 1;
    out.push({ name: n, code: codes[n.toLowerCase()] || '' });
  });
  return out;
}

/** Teachers assigned to teach a subject in a class (subjects/classes typed on the teacher's account). */
function _ttTeachersFor_(teachers, cls, subj) {
  var cn = _ttNorm_(cls), sn = _ttNorm_(subj.name), sc = _ttNorm_(subj.code);
  return teachers.filter(function (t) {
    var okS = (t.subjects || []).some(function (x) { var k = _ttNorm_(x); return k && (k === sn || (sc && k === sc)); });
    if (!okS) return false;
    return !(t.classes || []).length || (t.classes || []).some(function (c) { return _ttNorm_(c) === cn; });
  });
}

function getTTLoad(cls) {
  _ttAssertAcademic_();
  cls = String(cls || '');
  if (getClasses().indexOf(cls) < 0) throw new Error('Unknown class.');
  var saved = _ttGetSet_('load:' + cls) || {}, loads = saved.loads || {}, teachers = _ttTeachers_(), placed = {};
  _ttReadSlots_(false).forEach(function (x) { if (x.cls === cls && x.subject) placed[x.subject] = (placed[x.subject] || 0) + 1; });
  var subjects = _ttSubjectsForClass_(cls).map(function (s) {
    return { name: s.name, code: s.code, periods: loads.hasOwnProperty(s.name) ? +loads[s.name] : 0, placed: placed[s.name] || 0,
             teachers: _ttTeachersFor_(teachers, cls, s).map(function (t) { return { name: t.name, username: t.username }; }) };
  });
  return { cls: cls, subjects: subjects, maxPerDay: saved.maxPerDay || 2, capacity: TT_DAYS_ * TT_PERIODS_, saved: !!saved.loads };
}

function saveTTLoad(cls, loads, maxPerDay) {
  _ttAssertAcademic_();
  var blk = _ttBlocked_(); if (blk) return { success: false, message: blk };
  cls = String(cls || '');
  if (getClasses().indexOf(cls) < 0) return { success: false, message: 'Unknown class.' };
  var valid = {}; _ttSubjectsForClass_(cls).forEach(function (s) { valid[s.name] = 1; });
  var clean = {}, total = 0, cap = TT_DAYS_ * TT_PERIODS_;
  Object.keys(loads || {}).forEach(function (k) {
    if (!valid[k]) return;
    var n = parseInt(loads[k], 10); if (!(n >= 0)) n = 0; if (n > 10) n = 10;
    if (n > 0) { clean[k] = n; total += n; }
  });
  if (total > cap) return { success: false, message: 'The total (' + total + ' periods) is more than the ' + cap + ' periods available in a week.' };
  var mpd = parseInt(maxPerDay, 10); if (!(mpd >= 2 && mpd <= 4)) mpd = 2;   // a double period is 2 periods
  _ttPutSet_('load:' + cls, { loads: clean, maxPerDay: mpd, updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') });
  return { success: true, total: total, message: 'Periods saved for ' + cls + ' (' + total + ' / ' + cap + ').' };
}

/** Which neighbouring periods run on straight after each other (a double period cannot cross a break or lunch).
 *  Read from the period times in Timetable Settings, so it follows each school's own day. */
function _ttContig_() {
  var st = _ttGetSet_('draft') || _ttDefaults_(), per = (st.periods && st.periods.length === 9) ? st.periods : _ttDefaults_().periods;
  function parse(t) {
    var m = String(t || '').match(/(\d{1,2})[:.](\d{2})\s*[-\u2013\u2014]\s*(\d{1,2})[:.](\d{2})/);
    return m ? [+m[1] * 60 + +m[2], +m[3] * 60 + +m[4]] : null;
  }
  var c = {}, ok = true;
  for (var p = 1; p < TT_PERIODS_; p++) {
    var a = parse(per[p - 1]), b = parse(per[p]);
    if (!a || !b) { ok = false; break; }
    c[p] = (b[0] - a[1]) <= 10;                      // up to 10 minutes between periods still counts as continuous
  }
  if (!ok) c = { 1: true, 2: true, 3: true, 4: false, 5: true, 6: true, 7: false, 8: true };
  return c;
}

/**
 * Periods per week -> lessons:  2 = 1 double | 3 = 1 double + 1 single | 4 = 2 doubles | 5 = 2 doubles + 1 single ...
 * A double period = two periods in a row (same teacher), never across a break or lunch.
 *
 * target: a class name, or '__ALL__' (every class that has saved periods)
 * mode  : 'fill'    = only fill empty periods, keep what is already in the timetable (counts toward the totals)
 *         'rebuild' = clear the class first, then fill it again
 */
function generateTimetable(target, mode) {
  var me = _ttAssertAcademic_();
  var blk = _ttBlocked_(); if (blk) return { success: false, message: blk };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var allClasses = getClasses();
    var classes = (target === '__ALL__') ? allClasses.slice() : [String(target || '')];
    classes = classes.filter(function (c) { return allClasses.indexOf(c) >= 0; });
    if (!classes.length) return { success: false, message: 'Choose a class first.' };

    var cfgOf = {}, gen = [], skipped = [];
    classes.forEach(function (c) {
      var cfg = _ttGetSet_('load:' + c);
      if (cfg && cfg.loads && Object.keys(cfg.loads).length) { cfgOf[c] = cfg; gen.push(c); } else skipped.push(c);
    });
    if (!gen.length) return { success: false, message: 'Set the number of periods per week for the subjects first (nothing to fill).', skipped: skipped };

    var slots = _ttReadSlots_(false);
    if (mode === 'rebuild') slots = slots.filter(function (s) { return gen.indexOf(s.cls) < 0; });
    var teachers = _ttTeachers_(), contig = _ttContig_();
    var grid = {}, busy = {}, tload = {};
    function gk(c, d, p) { return c + '|' + d + '|' + p; }
    function bk(u, d, p) { return u + '|' + d + '|' + p; }
    function reg(s) {
      grid[gk(s.cls, s.day, s.period)] = s;
      if (s.username) { busy[bk(s.username, s.day, s.period)] = s.cls; tload[s.username] = (tload[s.username] || 0) + 1; }
    }
    function blockStart(p) { var b = p; while (b > 1 && contig[b - 1]) b--; return b; }
    slots.forEach(reg);

    var reports = [], placedTotal = 0;
    gen.forEach(function (cls) {
      var cfg = cfgOf[cls], maxPD = Math.max(2, +cfg.maxPerDay || 2), byName = {};
      _ttSubjectsForClass_(cls).forEach(function (s) { byName[s.name] = s; });
      var have = {}, perDay = {}, usedP = {}, dayFill = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
      slots.forEach(function (s) {
        if (s.cls !== cls) return;
        dayFill[s.day] = (dayFill[s.day] || 0) + 1;
        if (s.subject) { have[s.subject] = (have[s.subject] || 0) + 1; perDay[s.subject + '|' + s.day] = (perDay[s.subject + '|' + s.day] || 0) + 1; usedP[s.subject + '|' + s.period] = 1; }
      });
      var tasks = [];
      Object.keys(cfg.loads).forEach(function (name) {
        var sub = byName[name]; if (!sub) return;                       // subject no longer belongs to this class
        var need = (+cfg.loads[name] || 0) - (have[name] || 0);
        if (need > 0) tasks.push({ sub: sub, need: need, doubles: Math.floor(need / 2), singles: need % 2, split: 0,
                                   cands: _ttTeachersFor_(teachers, cls, sub) });
      });
      tasks.sort(function (a, b) { return (b.need - a.need) || (a.cands.length - b.cands.length); });
      var rep = { cls: cls, placed: 0, unplaced: [], noTeacher: [], split: [] };

      function freeTeacher(t, d, ps) {                                  // one teacher who is free for ALL the periods
        if (!t.cands.length) return { ok: true, tch: null };
        var free = t.cands.filter(function (c) { return ps.every(function (p) { return !busy[bk(c.username, d, p)]; }); });
        if (!free.length) return { ok: false };
        free.sort(function (a, b) { return (tload[a.username] || 0) - (tload[b.username] || 0); });
        return { ok: true, tch: free[0] };
      }
      function put(t, d, ps, tch) {
        ps.forEach(function (p) {
          var row = { day: d, period: p, cls: cls, subject: t.sub.name, code: t.sub.code || '',
                      teacher: tch ? tch.name : '', username: tch ? tch.username : '', label: '', room: '' };
          slots.push(row); reg(row);
          usedP[t.sub.name + '|' + p] = 1; dayFill[d] = (dayFill[d] || 0) + 1;
          rep.placed++; placedTotal++;
        });
        perDay[t.sub.name + '|' + d] = (perDay[t.sub.name + '|' + d] || 0) + ps.length;
      }
      function adj(t, d, p) { var g = grid[gk(cls, d, p)]; return g && g.subject === t.sub.name; }

      // 1) double periods first (they are the hardest to place)
      tasks.forEach(function (t) {
        for (var k = 0; k < t.doubles; k++) {
          var best = null;
          for (var d = 1; d <= TT_DAYS_; d++) for (var p = 1; p < TT_PERIODS_; p++) {
            if (!contig[p] || grid[gk(cls, d, p)] || grid[gk(cls, d, p + 1)]) continue;
            var ft = freeTeacher(t, d, [p, p + 1]); if (!ft.ok) continue;
            var sd = perDay[t.sub.name + '|' + d] || 0;
            var sc = sd * 60 + ((sd + 2) > maxPD ? 10000 : 0) + (dayFill[d] || 0) * 2 + d * 0.1 + p * 0.01;
            if ((p - blockStart(p)) % 2 !== 0) sc += 15;               // 1-2, 3-4, 5-6, 8-9 are preferred to 2-3 or 6-7
            if (adj(t, d, p - 1) || adj(t, d, p + 2)) sc += 25;
            if (usedP[t.sub.name + '|' + p]) sc += 6;
            if (!best || sc < best.sc) best = { d: d, p: p, sc: sc, tch: ft.tch };
          }
          if (best) put(t, best.d, [best.p, best.p + 1], best.tch);
          else t.split += 2;                                            // no free pair: fall back to two single periods
        }
      });
      // 2) single periods (the odd one out, and any double that could not be placed as a pair)
      tasks.forEach(function (t) {
        var n = t.singles + t.split, missing = 0, splitDone = 0;
        for (var k = 0; k < n; k++) {
          var best = null;
          for (var d = 1; d <= TT_DAYS_; d++) for (var p = 1; p <= TT_PERIODS_; p++) {
            if (grid[gk(cls, d, p)]) continue;
            var ft = freeTeacher(t, d, [p]); if (!ft.ok) continue;
            var sd = perDay[t.sub.name + '|' + d] || 0;
            var sc = sd * 60 + ((sd + 1) > maxPD ? 10000 : 0) + (dayFill[d] || 0) * 2 + d * 0.1 + p * 0.01;
            if (adj(t, d, p - 1) || adj(t, d, p + 1)) sc += 25;
            if (usedP[t.sub.name + '|' + p]) sc += 6;
            if (!best || sc < best.sc) best = { d: d, p: p, sc: sc, tch: ft.tch };
          }
          if (!best) { missing++; continue; }
          put(t, best.d, [best.p], best.tch);
          if (k >= t.singles) splitDone++;
        }
        if (!t.cands.length) rep.noTeacher.push(t.sub.name);
        if (t.split && splitDone) rep.split.push({ subject: t.sub.name, periods: splitDone });
        if (missing) rep.unplaced.push({ subject: t.sub.name, missing: missing, reason: t.cands.length ? 'no free period (class full or the teacher is busy)' : 'no free period (class is full)' });
      });
      reports.push(rep);
    });

    var sh = _ttSlotSheet_(false);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 9).clearContent();
    var out = slots.map(function (s) { return [s.day, s.period, s.cls, s.subject, s.code, s.teacher, s.username, s.label, s.room]; });
    if (out.length) sh.getRange(2, 1, out.length, 9).setValues(out);
    _ttTouch_();
    try { logActivity(me.username, 'Timetable auto-filled: ' + placedTotal + ' periods (' + gen.join(', ') + ')'); } catch (e) {}
    return { success: true, placed: placedTotal, classes: reports, skipped: skipped,
             message: 'Timetable filled: ' + placedTotal + ' period(s) placed. Review it in the Editor, then press Apply Timetable.' };
  } finally { lock.releaseLock(); }
}