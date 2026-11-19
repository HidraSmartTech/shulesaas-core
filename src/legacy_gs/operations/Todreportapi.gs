/**
 * ============================================================================
 * Todreportapi.gs — "Teacher's Duty Report" (TOD) kwa ajili ya SaaS Multi-Tenant
 * ============================================================================
 * Inasaidia Pre-Primary, Primary, O-Level na A-Level:
 *   - PS_CLASSES inabadilishwa kuwa getClasses() inayotegemea education_level
 *   - Class register sheet lookup inatumia class name badala ya index iliyowekwa
 *
 * TEACHER SIDE (Academics.html, ClassTeacher.html, Teacher.html — teacher session):
 *     getTODFormData, saveTODDutyReport, listMyTODDutyReports,
 *     deleteTODDutyReport, getTODSchoolInfo
 *
 * HEAD OF SCHOOL SIDE (HeadOfSchool.html — token session through psCall):
 *     api_todTeachers_, api_todRosterList_, api_todRosterSave_,
 *     api_todRosterDelete_, api_todReports_, api_todReport_
 *
 * Sheets (created automatically): "TOD Duty Reports" and "TOD Roster".
 * The old "TOD Reports" sheet (legacy format) is still READ so old reports are not lost.
 * ============================================================================
 */
var TOD_DUTY_SHEET   = 'TOD Duty Reports';
var TOD_ROSTER_SHEET = 'TOD Roster';
// positions that can be put on the duty roster (dashboards that have the TOD form)
var TOD_TEACHING_POSITIONS = ['Academic Master', 'Class Teacher', 'Subject Teacher', 'Depute Headmaster', 'Second Master'];

/* ──────────────────────── helpers ──────────────────────── */
function _todYmd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').substring(0, 10);
}

function _todSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TOD_DUTY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TOD_DUTY_SHEET);
    var h = ['ID', 'Date', 'Teacher', 'Username', 'Submitted At', 'Answers (JSON)', 'Attendance (JSON)', '% Present', 'Total Registered', 'Total Absent'];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setBackground('#0d1b36').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');   // keep yyyy-MM-dd as text
  }
  return sh;
}

function _todRosterSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TOD_ROSTER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TOD_ROSTER_SHEET);
    var headers = ['ID', 'Teacher', 'Username', 'Start Date', 'End Date', 'Notes', 'Created By', 'Created At'];
    var r = sh.getRange(1, 1, 1, headers.length);
    r.setValues([headers]);
    r.setBackground('#0d1b36').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('D:E').setNumberFormat('@');
    sh.setColumnWidth(2, 200); sh.setColumnWidth(6, 260);
  }
  return sh;
}

/** School name / council / motto / logo from School Settings + Head of School's name. */
function _todSchoolInfo_(withLogo) {
  var st = {};
  try { st = psSettings_(); } catch (e) {}
  var name = String(st.SchoolName || '').trim();
  if (name === PS_PLACEHOLDER_NAME) name = '';
  var head = '';
  try {
    var a = psTable_('Administrators').rows().filter(function (r) {
      return psRoleKey_(r.Role) === 'headofschool' && String(r.Status || '') !== 'Inactive';
    })[0];
    if (a) head = String(a.FullName || a.Username || '');
  } catch (e) {}
  return {
    councilName: String(st.SchoolCouncil || '').trim().toUpperCase(),
    schoolName:  name.toUpperCase(),
    motto:       String(st.SchoolMotto || '').trim(),
    headName:    head,
    logoUrl:     withLogo ? String(st.SchoolLogo || '') : ''
  };
}

/** Orodha ya madarasa: itumie getClasses() kama ipo (multi-level), vinginevyo PS_CLASSES. */
function _todClassNames_() {
  try { var cs = getClasses(); if (cs && cs.length) return cs; } catch (e) {}
  try { if (typeof PS_CLASSES !== 'undefined' && PS_CLASSES && PS_CLASSES.length) return PS_CLASSES; } catch (e) {}
  return [];
}

/** Jina la sheet ya daftari la darasa (class register). */
function _todClassSheet_(cls) {
  try { return psClassSheetName_(cls); } catch (e) { return String(cls); }
}

/**
 * Registered boys / girls of every class, counted from the REAL class registers.
 * Kept in the cache for 60 s (and dropped on any psCall write).
 */
function _todRoster_() {
  var cache = null, ck = 'TOD_ROSTER_' + (function () { try { return psCacheVer_(); } catch (e) { return 'v1'; } })();
  try { cache = CacheService.getScriptCache(); var hit = cache.get(ck); if (hit) return JSON.parse(hit); } catch (e) {}
  var ss = SpreadsheetApp.getActiveSpreadsheet(), out = [];
  _todClassNames_().forEach(function (c, i) {
    var rb = 0, rg = 0, sh = ss.getSheetByName(_todClassSheet_(c));
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r) {
        if (!String(r[0]).trim() && !String(r[1]).trim()) return;          // empty line
        if (String(r[3] || '').trim().toUpperCase().charAt(0) === 'F') rg++; else rb++;
      });
    }
    out.push({ classId: i + 1, className: c, rb: rb, rg: rg, ab: 0, ag: 0, sb: 0, sg: 0, pb: 0, pg: 0 });
  });
  try { if (cache) cache.put(ck, JSON.stringify(out), 60); } catch (e) {}
  return out;
}

function _todRowToObj_(r) {
  return {
    id: String(r[0]), date: _todYmd_(r[1]), teacherName: String(r[2] || ''), username: String(r[3] || ''),
    submittedAt: r[4] instanceof Date ? Utilities.formatDate(r[4], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(r[4] || ''),
    answers: String(r[5] || '{}'), attendanceRows: String(r[6] || '[]'),
    percentage: String(r[7] || ''), totalRegistered: String(r[8] || ''), totalAbsent: String(r[9] || '')
  };
}

function _todAll_() {
  var sh = _todSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues().filter(function (r) { return r[0]; }).map(_todRowToObj_);
}

function _todMe_() {
  var s = getTeacherSession();
  if (!s) throw new Error('Session expired. Please log in again.');
  return { name: s.fullName || s.username, username: s.username, position: s.position || '' };
}

function _todMyList_(all, username) {
  return all.filter(function (r) { return r.username === username; })
    .sort(function (a, b) { return a.date < b.date ? 1 : -1; }).slice(0, 30);
}

function _todCleanRows_(rows) {
  var n = function (v) { v = parseInt(v, 10); return isNaN(v) || v < 0 ? 0 : Math.min(v, 9999); };
  return rows.map(function (r) {
    var o = { classId: r.classId, className: String(r.className || r.cls || ''),
              rb: n(r.rb), rg: n(r.rg), ab: n(r.ab), ag: n(r.ag), sb: n(r.sb), sg: n(r.sg), pb: n(r.pb), pg: n(r.pg) };
    o.ab = Math.min(o.ab, o.rb); o.ag = Math.min(o.ag, o.rg);       // absent can never exceed registered
    return o;
  });
}

/* ═════════════════════════════════════════════════════════════════════════
   TEACHER SIDE — called from Academics / ClassTeacher / Teacher dashboards
   ═════════════════════════════════════════════════════════════════════════ */

/** Everything the form needs in ONE call: real roster, my report of that date, my last reports. */
function getTODFormData(dateStr) {
  var me = _todMe_();
  dateStr = _todYmd_(dateStr) || _todYmd_(new Date());
  var all = _todAll_(), mine = null;
  all.forEach(function (r) { if (r.date === dateStr && r.username === me.username) mine = r; });
  return { roster: _todRoster_(), mine: mine, teacherName: me.name, date: dateStr, list: _todMyList_(all, me.username) };
}

/** School details for printing (includes the logo). Called once, when a report is printed. */
function getTODSchoolInfo() {
  _todMe_();
  return _todSchoolInfo_(true);
}

function saveTODDutyReport(p) {
  var me = _todMe_();
  try {
    var chk = checkDashboardWriteAccess(['TOD Reports', 'Reports']);
    if (chk && chk.ok === false) return { success: false, message: chk.message };
  } catch (e) {}

  var date = _todYmd_(p && p.date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { success: false, message: 'Select a valid date.' };
  var rows;
  try { rows = JSON.parse(p.attendanceRows || '[]'); } catch (e) { return { success: false, message: 'Invalid attendance data.' }; }
  if (!rows || !rows.length) return { success: false, message: 'Attendance table is empty.' };
  rows = _todCleanRows_(rows);

  var reg = 0, abs = 0;
  rows.forEach(function (r) { reg += r.rb + r.rg; abs += r.ab + r.ag; });
  var pct = reg > 0 ? Math.round(1000 * (reg - abs) / reg) / 10 : 0;

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { success: false, message: 'The system is busy. Please try again in a moment.' }; }
  try {
    var sh = _todSheet_();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    var data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() : [];
    var found = -1;
    for (var i = 0; i < data.length; i++) {
      if (_todYmd_(data[i][1]) === date && String(data[i][3]) === me.username) { found = i; break; }
    }
    var rowVals = [found >= 0 ? data[found][0] : Utilities.getUuid().substring(0, 8), date, me.name, me.username, now,
                   String(p.answers || '{}'), JSON.stringify(rows), pct, reg, abs];
    if (found >= 0) {
      sh.getRange(found + 2, 1, 1, 10).setValues([rowVals]);
    } else {
      var n = sh.getLastRow() + 1;
      sh.getRange(n, 2).setNumberFormat('@');
      sh.getRange(n, 1, 1, 10).setValues([rowVals]);
    }
    try { logActivity(me.username, 'TOD duty report ' + date); } catch (e) {}
    return { success: true, updated: found >= 0, report: _todRowToObj_(rowVals), list: _todMyList_(_todAll_(), me.username),
             message: found >= 0 ? 'Duty report updated and sent to the Head of School.' : 'Duty report saved and sent to the Head of School.' };
  } finally {
    lock.releaseLock();
  }
}

function listMyTODDutyReports() {
  var me = _todMe_();
  return _todMyList_(_todAll_(), me.username);
}

function deleteTODDutyReport(id) {
  var me = _todMe_();
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { success: false, message: 'The system is busy. Please try again in a moment.' }; }
  try {
    var sh = _todSheet_();
    if (sh.getLastRow() < 2) return { success: false, message: 'Not found.' };
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        if (String(data[i][3]) !== me.username) return { success: false, message: 'You can delete only your own report.' };
        sh.deleteRow(i + 2);
        return { success: true, message: 'Duty report deleted.', list: _todMyList_(_todAll_(), me.username) };
      }
    }
    return { success: false, message: 'Report not found.' };
  } finally {
    lock.releaseLock();
  }
}

/* ═════════════════════════════════════════════════════════════════════════
   HEAD OF SCHOOL SIDE — psCall handlers
   Signature: (session, ...args)
   ═════════════════════════════════════════════════════════════════════════ */

/** Teachers that can be placed on the duty roster. */
function api_todTeachers_(s) {
  return psTable_('Users').rows()
    .filter(function (u) {
      return TOD_TEACHING_POSITIONS.indexOf(String(u.Position || '').trim()) >= 0 && String(u.Status || 'Active') !== 'Inactive';
    })
    .map(function (u) { return { name: u.Name, username: u.Username, position: u.Position }; });
}

function api_todRosterList_(s) {
  var sh = _todRosterSheet_();
  if (sh.getLastRow() < 2) return [];
  var out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().forEach(function (r) {
    if (!r[0]) return;
    out.push({ id: String(r[0]), teacher: String(r[1] || ''), username: String(r[2] || ''),
               start: _todYmd_(r[3]), end: _todYmd_(r[4]), notes: String(r[5] || '') });
  });
  out.sort(function (a, b) { return a.start < b.start ? -1 : (a.start > b.start ? 1 : 0); });
  return out;
}

/** entries: [{teacher, username, start:'yyyy-MM-dd', end:'yyyy-MM-dd', notes}] */
function api_todRosterSave_(s, entries) {
  if (!entries || !entries.length) return { success: false, message: 'Nothing to save.' };
  var sh = _todRosterSheet_();
  var existing = {};
  api_todRosterList_(s).forEach(function (e) { existing[(e.teacher + '|' + e.start + '|' + e.end).toLowerCase()] = true; });

  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var rows = [], skipped = 0;
  entries.forEach(function (e) {
    var t = String(e.teacher || '').trim(), st = String(e.start || '').substring(0, 10), en = String(e.end || '').substring(0, 10);
    if (!t || !/^\d{4}-\d{2}-\d{2}$/.test(st) || !/^\d{4}-\d{2}-\d{2}$/.test(en) || en < st) { skipped++; return; }
    var key = (t + '|' + st + '|' + en).toLowerCase();
    if (existing[key]) { skipped++; return; }
    existing[key] = true;
    rows.push([Utilities.getUuid().substring(0, 8), t, String(e.username || ''), st, en, String(e.notes || ''), s.name || s.username, now]);
  });
  if (!rows.length) return { success: false, message: 'No new duties were added (duplicates or invalid dates).' };

  var start = sh.getLastRow() + 1;
  sh.getRange(start, 4, rows.length, 2).setNumberFormat('@');
  sh.getRange(start, 1, rows.length, 8).setValues(rows);
  try { logActivity(s.username, 'TOD roster: added ' + rows.length + ' duty slot(s)'); } catch (e) {}
  return { success: true, added: rows.length, skipped: skipped,
           message: rows.length + ' duty slot(s) saved' + (skipped ? ' (' + skipped + ' skipped)' : '') + '.' };
}

function api_todRosterDelete_(s, id) {
  var sh = _todRosterSheet_();
  if (sh.getLastRow() < 2) return { success: false, message: 'Roster is empty.' };
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sh.deleteRow(i + 2);
      try { logActivity(s.username, 'TOD roster: removed duty slot ' + id); } catch (e) {}
      return { success: true, message: 'Duty removed.' };
    }
  }
  return { success: false, message: 'Duty not found.' };
}

/** Every submitted report: new format ("TOD Duty Reports") + old format ("TOD Reports", read only). */
function api_todReports_(s) {
  var out = [], seen = {};
  _todAll_().forEach(function (r) {
    seen[(r.teacherName + '|' + r.date).toLowerCase()] = true;
    out.push({ id: r.id, format: 'v2', date: r.date, teacher: r.teacherName, submittedAt: r.submittedAt,
               totalPresent: r.percentage + '%', totalAbsentees: r.totalAbsent, pdfLink: '' });
  });
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TOD Reports');
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 25).getValues().forEach(function (row) {
      var d = _todYmd_(row[1]);
      if (!d) return;
      var t = String(row[2] || '');
      if (seen[(t + '|' + d).toLowerCase()]) return;
      out.push({
        id: '', format: 'legacy', date: d, teacher: t,
        submittedAt: row[3] instanceof Date
          ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(row[3] || ''),
        totalPresent: String(row[5] || ''), totalAbsentees: String(row[6] || ''), pdfLink: String(row[24] || '')
      });
    });
  }
  out.sort(function (a, b) {
    if (a.date !== b.date) return a.date < b.date ? 1 : -1;
    return a.teacher < b.teacher ? -1 : 1;
  });
  return out;
}

function api_todReport_(s, teacher, dateStr, format) {
  if (format === 'v2') {
    var found = null;
    _todAll_().forEach(function (r) { if (r.date === dateStr && r.teacherName === teacher) found = r; });
    if (!found) return null;
    return { format: 'v2', date: found.date, teacherName: found.teacherName, answers: found.answers,
             attendanceRows: found.attendanceRows, school: _todSchoolInfo_(true) };
  }
  var r = null;
  try { r = getTODReportByDate(teacher, dateStr); } catch (e) {}       // old format (Code.gs)
  if (!r) return null;
  var out = { format: 'legacy' };
  Object.keys(r).forEach(function (k) {
    var v = r[k];
    out[k] = v instanceof Date
      ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(v == null ? '' : v);
  });
  return out;
}