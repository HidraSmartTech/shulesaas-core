/**
 * SecondMasterAPI.gs — backend for SecondMaster.html
 * =====================================================================
 * Add this file to the Apps Script project (next to Code.gs).
 *
 * Provides:
 *   - TOD roster ("TOD Roster" sheet, created automatically)
 *   - read-only access to ALL "TOD Reports" rows for the Second Master
 *
 * Everything else the Second Master dashboard uses (assignments, scores,
 * score tracking, submission report, attendance tracking) already exists
 * in Code.gs and is re-used unchanged.
 * =====================================================================
 */

var HOS_ROSTER_SHEET = 'TOD Roster';
var HOS_NON_TEACHING = ['accountant', 'librarian', 'lab technician'];

/* ── guard: only a logged-in Second Master may call these ─────────── */
function _hosAssert_() {
  var s = getTeacherSession();
  if (!s) throw new Error('Session expired. Please log in again.');
  var pos = String(s.position || '').toLowerCase();
  if (!/second\s*master/.test(pos)) {
    throw new Error('Only the Second Master can use this feature.');
  }
  return s;
}

function _hosYmd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').substring(0, 10);
}

function _hosRosterSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOS_ROSTER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(HOS_ROSTER_SHEET);
    var headers = ['ID', 'Teacher', 'Username', 'Start Date', 'End Date', 'Notes', 'Created By', 'Created At'];
    var r = sh.getRange(1, 1, 1, headers.length);
    r.setValues([headers]);
    r.setBackground('#0d1b36').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('D:E').setNumberFormat('@');   // keep yyyy-MM-dd as text
    sh.setColumnWidth(2, 200); sh.setColumnWidth(6, 260);
  }
  return sh;
}

/* ── teachers that can be placed on the roster ─────────────────────── */
function getTODTeachersHOS() {
  _hosAssert_();
  return getTeacherList()
    .filter(function (t) { return HOS_NON_TEACHING.indexOf(String(t.position || '').trim().toLowerCase()) < 0; })
    .map(function (t) { return { name: t.name, username: t.username, position: t.position || 'Teacher' }; });
}

/* ── roster ────────────────────────────────────────────────────────── */
function getTODRosterHOS() {
  _hosAssert_();
  var sh = _hosRosterSheet_();
  if (sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  var out = [];
  rows.forEach(function (r) {
    if (!r[0]) return;
    out.push({
      id: String(r[0]), teacher: String(r[1] || ''), username: String(r[2] || ''),
      start: _hosYmd_(r[3]), end: _hosYmd_(r[4]), notes: String(r[5] || '')
    });
  });
  out.sort(function (a, b) { return a.start < b.start ? -1 : (a.start > b.start ? 1 : 0); });
  return out;
}

/**
 * entries: [{teacher, username, start:'yyyy-MM-dd', end:'yyyy-MM-dd', notes}]
 */
function saveTODRosterHOS(entries) {
  var me = _hosAssert_();
  try {
    var chk = checkDashboardWriteAccess(['TOD Reports', 'Reports']);
    if (chk && chk.ok === false) return { success: false, message: chk.message };
  } catch (e) {}

  if (!entries || !entries.length) return { success: false, message: 'Nothing to save.' };
  var sh = _hosRosterSheet_();
  var existing = {};
  getTODRosterHOS().forEach(function (e) { existing[(e.teacher + '|' + e.start + '|' + e.end).toLowerCase()] = true; });

  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var rows = [], skipped = 0;
  entries.forEach(function (e) {
    var t = String(e.teacher || '').trim(), s = String(e.start || '').substring(0, 10), en = String(e.end || '').substring(0, 10);
    if (!t || !/^\d{4}-\d{2}-\d{2}$/.test(s) || !/^\d{4}-\d{2}-\d{2}$/.test(en) || en < s) { skipped++; return; }
    var key = (t + '|' + s + '|' + en).toLowerCase();
    if (existing[key]) { skipped++; return; }
    existing[key] = true;
    rows.push([Utilities.getUuid().substring(0, 8), t, String(e.username || ''), s, en, String(e.notes || ''), me.fullName || me.username, now]);
  });
  if (!rows.length) return { success: false, message: 'No new duties were added (duplicates or invalid dates).' };

  var start = sh.getLastRow() + 1;
  sh.getRange(start, 4, rows.length, 2).setNumberFormat('@');
  sh.getRange(start, 1, rows.length, 8).setValues(rows);
  try { logActivity(me.username, 'TOD roster: added ' + rows.length + ' duty slot(s)'); } catch (e) {}
  return { success: true, added: rows.length, skipped: skipped,
           message: rows.length + ' duty slot(s) saved' + (skipped ? ' (' + skipped + ' skipped)' : '') + '.' };
}

function deleteTODRosterHOS(id) {
  var me = _hosAssert_();
  var sh = _hosRosterSheet_();
  if (sh.getLastRow() < 2) return { success: false, message: 'Roster is empty.' };
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sh.deleteRow(i + 2);
      try { logActivity(me.username, 'TOD roster: removed duty slot ' + id); } catch (e) {}
      return { success: true, message: 'Duty removed.' };
    }
  }
  return { success: false, message: 'Duty not found.' };
}

/* ── reports (read only) ───────────────────────────────────────────── */
// Reports come from TWO sheets: the new "TOD Duty Reports" (format v2, see
// TODReportAPI.gs) and the old "TOD Reports" (legacy format).
function getAllTODReportsHOS() {
  _hosAssert_();
  var out = [], seen = {};
  _todAll_().forEach(function (r) {
    seen[(r.teacherName + '|' + r.date).toLowerCase()] = true;
    out.push({ id: r.id, format: 'v2', date: r.date, teacher: r.teacherName, submittedAt: r.submittedAt,
               totalPresent: r.percentage + '%', totalAbsentees: r.totalAbsent, pdfLink: '' });
  });
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TOD Reports');
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 25).getValues().forEach(function (row) {
      var d = _hosYmd_(row[1]);
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

function getTODReportDetailHOS(teacher, dateStr, format) {
  _hosAssert_();
  if (format === 'v2') {
    var found = null;
    _todAll_().forEach(function (r) { if (r.date === dateStr && r.teacherName === teacher) found = r; });
    if (!found) return null;
    return { format: 'v2', date: found.date, teacherName: found.teacherName, answers: found.answers,
             attendanceRows: found.attendanceRows, school: TOD_SCHOOL_INFO };
  }
  var r = getTODReportByDate(teacher, dateStr);
  if (!r) return null;
  var out = { format: 'legacy' };
  Object.keys(r).forEach(function (k) {
    var v = r[k];
    out[k] = v instanceof Date
      ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(v == null ? '' : v);
  });
  return out;
}