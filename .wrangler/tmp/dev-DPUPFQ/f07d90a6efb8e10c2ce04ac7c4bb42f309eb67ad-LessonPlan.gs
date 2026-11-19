/**
 * LessonPlan.gs — Lesson Plan panel (server side)
 * =====================================================================
 * Used by 4 accounts with the SAME procedure:
 *   Academic, Class Teacher, Subject Teacher, Depute Headmaster.
 * Admin switches it ON/OFF per account with the feature "Lesson Plan"
 * (Admin → Manage Assignments and Accounts; see RBAC.gs / DashboardAccess.gs).
 *
 * Needs these files in the same Apps Script project:
 *   LessonPlanSyllabus.gs  (LP_SYLLABUS data, Form I–IV, 12 subjects)
 *   LessonPlanPanel.html   (the panel; injected into the 4 dashboards by rbacInject_)
 * Data is saved in the master spreadsheet, tab "LessonPlans" (created automatically).
 * =====================================================================
 */
var LP_SHEET   = 'LessonPlans';
var LP_FEATURE = 'Lesson Plan';
var LP_KEYS    = ['academics', 'classTeacher', 'teacher', 'secondMaster'];
var LP_HEADERS = ['ID', 'Created', 'Updated', 'Username', 'Teacher', 'Account', 'Class', 'Subject', 'Date', 'Time', 'Main Activity', 'Data'];

/* ── injection into the dashboards (called from rbacInject_ in RBAC.gs) ── */
function lpInjectPanel_(key, html) {
  if (LP_KEYS.indexOf(key) < 0) return html;
  var panel;
  try { panel = HtmlService.createHtmlOutputFromFile('LessonPlanPanel').getContent(); }
  catch (e) { Logger.log('LessonPlanPanel missing: ' + e.message); return html; }
  html = String(html || '');
  var at = html.toLowerCase().lastIndexOf('</body>');
  return at >= 0 ? html.slice(0, at) + panel + '\n' + html.slice(at) : html + panel;
}

/* ── helpers ─────────────────────────────────────────────── */
function lpSession_() {
  var s = getTeacherSession();
  if (!s || !s.username) throw new Error('Session expired. Please log in again.');
  var key = rbacKeyForPosition_(s.position);
  if (LP_KEYS.indexOf(key) < 0) throw new Error('Lesson Plan is not available for this account.');
  if (rbacFeatures_(key)[LP_FEATURE] === false) throw new Error('Lesson Plan has been switched OFF by Admin.');
  return { s: s, key: key };
}

function lpSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(LP_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LP_SHEET);
    sh.appendRow(LP_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, LP_HEADERS.length).setBackground('#0d1b36').setFontColor('#ffffff').setFontWeight('bold');
    sh.getRange('I:I').setNumberFormat('@');   // keep dates as text (yyyy-MM-dd)
  }
  return sh;
}

function lpTz_() { return Session.getScriptTimeZone() || 'Africa/Dar_es_Salaam'; }

function lpDateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, lpTz_(), 'yyyy-MM-dd');
  return String(v || '').trim().substring(0, 10);
}

function lpFormNumber_(className) {
  var s = String(className || '').toLowerCase();
  var d = s.match(/(\d)/);
  if (d && +d[1] >= 1 && +d[1] <= 4) return +d[1];
  var words = { one: 1, two: 2, three: 3, four: 4, kwanza: 1, pili: 2, tatu: 3, nne: 4 };
  for (var w in words) if (new RegExp('\\b' + w + '\\b').test(s)) return words[w];
  var r = s.match(/(?:form|kidato(?:\s+cha)?|f)\s*(iv|iii|ii|i)\b/);
  return r ? { i: 1, ii: 2, iii: 3, iv: 4 }[r[1]] : 0;
}

function lpClassKey_(v) { return String(v || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }

/* 'g' = girl, 'b' = boy, '' = unknown. Accepts F/Female/Girl/KE and M/Male/Boy/ME. */
function lpSex_(v) {
  var s = String(v || '').trim().toLowerCase();
  if (!s) return '';
  if (/^(f|female|girl|g|ke|mwanamke|msichana)/.test(s)) return 'g';
  if (/^(m|male|boy|b|me|mwanaume|mvulana)/.test(s)) return 'b';
  return '';
}

function lpStr_(v, max) { return String(v == null ? '' : v).slice(0, max || 2000); }
function lpNum_(v) { v = Number(v); return isFinite(v) && v > 0 ? Math.floor(v) : 0; }

/* Keeps only known fields with sane sizes. */
function lpClean_(p) {
  p = p || {};
  function cnt(c) { c = c || {}; return { girls: lpNum_(c.girls), boys: lpNum_(c.boys) }; }
  return {
    id: lpStr_(p.id, 64), school: lpStr_(p.school, 200), teacher: lpStr_(p.teacher, 200),
    className: lpStr_(p.className, 100), subject: lpStr_(p.subject, 100), syllabusSubject: lpStr_(p.syllabusSubject, 100),
    lang: p.lang === 'sw' ? 'sw' : 'en', form: lpNum_(p.form), date: lpStr_(p.date, 10),
    start: lpStr_(p.start, 5), end: lpStr_(p.end, 5), timeText: lpStr_(p.timeText, 60),
    registered: cnt(p.registered), present: cnt(p.present),
    mainCompetence: lpStr_(p.mainCompetence), specificCompetence: lpStr_(p.specificCompetence),
    topic: lpStr_(p.topic, 300), mainActivity: lpStr_(p.mainActivity), verb: lpStr_(p.verb, 40),
    specificText: lpStr_(p.specificText), specificActivities: lpStr_(p.specificActivities),
    resources: lpStr_(p.resources), references: lpStr_(p.references), style: lpNum_(p.style) ? 1 : 0,
    stages: (Array.isArray(p.stages) ? p.stages : []).slice(0, 8).map(function (st) {
      st = st || {};
      return { name: lpStr_(st.name, 80), minutes: lpNum_(st.minutes), teaching: lpStr_(st.teaching), learning: lpStr_(st.learning), assessment: lpStr_(st.assessment) };
    }),
    remarks: lpStr_(p.remarks, 3000)
  };
}

function lpFindRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2 || !id) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

/* ── API used by LessonPlanPanel.html ────────────────────── */
function lpBootstrap() {
  var ctx = lpSession_(), s = ctx.s, sc = {}, all = [];
  try { sc = rbacSchool_(); } catch (e) {}
  try { all = getAcademicClassesRaw().map(function (c) { return String(c.name); }); } catch (e) {}
  return {
    school: sc.name || '', teacher: s.fullName || s.username, username: s.username,
    account: (typeof DASHBOARD_TITLES !== 'undefined' && DASHBOARD_TITLES[ctx.key]) || ctx.key,
    myClasses: s.classes || [], allClasses: all, mySubjects: s.subjects || [],
    subjects: Object.keys(LP_SYLLABUS).map(function (n) { return { name: n, lang: LP_SYLLABUS[n].lang }; }),
    today: Utilities.formatDate(new Date(), lpTz_(), 'yyyy-MM-dd')
  };
}

/* Syllabus rows of one subject: [form1Rows, form2Rows, form3Rows, form4Rows]. */
function lpGetSyllabus(subject) {
  lpSession_();
  var s = LP_SYLLABUS[String(subject || '')];
  return s ? s.forms : null;
}

/* Registered students (Form sheets) and present students (Attendance sheet) for a class on a date. */
function lpStudentCounts(className, dateStr) {
  lpSession_();
  var out = { registered: { girls: 0, boys: 0 }, present: { girls: 0, boys: 0 }, presentFound: false };
  var form = lpFormNumber_(className);
  if (!form) return out;
  var ss = SpreadsheetApp.getActiveSpreadsheet(), want = lpClassKey_(className);

  function pick(rows, classCol) {
    var same = rows.filter(function (r) { return lpClassKey_(r[classCol]) === want; });
    return same.length ? same : rows;
  }

  var fs = ss.getSheetByName('Form' + form);
  if (fs && fs.getLastRow() > 1) {
    var rows = fs.getDataRange().getValues().slice(1).filter(function (r) { return r[0] || r[1]; });
    pick(rows, 5).forEach(function (r) {
      var g = lpSex_(r[3]);
      if (g === 'g') out.registered.girls++; else if (g === 'b') out.registered.boys++;
    });
  }

  var as = ss.getSheetByName('Attendance');
  if (as && as.getLastRow() > 1 && dateStr) {
    var day = String(dateStr).substring(0, 10);
    var arows = as.getDataRange().getValues().slice(1).filter(function (r) {
      return lpDateStr_(r[0]) === day && lpFormNumber_(r[1]) === form;
    });
    var use = pick(arows, 1);
    if (use.length) {
      out.presentFound = true;
      use.forEach(function (r) {
        var here = String(r[6]).trim().toUpperCase() === 'P' || String(r[7]).trim().toUpperCase() === 'P';
        if (!here) return;
        var g = lpSex_(r[4]);
        if (g === 'g') out.present.girls++; else if (g === 'b') out.present.boys++;
      });
    }
  }
  return out;
}

function lpSave(plan) {
  var ctx = lpSession_();
  var acc = checkDashboardWriteAccess([LP_FEATURE]);
  if (!acc.ok) throw new Error(acc.message);
  plan = lpClean_(plan);
  if (!plan.className || !plan.subject || !plan.date) throw new Error('Class, subject and date are required.');
  var user = ctx.s.username, now = new Date(), sh = lpSheet_();
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var row = plan.id ? lpFindRow_(sh, plan.id) : 0, created = now;
    if (row) {
      var cur = sh.getRange(row, 1, 1, 4).getValues()[0];
      if (String(cur[3]) !== String(user)) throw new Error('You can only edit your own lesson plans.');
      created = cur[1] || now;
    } else {
      plan.id = Utilities.getUuid();
    }
    var json = JSON.stringify(plan);
    if (json.length > 45000) throw new Error('This lesson plan is too long to save. Shorten some fields.');
    var vals = [plan.id, created, now, user, ctx.s.fullName || user, (typeof DASHBOARD_TITLES !== 'undefined' && DASHBOARD_TITLES[ctx.key]) || ctx.key,
      plan.className, plan.subject, plan.date, plan.timeText, plan.mainActivity.slice(0, 300), json];
    if (row) sh.getRange(row, 1, 1, vals.length).setValues([vals]);
    else sh.appendRow(vals);
  } finally { lock.releaseLock(); }
  try { logActivity(user, 'Lesson plan saved: ' + plan.subject + ' – ' + plan.className + ' (' + plan.date + ')'); } catch (e) {}
  return { id: plan.id, updated: Utilities.formatDate(now, lpTz_(), 'yyyy-MM-dd HH:mm') };
}

function lpList() {
  var ctx = lpSession_(), sh = lpSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, LP_HEADERS.length).getValues(), out = [];
  rows.forEach(function (r) {
    if (String(r[3]) !== String(ctx.s.username)) return;
    out.push({ id: String(r[0]), updated: r[2] instanceof Date ? Utilities.formatDate(r[2], lpTz_(), 'yyyy-MM-dd HH:mm') : String(r[2]),
      className: String(r[6]), subject: String(r[7]), date: lpDateStr_(r[8]), time: String(r[9]), mainActivity: String(r[10]) });
  });
  out.sort(function (a, b) { return (b.date + b.updated).localeCompare(a.date + a.updated); });
  return out;
}

function lpGet(id) {
  var ctx = lpSession_(), sh = lpSheet_(), row = lpFindRow_(sh, id);
  if (!row) throw new Error('Lesson plan not found.');
  var r = sh.getRange(row, 1, 1, LP_HEADERS.length).getValues()[0];
  if (String(r[3]) !== String(ctx.s.username)) throw new Error('You can only open your own lesson plans.');
  var p = JSON.parse(String(r[11] || '{}'));
  p.id = String(r[0]);
  return p;
}

function lpDelete(id) {
  var ctx = lpSession_();
  var acc = checkDashboardWriteAccess([LP_FEATURE]);
  if (!acc.ok) throw new Error(acc.message);
  var sh = lpSheet_(), row = lpFindRow_(sh, id);
  if (!row) return 'Already deleted.';
  if (String(sh.getRange(row, 4).getValue()) !== String(ctx.s.username)) throw new Error('You can only delete your own lesson plans.');
  sh.deleteRow(row);
  try { logActivity(ctx.s.username, 'Lesson plan deleted: ' + id); } catch (e) {}
  return 'Deleted.';
}

/* Converts the printable lesson plan (built by the panel) into a PDF file. */
function lpPdf(html, fileName) {
  lpSession_();
  html = String(html || '');
  if (!html || html.length > 300000) throw new Error('Nothing to convert.');
  var name = String(fileName || 'Lesson_Plan').replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '_').slice(0, 80) || 'Lesson_Plan';
  var pdf = Utilities.newBlob(html, 'text/html', name + '.html').getAs('application/pdf');
  return { name: name + '.pdf', data: Utilities.base64Encode(pdf.getBytes()) };
}
