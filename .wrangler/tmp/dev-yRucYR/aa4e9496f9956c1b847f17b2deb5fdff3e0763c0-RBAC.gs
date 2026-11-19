/**
 * RBAC.gs — Centralized Role-Based Access Control (Admin decides everything)
 * -------------------------------------------------------------------------
 * Reuses the existing Admin "Manage Assignments and Accounts" (ex Monitor) policy store:
 * policy.features[featureName] = true/false per account key.
 *   Keys: academics, classTeacher, teacher (Subject Teacher), secondMaster (Depute Headmaster), parents
 * Features that control account creation / account management are OFF until Admin turns them ON.
 */
const RBAC_PANELS = {
  academics: {
    'Student Admissions':['admissions'], 'Manage Classes':['classes'], 'Map Students':['mapstudents'],
    'Year Progression':['progression'], 'Attendance Tracking':['attendance'], 'Manage Subjects':['subjects'],
    'Exam Setup':['examtab-setup'], 'Score Deadline':['examtab-deadline'], 'Publish Results':['examtab-publish'], 'Report Cards':['examtab-cards'], 'Submit Scores':['scores'],
    'Tracking Scores':['tracking'], 'Timetable Settings':['timetable'], 'TOD Report':['tod'], 'Messages':['messages'], 'Lesson Plan':['lessonplan'], 'Academic Settings':['acadsettings'],
    'Manage Staff Accounts':['staffacc'], 'Manage Parent Accounts':['parentacc'],
    'Create Class Teacher Account':[], 'Create Subject Teacher Account':[], 'Create Depute Headmaster Account':[], 'Create Parent Account':[]
  },
  classTeacher: {
    'My Assignments':['assignments'], 'My Timetable':['timetable'], 'Behavioural Criteria':['behaviour'],
    'Submit Scores':['scores'], 'Score Tracking':['tracking'], 'Submission Report':['submission-report'],
    'Attendance':['attendance'], 'TOD Report':['tod'], 'Messages':['messages'], 'Lesson Plan':['lessonplan']
  },
  teacher: {
    'My Assignments':['assignments'], 'My Timetable':['timetable'], 'Submit Scores':['scores'],
    'Score Tracking':['tracking'], 'TOD Report':['tod'], 'Submission Report':['submit-report'], 'Messages':['messages'], 'Lesson Plan':['lessonplan']
  },
  secondMaster: {
    'My Assignments':['assignments'], 'Submit Scores':['scores'], 'Score Tracking':['tracking'],
    'Submission Report':['submit-report'], 'Attendance Tracking':['attendance'], 'TOD Tracking':['todtrack'], 'Lesson Plan':['lessonplan']
  },
  parents: {
    'Report Card':['report'], 'Performance':['performance'], 'Attendance':['attendance'],
    'Announcements':['announcements'], 'Student Info':['info']
  }
};

function rbacDefaultOff_(name) { return /^(Manage (Staff|Parent) Accounts|Create )/.test(name); }

function rbacKeyForPosition_(pos) {
  pos = String(pos || '').toLowerCase();
  if (/second\s*master|depute|deputy/.test(pos)) return 'secondMaster';
  if (pos.indexOf('academic') >= 0) return 'academics';
  if (pos.indexOf('class teacher') >= 0) return 'classTeacher';
  return 'teacher';
}

/** Effective ON/OFF map for one account key. */
function rbacFeatures_(key) {
  var out = {};
  Object.keys(RBAC_PANELS[key] || {}).forEach(function (n) { out[n] = !rbacDefaultOff_(n); });
  try {
    var p = getDashboardAccessPolicies()[key];
    if (p && p.features) Object.keys(p.features).forEach(function (n) {
      if (n in out) { var v = p.features[n]; out[n] = !(v === false || String(v).toLowerCase() === 'false'); }
    });
  } catch (e) { Logger.log('rbacFeatures_: ' + e.message); }
  return out;
}

/** Called by dashboards (polling) so Admin changes apply live. */
function rbacMyFeatures(key) {
  var k = key;
  try { var s = getTeacherSession(); if (s && s.username) k = rbacKeyForPosition_(s.position); } catch (e) {}
  if (!RBAC_PANELS[k]) return { features: {} };
  return { key: k, features: rbacFeatures_(k) };
}

/** Injects the client-side enforcer (hides disabled menus/buttons, blocks goTo) into a dashboard page. */
function rbacInject_(key, html) {
  html = rbacBrand_(html);
  if (!RBAC_PANELS[key]) return html;
  if (typeof lpInjectPanel_ === 'function') html = lpInjectPanel_(key, html);   // Lesson Plan panel (LessonPlan.gs)
  var client = '(function(){var KEY=' + JSON.stringify(key) + ',MAP=' + JSON.stringify(RBAC_PANELS[key]) +
    ',FEAT=' + JSON.stringify(rbacFeatures_(key)) + ',OFFP={};' +
    'function calc(){OFFP={};Object.keys(MAP).forEach(function(f){if(FEAT[f]===false)MAP[f].forEach(function(p){OFFP[p]=1;});});}' +
    'function apply(){calc();window.RBAC_FEATURES=FEAT;Object.keys(MAP).forEach(function(f){MAP[f].forEach(function(p){' +
    'var off=!!OFFP[p];document.querySelectorAll("[data-p=\\""+p+"\\"],[onclick*=\\"goTo(\'"+p+"\')\\"]").forEach(function(el){' +
    'if(off){if(el.dataset.rbacd===undefined)el.dataset.rbacd=el.style.display;el.style.display="none";}' +
    'else if(el.dataset.rbacd!==undefined){el.style.display=el.dataset.rbacd;delete el.dataset.rbacd;}});});});' +
    'try{if(window.rbacOnApply)window.rbacOnApply(FEAT);}catch(e){}}' +
    'var og=window.goTo;if(typeof og==="function"){window.goTo=function(id){calc();if(OFFP[id]){return og.call(this,"dashboard");}return og.apply(this,arguments);};}' +
    'function poll(){try{google.script.run.withSuccessHandler(function(r){if(r&&r.features){FEAT=r.features;apply();}}).withFailureHandler(function(){}).rbacMyFeatures(KEY);}catch(e){}}' +
    'if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apply);else apply();' +
    'setTimeout(apply,400);setTimeout(poll,1500);setInterval(poll,30000);})();';
  var tag = '<script>' + client.replace(/<\//g, '<\\/') + '<\/script>';
  html = String(html || '');
  // Insert before the LAST </body> only: some pages contain "</body>" inside JS strings (print windows).
  var at = html.toLowerCase().lastIndexOf('</body>');
  return at >= 0 ? html.slice(0, at) + tag + '\n' + html.slice(at) : html + tag;
}

/* ───────── Account helpers ───────── */
function rbacIsInactiveStaff_(username) {
  try {
    var sh = ensureUsersSheet(), last = sh.getLastRow(); if (last < 2) return false;
    var rows = sh.getRange(2, 2, last - 1, 9).getValues(); // B..J
    for (var i = 0; i < rows.length; i++) if (String(rows[i][0]) === String(username)) return String(rows[i][8]).toLowerCase() === 'inactive';
  } catch (e) {}
  return false;
}
function rbacIsInactiveParent_(username) {
  try {
    var sh = ensureParentsSheet(), last = sh.getLastRow(); if (last < 2) return false;
    var rows = sh.getRange(2, 8, last - 1, 6).getValues(); // H..M
    for (var i = 0; i < rows.length; i++) if (String(rows[i][0]).trim().toLowerCase() === String(username).trim().toLowerCase()) return String(rows[i][5]).toLowerCase() === 'inactive';
  } catch (e) {}
  return false;
}
/** Only ONE Academic account may exist (Admin manages it). */
function rbacAcademicConflict_(position, excludeUsername) {
  if (rbacKeyForPosition_(position) !== 'academics') return false;
  return getTeacherList().some(function (t) {
    return rbacKeyForPosition_(t.position) === 'academics' && String(t.username) !== String(excludeUsername || '');
  });
}

/* ───────── Academic account management (permission-checked) ───────── */
function acadCheck_(feature) {
  var s = getTeacherSession();
  if (!s || !s.username) throw new Error('Session expired. Please log in again.');
  if (rbacKeyForPosition_(s.position) !== 'academics') throw new Error('Only the Academic account can do this.');
  if (rbacFeatures_('academics')[feature] !== true) throw new Error('Not enabled by Admin: ' + feature);
  return s;
}
function acadStaffRole_(position) {
  position = String(position || '').trim();
  if (position === 'Class Teacher') return 'Create Class Teacher Account';
  if (position === 'Teacher') return 'Create Subject Teacher Account';
  if (position === 'Second Master') return 'Create Depute Headmaster Account';
  throw new Error('Academic can only manage Class Teacher, Subject Teacher and Depute Headmaster accounts.');
}
function acadIsManagedStaff_(username) {
  var t = getTeacherList().find(function (x) { return String(x.username) === String(username); });
  return t && (t.position === 'Class Teacher' || t.position === 'Teacher' || t.position === 'Second Master') ? t : null;
}
function acadGetStaff() {
  var s = acadCheck_('Manage Staff Accounts');
  return getTeacherList().filter(function (t) { return t.position === 'Class Teacher' || t.position === 'Teacher' || t.position === 'Second Master'; })
    .map(function (t) {
      return { name: t.name, username: t.username, classes: t.classes.join(','), subjects: t.subjects.join(','),
        sex: t.sex, position: t.position, department: t.department, picUrl: t.picUrl, active: !rbacIsInactiveStaff_(t.username) };
    });
}
function acadAddStaff(name, username, password, classes, subjects, sex, position, department, picUrl) {
  var s = acadCheck_('Manage Staff Accounts'); acadCheck_(acadStaffRole_(position));
  if (getTeacherList().some(function (t) { return String(t.username) === String(username); })) return 'Error: username already exists';
  return addTeacher(name, username, password, classes, subjects, s.username, sex, position, department, picUrl);
}
function acadUpdateStaff(oldUsername, name, username, password, classes, subjects, sex, position, department, picUrl) {
  var s = acadCheck_('Manage Staff Accounts'); acadCheck_(acadStaffRole_(position));
  if (!acadIsManagedStaff_(oldUsername)) throw new Error('You cannot modify this account.');
  var wasInactive = rbacIsInactiveStaff_(oldUsername);
  var r = updateTeacher(oldUsername, name, username, password, classes, subjects, s.username, sex, position, department, picUrl);
  if (wasInactive) acadSetActiveRaw_('staff', username, false);
  return r;
}
function acadDeleteStaff(username) {
  var s = acadCheck_('Manage Staff Accounts');
  if (!acadIsManagedStaff_(username)) throw new Error('You cannot delete this account.');
  return deleteTeacher(username, s.username);
}
function acadGetParents() { acadCheck_('Manage Parent Accounts'); return getParentList().map(function (r) { return r.slice(0, 12).concat([rbacIsInactiveParent_(r[7]) ? 'inactive' : 'active']); }); }
function acadAddParent(data) { acadCheck_('Manage Parent Accounts'); acadCheck_('Create Parent Account'); return addParent(data); }
function acadUpdateParent(oldUsername, data) { acadCheck_('Manage Parent Accounts'); return updateParent(oldUsername, data); }
function acadDeleteParent(username) { acadCheck_('Manage Parent Accounts'); return deleteParent(username); }

function acadSetActiveRaw_(type, username, active) {
  var sh = type === 'staff' ? ensureUsersSheet() : ensureParentsSheet();
  var userCol = type === 'staff' ? 2 : 8, statusCol = type === 'staff' ? 10 : 13, last = sh.getLastRow();
  if (last < 2) return false;
  var names = sh.getRange(2, userCol, last - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) if (String(names[i][0]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
    sh.getRange(i + 2, statusCol).setValue(active ? 'active' : 'inactive'); return true;
  }
  return false;
}
/** Activate / deactivate a Class Teacher, Subject Teacher or Parent account. */
function acadSetActive(type, username, active) {
  var s;
  if (type === 'staff') { s = acadCheck_('Manage Staff Accounts'); if (!acadIsManagedStaff_(username)) throw new Error('You cannot modify this account.'); }
  else { s = acadCheck_('Manage Parent Accounts'); }
  if (!acadSetActiveRaw_(type, username, !!active)) return 'Account not found.';
  try { logActivity(s.username, (active ? 'Activated ' : 'Deactivated ') + type + ' account: ' + username); } catch (e) {}
  return 'Account ' + username + (active ? ' activated.' : ' deactivated.');
}

/* ───────── AUTO-CREATE ALL SHEET TABS ─────────
 * setupAllSheets()      → run manually once from the Apps Script editor (Run ▶) to create every tab.
 * rbacAutoSetup_()      → runs automatically once (from doGet) the first time the web app is opened.
 * Existing tabs and data are never changed; only missing tabs / the "Status" header are added.
 */
function setupAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), created = [], failed = [];
  function mk(name, headers, color) {
    var sh = ss.getSheetByName(name);
    if (sh) return sh;
    sh = ss.insertSheet(name);
    sh.appendRow(headers); sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setBackground(color || '#0d1b36').setFontColor('#ffffff').setFontWeight('bold');
    created.push(name); return sh;
  }
  function run(label, fn) { try { var before = ss.getSheets().length; fn(); if (ss.getSheets().length > before) created.push(label); } catch (e) { failed.push(label + ': ' + e.message); } }

  // existing ensure* helpers (same headers the panels already expect)
  run('Users', ensureUsersSheet);
  run('Parents', ensureParentsSheet);
  run('Attendance', ensureAttendanceSheet);
  run('AcademicYears', ensureAcademicYearsSheet);
  run('Alumni', ensureAlumniSheet);
  run('Behaviors', ensureBehaviorSheet_);
  run('DashboardAccess', ensureDashboardAccessSheet_);
  run('Scores', ensureScoresSheet);
  [1, 2, 3, 4].forEach(function (n) { run('Form' + n + '_Scores', function () { getScoresSheet('Form ' + n); }); });

  // tabs that were only created lazily before
  [1, 2, 3, 4].forEach(function (n) { mk('Form' + n, ['AdmNo','FirstName','LastName','Sex','DOB','Class','Parent','Phone','Address']); });
  run('Activity',  function () { mk('Activity', ['Date','User','Action']); });
  run('AcademicClasses',  function () { mk('AcademicClasses', ['ID','Name','Stream','Teacher','Capacity','Students']); });
  run('AcademicSubjects', function () { mk('AcademicSubjects', ['Name','Code','Category','Department','Classes']); });
  run('Examinations', function () { mk('Examinations', ['ID','Name','Type','Year','StartDate','EndDate','Classes','Remarks','Status']); });
  run('Messages', function () { mk('Messages', ['From','To','Subject','Body','Date','FromName','ToName']); });
  run('TODReports', function () { mk('TODReports', ['Date','Teacher','Class','Report']); });
  run('AcademicSettings', function () { mk('AcademicSettings', ['Key','Part','Value']); });
  run('SystemConfig', function () { mk('SystemConfig', ['Key','Value']); });
  run('SchoolConfig', function () { mk('SchoolConfig', ['Key','Value']); });
  run('AdminProfile', function () { mk('AdminProfile', ['Key','Value']); });
  run('DashboardAccessEvents', function () { mk('DashboardAccessEvents', ['Time','Panel','Action','By']); });

  // account Status column (blank = active, "inactive" = deactivated by Academic)
  run('Users.Status', function () { var s = ss.getSheetByName('Users'); if (s && !String(s.getRange(1, 10).getValue())) s.getRange(1, 10).setValue('Status').setFontWeight('bold'); });
  run('Parents.Status', function () { var s = ss.getSheetByName('Parents'); if (s && !String(s.getRange(1, 13).getValue())) s.getRange(1, 13).setValue('Status').setFontWeight('bold'); });

  var msg = 'Sheets created: ' + (created.length ? created.join(', ') : 'none (all existed)') + (failed.length ? ' | Problems: ' + failed.join('; ') : '');
  Logger.log(msg); return msg;
}

function rbacAutoSetup_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SMS_SHEETS_READY_V1') === 'yes') return;
  setupAllSheets();
  props.setProperty('SMS_SHEETS_READY_V1', 'yes');
}

/* ───────── Academic Profile (Academic sets own Department / Classes / Subjects) ───────── */
function acadProfileSession_() {
  var s = getTeacherSession();
  if (!s || !s.username) throw new Error('Session expired. Please log in again.');
  if (rbacKeyForPosition_(s.position) !== 'academics') throw new Error('Only the Academic account can do this.');
  return s;
}
function acadGetMyProfile() {
  var s = acadProfileSession_();
  var t = getTeacherList().find(function (x) { return String(x.username) === String(s.username); });
  if (!t) throw new Error('Account not found.');
  return { name: t.name, sex: t.sex || '', department: t.department || '', picUrl: t.picUrl || '',
           classes: (t.classes || []).join(','), subjects: (t.subjects || []).join(',') };
}
/** Profile picture: uploaded (resized in the browser) and stored on the Users sheet (column I). '' removes it. */
function acadSaveMyPicture(pic) {
  var s = acadProfileSession_();
  pic = String(pic || '');
  if (pic && (pic.length > 48000 || !/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+\/=]+$/.test(pic))) throw new Error('Invalid or too large picture. Please choose another image.');
  var sh = ensureUsersSheet(), data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(s.username)) {
      var cell = sh.getRange(i + 1, 9); cell.setNumberFormat('@'); cell.setValue(pic);
      try { logActivity(s.username, pic ? 'Updated profile picture' : 'Removed profile picture'); } catch (e) {}
      return pic ? 'Picture updated.' : 'Picture removed.';
    }
  }
  throw new Error('Account not found.');
}
function acadSaveMyProfile(name, sex, department, classes, subjects) {
  var s = acadProfileSession_();
  name = String(name || '').trim();
  if (!name) throw new Error('Full name is required.');
  function clean(v) { return String(v || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean).join(','); }
  var sh = ensureUsersSheet(), data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(s.username)) {
      sh.getRange(i + 1, 1).setValue(name);               // Name
      sh.getRange(i + 1, 4).setValue(clean(classes));     // Classes
      sh.getRange(i + 1, 5).setValue(clean(subjects));    // Subjects
      sh.getRange(i + 1, 6).setValue(sex || '');          // Sex
      sh.getRange(i + 1, 8).setValue(department || '');   // Department
      try { logActivity(s.username, 'Updated own profile'); } catch (e) {}
      return 'Profile saved successfully!';
    }
  }
  throw new Error('Account not found.');
}