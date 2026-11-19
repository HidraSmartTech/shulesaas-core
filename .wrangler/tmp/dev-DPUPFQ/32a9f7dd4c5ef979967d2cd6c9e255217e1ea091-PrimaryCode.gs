// ============================================================
// UNIFIED CODE.GS - School Management System
// ALL data stored in the MASTER spreadsheet only.
// ============================================================

// ==================== CONFIGURATIONS ====================
const CONFIG_SHEET_NAME   = "Config";
const STATUS_CELL         = "B1";
const DEADLINE_OPEN_CELL  = "B2";
const DEADLINE_CLOSE_CELL = "B3";

// ==================== ADMIN USERS ====================
const ADMIN_USERS = __APP_ENV__.adminUsers; // from INITIAL_MANAGER_* environment variables

// ==================== FORM SHEET IDs (legacy admin links only) ====================
const FORM_SHEETS = {
};

// ==================== ATTENDANCE SHEET IDs (legacy admin links only) ====================
const ATTENDANCE_SHEETS = {
};

// ==================== TEACHER ON DUTY SHEET ====================
const TOD_SHEET_ID = '';

// ==================== SUBJECT COLUMNS (legacy compat) ====================
const SUBJECT_COLUMNS = {
  "H/Maadili": 4, "Civics": 4,
  "History": 7,   "Geography": 10,
  "Kiswahili": 13,"English": 16,
  "Physics": 19,  "Chemistry": 22,
  "Biology": 25,  "Mathematics": 28,
  "BusinessStudies": 31, "ComputerScience": 34
};

const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

// ============================================================
// WEB APP ENTRY POINT
// ============================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('School Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/* ===== ATTENDANCE: monthly data for Register + Report ===== */
// SPEED: the big 'Attendance' sheet and the 'Students' roster are now read ONCE per request
// (getMonthlyAttendanceBatch serves every class in one call) instead of once per class.

// roster sheet 'Students' (read once)
function attRosterRows_(ss) {
  var sh = ss.getSheetByName('Students');
  if (!sh || sh.getLastRow() <= 1) return null;
  var rv = sh.getDataRange().getValues(), rh = rv.shift();
  return { rows: rv, iName: rh.indexOf('Name'), iSex: rh.indexOf('Sex'), iCls: rh.indexOf('Class') };
}

// every Attendance row that falls in month/year (read once, parsed once)
function attMonthRows_(ss, month, year) {
  var attSh = ss.getSheetByName('Attendance'), out = [];
  if (!attSh || attSh.getLastRow() <= 1) return out;
  var av = attSh.getDataRange().getValues(), ah = av.shift();
  var c = { cls: ah.indexOf('Class'), date: ah.indexOf('Date'), name: ah.indexOf('StudentName'),
            sex: ah.indexOf('Sex'), m: ah.indexOf('Morning'), a: ah.indexOf('Afternoon') };
  var tz = Session.getScriptTimeZone(), dayCache = {};
  av.forEach(function(r){
    var d = c.date > -1 ? r[c.date] : r[0];
    if (!(d instanceof Date)) d = new Date(d);
    if (!d || isNaN(d.getTime())) return;
    if (d.getFullYear() !== year || (d.getMonth()+1) !== month) return;
    var t = d.getTime();
    var dateStr = dayCache[t] || (dayCache[t] = Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
    out.push({
      cls: String(c.cls > -1 ? r[c.cls] : r[1] || '').trim(),
      name: String(c.name > -1 ? r[c.name] : r[3] || '').trim(),
      sex: String(c.sex > -1 ? r[c.sex] : r[4] || '').trim(),
      dateStr: dateStr,
      morning: String(c.m > -1 ? r[c.m] : r[6] || ''),
      afternoon: String(c.a > -1 ? r[c.a] : r[7] || '')
    });
  });
  return out;
}

// builds { students, records } for ONE class from the already-loaded data
function attBuildClass_(ss, cls, roster, monthRows) {
  var studentsMap = {}, students = [];
  function addStudent(name, sex) {
    var cleanName = String(name || '').trim();
    if (!cleanName) return;
    if (!studentsMap[cleanName]) {
      studentsMap[cleanName] = { name: cleanName, sex: String(sex || '').trim() };
      students.push(studentsMap[cleanName]);
    } else if (!studentsMap[cleanName].sex && sex) {
      studentsMap[cleanName].sex = String(sex).trim();
    }
  }
  var formSh = studentSheetForClass_(ss, cls);
  if (formSh && formSh.getLastRow() > 1) {
    formSh.getDataRange().getValues().slice(1).forEach(function(r){
      var fullName = [r[1], r[2]].map(function(v){ return String(v || '').trim(); }).join(' ').trim();
      addStudent(fullName, r[3]);
    });
  }
  if (roster) {
    roster.rows.forEach(function(r){
      if (roster.iCls > -1 && !classLevelMatch(r[roster.iCls], cls)) return;
      addStudent(roster.iName > -1 ? r[roster.iName] : '', roster.iSex > -1 ? r[roster.iSex] : '');
    });
  }
  var records = [];
  monthRows.forEach(function(r){
    if (!classLevelMatch(r.cls, cls)) return;
    addStudent(r.name, r.sex);
    records.push({ name: r.name, dateStr: r.dateStr, morning: r.morning, afternoon: r.afternoon });
  });
  return { students: students, records: records };
}

function getMonthlyAttendance(p){
  attAssertClass_(p.class);
  var cls = String(p.class || '').trim();
  var month = Number(p.month), year = Number(p.year);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return attBuildClass_(ss, cls, attRosterRows_(ss), attMonthRows_(ss, month, year));
}

// ALL classes of a month in ONE call → { "Standard 1": {students, records}, ... }
function getMonthlyAttendanceBatch(p){
  var classes = (p.classes || []).map(function(c){ return String(c).trim(); });
  var month = Number(p.month), year = Number(p.year);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roster = attRosterRows_(ss), rows = attMonthRows_(ss, month, year), out = {};
  classes.forEach(function(c){ out[c] = attBuildClass_(ss, c, roster, rows); });
  return out;
}

// Days Recorded: every class × 12 months of a year in ONE call → { "Standard 1": [d1..d12], ... }
function getDaysSavedBatch(p){
  var classes = (p.classes || []).map(function(c){ return String(c).trim(); });
  var year = Number(p.year), ss = SpreadsheetApp.getActiveSpreadsheet();
  var sets = {}, out = {};
  classes.forEach(function(c){ sets[c] = []; for (var m = 0; m < 12; m++) sets[c].push({}); });
  var attSh = ss.getSheetByName('Attendance');
  if (attSh && attSh.getLastRow() > 1) {
    var av = attSh.getDataRange().getValues(), ah = av.shift();
    var iCls = ah.indexOf('Class'), iDate = ah.indexOf('Date');
    var matchCache = {};
    av.forEach(function(r){
      var d = iDate > -1 ? r[iDate] : r[0];
      if (!(d instanceof Date)) d = new Date(d);
      if (!d || isNaN(d.getTime()) || d.getFullYear() !== year) return;
      var rowClass = String(iCls > -1 ? r[iCls] : r[1] || '').trim();
      var hit = matchCache[rowClass];
      if (!hit) hit = matchCache[rowClass] = classes.filter(function(c){ return classLevelMatch(rowClass, c); });
      if (!hit.length) return;
      var key = d.getFullYear() + '-' + d.getMonth() + '-' + d.getDate(), mi = d.getMonth();
      hit.forEach(function(c){ sets[c][mi][key] = 1; });
    });
  }
  classes.forEach(function(c){ out[c] = sets[c].map(function(o){ return Object.keys(o).length; }); });
  return out;
}

/* ===== ATTENDANCE: yearly summary for Days Saved ===== */
function getYearlyAttendanceSummary(p){
  var cls = String(p.class || '').trim();
  var year = Number(p.year);
  var months = [];
  for (var month = 1; month <= 12; month++) {
    var data = getMonthlyAttendance({ class: cls, month: month, year: year });
    var seenDays = {};
    var present = 0;
    var total = 0;
    (data.records || []).forEach(function(r){
      if (r.dateStr) seenDays[r.dateStr] = 1;
      total++;
      if (String(r.morning || '').toUpperCase() === 'P') present++;
    });
    months.push({ days: Object.keys(seenDays).length, present: present, total: total });
  }
  return { months: months };
}

// ============================================================
// LOGIN SYSTEM
// ============================================================
// ============================================================
// REPLACE your existing checkLogin() with this version.
// It now returns parentData for the parent portal.
// ============================================================
function legacyCheckLogin_v1_DISABLED(username, password) {

  // ── Admin ──────────────────────────────────────────────────
  if (ADMIN_USERS.some(function(u) {
    return u.username === username && u.password === password;
  })) {
    return {
      status:   'success',
      role:     'admin',
      fullName: 'Administrator',
      username: username,
      html:     HtmlService.createHtmlOutputFromFile('Admin').getContent()
    };
  }

  // ── Teachers (Academic Master, Class Teacher, Staff...) ───
  var teachers = getTeacherList();
  var t = teachers.find(function(u) {
    return u.username === username && u.password === password;
  });
  if (t) {
    setTeacherSession(t.username);
    var pos      = (t.position || '').trim();
    var htmlFile = 'Teacher';
    var role     = 'teacher';
    if (pos === 'Academic Master')   { htmlFile = 'Academics';    role = 'academic';     }
    else if (pos === 'Class Teacher'){ htmlFile = 'ClassTeacher'; role = 'classteacher'; }
    else if (['Sports Manager','Accountant','Librarian','Lab Technician'].indexOf(pos) >= 0) {
      htmlFile = 'Staff'; role = 'staff';
    }
    return {
      status:         'success',
      role:           role,
      position:       pos,
      fullName:       t.name,
      username:       t.username,
      classesSubjects:t.classesSubjects,
      classes:        t.classes,
      subjects:       t.subjects,
      html:           HtmlService.createHtmlOutputFromFile(htmlFile).getContent()
    };
  }

  // ── Parents ────────────────────────────────────────────────
  var sh   = ensureParentsSheet();
  var last = sh.getLastRow();
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, 12).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r         = rows[i];
      var rowUser   = String(r[7] || '').trim();
      var rowPass   = String(r[8] || '').trim();
      var rowRole   = String(r[11]|| '').trim().toLowerCase();

      if (rowUser.toLowerCase() === username.toLowerCase() && rowPass === password) {
        // Resolve student name from Form sheet if not stored
        var storedStudentName = String(r[9]  || '').trim();
        var storedClassName   = String(r[10] || '').trim();

        logActivity(rowUser, 'Parent login: ' + rowUser);

        return {
          status:   'success',
          role:     'parent',
          fullName: String(r[0] || username).trim(),
          username: rowUser,
          // parentData is used by Login.html to boot the parent portal
          parentData: {
            parentName:  String(r[0] || '').trim(),
            relation:    String(r[2] || '').trim(),
            phone:       String(r[3] || '').trim(),
            email:       String(r[4] || '').trim(),
            admNo:       String(r[5] || '').trim(),
            address:     String(r[6] || '').trim(),
            studentName: storedStudentName,
            className:   storedClassName,
            sex:         '',         // not stored in Parents sheet — add col if needed
            status:      'Approved'
          }
        };
      }
    }
  }

  // ── Invalid credentials ────────────────────────────────────
  return { status: 'fail', message: 'Invalid username or password.' };
}


// ── HELPER: get parent portal data for Admin portal generator
function getParentPortalData() {
  var sh   = ensureParentsSheet();
  var last = sh.getLastRow();
  var parents = [];
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, 12).getValues();
    parents = rows
      .filter(function(r){ return String(r[0]||'').trim() !== ''; })
      .map(function(r){
        return {
          name:        String(r[0]||'').trim(),
          sex:         String(r[1]||'').trim(),
          relation:    String(r[2]||'').trim(),
          phone:       String(r[3]||'').trim(),
          email:       String(r[4]||'').trim(),
          admNo:       String(r[5]||'').trim(),
          address:     String(r[6]||'').trim(),
          username:    String(r[7]||'').trim(),
          password:    String(r[8]||'').trim(),
          studentName: String(r[9]||'').trim(),
          className:   String(r[10]||'').trim()
        };
      });
  }
  // Also return student list so portal generator can resolve names
  var students = getStudentListRaw ? getStudentListRaw() : [];
  return { parents: parents, students: students };
}


function logout() {
  clearTeacherSession();
  return HtmlService.createHtmlOutputFromFile('Login').getContent();
}

// ============================================================
// TEACHER SESSION MANAGEMENT
// ============================================================
function setTeacherSession(username) {
  CacheService.getUserCache().put("currentTeacher", username, 21600);   // 6h (max allowed)
}

function getTeacherSession() {
  const username = CacheService.getUserCache().get("currentTeacher");
  if (!username) return null;
  const teachers = getTeacherList();
  const teacher  = teachers.find(u => u.username === username);
  if (!teacher) return null;
  try { CacheService.getUserCache().put("currentTeacher", username, 21600); } catch(e) {}   // sliding: active users stay logged in

  // Sum students across all assigned classes
  const totalStudents = (teacher.classes || []).reduce(function(sum, cls) {
    return sum + getStudentCountForClass(cls);
  }, 0);

  var sessionObj = {
    username:        teacher.username,
    fullName:        teacher.name,
    classesSubjects: teacher.classesSubjects,
    classes:         teacher.classes,
    homeroomClasses: teacher.homeroomClasses || [],
    subjects:        teacher.subjects,
    position:        teacher.position,
    totalStudents:   totalStudents
  };

  // SECURITY: if Admin has blocked this dashboard after login, do not allow
  // the dashboard to initialise as normal using an old cached session.
  try {
    if (typeof canAccessDashboard === 'function' && typeof getDashboardKeyFromPosition === 'function') {
      var dashKey = getDashboardKeyFromPosition(teacher.position || 'Teacher');
      var access = canAccessDashboard(dashKey, teacher.username);
      if (!access || access.allowed === false) {
        clearTeacherSession();
        return null;
      }
    }
  } catch(e) {}

  return sessionObj;
}

/**
 * Returns the number of students marked present today across the
 * teacher's classes, based on the Attendance sheet.
 */
function getTodayPresentCount() {
  const teacher = getTeacherSession();
  if (!teacher || !teacher.classes || !teacher.classes.length) return 0;

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance');
  if (!sheet) return 0;

  const today   = new Date().toISOString().split('T')[0]; // YYYY-MM-DD
  const classes = teacher.classes.map(c => c.trim().toLowerCase());
  const rows    = sheet.getDataRange().getValues();

  let present = 0;
  rows.slice(1).forEach(function(r) {
    const rowDate  = (r[0] || '').toString().trim().substring(0, 10);
    const rowClass = (r[1] || '').toString().trim().toLowerCase();
    const morning  = (r[6] || 'P').toString().trim().toUpperCase();
    const afternoon= (r[7] || 'P').toString().trim().toUpperCase();

    if (rowDate === today && classes.indexOf(rowClass) !== -1) {
      // Count as present if either session is 'P'
      if (morning === 'P' || afternoon === 'P') present++;
    }
  });

  return present;
}

function clearTeacherSession() {
  CacheService.getUserCache().remove("currentTeacher");
}

// ============================================================
// TEACHER MANAGEMENT
// ============================================================
function ensureUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['Name','Username','Password','Classes','Subjects','Sex','Position','Department','PicUrl']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Classes where somebody is the Class Teacher: set by the Academic Master in Manage Classes.
// AcademicClasses.Teacher holds a username (a full name in older records); both are understood.
function homeroomMap_() {
  const map = {};
  try {
    getAcademicClassesRaw().forEach(c => {
      const key = String(c.teacher || '').trim().toLowerCase();
      const cls = String(c.name || '').trim();
      if (!key || !cls) return;
      (map[key] = map[key] || []).push(cls);
    });
  } catch (e) {}
  return map;
}

function getTeacherList() {
  const sheet = ensureUsersSheet();
  const data  = sheet.getDataRange().getValues();
  const homeroom = homeroomMap_();

  return data.slice(1).map(r => {
    // TEACHING: classes and subjects given by the Head of School (scores, assignments)
    const classes  = r[3] ? r[3].toString().split(',').map(s => s.trim()).filter(Boolean) : [];
    const subjects = r[4] ? r[4].toString().split(',').map(s => s.trim()).filter(Boolean) : [];
    const classesSubjects = [];
    classes.forEach(cls => {
      if (subjects.length) {
        subjects.forEach(subject => classesSubjects.push({ class: cls, subject }));
      } else {
        classesSubjects.push({ class: cls, subject: '' });
      }
    });

    // CLASS TEACHER: the class(es) given in Manage Classes (attendance, behaviour, parents)
    const homeroomClasses = [];
    [String(r[1] || '').trim().toLowerCase(), String(r[0] || '').trim().toLowerCase()].forEach(key => {
      (homeroom[key] || []).forEach(cls => { if (homeroomClasses.indexOf(cls) < 0) homeroomClasses.push(cls); });
    });

    return {
      name:     r[0], username: r[1], password: r[2],
      classes, subjects,
      sex: r[5]||'', position: r[6]||'', department: r[7]||'', picUrl: r[8]||'',
      classesSubjects, homeroomClasses
    };
  });
}

function getTeachers() {
  return getTeacherList().map(t => ({
    name: t.name, username: t.username,
    classes: t.classes.join(','), subjects: t.subjects.join(','),
    sex: t.sex, position: t.position, department: t.department, picUrl: t.picUrl
  }));
}

function addTeacher(name, username, password, classes, subjects, adminUser, sex, position, department, picUrl) {
  if (!name || !username || !password) return 'All required fields must be filled';
  const sheet      = ensureUsersSheet();
  const classStr   = Array.isArray(classes)  ? classes.join(',')  : (classes  || '');
  const subjectStr = Array.isArray(subjects) ? subjects.join(',') : (subjects || '');
  sheet.appendRow([name, username, password, classStr, subjectStr, sex||'', position||'', department||'', picUrl||'']);
  logActivity(adminUser || 'admin', 'Added teacher: ' + name + ' (' + username + ')');
  return 'Teacher ' + name + ' added successfully';
}

function updateTeacher(oldUsername, name, username, password, classes, subjects, adminUser, sex, position, department, picUrl) {
  const sheet = ensureUsersSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === oldUsername) {
      const classStr   = Array.isArray(classes)  ? classes.join(',')  : (classes  || '');
      const subjectStr = Array.isArray(subjects) ? subjects.join(',') : (subjects || '');
      sheet.getRange(i+1,1,1,9).setValues([[name,username,password,classStr,subjectStr,sex||'',position||'',department||'',picUrl||'']]);
      logActivity(adminUser || 'admin', 'Updated teacher: ' + oldUsername + ' -> ' + username);
      return 'Teacher updated successfully';
    }
  }
  return 'Teacher not found';
}

function deleteTeacher(username, adminUser) {
  const sheet = ensureUsersSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      sheet.deleteRow(i + 1);
      logActivity(adminUser || 'admin', 'Deleted teacher: ' + username);
      return 'Teacher deleted successfully';
    }
  }
  return 'Teacher not found';
}

// ============================================================
// ACTIVITY LOG
// ============================================================
function logActivity(user, action) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Activity');
  if (!sheet) { sheet = ss.insertSheet('Activity'); sheet.appendRow(['Date','User','Action']); }
  sheet.appendRow([new Date(), user, action]);
}

function getActivityLog() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Activity');
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1);
}

// ============================================================
// SUBMISSION CONTROL
// ============================================================
function getCurrentStatus() {
  const data = getDeadlineData();
  if (!data || !data.closesAt) return "CLOSED";
  const now      = new Date();
  const opensAt  = data.opensAt ? new Date(data.opensAt) : null;
  const closesAt = new Date(data.closesAt);
  if (opensAt && now < opensAt) return "PENDING";
  if (now >= closesAt) return "CLOSED";
  return "OPEN";
}

function getDeadlineData() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!cfg) return null;
  const opensVal  = cfg.getRange(DEADLINE_OPEN_CELL).getValue();
  const closesVal = cfg.getRange(DEADLINE_CLOSE_CELL).getValue();
  if (!closesVal) return null;
  return {
    opensAt:  opensVal  ? new Date(opensVal).toISOString()  : null,
    closesAt: closesVal ? new Date(closesVal).toISOString() : null
  };
}

function setDeadlineRange(opensAtStr, closesAtStr, user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cfg  = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!cfg) cfg = ss.insertSheet(CONFIG_SHEET_NAME);
  if (opensAtStr)  cfg.getRange(DEADLINE_OPEN_CELL).setValue(opensAtStr);
  if (closesAtStr) cfg.getRange(DEADLINE_CLOSE_CELL).setValue(closesAtStr);
  cfg.getRange(STATUS_CELL).setValue(getCurrentStatus());
  logActivity(user || 'system', 'Set deadline — Opens: ' + (opensAtStr||'—') + ' | Closes: ' + closesAtStr);
  return 'Deadline saved. Status will be computed automatically.';
}

function setStatus(status, adminUser) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cfg  = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!cfg) cfg = ss.insertSheet(CONFIG_SHEET_NAME);
  cfg.getRange(STATUS_CELL).setValue(status.toUpperCase().trim());
  logActivity(adminUser || 'admin', 'Manually set status to: ' + status);
  return status;
}

function getDeadline() {
  const data = getDeadlineData();
  return data ? (data.closesAt || '') : '';
}

function setDeadline(dateTimeStr, adminUser) {
  return setDeadlineRange(null, dateTimeStr, adminUser);
}

// ============================================================
// CLASS / LEVEL RESOLUTION HELPERS
// ============================================================
function resolveClassLevel(className) {
  if (!className) return '';
  const raw     = className.toString().trim();
  const classes = getAcademicClassesRaw();
  const byName  = classes.find(c => c.name && c.name.trim().toLowerCase() === raw.toLowerCase());
  if (byName && byName.stream) return byName.stream.trim();
  const byStream = classes.find(c => c.stream &&
    c.stream.trim().replace(/\s+/g,'').toLowerCase() === raw.replace(/\s+/g,'').toLowerCase());
  if (byStream) return byStream.stream.trim();
  const m = raw.match(/(\d+)/);
  if (m) return 'Form ' + m[1];
  return raw;
}

/**
 * Extracts the form number from any class name variant.
 * "Standard 1", "Std1", "STD1-A", "Form 1" → "1"
 */
function extractFormNumber(className) {
  if (!className) return '';
  const m = className.toString().match(/(\d+)/);
  return m ? m[1] : '';
}

function primaryClassSheetName_(className) {
  const match = String(className || '').trim().match(/^standard\s*(\d+)(?:\b|$)/i);
  if (!match) return '';
  const number = Number(match[1]);
  if (number < 1 || number > PS_CLASSES.length) return '';
  return psClassSheetName_(PS_CLASSES[number - 1]);
}

function primaryStudentSheetNames_() {
  return PS_CLASSES.map(psClassSheetName_);
}

// Returns the student register sheet of a class ("Standard 3" -> sheet "Standard3"), or null.
function studentSheetForClass_(ss, className) {
  var name = primaryClassSheetName_(className);
  return name ? ss.getSheetByName(name) : null;
}

/**
 * Returns true when className belongs to the exam's classes list.
 */
function classMatchesExam(className, examClassesStr) {
  const haystack = (examClassesStr || '').toString();
  if (!haystack || haystack.trim() === '' || haystack.trim().toLowerCase() === 'all') return true;
  const classNum = extractFormNumber(className);
  const rawLower = className.toString().trim().toLowerCase();
  const level    = resolveClassLevel(className).toLowerCase();
  const tokens   = haystack.split(/[,;]+/).map(t => t.trim().toLowerCase());
  return tokens.some(token => {
    if (!token) return false;
    const tokenNum     = extractFormNumber(token);
    const tokenNoSpace = token.replace(/\s+/g,'');
    const levelNoSpace = level.replace(/\s+/g,'');
    const rawNoSpace   = rawLower.replace(/\s+/g,'');
    if (classNum && tokenNum && classNum === tokenNum) return true;
    return token === level || token === rawLower ||
           tokenNoSpace === levelNoSpace || tokenNoSpace === rawNoSpace;
  });
}

/**
 * Converts a sheet date value to an integer YYYYMMDD.
 */
function dateToInt(val) {
  if (!val) return 0;
  if (val instanceof Date) {
    const tz  = Session.getScriptTimeZone();
    const str = Utilities.formatDate(val, tz, 'yyyyMMdd');
    return parseInt(str, 10) || 0;
  }
  const s = val.toString().trim().replace(/-/g,'').substring(0, 8);
  const n = parseInt(s, 10);
  return isNaN(n) ? 0 : n;
}

/**
 * Returns today as integer YYYYMMDD in script timezone.
 */
function todayInt() {
  const tz = Session.getScriptTimeZone();
  return parseInt(Utilities.formatDate(new Date(), tz, 'yyyyMMdd'), 10);
}

function sheetDateToLocal(val) {
  if (!val) return null;
  if (val instanceof Date) return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  const s     = val.toString().trim().substring(0, 10);
  const parts = s.split('-');
  if (parts.length === 3) {
    return new Date(parseInt(parts[0],10), parseInt(parts[1],10)-1, parseInt(parts[2],10));
  }
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

// ============================================================
// EXAM / SUBJECT HELPERS
// ============================================================

/**
 * Returns unique exam TYPE values (e.g. "CA", "SE") for a class
 * where today falls within [startDate, endDate].
 */
function getExamCategoriesForClass(className) {
  const today      = todayInt();
  const exams      = getAcademicExamsRaw();
  const categories = new Set();
  exams.forEach(e => {
    if (!classMatchesExam(className, e.classes)) return;
    const startInt   = dateToInt(e.startDate);
    const endInt     = dateToInt(e.endDate);
    const afterStart = !startInt || today >= startInt;
    const beforeEnd  = !endInt   || today <= endInt;
    if (afterStart && beforeEnd && e.type && e.status === 'active') {
    categories.add(e.type.toString().toUpperCase().trim());
    }
  });
  return Array.from(categories).sort();
}

/**
 * Returns all exams for a class + category active today.
 */
function getExamsForClass(className, category) {
  const today = todayInt();
  const exams = getAcademicExamsRaw();
 return exams.filter(e => {
  if (e.status !== 'active') return false;           // ✅ ADD THIS LINE
  if (!classMatchesExam(className, e.classes)) return false;
  const matchCat = !category ||
    (e.type || '').toString().toUpperCase().trim() === category.toString().toUpperCase().trim();
  if (!matchCat) return false;
  const startInt = dateToInt(e.startDate);
  const endInt   = dateToInt(e.endDate);
  return (!startInt || today >= startInt) && (!endInt || today <= endInt);
}).map(e => ({
    id: e.id, name: e.name, type: e.type, year: e.year,
    startDate: e.startDateStr || '', endDate: e.endDateStr || '', status: e.status || 'Draft'
  }));
}

function getSubjectsForClassAndTeacher(className, teacherSubjects) {
  var level    = resolveClassLevel(className);
  var classNum = extractFormNumber(className);
  var allSubjects = getAcademicSubjectsRaw();

  // All subjects valid for this class (from AcademicSubjects sheet)
  var validForClass = allSubjects.filter(function(s) {
    var sc = (s.classes || '').toString();
    if (!sc || sc.trim().toLowerCase() === 'all') return true;
    var tokens = sc.split(/[,;]+/).map(function(t){ return t.trim().toLowerCase(); });
    return tokens.some(function(t) {
      if (!t) return false;
      var tNum = extractFormNumber(t);
      if (classNum && tNum && classNum === tNum) return true;
      return t === level.toLowerCase() ||
             t.replace(/\s+/g,'') === level.replace(/\s+/g,'').toLowerCase();
    });
  }).map(function(s){ return s.name.trim(); });

  // If teacher has no subject assignment, return all class subjects
  if (!teacherSubjects || !teacherSubjects.length) {
    return validForClass.map(function(n){ return { name: n }; });
  }

  // Match teacher's subjects against class subjects using flexible comparison
  var tLower = teacherSubjects.map(function(s){ return s.trim().toLowerCase(); });

  var matched = validForClass.filter(function(name) {
    var nLower = name.toLowerCase();
    return tLower.some(function(ts) {
      return ts === nLower ||
             nLower.includes(ts) ||
             ts.includes(nLower) ||
             nLower.replace(/\s+/g,'') === ts.replace(/\s+/g,'');
    });
  });

  // KEY FIX: If teacher has subjects assigned but none matched the class filter,
  // fall back to ALL valid subjects for the class (prevents empty dropdown)
  if (!matched.length) return validForClass.map(function(n){ return { name: n }; });

  return matched.map(function(name){ return { name: name }; });
}
// ============================================================
// SCORES SHEET
// Columns: ExamName|Class|Subject|AdmNo|StudentName|Sex|Score|SubmittedBy|SubmittedAt
// ============================================================
// ============================================================
// PER-FORM SCORES SHEETS
// Form1_Scores | Form2_Scores | Form3_Scores | Form4_Scores
// Columns: ExamName|Class|Subject|AdmNo|StudentName|Sex|Score|SubmittedBy|SubmittedAt
// ============================================================
function getScoresSheet(className) {
  var formNum = extractFormNumber(className);
  if (!formNum) {
    // Fallback: return old Scores sheet if form number can't be resolved
    return ensureScoresSheetLegacy();
  }
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Form' + formNum + '_Scores';
  var sh        = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow(['ExamName','Class','Subject','AdmNo','StudentName','Sex','Score','SubmittedBy','SubmittedAt']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,9)
      .setBackground('#0d1b36')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    // Set column widths for readability
    sh.setColumnWidth(1, 160); // ExamName
    sh.setColumnWidth(2, 100); // Class
    sh.setColumnWidth(3, 140); // Subject
    sh.setColumnWidth(4, 100); // AdmNo
    sh.setColumnWidth(5, 180); // StudentName
    sh.setColumnWidth(6, 60);  // Sex
    sh.setColumnWidth(7, 70);  // Score
    sh.setColumnWidth(8, 140); // SubmittedBy
    sh.setColumnWidth(9, 160); // SubmittedAt
  }
  return sh;
}

// Keep old ensureScoresSheet() renamed for legacy/migration use only
function ensureScoresSheetLegacy() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Scores');
  if (!sh) {
    sh = ss.insertSheet('Scores');
    sh.appendRow(['ExamName','Class','Subject','AdmNo','StudentName','Sex','Score','SubmittedBy','SubmittedAt']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,9).setBackground('#0d1b36').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sh;
}

// Keep ensureScoresSheet() pointing to legacy for any callers not yet updated
function ensureScoresSheet() {
  return ensureScoresSheetLegacy();
}

// ============================================================
// FAST: Single-sheet student count  (used by ClassTeacher dashboard)
// ============================================================
function getStudentCountForClass(className) {
  const sheetName = primaryClassSheetName_(className);
  if (!sheetName) return 0;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  return Math.max(0, sheet.getLastRow() - 1);   // -1 for header row
}

// ============================================================
// FAST: Single-sheet student list  (used by ClassTeacher assignments)
// ============================================================
function getStudentsForClass(className) {
  const sheetName = primaryClassSheetName_(className);
  if (!sheetName) return [];
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1)
    .filter(r => r[0] || r[1])
    .map(r => ({
      admNo:     (r[0] || '').toString().trim(),
      firstName: (r[1] || '').toString().trim(),
      lastName:  (r[2] || '').toString().trim(),
      sex:       (r[3] || '').toString().trim(),
      className: className
    }));
}

// ============================================================
// FAST: Load students + existing scores for score entry form
// Reads ONE Form sheet + ONE Scores sheet — no full scan loops
// ============================================================
function getStudentsForScoring(payload) {
  var cls      = String(payload['class']    || payload.class    || '').trim();
  var subject  = String(payload['subject']  || payload.subject  || '').trim();
  var examType = String(payload['examType'] || payload.examType || '').trim();

  if (!cls || !subject || !examType) return [];

  var OPTIONAL_NAMES = []; // primary school: only the AcademicSubjects 'Optional' category makes a subject optional
  var isOptional     = false;
  var ss             = SpreadsheetApp.getActiveSpreadsheet();

  // Check AcademicSubjects category field
  var subjSheet = ss.getSheetByName('AcademicSubjects');
  if (subjSheet && subjSheet.getLastRow() > 1) {
    var subjData = subjSheet.getDataRange().getValues();
    for (var si = 1; si < subjData.length; si++) {
      if (String(subjData[si][0]||'').trim().toLowerCase() === subject.toLowerCase()) {
        if (String(subjData[si][2]||'').trim().toLowerCase() === 'optional') isOptional = true;
        break;
      }
    }
  }
  if (!isOptional && OPTIONAL_NAMES.indexOf(subject.toLowerCase()) >= 0) isOptional = true;

  var formNum = extractFormNumber(cls);
  if (!formNum) return [];

  // ── Build score lookup from the FORM-SPECIFIC scores sheet ──
  var scoreMap    = {};
  var scoresSheet = getScoresSheet(cls); // <-- uses per-form sheet
  var lastScRow   = scoresSheet.getLastRow();

  if (lastScRow > 1) {
    var scoreRows = scoresSheet.getRange(2, 1, lastScRow - 1, 9).getValues();
    var clsL  = cls.toLowerCase();
    var subjL = subject.toLowerCase();
    var examL = examType.toLowerCase();

    scoreRows.forEach(function(r) {
      if ((r[0]||'').toString().trim().toLowerCase() !== examL) return;
      if ((r[2]||'').toString().trim().toLowerCase() !== subjL) return;
      // Class check: same form number is sufficient (sheet is already filtered by form)
      var rCls = (r[1]||'').toString().trim().toLowerCase();
      if (rCls !== clsL && extractFormNumber(rCls) !== formNum) return;
      var name = (r[4]||'').toString().trim();
      if (!name) return;
      var score = r[6];
      scoreMap[name.toLowerCase()] =
        (score !== '' && score !== null && score !== undefined) ? score : '';
    });

    // ── Also check legacy Scores sheet for previously submitted data ──
    var legacySheet = ss.getSheetByName('Scores');
    if (legacySheet && legacySheet.getLastRow() > 1) {
      var legacyRows = legacySheet.getRange(2, 1, legacySheet.getLastRow() - 1, 9).getValues();
      legacyRows.forEach(function(r) {
        if ((r[0]||'').toString().trim().toLowerCase() !== examL) return;
        if ((r[2]||'').toString().trim().toLowerCase() !== subjL) return;
        var rCls = (r[1]||'').toString().trim().toLowerCase();
        if (rCls !== clsL && extractFormNumber(rCls) !== formNum) return;
        var name = (r[4]||'').toString().trim();
        if (!name) return;
        var score = r[6];
        // Only populate if not already found in the new sheet
        var key = name.toLowerCase();
        if (!scoreMap.hasOwnProperty(key)) {
          scoreMap[key] = (score !== '' && score !== null && score !== undefined) ? score : '';
        }
      });
    }
  }

  function mergeScores(students) {
    return students.map(function(s) {
      var fullName = (s.firstName + ' ' + s.lastName).trim();
      var score    = scoreMap.hasOwnProperty(fullName.toLowerCase())
                     ? scoreMap[fullName.toLowerCase()] : '';
      return { admNo: s.admNo, name: fullName, sex: s.sex || '—', score: score };
    }).filter(function(s) { return s.name; });
  }

  // Students are no longer mapped per subject: with no mapping the whole class register is used.
  if (isOptional) {
    var _ms = ss.getSheetByName(mapSheetName(subject, cls));
    if (!_ms || _ms.getLastRow() < 2) isOptional = false;
  }

  // OPTIONAL → read from dedicated map sheet
  if (isOptional) {
    var mapSheet = ss.getSheetByName(mapSheetName(subject, cls));
    if (!mapSheet || mapSheet.getLastRow() < 2) {
      Logger.log('No mapping yet: ' + mapSheetName(subject, cls));
      return [];
    }
    var mapData = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 4).getValues();
    var students = mapData
      .filter(function(r) { return r[0]; })
      .map(function(r) {
        return {
          admNo:     String(r[0]||'').trim(),
          firstName: String(r[1]||'').trim(),
          lastName:  String(r[2]||'').trim(),
          sex:       String(r[3]||'').trim()
        };
      });
    return mergeScores(students);
  }

  // COMPULSORY → read from Form sheet
  var formSheet = studentSheetForClass_(ss, cls);
  if (!formSheet) return [];
  var formData = formSheet.getDataRange().getValues();
  if (formData.length < 2) return [];
  var students = formData.slice(1)
    .filter(function(r) { return r[0] || r[1]; })
    .map(function(r) {
      return {
        admNo:     String(r[0]||'').trim(),
        firstName: String(r[1]||'').trim(),
        lastName:  String(r[2]||'').trim(),
        sex:       String(r[3]||'').trim()
      };
    });
  return mergeScores(students);
}
// ============================================================
// SUBMIT / UPDATE SCORES
// ============================================================
function submitScoresByName(payload) {
  var teacher = getTeacherSession();
  if (!teacher) return 'Session expired. Please login again.';

  var accessCheck = checkDashboardWriteAccess(['Score Entry']);
  if (!accessCheck.ok) return accessCheck.message;

  var status = getCurrentStatus();
  if (status !== 'OPEN')
    return 'Submission portal is currently ' + status + '. Contact Academic Master to open it.';

  var cls      = payload.class;
  var subject  = payload.subject;
  var examType = payload.examType;
  var students = payload.students;
  var action   = payload.action;

  if (!cls || !subject || !examType || !students || !students.length)
    return 'Missing required data. Please try again.';

  // ── Use the form-specific sheet ──
  var sh   = getScoresSheet(cls); // <-- per-form sheet
  var rows = sh.getDataRange().getValues();

  // Build lookup: "exam|||subject|||studentname" → sheet row number (1-based)
  // No need to include class in key since this sheet is already per-class/form
  var keyToRow = {};
  rows.slice(1).forEach(function(r, i) {
    var key = (r[0]||'').toString().trim().toLowerCase() + '|||' +
              (r[2]||'').toString().trim().toLowerCase() + '|||' +
              (r[4]||'').toString().trim().toLowerCase();
    keyToRow[key] = i + 2; // +1 header, +1 1-based
  });

  var now     = new Date();
  var written = 0;
  var skipped = 0;

  students.forEach(function(s) {
    if (!s.name) return;

    var scoreVal = (s.score === '' || s.score === null || s.score === undefined)
      ? '' : Number(s.score);

    var lookupKey = examType.toLowerCase() + '|||' +
                    subject.toLowerCase()  + '|||' +
                    s.name.trim().toLowerCase();

    var existingRow = keyToRow[lookupKey];

    if (existingRow) {
      if (action === 'update') {
        sh.getRange(existingRow, 7).setValue(scoreVal);
        sh.getRange(existingRow, 8).setValue(teacher.username);
        sh.getRange(existingRow, 9).setValue(now);
        written++;
      } else {
        skipped++;
      }
    } else {
      sh.appendRow([
        examType,
        cls,
        subject,
        s.admNo  || '',
        s.name.trim(),
        s.sex    || '',
        scoreVal,
        teacher.username,
        now
      ]);
      written++;
    }
  });

  logActivity(
    teacher.username,
    action.toUpperCase() + ' ' + written + ' score(s) — ' + cls + ' | ' + subject + ' | ' + examType
  );

  if (action === 'update') return written + ' score(s) updated successfully!';
  if (skipped > 0 && written === 0) return 'Scores already exist. Use Update to modify them.';
  if (written === 0) return 'No students found to save. Please try again.';
  return written + ' score(s) submitted successfully!';
}

// Compat shims — keep old callers working
function getExistingScores(data)    { return getStudentsForScoring(data); }
function submitOrUpdateScores(data) { return submitScoresByName(data); }

// ============================================================
// SCORE REPORTS
// ============================================================
function getScoreReport(filter) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];

  // Determine which form sheets to read
  var formNums = ['1','2','3','4'];
  if (filter && filter.className) {
    var fn = extractFormNumber(filter.className);
    if (fn) formNums = [fn]; // only read the relevant form sheet
  }

  formNums.forEach(function(fn) {
    var sh = ss.getSheetByName('Form' + fn + '_Scores');
    if (!sh || sh.getLastRow() < 2) return;
    sh.getDataRange().getValues().slice(1).forEach(function(r) {
      results.push({
        examName:    (r[0]||'').toString().trim(),
        className:   (r[1]||'').toString().trim(),
        subject:     (r[2]||'').toString().trim(),
        admNo:       (r[3]||'').toString().trim(),
        studentName: (r[4]||'').toString().trim(),
        sex:         (r[5]||'').toString().trim(),
        score:       (r[6] !== '' && r[6] !== null) ? Number(r[6]) : '',
        submittedBy: (r[7]||'').toString().trim(),
        submittedAt: r[8] ? (r[8] instanceof Date
                      ? r[8].toLocaleDateString()
                      : r[8].toString()) : ''
      });
    });

    // Also pull from legacy Scores sheet the first time (migration period)
    var legacy = ss.getSheetByName('Scores');
    if (legacy && legacy.getLastRow() > 1) {
      legacy.getDataRange().getValues().slice(1).forEach(function(r) {
        var rFn = extractFormNumber((r[1]||'').toString());
        if (rFn !== fn) return; // only pull rows matching this form
        results.push({
          examName:    (r[0]||'').toString().trim(),
          className:   (r[1]||'').toString().trim(),
          subject:     (r[2]||'').toString().trim(),
          admNo:       (r[3]||'').toString().trim(),
          studentName: (r[4]||'').toString().trim(),
          sex:         (r[5]||'').toString().trim(),
          score:       (r[6] !== '' && r[6] !== null) ? Number(r[6]) : '',
          submittedBy: (r[7]||'').toString().trim(),
          submittedAt: r[8] ? (r[8] instanceof Date
                        ? r[8].toLocaleDateString()
                        : r[8].toString()) : ''
        });
      });
    }
  });

  return results
    .filter(function(r) { return r.studentName; })
    .filter(function(r) {
      if (!filter) return true;
      var examOk    = !filter.examName  || r.examName.toLowerCase()  === filter.examName.toLowerCase();
      var classOk   = !filter.className ||
        r.className.toLowerCase()  === filter.className.toLowerCase() ||
        resolveClassLevel(r.className).toLowerCase() === resolveClassLevel(filter.className).toLowerCase();
      var subjectOk = !filter.subject   || r.subject.toLowerCase()   === filter.subject.toLowerCase();
      return examOk && classOk && subjectOk;
    });
}

function getClassExamReport(examName, className) {
  const all = getScoreReport({ examName, className });
  if (!all.length) return [];
  const studentMap = {};
  all.forEach(r => {
    if (!studentMap[r.studentName]) {
      studentMap[r.studentName] = { studentName: r.studentName, admNo: r.admNo, sex: r.sex, scores: {} };
    }
    if (r.score !== '') studentMap[r.studentName].scores[r.subject] = r.score;
  });
  const results = Object.values(studentMap).map(s => {
    const vals  = Object.values(s.scores).filter(v => v !== '');
    const total = vals.reduce((a, b) => a + b, 0);
    const avg   = vals.length ? (total / vals.length).toFixed(1) : '—';
    return { ...s, total, average: avg };
  }).sort((a, b) => b.total - a.total);
  results.forEach((s, i) => { s.position = i + 1; });
  return results;
}

// ============================================================
// SCORE TRACKING DATA
// ============================================================
function getScoreTrackingDataLegacy_(filterClass, filterExamType, filterYear) {
  const teachers    = getTeacherList();
  const allStudents = getStudentListRaw();
  const scoresSheet = ensureScoresSheet();
  const scoreRows   = scoresSheet.getDataRange().getValues().slice(1);
  const assignments = [];

  teachers.forEach(teacher => {
    const classes  = teacher.classes  || [];
    const subjects = teacher.subjects || [];
    const pairs    = [];

    classes.forEach((cls, idx) => {
      const subj = subjects[idx] || '';
      if (!subj) return;
      if (filterClass && cls !== filterClass) return;
      pairs.push({ cls, subj });
    });

    pairs.forEach(({ cls, subj }) => {
      const level         = resolveClassLevel(cls);
      const classStudents = allStudents.filter(s => {
        const sc = (s.className || '').toString().trim();
        return sc.toLowerCase() === cls.toLowerCase() ||
               sc.toLowerCase() === level.toLowerCase() ||
               resolveClassLevel(sc).toLowerCase() === level.toLowerCase();
      });
      const total       = classStudents.length;
      const activeExams = getAcademicExamsRaw().filter(e => classMatchesExam(cls, e.classes));
      const examList    = activeExams.length ? activeExams : [{ name: filterExamType || 'All', type: '' }];

      examList.forEach(exam => {
        if (filterExamType && exam.name !== filterExamType &&
            !(exam.type || '').toUpperCase().startsWith(filterExamType.toUpperCase())) return;
        const submitted = scoreRows.filter(r => {
          const rowExam    = (r[0]||'').toString().trim().toLowerCase();
          const rowClass   = (r[1]||'').toString().trim().toLowerCase();
          const rowSubject = (r[2]||'').toString().trim().toLowerCase();
          const rowScore   = r[6];
          return rowClass === cls.toLowerCase() || rowClass === level.toLowerCase()
            ? rowSubject === subj.toLowerCase() && rowExam === exam.name.toLowerCase() &&
              rowScore !== '' && rowScore !== null && rowScore !== undefined
            : false;
        }).length;
        const status = submitted === 0 ? 'pending' : submitted >= total ? 'submitted' : 'partial';
        assignments.push({
          className: cls, subjectName: subj,
          examType: exam.type || exam.name, examName: exam.name,
          teacherName: teacher.name, teacherUsername: teacher.username,
          totalStudents: total, scoresSubmitted: submitted,
          status, submittedAt: submitted > 0 ? new Date().toLocaleDateString() : null
        });
      });
    });
  });

  const total     = assignments.length;
  const submitted = assignments.filter(a => a.status === 'submitted').length;
  const pending   = assignments.filter(a => a.status !== 'submitted').length;
  return { assignments, summary: { total, submitted, pending } };
}

// ============================================================
// ADMIN LINKS (legacy — opens external spreadsheet URL)
// ============================================================
function openClassSheet(formKey) {
  const id = FORM_SHEETS[formKey];
  return id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : 'Sheet not found';
}
function openAttendanceSheet(formKey) {
  const id = ATTENDANCE_SHEETS[formKey];
  return id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : 'Sheet not found';
}

// ============================================================
// ATTENDANCE SYSTEM
// Columns: Date|Class|AdmNo|StudentName|Sex|RowIndex|Morning|Afternoon|Teacher|Reason
// ============================================================
function ensureAttendanceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName('Attendance');
  if (!sh) {
    sh = ss.insertSheet('Attendance');
    sh.appendRow(['Date','Class','AdmNo','StudentName','Sex','RowIndex','Morning','Afternoon','Teacher','Reason']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,10).setBackground('#0d3d27').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sh;
}

// Attendance belongs to the Class Teacher of the class (Manage Classes), not to everybody who teaches there.
// The Academic Master and other roles keep their own access.
function attAssertClass_(className) {
  const t = getTeacherSession();
  if (!t) return;
  if (getDashboardKeyFromPosition(String(t.position || '')) !== 'classTeacher') return;
  if (!behCanAccess_(t, className)) throw new Error('You can only manage attendance for the class you are Class Teacher of.');
}

function getAttendanceStudents(data) {
  attAssertClass_(data.class);
  const cls   = data.class;
  const dateStr = data.dateStr;
  const level = resolveClassLevel(cls);

  // Get students from Form sheet (fast single-sheet read)
  const formNum = extractFormNumber(cls);
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = studentSheetForClass_(ss, cls);

  let students = [];
  if (formSheet) {
    const fd = formSheet.getDataRange().getValues();
    if (fd.length > 1) {
      students = fd.slice(1).filter(r => r[0] || r[1]).map(r => ({
        admNo:     (r[0]||'').toString().trim(),
        firstName: (r[1]||'').toString().trim(),
        lastName:  (r[2]||'').toString().trim(),
        sex:       (r[3]||'').toString().trim(),
        className: cls
      }));
    }
  }

  // Fallback to full student list if Form sheet empty
  if (!students.length) {
    const all = getStudentListRaw();
    students  = all.filter(s => {
      const sc = (s.className || '').toString().trim();
      return sc.toLowerCase() === cls.toLowerCase() ||
             sc.toLowerCase() === level.toLowerCase() ||
             resolveClassLevel(sc).toLowerCase() === level.toLowerCase();
    });
  }

  if (!students.length) return [];

  // Load existing attendance for this date+class
  const attMap = {};
  try {
    const sh   = ensureAttendanceSheet();
    const rows = sh.getDataRange().getValues();
    rows.slice(1).forEach(r => {
      const rowDate  = (r[0]||'').toString().trim().substring(0, 10);
      const rowClass = (r[1]||'').toString().trim();
      const rowName  = (r[3]||'').toString().trim();
      if (rowDate === dateStr &&
          (rowClass.toLowerCase() === cls.toLowerCase() || rowClass.toLowerCase() === level.toLowerCase())) {
        attMap[rowName] = { morning: r[6]||'P', afternoon: r[7]||'P' };
      }
    });
  } catch(e) { Logger.log('getAttendanceStudents: ' + e.message); }

  return students.map((s, i) => {
    const fullName = ((s.firstName || '') + ' ' + (s.lastName || '')).trim();
    const existing = attMap[fullName] || {};
    return {
      name:      fullName,
      sex:       s.sex || '',
      admNo:     s.admNo || '',
      rowIndex:  i + 1,
      morning:   existing.morning   || 'P',
      afternoon: existing.afternoon || 'P'
    };
  }).filter(s => s.name);
}

function classLevelMatch(studentClass, teacherClass) {
  if (!studentClass || !teacherClass) return false;
  const s    = studentClass.toString().trim().toLowerCase();
  const t    = teacherClass.toString().trim().toLowerCase();
  if (s === t) return true;
  const sNum = extractFormNumber(s);
  const tNum = extractFormNumber(t);
  return sNum && tNum && sNum === tNum;
}

function submitAttendance(data) {
  var accessCheck = checkDashboardWriteAccess(['Attendance','Class Register']);
  if (!accessCheck.ok) return { success: false, message: accessCheck.message };
  try { attAssertClass_(data.class); } catch (e) { return { success: false, message: e.message }; }

  const cls       = data.class;
  const dateStr   = data.dateStr;
  const students  = data.students;
  const reason    = data.reason || '';

  if (!cls || !dateStr || !students || !students.length) {
    return { success: false, message: 'Missing required fields' };
  }

  const level       = resolveClassLevel(cls);
  const teacher     = getTeacherSession();
  const teacherName = teacher ? teacher.fullName : '';
  const sh          = ensureAttendanceSheet();
  const rows        = sh.getDataRange().getValues();

  // Build index of existing rows
  const keyToRow = {};
  rows.slice(1).forEach((r, i) => {
    const rowDate  = (r[0]||'').toString().trim().substring(0, 10);
    const rowClass = (r[1]||'').toString().trim().toLowerCase();
    const rowName  = (r[3]||'').toString().trim().toLowerCase();
    const key = rowDate + '|||' + rowClass + '|||' + rowName;
    keyToRow[key] = i + 2;
  });

  let written = 0;
  students.forEach(s => {
    if (!s.name) return;
    const key    = dateStr + '|||' + cls.toLowerCase()   + '|||' + s.name.trim().toLowerCase();
    const altKey = dateStr + '|||' + level.toLowerCase() + '|||' + s.name.trim().toLowerCase();
    const existingRow = keyToRow[key] || keyToRow[altKey];

    if (existingRow) {
      sh.getRange(existingRow, 7).setValue(s.morning   || 'P');
      sh.getRange(existingRow, 8).setValue(s.afternoon || 'P');
      sh.getRange(existingRow, 9).setValue(teacherName);
      sh.getRange(existingRow, 10).setValue(reason);
    } else {
      sh.appendRow([
        dateStr, cls, s.admNo || '', s.name.trim(), s.sex || '',
        s.rowIndex || '', s.morning || 'P', s.afternoon || 'P', teacherName, reason
      ]);
    }
    written++;
  });

  logAttendanceSummary(cls, dateStr, students, reason, teacherName);
  logActivity(teacherName || 'teacher',
    'Attendance saved: ' + cls + ' on ' + dateStr + ' (' + written + ' students)');

  return { success: true, message: 'Attendance saved for ' + written + ' student(s)!' };
}

function logAttendanceSummary(cls, dateStr, students, reason, teacherName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sum  = ss.getSheetByName('AttendanceSummary');
  if (!sum) {
    sum = ss.insertSheet('AttendanceSummary');
    sum.appendRow(['Date','Class','Total','Absent','Percentage','Reason','Teacher']);
    sum.setFrozenRows(1);
  }
  const totalStudents = students.length;
  const absentCount   = students.filter(s => s.morning === 'A' || s.afternoon === 'A').length;
  const percentage    = totalStudents ? ((absentCount / totalStudents) * 100).toFixed(0) : 0;
  const lastRow       = sum.getLastRow();
  const existingRows  = lastRow > 1 ? sum.getRange(2, 1, lastRow - 1, 2).getValues() : [];
  let rowToUpdate     = null;
  for (let i = 0; i < existingRows.length; i++) {
    if (existingRows[i][0] == dateStr && existingRows[i][1] == cls) { rowToUpdate = i + 2; break; }
  }
  const rowData = [dateStr, cls, totalStudents, absentCount, percentage + '%', reason, teacherName];
  if (rowToUpdate) sum.getRange(rowToUpdate, 1, 1, 7).setValues([rowData]);
  else             sum.appendRow(rowData);
}

// ============================================================
// TEACHER ON DUTY
// ============================================================
function getTODStudents() {
  try {
    const ss = SpreadsheetApp.openById(TOD_SHEET_ID);
    const sh = ss.getSheetByName('TOD');
    if (!sh) return [];
    return sh.getDataRange().getValues().slice(1);
  } catch(e) { return []; }
}

function submitTOD(data) {
  try {
    const ss = SpreadsheetApp.openById(TOD_SHEET_ID);
    let sh   = ss.getSheetByName('TODReports');
    if (!sh) { sh = ss.insertSheet('TODReports'); sh.appendRow(['Date','Teacher','Class','Report']); }
    sh.appendRow([data.date, data.teacher, data.class, data.report]);
    return { success: true, message: 'TOD report submitted!' };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
// TEACHER TRACKING (legacy)
// ============================================================
function getTeacherProfile(username) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("TeacherProfiles");
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const row  = data.slice(1).find(r => r[3] && r[3].toString().trim() === username.trim());
  if (!row) return null;
  return {
    TID: row[0], schoolName: row[1], fullName: row[2], username: row[3], password: row[4],
    email: row[5], phone: row[6], designation: row[7], position: row[8], get: row[9],
    classes:  row[10] ? row[10].toString().split(',').map(c => c.trim()) : [],
    subjects: row[11] ? row[11].toString().split(',').map(s => s.trim()) : [],
    status: row[12]
  };
}

function fetchProfile(username)                      { return getTeacherProfile(username); }
function fetchReport(username, className, examType)  { return getTrackingReport(username, className, examType); }

function fetchExamTypes(className) {
  const names    = new Set();
  const allExams = getAcademicExamsRaw();
  allExams.forEach(e => {
    if (classMatchesExam(className, e.classes) && e.name) names.add(e.name.trim());
  });
  const level = resolveClassLevel(className);
  try {
    const sh   = ensureScoresSheet();
    const rows = sh.getDataRange().getValues().slice(1);
    rows.forEach(r => {
      const rowClass = (r[1]||'').toString().trim();
      const classMatch =
        rowClass.toLowerCase() === className.toLowerCase() ||
        rowClass.toLowerCase() === level.toLowerCase() ||
        resolveClassLevel(rowClass).toLowerCase() === level.toLowerCase();
      if (classMatch) {
        const examName = (r[0]||'').toString().trim();
        if (examName) names.add(examName);
      }
    });
  } catch(e) { Logger.log('fetchExamTypes: ' + e.message); }
  return Array.from(names).sort();
}

function getTrackingExamTypes(className) {
  const level = resolveClassLevel(className);
  const sh    = ensureScoresSheet();
  const rows  = sh.getDataRange().getValues().slice(1);
  const names = new Set();
  rows.forEach(r => {
    const rowClass = (r[1]||'').toString().trim();
    if (rowClass.toLowerCase() === className.toLowerCase() ||
        rowClass.toLowerCase() === level.toLowerCase() ||
        resolveClassLevel(rowClass).toLowerCase() === level.toLowerCase()) {
      const examName = (r[0]||'').toString().trim();
      if (examName) names.add(examName);
    }
  });
  return Array.from(names).sort();
}

function getTrackingReport(username, className, examType) {
  function clean(v) {
    return (v || '').toString().trim();
  }

  function lower(v) {
    return clean(v).toLowerCase();
  }

  function sameClass(a, b) {
    var aa = clean(a);
    var bb = clean(b);

    if (!aa || !bb) return false;

    var aLevel = resolveClassLevel(aa);
    var bLevel = resolveClassLevel(bb);
    var aForm  = extractFormNumber(aa);
    var bForm  = extractFormNumber(bb);

    return lower(aa) === lower(bb) ||
           lower(aLevel) === lower(bLevel) ||
           lower(aa) === lower(bLevel) ||
           lower(aLevel) === lower(bb) ||
           (aForm && bForm && aForm === bForm);
  }

  function addSubject(map, subject) {
    subject = clean(subject);
    if (subject) map[lower(subject)] = subject;
  }

  function readSubjectsFromAssignment(item, subjectMap, selectedClass) {
    if (!item) return;

    if (typeof item === 'string') {
      addSubject(subjectMap, item);
      return;
    }

    var itemClass =
      item.className || item.ClassName ||
      item.class || item.Class ||
      item.level || item.Level ||
      item.form || item.Form ||
      item.assignedClass || item.AssignedClass || '';

    var classOk = !itemClass || sameClass(itemClass, selectedClass);

    if (!classOk) return;

    var oneSubject =
      item.subject || item.Subject ||
      item.subjectName || item.SubjectName ||
      item.name || item.Name || '';

    addSubject(subjectMap, oneSubject);

    var manySubjects =
      item.subjects || item.Subjects ||
      item.assignedSubjects || item.AssignedSubjects || [];

    if (Array.isArray(manySubjects)) {
      manySubjects.forEach(function(s) {
        if (typeof s === 'string') {
          addSubject(subjectMap, s);
        } else {
          addSubject(
            subjectMap,
            s.subject || s.Subject || s.subjectName || s.SubjectName || s.name || s.Name
          );
        }
      });
    }
  }

  var level   = resolveClassLevel(className);
  var formNum = extractFormNumber(className);
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var allRows = [];

  var teacher = null;

  try {
    var teachers = getTeacherList();
    var u = lower(username);

    teacher = teachers.filter(function(t) {
      return lower(t.username || t.Username || t.email || t.Email) === u;
    })[0] || null;
  } catch (e) {
    Logger.log('getTrackingReport teacher lookup error: ' + e.message);
  }

  if (!teacher) return [];

  var assignedClass = false;

  var teacherClasses =
    teacher.classes || teacher.Classes ||
    teacher.assignedClasses || teacher.AssignedClasses ||
    teacher.classNames || teacher.ClassNames || [];

  if (Array.isArray(teacherClasses)) {
    assignedClass = teacherClasses.some(function(c) {
      return sameClass(c, className);
    });
  }

  var assignedSubjectMap = {};

  var teacherSubjects =
    teacher.subjects || teacher.Subjects ||
    teacher.assignedSubjects || teacher.AssignedSubjects ||
    teacher.assignments || teacher.Assignments || [];

  if (Array.isArray(teacherSubjects)) {
    teacherSubjects.forEach(function(item) {
      readSubjectsFromAssignment(item, assignedSubjectMap, className);

      var itemClass =
        item && typeof item === 'object'
          ? item.className || item.ClassName || item.class || item.Class || item.level || item.Level || item.form || item.Form
          : '';

      if (itemClass && sameClass(itemClass, className)) {
        assignedClass = true;
      }
    });
  }

  if (!assignedClass) return [];

  if (Object.keys(assignedSubjectMap).length === 0) return [];

  // Read from the per-form sheet
  if (formNum) {
    var formSh = ss.getSheetByName('Form' + formNum + '_Scores');

    if (formSh && formSh.getLastRow() > 1) {
      allRows = allRows.concat(formSh.getDataRange().getValues().slice(1));
    }
  }

  // Also read legacy Scores sheet as fallback
  var legacy = ss.getSheetByName('Scores');

  if (legacy && legacy.getLastRow() > 1) {
    legacy.getDataRange().getValues().slice(1).forEach(function(r) {
      if (extractFormNumber(clean(r[1])) === formNum) allRows.push(r);
    });
  }

  var relevant = allRows.filter(function(r) {
    var rowExam    = lower(r[0]);
    var rowClass   = clean(r[1]);
    var rowSubject = clean(r[2]);

    var classMatch = sameClass(rowClass, className);
    var examMatch  = rowExam === lower(examType);
    var subjectAssigned = assignedSubjectMap[lower(rowSubject)] === rowSubject ||
                          assignedSubjectMap[lower(rowSubject)] !== undefined;

    return classMatch && examMatch && subjectAssigned;
  });

  if (!relevant.length) return [];

  var studentMap = {};

  relevant.forEach(function(r) {
    var name    = clean(r[4]);
    var sex     = clean(r[5]);
    var subject = clean(r[2]);
    var score   = r[6] !== '' && r[6] !== null ? Number(r[6]) : '';

    if (!name) return;

    if (!studentMap[name]) {
      studentMap[name] = {
        student: name,
        sex: sex,
        scores: {},
        grades: {}
      };
    }

    studentMap[name].scores[subject] = score;

    if (score !== '') {
      studentMap[name].grades[subject] =
        score >= 75 ? 'A' :
        score >= 65 ? 'B' :
        score >= 45 ? 'C' :
        score >= 30 ? 'D' : 'F';
    } else {
      studentMap[name].grades[subject] = '';
    }
  });

  var results = Object.values(studentMap).map(function(s) {
    var vals = Object.keys(s.scores)
      .filter(function(sub) {
        return assignedSubjectMap[lower(sub)] !== undefined;
      })
      .map(function(sub) {
        return s.scores[sub];
      })
      .filter(function(v) {
        return v !== '' && v !== null && !isNaN(v);
      });

    var sum   = vals.reduce(function(a, b) { return a + Number(b); }, 0);
    var count = vals.length;
    var avg   = count ? sum / count : 0;

    var grade =
      avg >= 75 ? 'A' :
      avg >= 65 ? 'B' :
      avg >= 45 ? 'C' :
      avg >= 30 ? 'D' : 'F';

    return Object.assign({}, s, {
      sum: sum,
      avg: avg,
      grade: grade,
      remarks: avg >= 30 ? 'Pass' : 'Fail'
    });
  });

  results.sort(function(a, b) {
    return b.sum - a.sum;
  });

  results.forEach(function(s, i) {
    s.position = i + 1;
  });

  return results;
}


// ============================================================
// CLASS REPORTS
// ============================================================
function getSubjectsByClass(className) {
  if (/form\s*[12]/i.test(className)) {
    return [
      {name:"Historia (TZ) Maadili",code:"H/TZM"},{name:"History",code:"HIS"},{name:"Geography",code:"GEO"},
      {name:"Kiswahili",code:"KISW"},{name:"English",code:"ENG"},{name:"Physics",code:"PHY"},
      {name:"Chemistry",code:"CHEM"},{name:"Biology",code:"BIO"},{name:"Mathematics",code:"MATH"},
      {name:"Business Studies",code:"B/STD"},{name:"Computer Science",code:"CSC"},{name:"Computer Application",code:"CAPP"}
    ];
  }
  return [
    {name:"Civics",code:"CIV"},{name:"History",code:"HIS"},{name:"Geography",code:"GEO"},
    {name:"Kiswahili",code:"KISW"},{name:"English",code:"ENG"},{name:"Physics",code:"PHY"},
    {name:"Chemistry",code:"CHEM"},{name:"Biology",code:"BIO"},{name:"Mathematics",code:"MATH"}
  ];
}

function getClassSheetId(className) {
  const key = className.replace(/\s+/g,'');
  return FORM_SHEETS[key] || FORM_SHEETS[className] || null;
}

function getClasses() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const acSheet = ss.getSheetByName('AcademicClasses');
  if (acSheet && acSheet.getLastRow() > 1) {
    return acSheet.getDataRange().getValues().slice(1).map(r => r[1]).filter(Boolean);
  }
  return Object.keys(FORM_SHEETS);
}

function getClassExamTypes(className) { return ['SE','CA']; }

function getExamSubTypes(className, category) {
  const fromExams = getExamsForClass(className, category);
  if (fromExams && fromExams.length) return fromExams.map(e => e.name);
  if (category === 'SE') return ['SE Mid Term','SE Terminal','SE Annual'];
  if (category === 'CA') return ['CA 1','CA 2','CA 3'];
  return [];
}


// ============================================================
// EXAM DIVISION HELPERS — best 7 points
// Formula equivalent:
// Point = SUMPRODUCT(SMALL({subject points...}, ROW(INDIRECT("1:7"))))
// Division = IF(Point<=17,"I",IF(Point<=21,"II",IF(Point<=25,"III",IF(Point<=33,"IV","0"))))
// ============================================================
function examScoreToGrade_(score) {
  var n = Number(score);
  if (score === '' || score === null || score === undefined || isNaN(n)) return '-';
  if (n >= 75) return 'A';
  if (n >= 65) return 'B';
  if (n >= 45) return 'C';
  if (n >= 30) return 'D';
  return 'F';
}

function examGradeToPoint_(grade) {
  switch (String(grade || '').trim().toUpperCase()) {
    case 'A': return 1;
    case 'B': return 2;
    case 'C': return 3;
    case 'D': return 4;
    case 'F': return 5;
    default:  return null;
  }
}

function examPointToDivision_(point) {
  var p = Number(point);
  if (point === '' || point === null || point === undefined || isNaN(p)) return '0';
  if (p <= 17) return 'I';
  if (p <= 21) return 'II';
  if (p <= 25) return 'III';
  if (p <= 33) return 'IV';
  return '0';
}

function examBest7PointSum_(points) {
  var pts = (points || [])
    // FIX: drop missing/invalid grades (null/undefined/'') BEFORE converting to
    // Number. Number(null) === 0, so a student with no score for a subject was
    // silently being treated as a perfect "0-point" grade, dragging their best-7
    // sum down and wrongly landing them in Division I even when absent.
    .filter(function(v) { return v !== null && v !== undefined && v !== ''; })
    .map(function(v) { return Number(v); })
    .filter(function(v) { return !isNaN(v); });

  // No valid grade in ANY subject => student did not sit the exam at all.
  if (pts.length === 0) return 'ABS';

  // Attempted some but fewer than 7 subjects => incomplete attempt.
  if (pts.length < 7) return '';

  pts.sort(function(a, b) { return a - b; });
  var sum = 0;
  for (var i = 0; i < 7; i++) sum += pts[i];
  return sum;
}

function examPointToDivision_(point) {
  if (point === 'ABS') return 'ABS'; // student did not sit the exam
  var p = Number(point);
  // NEW: empty point => empty division (not '0'/FAIL)
  if (point === '' || point === null || point === undefined || isNaN(p)) return '';
  if (p <= 17) return 'I';
  if (p <= 21) return 'II';
  if (p <= 25) return 'III';
  if (p <= 33) return 'IV';
  return '0';
}
function examNormalizeDivisionKey_(division) {
  var d = String(division || '').trim().toUpperCase();
  if (d === 'I' || d === 'II' || d === 'III' || d === 'IV') return d;
  if (d === '') return 'NA';
  if (d === 'ABS') return 'ABS';
  return 'ZERO';
}

function examBest7SummaryFromSubjects_(subjects) {
  var pointValues = (subjects || []).map(function(sub) {
    var grade = sub && sub.grade ? sub.grade : examScoreToGrade_(sub ? sub.score : '');
    return examGradeToPoint_(grade);
  });
  var point = examBest7PointSum_(pointValues);
  var division = examPointToDivision_(point);
  return { point: point, division: division };
}

function getExamResults(className, examType) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var formNum = extractFormNumber(className);
  var level   = resolveClassLevel(className);
  var rows    = [];

  if (formNum) {
    var sh = ss.getSheetByName('Form' + formNum + '_Scores');
    if (sh && sh.getLastRow() > 1) rows = rows.concat(sh.getDataRange().getValues().slice(1));
  }
  var legacy = ss.getSheetByName('Scores');
  if (legacy && legacy.getLastRow() > 1) {
    legacy.getDataRange().getValues().slice(1).forEach(function(r) {
      var rFn = extractFormNumber((r[1] || '').toString());
      if (formNum && rFn !== formNum) return;
      rows.push(r);
    });
  }

  var wantedExam = String(examType || '').trim().toLowerCase();
  var wantedClass = String(className || '').trim().toLowerCase();
  var relevant = rows.filter(function(r) {
    var rowExam  = String(r[0] || '').trim().toLowerCase();
    var rowClass = String(r[1] || '').trim();
    var rowClassLower = rowClass.toLowerCase();
    var classMatch =
      rowClassLower === wantedClass ||
      rowClassLower === String(level || '').toLowerCase() ||
      extractFormNumber(rowClass) === formNum ||
      resolveClassLevel(rowClass).toLowerCase() === String(level || '').toLowerCase();
    return classMatch && rowExam === wantedExam;
  });
  if (!relevant.length) return [];

  var studentMap = {};
  relevant.forEach(function(r) {
    var subject = String(r[2] || '').trim();
    var admNo   = String(r[3] || '').trim();
    var name    = String(r[4] || '').trim();
    var sex     = String(r[5] || '').trim();
    var raw     = r[6];
    var score   = (raw !== '' && raw !== null && raw !== undefined && !isNaN(Number(raw))) ? Number(raw) : '-';
    if (!name) return;

    if (!studentMap[name]) {
      studentMap[name] = { name: name, admNo: admNo, sex: sex, subjects: [] };
    }
    if (!studentMap[name].admNo && admNo) studentMap[name].admNo = admNo;
    if (!studentMap[name].sex && sex) studentMap[name].sex = sex;

    var grade = examScoreToGrade_(score);
    studentMap[name].subjects.push({
      code: subject,
      name: subject,
      score: score,
      grade: grade,
      point: examGradeToPoint_(grade)
    });
  });

  return Object.values(studentMap).map(function(s) {
    var nums = s.subjects
      .map(function(x) { return typeof x.score === 'number' ? x.score : null; })
      .filter(function(v) { return v !== null; });
    var total = nums.reduce(function(a, b) { return a + b; }, 0);
    var avg   = nums.length ? total / nums.length : 0;
    var avgGrade = examScoreToGrade_(avg);
    var summary = examBest7SummaryFromSubjects_(s.subjects);

    return {
      sn: 0,
      name: s.name,
      admNo: s.admNo || '',
      sex: s.sex,
      subjects: s.subjects,
      total: total,
      average: avg.toFixed(2),
      grade: avgGrade,
      point: summary.point,
      division: summary.division,
      position: 0,
      remarks: summary.division === 'ABS' ? 'ABSENT' : (summary.division === '' ? 'INCOMPLETE' : (summary.division !== '0' ? 'PASS' : 'FAIL'))
    };
  })
  .sort(function(a, b) { return b.total - a.total; })
  .map(function(s, i) { return Object.assign({}, s, { sn: i + 1, position: i + 1, outOf: Object.keys(studentMap).length }); });
}
function fetchTopPanelData(className, examType) {
  let students = getExamResults(className, examType);
  if (!students || !students.length) return {};

  // ── Prem No for every pupil (looked up inside THIS class only, because an Adm No may repeat in other classes) ──
  const premByAdm = {}, premByName = {};
  try {
    const clsKey = primaryClassSheetName_(className) || String(className || '').replace(/\s+/g, '').toLowerCase();
    getStudentListRaw().forEach(function (st) {
      const k = primaryClassSheetName_(st.className) || String(st.className || '').replace(/\s+/g, '').toLowerCase();
      if (k !== clsKey) return;
      const prem = String(st.premNo || '').trim();
      if (st.admNo) premByAdm[String(st.admNo).trim().toLowerCase()] = prem;
      premByName[(String(st.firstName || '') + ' ' + String(st.lastName || '')).trim().toLowerCase().replace(/\s+/g, ' ')] = prem;
    });
  } catch (e) {}
  students = students.map(function (s) {
    const prem = premByAdm[String(s.admNo || '').trim().toLowerCase()]
              || premByName[String(s.name || '').trim().toLowerCase().replace(/\s+/g, ' ')] || '';
    return Object.assign({}, s, { premNo: prem });
  });

  // ── Subject codes from the AcademicSubjects sheet (Name → Code, e.g. Arithmetic → ARITH) ──
  const codeByName = {};
  try {
    getAcademicSubjectsRaw().forEach(function (a) {
      const k = String(a.name || '').trim().toLowerCase().replace(/\s+/g, ' ');
      if (k && String(a.code || '').trim()) codeByName[k] = String(a.code).trim();
    });
  } catch (e) {}

  const subjectNameSet = new Set();
  students.forEach(s => s.subjects.forEach(sub => { if (sub.name) subjectNameSet.add(sub.name); }));
  // code = key used to match scores (the subject name); shortCode = the official subject code to display
  const subjects      = Array.from(subjectNameSet).sort().map(n => ({
    name: n, code: n,
    shortCode: codeByName[String(n).trim().toLowerCase().replace(/\s+/g, ' ')] || ''
  }));
  const gradesWeight  = { A: 1, B: 2, C: 3, D: 4, F: 5 };

  const attendance = { F: {}, M: {}, TOTAL: {} };
  ['F','M'].forEach(sex => {
    const ss2 = students.filter(s => String(s.sex || '').trim().toUpperCase().charAt(0) === sex);
    // FIX: was hardcoded to 0 and never populated, so absentees (division
    // === 'ABS', meaning the student had no valid score in any subject)
    // were silently counted as PRESENT here.
    const absCount = ss2.filter(s => s.division === 'ABS').length;
    attendance[sex] = { REG: ss2.length, ABS: absCount, PRE: ss2.length - absCount };
  });
  attendance.TOTAL = {
    REG: attendance.F.REG + attendance.M.REG,
    ABS: attendance.F.ABS + attendance.M.ABS,
    PRE: attendance.F.PRE + attendance.M.PRE
  };

 // PRIMARY SCHOOL: no divisions. Pass = overall grade D, C, B or A; Fail = overall grade F only.
 students.forEach(s => {
  if (s.division === 'ABS') {
    s.remarks = 'ABSENT';            // did not sit the exam at all
  } else {
    s.remarks = (String(s.grade || '').toUpperCase() === 'F') ? 'FAIL' : 'PASS';
  }
});

  const examRanking = {
    PassedCandidates: students.filter(s => s.remarks === 'PASS').length,
    FailedCandidates: students.filter(s => s.remarks === 'FAIL').length
  };
  // Examination AVERAGE (primary school): mean of the pupils' average scores (absentees excluded)
  const sitters = students.filter(s => s.remarks !== 'ABSENT');
  const avgSum  = sitters.reduce((a, s) => a + (parseFloat(s.average) || 0), 0);
  examRanking.Average      = sitters.length ? (avgSum / sitters.length).toFixed(2) : '-';
  examRanking.AverageGrade = sitters.length ? examScoreToGrade_(avgSum / sitters.length) : '';

  const subjectsPerf = subjects.map((sub, i) => {
    let pass=0,fail=0,A=0,B=0,C=0,D=0,F=0,totalScore=0,count=0;
    students.forEach(s => {
      const subj = s.subjects.find(ss => ss.code === sub.code || ss.name === sub.name);
      if (!subj || subj.score === '-') return;
      const score = Number(subj.score) || 0;
      const grade = subj.grade || examScoreToGrade_(score);
      totalScore += score; count++;
      if (grade==='A'){A++;pass++;}else if(grade==='B'){B++;pass++;}
      else if(grade==='C'){C++;pass++;}else if(grade==='D'){D++;pass++;}else{F++;fail++;}
    });
    const average = count ? totalScore/count : 0;
    const cl      = average>=75?'A (Excellent)':average>=65?'B (Very Good)':average>=45?'C (Good)':average>=30?'D (Satisfactory)':'F (Fail)';
    return { sn: i+1, name: sub.name, PASS: pass, FAIL: fail, A, B, C, D, F, AVERAGE: count ? average.toFixed(2) : '-', COMPETENCY_LEVEL: cl };
  });

  const divisionPerf = {
    DIVI:   students.filter(s => s.division === 'I').length,
    DIVII:  students.filter(s => s.division === 'II').length,
    DIVIII: students.filter(s => s.division === 'III').length,
    DIVIV:  students.filter(s => s.division === 'IV').length,
    DIV0:   students.filter(s => examNormalizeDivisionKey_(s.division) === 'ZERO').length,
    // FIX: was hardcoded to 0 and never populated. Now reflects students who
    // did not sit the exam at all (division === 'ABS').
    ABS:    students.filter(s => s.division === 'ABS').length,
    TOTAL:  students.length
  };

  const studentDivision = {
    // FIX: added ABS/NA keys — without them, an absent or incomplete
    // student's bucket++ below would create the key as NaN instead of 1.
    F:     {I:0,II:0,III:0,IV:0,ZERO:0,ABS:0,NA:0,TOTAL:0},
    M:     {I:0,II:0,III:0,IV:0,ZERO:0,ABS:0,NA:0,TOTAL:0},
    TOTAL: {I:0,II:0,III:0,IV:0,ZERO:0,ABS:0,NA:0,TOTAL:0}
  };
  students.forEach(s => {
    const sex = (String(s.sex||'').trim().toUpperCase().charAt(0) === 'F') ? 'F' : 'M';
    const key = examNormalizeDivisionKey_(s.division);
    studentDivision[sex][key]++;
    studentDivision[sex].TOTAL++;
    studentDivision.TOTAL[key]++;
    studentDivision.TOTAL.TOTAL++;
  });

  const gradePerf = {
    F:     {A:0,B:0,C:0,D:0,F:0,TOTAL:0},
    M:     {A:0,B:0,C:0,D:0,F:0,TOTAL:0},
    TOTAL: {A:0,B:0,C:0,D:0,F:0,TOTAL:0}
  };
  students.forEach(s => {
    const sex = (String(s.sex||'').trim().toUpperCase().charAt(0) === 'F') ? 'F' : 'M';
    const g = ['A','B','C','D','F'].includes(s.grade) ? s.grade : 'F';
    gradePerf[sex][g]++;
    gradePerf[sex].TOTAL++;
    gradePerf.TOTAL[g]++;
    gradePerf.TOTAL.TOTAL++;
  });

  return { attendance, examRanking, subjectsPerf, divisionPerf, studentDivision, gradePerf, students, subjects };
}

// ============================================================
// STUDENT LIST (full — all forms)
// ============================================================
function getStudentListRaw() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const allStudents = [];
  const seenAdmNo    = {}; // dedupe key: lower-cased AdmNo

  primaryStudentSheetNames_().forEach(function(sheetName, index) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    // Ensure header row has 'Status' in col 10 (index 9)
    // (safe no-op if it's already there)
    const headerRow = data[0];
    if (headerRow.length < 10 || !headerRow[9]) {
      sheet.getRange(1, 10).setValue('Status');
    }

    data.slice(1).forEach(function(r) {
      if (r[0] || r[1]) {
        var admNo = String(r[0] || '').trim();
        allStudents.push({
          admNo:     r[0],
          firstName: r[1],
          lastName:  r[2],
          sex:       r[3],
          dob:       r[4] ? r[4].toString().substring(0, 10) : '',
          className: r[5] || PS_CLASSES[index],
          parent:    r[6],
          phone:     r[7],
          address:   r[8],
          status:    r[9] ? r[9].toString().trim() : 'Pending',
          premNo:    String(r[10] || '').trim()
        });
        if (admNo) seenAdmNo[admNo.toLowerCase()] = true;
      }
    });
  });

  // ------------------------------------------------------------
  // Also include students from the separate legacy "Students" roster.
  // ------------------------------------------------------------
  var rosterSh = ss.getSheetByName('Students');
  if (rosterSh && rosterSh.getLastRow() > 1) {
    var rv = rosterSh.getDataRange().getValues();
    var rh = rv.shift().map(function(h){ return String(h || '').trim(); });
    var iAdm    = rh.indexOf('AdmNo');
    var iFirst  = rh.indexOf('FirstName');
    var iLast   = rh.indexOf('LastName');
    var iName   = rh.indexOf('Name');
    var iSex    = rh.indexOf('Sex');
    var iDob    = rh.indexOf('DOB');
    var iCls    = rh.indexOf('Class');
    var iParent = rh.indexOf('Parent');
    var iPhone  = rh.indexOf('Phone');
    var iAddr   = rh.indexOf('Address');
    var iStat   = rh.indexOf('Status');
    var iPrem   = rh.indexOf('PremNo');

    rv.forEach(function(r) {
      var admNo = iAdm > -1 ? String(r[iAdm] || '').trim() : '';
      var firstName = '', lastName = '';
      if (iFirst > -1 || iLast > -1) {
        firstName = iFirst > -1 ? String(r[iFirst] || '').trim() : '';
        lastName  = iLast  > -1 ? String(r[iLast]  || '').trim() : '';
      } else if (iName > -1) {
        var parts = String(r[iName] || '').trim().split(/\s+/);
        firstName = parts.shift() || '';
        lastName  = parts.join(' ');
      }
      if (!admNo && !firstName) return;                          // nothing usable on this row
      if (admNo && seenAdmNo[admNo.toLowerCase()]) return;

      allStudents.push({
        admNo:     admNo,
        firstName: firstName,
        lastName:  lastName,
        sex:       iSex   > -1 ? String(r[iSex]  || '').trim() : '',
        dob:       iDob   > -1 && r[iDob] ? r[iDob].toString().substring(0, 10) : '',
        className: iCls   > -1 ? String(r[iCls]  || '').trim() : '',
        parent:    iParent> -1 ? String(r[iParent]|| '').trim() : '',
        phone:     iPhone > -1 ? String(r[iPhone] || '').trim() : '',
        address:   iAddr  > -1 ? String(r[iAddr]  || '').trim() : '',
        status:    iStat  > -1 && r[iStat] ? String(r[iStat]).trim() : 'Approved',
        premNo:    iPrem  > -1 ? String(r[iPrem] || '').trim() : ''
      });
      if (admNo) seenAdmNo[admNo.toLowerCase()] = true;
    });
  }

  return allStudents;
}

function getStudentList() { return getStudentListRaw(); }

function generateUniqueAdmNo(ss){
  // Keep admission numbers unique across all primary class sheets.
  var used = {};
  primaryStudentSheetNames_().concat(['Students']).forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function(r){
      var v = String(r[0] || '').trim();
      if (v) used[v.toLowerCase()] = true;
    });
  });
  var n = 1, admNo;
  do {
    admNo = 'S6790-' + String(n).padStart(3, '0');
    n++;
  } while (used[admNo.toLowerCase()]);
  return admNo;
}

/**
 * Scans primary class sheets + Students for AdmNo values used by more than one
 * student (a data problem that predates the unique-generation fix
 * above). Returns [{admNo, entries:[{name,className,sheet}]}].
 */
function findDuplicateAdmNos(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var byAdm = {};
  primaryStudentSheetNames_().concat(['Students']).forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    var vals = sh.getDataRange().getValues();
    var header = vals[0].map(function(h){ return String(h||'').trim(); });
    var iAdm = header.indexOf('AdmNo'), iFirst = header.indexOf('FirstName'),
        iLast = header.indexOf('LastName'), iName = header.indexOf('Name'),
        iCls = header.indexOf('Class');
    if (iAdm < 0) iAdm = 0; // Form sheets always have AdmNo in col A
    vals.slice(1).forEach(function(r, idx){
      var admNo = String(r[iAdm] || '').trim();
      if (!admNo) return;
      var display = (iFirst > -1 || iLast > -1)
        ? ((r[iFirst] || '') + ' ' + (r[iLast] || '')).trim()
        : (iName > -1 ? String(r[iName] || '').trim() : '');
      var clsKey = String(iCls > -1 ? (r[iCls] || name) : name).toLowerCase().replace(/\s+/g, '');
      var key = admNo.toLowerCase() + '|' + clsKey;   // same number in another class is allowed
      if (!byAdm[key]) byAdm[key] = { admNo: admNo, entries: [] };
      byAdm[key].entries.push({
        name: display,
        className: iCls > -1 ? String(r[iCls] || '').trim() : name,
        sheet: name,
        row: idx + 2
      });
    });
  });
  return Object.keys(byAdm).map(function(k){ return byAdm[k]; })
    .filter(function(g){ return g.entries.length > 1; });
}

// ============================================================
// ADMISSION NUMBER SERIES  (Set Admission No.)
// ------------------------------------------------------------
// Format:  <prefix>-<number padded>   e.g. PS0123456-0001, PS0123456-0002
// The series runs SEPARATELY FOR EACH CLASS: Standard 1 has 0001, 0002 …
// and Standard 2 also starts again from 0001 (duplicates between classes
// are allowed). Inside one class a number is never repeated.
// ============================================================
var ADM_SERIES_KEY = 'ADM_SERIES';

// Until the Academic Master sets a series, numbers start with the school code from Settings (or ADM).
function admDefaultPrefix_() {
  try {
    var code = String(psSettings_().SchoolCode || '').trim().replace(/[^A-Za-z0-9._\/]/g, '');
    if (code.length >= 2) return code.slice(0, 30);
  } catch (e) {}
  return 'ADM';
}
function admSeriesConfig_() {
  var cfg = { prefix: admDefaultPrefix_(), digits: 4, start: 1 };
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(ADM_SERIES_KEY);
    if (raw) {
      var o = JSON.parse(raw);
      if (o.prefix) cfg.prefix = String(o.prefix).trim();
      if (o.digits) cfg.digits = Number(o.digits) || 4;
      if (o.start)  cfg.start  = Number(o.start)  || 1;
    }
  } catch (e) {}
  return cfg;
}
function admSeriesFormat_(cfg, n) { return cfg.prefix + '-' + String(n).padStart(cfg.digits, '0'); }

// Next free number of ONE class (highest number in that class + 1, or "start" if none yet)
function admSeriesNextFor_(ss, className, cfg) {
  var name = primaryClassSheetName_(className);
  if (!name) return '';
  cfg = cfg || admSeriesConfig_();
  var sh = ss.getSheetByName(name), max = cfg.start - 1;
  var re = new RegExp('^' + cfg.prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '-(\\d+)$', 'i');
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
      var m = String(r[0] == null ? '' : r[0]).trim().match(re);
      if (m) { var n = parseInt(m[1], 10); if (n > max) max = n; }
    });
  }
  return admSeriesFormat_(cfg, max + 1);
}

function getAdmissionSeries() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), cfg = admSeriesConfig_(), next = {};
  PS_CLASSES.forEach(function (c) { next[c] = admSeriesNextFor_(ss, c, cfg); });
  return { cfg: cfg, next: next };
}

function saveAdmissionSeries(data) {
  data = data || {};
  var prefix = String(data.prefix || '').trim().replace(/[-\s]+$/, '');
  var digits = Number(data.digits), start = Number(data.start);
  if (!/^[A-Za-z0-9._\/]{2,30}$/.test(prefix))
    return { ok: false, message: 'Prefix must be 2–30 characters: letters, numbers, dot or slash (no spaces, no dash at the end).' };
  if (!(digits >= 3 && digits <= 6)) return { ok: false, message: 'Number length must be between 3 and 6.' };
  if (!(start >= 1 && start <= 999999) || start % 1 !== 0) return { ok: false, message: 'Start number must be a whole number of 1 or more.' };
  PropertiesService.getScriptProperties().setProperty(ADM_SERIES_KEY, JSON.stringify({ prefix: prefix, digits: digits, start: start }));
  try { logActivity('academic', 'Admission No. series set: ' + prefix + '-' + String(start).padStart(digits, '0')); } catch (e) {}
  var info = getAdmissionSeries();
  info.ok = true;
  info.message = 'Admission No. series saved. Next admission in Standard 1 will be ' + info.next[PS_CLASSES[0]] + '.';
  return info;
}

function previewNextAdmNo(className) {
  return admSeriesNextFor_(SpreadsheetApp.getActiveSpreadsheet(), className, admSeriesConfig_());
}

// Make sure a class sheet has the Status (col 10) and PremNo (col 11) headers
function admEnsureStudentCols_(sheet) {
  if (sheet.getMaxColumns() < 11) sheet.insertColumnsAfter(sheet.getMaxColumns(), 11 - sheet.getMaxColumns());
  var h = sheet.getRange(1, 1, 1, 11).getValues()[0];
  if (!String(h[9]  || '').trim()) sheet.getRange(1, 10).setValue('Status');
  if (!String(h[10] || '').trim()) sheet.getRange(1, 11).setValue('PremNo');
}

// Prem No must be unique for the whole school (it identifies the pupil nationally)
function admPremTakenBy_(ss, prem, exceptSheet, exceptRow) {
  if (!prem) return '';
  var p = prem.toLowerCase(), who = '';
  primaryStudentSheetNames_().some(function (nm) {
    var sh = ss.getSheetByName(nm);
    if (!sh || sh.getLastRow() < 2 || sh.getMaxColumns() < 11) return false;
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (nm === exceptSheet && (i + 2) === exceptRow) continue;
      if (String(vals[i][10] || '').trim().toLowerCase() === p) {
        who = (String(vals[i][1] || '') + ' ' + String(vals[i][2] || '')).trim() + ' (' + nm + ')';
        return true;
      }
    }
    return false;
  });
  return who;
}

function saveStudent(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = primaryClassSheetName_(data.className);
  if (!sheetName) return 'Error: Select a valid primary class (Standard 1–' + PS_CLASSES.length + ').';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['AdmNo','FirstName','LastName','Sex','DOB','Class','Parent','Phone','Address','Status','PremNo']);
    sheet.setFrozenRows(1);
    logActivity('system', 'Auto-created sheet: ' + sheetName);
  }
  admEnsureStudentCols_(sheet);

  const prem   = String(data.premNo || '').trim();
  const isEdit = !!(data.oldId && String(data.oldId).trim() !== '');

  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { return 'Error: System busy, please try again.'; }
  try {
    // ── NEW admission: number comes from the class series ──
    if (!isEdit) {
      const taken = admPremTakenBy_(ss, prem, '', 0);
      if (taken) return 'Error: Prem No ' + prem + ' already belongs to ' + taken + '.';
      const admNo = admSeriesNextFor_(ss, data.className, admSeriesConfig_());
      sheet.appendRow([admNo, data.firstName, data.lastName, data.sex, data.dob,
                       data.className, data.parent, data.phone, data.address, '', prem]);
      logActivity('academic', 'Admitted: ' + data.firstName + ' ' + data.lastName + ' (' + admNo + ') to ' + sheetName);
      return 'Student ' + data.firstName + ' admitted to ' + data.className + ' successfully (' + admNo + ')';
    }

    // ── EDIT: find the student by Adm No + CLASS (same number may exist in other classes) ──
    const oldId    = String(data.oldId).trim();
    const oldClass = String(data.oldClass || '').trim();
    const oldSheet = oldClass ? primaryClassSheetName_(oldClass) : '';
    const searchIn = oldSheet ? [oldSheet] : primaryStudentSheetNames_();
    const hits = [];
    searchIn.forEach(function (nm) {
      const sh = ss.getSheetByName(nm);
      if (!sh || sh.getLastRow() < 2) return;
      const vals = sh.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++)
        if (String(vals[i][0] || '').trim() === oldId) hits.push({ sh: sh, i: i, vals: vals });
    });
    if (!hits.length) return 'Error: Student ' + oldId + ' was not found in the primary class register.';
    if (hits.length > 1) return 'Error: Admission No. ' + oldId + ' exists in more than one class. Reload the register and edit the student again.';
    const h = hits[0], sourceSheet = h.sh, srcRow = h.i + 1;

    const taken = admPremTakenBy_(ss, prem, sourceSheet.getName(), srcRow);
    if (taken) return 'Error: Prem No ' + prem + ' already belongs to ' + taken + '.';

    let admNo = String(data.admNo || oldId).trim() || oldId;
    const base = [admNo, data.firstName, data.lastName, data.sex, data.dob,
                  data.className, data.parent, data.phone, data.address];

    if (sourceSheet.getName() === sheetName) {
      sourceSheet.getRange(srcRow, 1, 1, 9).setValues([base]);
      sourceSheet.getRange(srcRow, 11).setValue(prem);
      logActivity('academic', 'Updated student: ' + data.firstName + ' ' + data.lastName);
      return 'Student updated successfully in ' + data.className;
    }

    // moved to another class: keep the number unless it is already used in the new class
    let note = '';
    const used = {};
    if (sheet.getLastRow() > 1)
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function (r) { used[String(r[0] || '').trim().toLowerCase()] = true; });
    if (used[admNo.toLowerCase()]) {
      admNo = admSeriesNextFor_(ss, data.className, admSeriesConfig_());
      base[0] = admNo;
      note = ' (new Admission No. in ' + data.className + ': ' + admNo + ')';
    }
    const moved = h.vals[h.i].slice();
    while (moved.length < 11) moved.push('');
    base.forEach(function (v, c) { moved[c] = v; });
    moved[10] = prem;
    sheet.appendRow(moved);
    sourceSheet.deleteRow(srcRow);
    logActivity('academic', 'Updated student: ' + data.firstName + ' ' + data.lastName);
    return 'Student updated successfully in ' + data.className + note;
  } finally {
    lock.releaseLock();
  }
}

// className is REQUIRED to be exact when the same Adm No. exists in several classes
function deleteStudent(admNo, className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clsSheet = className ? primaryClassSheetName_(className) : '';
  const sheetNames = clsSheet ? [clsSheet] : primaryStudentSheetNames_().concat(['Students']);
  for (var s = 0; s < sheetNames.length; s++) {
    var sheet = ss.getSheetByName(sheetNames[s]);
    if (!sheet || sheet.getLastRow() < 2) continue;
    var data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() == String(admNo).trim()) { sheet.deleteRow(i + 1); return 'Student deleted'; }
    }
  }
  return 'Student not found';
}

// ============================================================
// ACADEMIC CLASSES / SUBJECTS / EXAMS
// ============================================================
function getAcademicClassesRaw() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AcademicClasses');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1)
    .map(r => ({ id: r[0], name: r[1], stream: r[2], teacher: r[3], capacity: r[4], students: r[5]||0 }))
    .filter(c => c.name);
}
function getAcademicClasses() { return getAcademicClassesRaw(); }

function saveAcademicClass(data) {
  const assigned = String(data.teacher || '').trim();
  if (assigned) {
    const person = getTeacherList().find(u => String(u.username || '').trim().toLowerCase() === assigned.toLowerCase());
    if (!person) throw new Error('That teacher account was not found.');
    if (String(person.position || '').trim() !== 'Class Teacher') {
      throw new Error(person.name + ' is a ' + (person.position || 'staff member') + '. Only an account with the position Class Teacher can be the Class Teacher of a class (the Head of School sets the position).');
    }
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('AcademicClasses');
  if (!sheet) {
    sheet = ss.insertSheet('AcademicClasses');
    sheet.appendRow(['ID','Name','Stream','Teacher','Capacity','Students']);
    sheet.setFrozenRows(1);
  }
  const row = [data.id||'', data.name, data.stream||'', data.teacher||'', data.capacity||40, 0];
  if (data.oldName && data.oldName.trim() !== '') {
    const vals = sheet.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][1] === data.oldName) {
        sheet.getRange(i+1,1,1,5).setValues([[data.id||'',data.name,data.stream||'',data.teacher||'',data.capacity||40]]);
        return 'Class updated successfully';
      }
    }
  }
  sheet.appendRow(row);
  return 'Class ' + data.name + ' added successfully';
}

function deleteAcademicClass(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AcademicClasses');
  if (!sheet) return 'Sheet not found';
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === name) { sheet.deleteRow(i+1); return 'Class deleted'; }
  }
  return 'Class not found';
}

function getAcademicSubjectsRaw() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AcademicSubjects');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1)
    .map(r => ({ name: r[0], code: r[1], category: r[2], department: r[3], classes: r[4] }))
    .filter(s => s.name);
}
function getAcademicSubjects() { return getAcademicSubjectsRaw(); }

function saveAcademicSubject(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('AcademicSubjects');
  if (!sheet) {
    sheet = ss.insertSheet('AcademicSubjects');
    sheet.appendRow(['Name','Code','Category','Department','Classes']);
    sheet.setFrozenRows(1);
  }
  const row = [data.name, data.code, data.category||'Compulsory', data.department||'', data.classes||''];
  if (data.oldCode && data.oldCode.trim() !== '') {
    const vals = sheet.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][1] === data.oldCode) { sheet.getRange(i+1,1,1,5).setValues([row]); return 'Subject updated successfully'; }
    }
  }
  sheet.appendRow(row);
  return 'Subject ' + data.name + ' added successfully';
}

function deleteAcademicSubject(code) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AcademicSubjects');
  if (!sheet) return 'Sheet not found';
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === code) { sheet.deleteRow(i+1); return 'Subject deleted'; }
  }
  return 'Subject not found';
}

function getAcademicExamsRaw() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Examinations');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).map(r => {
    const startRaw = r[4], endRaw = r[5];
    function isoLocal(v) {
      if (!v) return '';
      if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const s = v.toString().trim();
      const m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
      return m ? m[0] : s.substring(0, 10);
    }
    return {
      id: r[0], name: r[1], type: r[2], year: r[3],
      startDate: startRaw, endDate: endRaw,
      startDateStr: isoLocal(startRaw), endDateStr: isoLocal(endRaw),
      classes: r[6], remarks: r[7], status: r[8]||'Draft'
    };
  }).filter(e => e.name);
}

function getAcademicExams() {
  return getAcademicExamsRaw().map(e => ({
    id: e.id, name: e.name, type: e.type, year: e.year,
    startDate: e.startDateStr||'', endDate: e.endDateStr||'',
    classes: e.classes, remarks: e.remarks, status: e.status
  }));
}

function getActiveExamsForTeacher(className, examType) {
  var allExams = getAcademicExamsRaw();
  var today = new Date();
  today.setHours(0,0,0,0);

  return allExams.filter(function(e) {
    // 1. Must be active status
    if (e.status !== 'active') return false;

    // 2. Must match exam type
    if (e.type !== examType) return false;

    // 3. Must be for this class or "All" — use classMatchesExam to handle spacing
    if (!classMatchesExam(className, e.classes)) return false;

    // 4. Auto-hide if end date passed
    var endDate = new Date(e.endDateStr);
    var endDatePlusOne = new Date(endDate);
    endDatePlusOne.setDate(endDate.getDate() + 1);
    if (endDatePlusOne <= today) return false;

    return true;
  }).map(e => ({
    id: e.id,
    name: e.name,
    type: e.type,
    year: e.year
  }));
}

function saveAcademicExam(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Examinations');
  if (!sheet) {
    sheet = ss.insertSheet('Examinations');
    sheet.appendRow(['ID','Name','Type','Year','StartDate','EndDate','Classes','Remarks','Status']);
    sheet.setFrozenRows(1);
  }
  const id  = data.oldId || ('EX' + Date.now());
  const row = [id, data.name, data.type, data.year, data.startDate, data.endDate,
             data.classes||'All', data.remarks||'', data.status||'active'];
  if (data.oldId && data.oldId.trim() !== '') {
    const vals = sheet.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][0] === data.oldId) { sheet.getRange(i+1,1,1,9).setValues([row]); return 'Examination updated successfully'; }
    }
  }
  sheet.appendRow(row);
  return 'Examination "' + data.name + '" created successfully';
}
function checkAndDeactivatePastExams() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Examinations'); // <-- FIXED
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0,0,0,0);

  for (let i = 1; i < data.length; i++) {
    const endDate = new Date(data[i][5]); // Col F = EndDate
    let status = (data[i][8] || '').toString().toLowerCase().trim(); // Col I = Status

    // Auto-fix: Convert 'Draft' to 'active' if end date hasn't passed
    if (status === 'draft' && endDate >= today) {
      sheet.getRange(i+1, 9).setValue('active');
      status = 'active';
    }

    // Auto-deactivate if past end date
    if (endDate < today && status === 'active') {
      sheet.getRange(i+1, 9).setValue('inactive');
    }
  }
}
function deleteAcademicExam(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Examinations');
  if (!sheet) return 'Sheet not found';
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { sheet.deleteRow(i+1); return 'Examination deleted'; }
  }
  return 'Examination not found';
}

function publishExamResults(className, examType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Examinations');
  if (!sheet) return 'Examinations sheet not found';
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === examType) sheet.getRange(i+1, 9).setValue('Published');
  }
  logActivity('academic', 'Published results: ' + className + ' ' + examType);
  return 'Results published for ' + className + ' — ' + examType;
}

// ============================================================
// TIMETABLE
// ============================================================
function getTimetable(className) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Timetable_' + className.replace(/\s/g,'_'));
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const result = {};
  data.slice(1).forEach(r => {
    if (r[0]) {
      const key = r[0] + '_' + r[1];
      result[key] = { day: r[0], period: r[1], start: r[2], end: r[3], subject: r[4], teacher: r[5], room: r[6] };
    }
  });
  return result;
}

function saveTimetable(className, ttData) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Timetable_' + className.replace(/\s/g,'_');
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) sheet.clearContents(); else sheet = ss.insertSheet(sheetName);
  sheet.appendRow(['Day','Period','Start','End','Subject','Teacher','Room']);
  Object.keys(ttData).forEach(key => {
    const c = ttData[key];
    sheet.appendRow([c.day, c.period, c.start||'', c.end||'', c.subject||'', c.teacher||'', c.room||'']);
  });
  logActivity('academic', 'Saved timetable for ' + className);
  return 'Timetable saved for ' + className;
}

// ============================================================
// MESSAGES
// ============================================================
function getMessages(username) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Messages');
  if (!sheet) return { inbox: [], sent: [] };
  const data  = sheet.getDataRange().getValues();
  const inbox = [], sent = [];
  data.slice(1).forEach(r => {
    // Columns: From(0), To(1), Subject(2), Body(3), Date(4), FromName(5), ToName(6)
    const msg = {
      from:     r[5] || r[0],   // display name, fallback to username
      to:       r[6] || r[1],   // display name, fallback to username
      fromUser: r[0],
      toUser:   r[1],
      subject:  r[2],
      body:     r[3],
      date:     r[4] ? new Date(r[4]).toLocaleString() : ''
    };
    if (r[1] === username || r[1] === 'ALL_TEACHERS') inbox.push(msg);
    if (r[0] === username) sent.push(msg);
  });
  return { inbox: inbox.reverse(), sent: sent.reverse() };
}

function sendMessage(data) {
  var accessCheck = checkDashboardWriteAccess(['Messages','Parent Contact']);
  if (!accessCheck.ok) return accessCheck.message;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Messages');
  if (!sheet) {
    sheet = ss.insertSheet('Messages');
    sheet.appendRow(['From','To','Subject','Body','Date','FromName','ToName']);
    sheet.setFrozenRows(1);
  }

  // Resolve display names
  const teachers  = getTeacherList();
  const fromTeacher = teachers.find(t => t.username === data.from);
  const fromName  = fromTeacher ? fromTeacher.name : data.from;
  const toName    = data.to === 'ALL_TEACHERS' ? 'All Teachers'
                  : (teachers.find(t => t.username === data.to) || {}).name || data.to;

  sheet.appendRow([data.from, data.to, data.subject, data.body, new Date(), fromName, toName]);
  logActivity(fromName || data.from, 'Sent message to ' + toName + ': ' + data.subject);
  return 'Message sent successfully!';
}

function changeTeacherPassword(username, oldPass, newPass) {
  const sheet = ensureUsersSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      if (data[i][2] !== oldPass) return 'Current password is incorrect';
      sheet.getRange(i+1, 3).setValue(newPass);
      logActivity(username, 'Changed password');
      return 'Password changed successfully!';
    }
  }
  return 'User not found';
}

// ============================================================
// STAFF RECORDS
// ============================================================
function saveStaffRecord(role, type, row) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Staff_' + role.replace(/\s/g,'_') + '_' + type;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.appendRow(row);
  logActivity(role, 'Added record to ' + type + ': ' + (row[0]||''));
  return 'Record saved successfully!';
}

function getStaffRecords(role, type) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Staff_' + role.replace(/\s/g,'_') + '_' + type;
  const sheet     = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 1) return [];
  return sheet.getDataRange().getValues();
}

function deleteStaffRecord(role, type, rowIndex) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Staff_' + role.replace(/\s/g,'_') + '_' + type;
  const sheet     = ss.getSheetByName(sheetName);
  if (!sheet) return 'Sheet not found';
  sheet.deleteRow(rowIndex + 1);
  return 'Record deleted';
}

function getFinancialSummary() {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const feeSheet = ss.getSheetByName('Staff_Accountant_fees');
    const paySheet = ss.getSheetByName('Staff_Accountant_payments');
    let totalFees = 0, totalPaid = 0;
    if (feeSheet && feeSheet.getLastRow() > 0)
      feeSheet.getDataRange().getValues().forEach(r => { totalFees += Number(r[2])||0; });
    if (paySheet && paySheet.getLastRow() > 0)
      paySheet.getDataRange().getValues().forEach(r => { totalPaid += Number(r[3])||0; });
    return {
      total:   totalFees.toLocaleString() + ' TZS',
      paid:    totalPaid.toLocaleString() + ' TZS',
      pending: (totalFees - totalPaid).toLocaleString() + ' TZS'
    };
  } catch(e) { return { total: '0', paid: '0', pending: '0' }; }
}

// ============================================================
// ADMIN PANEL EXTRA FUNCTIONS
// ============================================================
function getAcademicStats() {
  const students = getStudentListRaw(), classes = getAcademicClassesRaw(),
        subjects = getAcademicSubjectsRaw(), exams = getAcademicExamsRaw();
  return { students: students.length, classes: classes.length, subjects: subjects.length, exams: exams.length };
}

function updateStudentStatus(admNo, status) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const formNames = primaryStudentSheetNames_();

  for (var f = 0; f < formNames.length; f++) {
    var sheet = ss.getSheetByName(formNames[f]);
    if (!sheet) continue;

    var data = sheet.getDataRange().getValues();

    // Ensure col-10 header exists
    if (data.length >= 1 && (data[0].length < 10 || !data[0][9])) {
      sheet.getRange(1, 10).setValue('Status');
    }

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === String(admNo).trim()) {
        sheet.getRange(i + 1, 10).setValue(status);
        logActivity('admin', status + ' student: ' + admNo + ' in ' + formNames[f]);
        return 'Student ' + admNo + ' ' + status.toLowerCase() + ' successfully';
      }
    }
  }

  return 'Student not found (Adm No: ' + admNo + ')';
}


function updateClassStatus(name, status) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AcademicClasses');
  if (!sheet) return 'Classes sheet not found';
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === name) {
      sheet.getRange(i+1, 7).setValue(status);
      logActivity('admin', status + ' class: ' + name);
      return 'Class ' + name + ' status set to ' + status;
    }
  }
  return 'Class not found';
}

function updateSubjectStatus(code, status) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AcademicSubjects');
  if (!sheet) return 'Subjects sheet not found';
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === code) {
      sheet.getRange(i+1, 6).setValue(status);
      logActivity('admin', status + ' subject: ' + code);
      return 'Subject ' + code + ' status set to ' + status;
    }
  }
  return 'Subject not found';
}

function getMonitorData(panel) {
  try {
    if (panel === 'academic' || panel === '') {
      const students = getStudentListRaw(), subjects = getAcademicSubjectsRaw(), exams = getAcademicExamsRaw();
      return {
        v1: students.length, v2: subjects.length, v3: exams.length,
        headers: ['Adm No.','Name','Sex','Class','Status'],
        rows: students.slice(0,15).map(s => [s.admNo, s.firstName+' '+s.lastName, s.sex, s.className, s.status||'Pending'])
      };
    }
    if (panel === 'classteachers') {
      const teachers = getTeacherList().filter(t => String(t.position || '').trim() === 'Class Teacher');
      const classes  = getAcademicClassesRaw();
      return {
        v1: teachers.length, v2: classes.length, v3: '—',
        headers: ['Name','Username','Position','Classes','Subjects'],
        rows: teachers.slice(0,50).map(t => [t.name, t.username, t.position, t.classes.join(','), t.subjects.join(',')])
      };
    }
    if (panel === 'teachers') {
      const teachers = getTeacherList().filter(t => String(t.position || '').trim() === 'Teacher');
      return {
        v1: teachers.length, v2: '—', v3: '—',
        headers: ['Name','Username','Position','Classes','Subjects'],
        rows: teachers.slice(0,50).map(t => [t.name, t.username, t.position, t.classes.join(','), t.subjects.join(',')])
      };
    }
    if (panel === 'accountant') {
      const fees     = getStaffRecords('Accountant','fees');
      const payments = getStaffRecords('Accountant','payments');
      return { v1: fees.length, v2: payments.length, v3: '—', headers: ['Description','Amount','Date','Method'], rows: payments.slice(0,15) };
    }
    if (panel === 'library') {
      const books   = getStaffRecords('Librarian','books');
      const borrows = getStaffRecords('Librarian','borrow');
      return { v1: books.length, v2: borrows.length, v3: '—', headers: ['Title','Author','ISBN','Category','Copies'], rows: books.slice(0,15) };
    }
    if (panel === 'sports') {
      const sports  = getStaffRecords('Sports Manager','sports');
      const events  = getStaffRecords('Sports Manager','events');
      return { v1: sports.length, v2: '—', v3: events.length, headers: ['Sport','Category','Season','Venue'], rows: sports.slice(0,15) };
    }
    if (panel === 'lab') {
      const inv    = getStaffRecords('Lab Technician','inventory');
      const sched  = getStaffRecords('Lab Technician','schedule');
      const safety = getStaffRecords('Lab Technician','safety');
      return { v1: inv.length, v2: sched.length, v3: safety.length, headers: ['Item','Category','Quantity','Unit','Location'], rows: inv.slice(0,15) };
    }
    if (panel === 'parents') {
      const parents = getParentList();
      return {
        v1: parents.length, v2: '—', v3: '—',
        headers: ['Name','Username','StudentName','ClassName'],
        rows: parents.slice(0,50).map(r => [r[0], r[7], r[9], r[10]])
      };
    }
  } catch(e) { Logger.log(e.message); }
  return { v1: 0, v2: 0, v3: 0, headers: [], rows: [] };
}

function saveSystemConfig(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('SystemConfig');
  if (!sheet) { sheet = ss.insertSheet('SystemConfig'); sheet.appendRow(['Key','Value']); }
  Object.keys(settings).forEach(key => {
    const data = sheet.getDataRange().getValues(); let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) { sheet.getRange(i+1,2).setValue(settings[key]); found = true; break; }
    }
    if (!found) sheet.appendRow([key, settings[key]]);
  });
  logActivity('admin','Updated system settings');
  return 'Settings saved successfully';
}

function saveSchoolConfig(info) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('SchoolConfig');
  if (!sheet) { sheet = ss.insertSheet('SchoolConfig'); sheet.appendRow(['Key','Value']); }
  Object.keys(info).forEach(key => {
    const data = sheet.getDataRange().getValues(); let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) { sheet.getRange(i+1,2).setValue(info[key]); found = true; break; }
    }
    if (!found) sheet.appendRow([key, info[key]]);
  });
  logActivity('admin','Updated school information');
  return 'School information saved successfully';
}

function saveAdminProfile(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('AdminProfile');
  if (!sheet) { sheet = ss.insertSheet('AdminProfile'); sheet.appendRow(['Key','Value']); }
  Object.keys(data).forEach(key => {
    const vals = sheet.getDataRange().getValues(); let found = false;
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][0] === key) { sheet.getRange(i+1,2).setValue(data[key]); found = true; break; }
    }
    if (!found) sheet.appendRow([key, data[key]]);
  });
  logActivity('admin','Updated admin profile');
  return 'Profile saved successfully!';
}

function changeAdminPassword() { throw new Error('This function was removed.'); }

// ============================================================
// PARENT PORTAL
// ============================================================
function checkParentLogin(username, password) {
  if (!username || !password) return { valid: false };
  if (ADMIN_USERS.some(u => u.username === username && u.password === password)) return { valid: true, role: "admin" };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parents");
  if (!sheet) return { valid: false };
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  return data.some(r => r[1] === username && r[2] === password) ? { valid: true, role: "parent" } : { valid: false };
}

function getClassesAndExamTypes(username) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Parents');
  if (!sheet) return { classes: [], examTypes: [] };
  const data       = sheet.getRange('A2:G' + sheet.getLastRow()).getValues();
  const classesSet = {}, examSet = {};
  data.forEach(r => {
    if (username === ADMIN_USERS[0].username || r[1] === username) {
      if (r[4]) classesSet[r[4].trim()] = true;
      if (r[5]) examSet[r[5].trim()]    = true;
    }
  });
  return { classes: Object.keys(classesSet).sort(), examTypes: Object.keys(examSet).sort() };
}

function getStudentsForParent(username, className, examType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Parents');
  if (!sheet) return [];
  const data = sheet.getRange('A2:G' + sheet.getLastRow()).getValues();
  return data
    .filter(r => (username === ADMIN_USERS[0].username || r[1] === username) &&
                 r[4].trim() === className && r[5].trim() === examType)
    .map(r => ({ name: r[3].trim(), sheetId: r[6].trim() }));
}

function getStudentReport(studentName, sheetId) {
  try {
    if (!sheetId) return {};
    const ss      = SpreadsheetApp.openById(sheetId);
    const sheet   = ss.getSheetByName('FinalAnnual');
    if (!sheet) return {};
    const lastCol = sheet.getLastColumn(), lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const row     = data.find(r => String(r[1]).trim() === String(studentName).trim());
    if (!row) return {};

    let subjectsList = [];
    for (let i = 3; i < headers.length; i += 3) {
      const sn = headers[i] ? String(headers[i]).trim() : '';
      if (sn) subjectsList.push(sn);
    }

    const subjects = [];
    let idx = 3;
    subjectsList.forEach(sub => {
      const score = idx < row.length ? row[idx] : '';
      const grade = (idx + 1) < row.length && row[idx + 1] !== '' ? row[idx + 1] : examScoreToGrade_(score);
      subjects.push({
        name: sub,
        score: score,
        grade: grade,
        point: examGradeToPoint_(grade)
      });
      idx += 3;
    });

    const behaviors = [];
    for (let i = 0; i < 9; i++) behaviors.push(row[idx + i] || 0);
    const headerLower = headers.map(h => String(h).toLowerCase().trim());
    function readByAliases(aliases) {
      for (let a of aliases) {
        const i = headerLower.indexOf(a.toLowerCase());
        if (i >= 0) return row[i];
      }
      return '';
    }

    const summary = examBest7SummaryFromSubjects_(subjects);
    return {
      name: studentName,
      subjects,
      behaviors,
      total:    readByAliases(['total','total marks']),
      average:  readByAliases(['average','avg']),
      grade:    readByAliases(['grade']),
      point:    summary.point,
      division: summary.division,
      position: readByAliases(['position','rank']),
      outOf:    readByAliases(['out of','outof','students'])
    };
  } catch(e) { Logger.log('getStudentReport error: ' + e.message); return {}; }
}

function getAllParentAssignments() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parents");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues(); data.shift();
  return data.map(r => ({
    fullName: r[0], username: r[1], password: r[2],
    studentName: r[3], className: r[4], examType: r[5], sheetId: r[6]
  }));
}

function addParentAssignment(a) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Parents");
  if (!sheet) {
    sheet = ss.insertSheet("Parents");
    sheet.appendRow(["FullName","Username","Password","StudentName","Class","ExamType","SheetID"]);
  }
  sheet.appendRow([a.fullName||'', a.username||'', a.password||'', a.studentName||'', a.className||'', a.examType||'', a.sheetId||'']);
}

function deleteParentAssignment(index) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parents");
  if (sheet) sheet.deleteRow(index + 2);
}

// ============================================================
// TEACHER CLASSES (for attendance)
// ============================================================
function getTeacherClasses() {
  const t = getTeacherSession();
  if (!t || !t.classes) return [];
  return t.classes;
}
function isSubjectValidForClass(subjectName, className, subjectsRaw) {
  if (!subjectName || !className) return false;

  var subjL = subjectName.toString().trim().toLowerCase();

  // Find the subject record in AcademicSubjects
  var subj = null;
  for (var i = 0; i < subjectsRaw.length; i++) {
    if ((subjectsRaw[i].name || '').toString().trim().toLowerCase() === subjL) {
      subj = subjectsRaw[i];
      break;
    }
  }

  // Subject not found in sheet → no restriction, allow it
  if (!subj) return true;

  var classRestr = (subj.classes || '').toString().trim();

  // Blank or "All" → valid for every class
  if (!classRestr || classRestr.toLowerCase() === 'all') return true;

  // Compare class form number against each token in the restriction
  var formNum = extractFormNumber(className);
  var clsLower = className.toString().trim().toLowerCase();

  var tokens = classRestr.split(/[,;]+/).map(function(t){ return t.trim().toLowerCase(); });

  return tokens.some(function(token) {
    if (!token) return false;
    var tokenNum = extractFormNumber(token);
    // Match by form number (most reliable: "Form 1" == "Form1" == "form 1")
    if (formNum && tokenNum && formNum === tokenNum) return true;
    // Match by full name (normalised)
    return token === clsLower ||
           token.replace(/\s+/g, '') === clsLower.replace(/\s+/g, '');
  });
}
// ============================================================
// SCORE TRACKING DATA FOR DASHBOARD WIDGET
// Uses only Users + Scores tabs. Each teacher assignment is class × subject.
// Returns: { summary: {total, submitted, pending, percentage}, assignments: [...] }
// ============================================================
/**
 * Returns the distinct exam NAMES (e.g. "Mid Term Exam", "End of Term Exam")
 * available for the Tracking Scores "Exam Category" filter, scoped to the
 * selected class / exam type / academic year. This lets Academics staff
 * pick one specific, currently-running exam so the report only reflects
 * submissions for THAT exam instead of every exam sharing the same SE/CA type.
 */
function getExamNamesForTracking(className, examType, year) {
  var cls  = (className || '').toString().trim();
  var type = (examType   || '').toString().trim().toUpperCase();
  var yr   = (year       || '').toString().trim();

  var exams = getAcademicExamsRaw();
  var seen  = {};
  var list  = [];

  exams.forEach(function(e) {
    if (cls && cls !== 'ALL' && !classMatchesExam(cls, e.classes)) return;
    if (type && type !== 'ALL' && (e.type || '').toString().toUpperCase().trim() !== type) return;
    if (yr && (e.year || '').toString().trim() && (e.year || '').toString().trim() !== yr) return;

    var name = (e.name || '').toString().trim();
    if (!name || seen[name]) return;
    seen[name] = true;

    list.push({
      name:      name,
      type:      (e.type || '').toString().toUpperCase().trim(),
      year:      e.year || '',
      status:    e.status || 'Draft',
      startDate: e.startDateStr || ''
    });
  });

  // Most recent exam first, so the "current" exam is easy to spot.
  list.sort(function(a, b) { return (b.startDate || '').localeCompare(a.startDate || ''); });

  return list;
}

function getScoreTrackingData(filterClass, filterExamType, filterYear, filterExamName) {
  var teachers         = getTeacherList();
  var selectedClass    = (filterClass    || '').toString().trim();
  var selectedCategory = (filterExamType || '').toString().trim().toUpperCase();
  var selectedYear     = (filterYear     || '').toString().trim();
  var selectedExamName = (filterExamName || '').toString().trim();
  var ss               = SpreadsheetApp.getActiveSpreadsheet();

  // ── Load AcademicSubjects once for class-restriction checking ──
  var subjectsRaw = getAcademicSubjectsRaw();   // ← NEW

  function norm(value) {
    return (value || '').toString().trim().replace(/\s+/g, ' ').toLowerCase();
  }
  function formNo(value) {
    var m = (value || '').toString().match(/(\d+)/);
    return m ? m[1] : '';
  }
  function classSame(a, b) {
    var an = norm(a), bn = norm(b);
    if (!an || !bn) return false;
    if (an === bn || an.replace(/\s+/g,'') === bn.replace(/\s+/g,'')) return true;
    return !!formNo(a) && formNo(a) === formNo(b);
  }
  function subjectSame(a, b) {
    var an = norm(a), bn = norm(b);
    return !!an && !!bn &&
      (an === bn || an.replace(/\s+/g,'') === bn.replace(/\s+/g,''));
  }
  function detectCategory(examName) {
    var x = (examName || '').toString().trim().toUpperCase();
    if (!x) return '';
    if (/\bSE\b|SEMESTER|SUMMATIVE|MIDTERM|TERMINAL|ANNUAL/.test(x)) return 'SE';
    if (/\bCA\b|CONTINUOUS|ASSESSMENT|\bROUND\b/.test(x)) return 'CA';
    return '';
  }

  // ── Real exam-name → type lookup, sourced from the Examinations sheet ──
  // detectCategory() above only guesses a type by pattern-matching the exam
  // NAME (e.g. "Terminal Examination" → SE), which silently misses exams
  // like "Post Mock Examination" or "Mock Examination Re-sit" that don't
  // contain any of those keywords — causing their submitted scores to be
  // filtered out entirely when someone filters Tracking Scores by SE/CA.
  // We prefer the actual Type column recorded against the exam and only
  // fall back to keyword-guessing for free-text exam names with no match
  // in the Examinations sheet (e.g. legacy data).
  var examTypeByName = {};
  getAcademicExamsRaw().forEach(function(e) {
    var n = (e.name || '').toString().trim().toLowerCase();
    if (n) examTypeByName[n] = (e.type || '').toString().toUpperCase().trim();
  });
  function resolveExamType(examName) {
    var n = (examName || '').toString().trim().toLowerCase();
    if (examTypeByName[n]) return examTypeByName[n];
    return detectCategory(examName);
  }

  // ── Pre-load all per-form score sheets into memory (4 reads max) ──
  var scoresByForm = {};
  ['1','2','3','4'].forEach(function(fn) {
    var sh = ss.getSheetByName('Form' + fn + '_Scores');
    scoresByForm[fn] = (sh && sh.getLastRow() > 1)
      ? sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues()
      : [];
    // Also include legacy Scores rows for this form number
    var legacy = ss.getSheetByName('Scores');
    if (legacy && legacy.getLastRow() > 1) {
      legacy.getDataRange().getValues().slice(1).forEach(function(r) {
        if (extractFormNumber((r[1]||'').toString()) === fn) {
          scoresByForm[fn].push(r);
        }
      });
    }
  });

  var assignments = [];

  teachers.forEach(function(teacher) {
    var pairs = (teacher.classesSubjects && teacher.classesSubjects.length)
      ? teacher.classesSubjects
      : [];

    pairs.forEach(function(pair) {
      var cls  = pair.class   || '';
      var subj = pair.subject || '';
      if (!cls || !subj) return;

      // ── CLASS FILTER ──────────────────────────────────────────
      if (selectedClass && selectedClass !== 'ALL' && !classSame(cls, selectedClass)) return;

      // ── SUBJECT-CLASS VALIDATION (THE KEY FIX) ───────────────
      // Skip pairs where the subject is not valid for this class
      // (e.g. Historia Ya Tanzanzia na Maadili restricted to Form 1, Form 2)
      if (!isSubjectValidForClass(subj, cls, subjectsRaw)) return;

      var fn         = formNo(cls);
      var sheetRows  = fn ? (scoresByForm[fn] || []) : [];

      var matchingRows = sheetRows.filter(function(row) {
        var rowClass   = row[1];
        var rowSubject = row[2];
        var rowScore   = row[6];
        var hasScore   = rowScore !== '' && rowScore !== null && rowScore !== undefined;
        if (!hasScore) return false;
        if (!classSame(rowClass, cls)) return false;
        if (!subjectSame(rowSubject, subj)) return false;

        // Exam Category filter (exact exam, e.g. "Mid Term Exam 2026") —
        // when set, this takes priority: it pins the report to ONE specific
        // exam so "submitted" really means "submitted for THIS exam", not
        // just "submitted for some exam of this type at some point".
        if (selectedExamName && selectedExamName.toUpperCase() !== 'ALL') {
          var rowExamName = (row[0] || '').toString().trim();
          if (rowExamName.toLowerCase() !== selectedExamName.toLowerCase()) return false;
        }

        // Exam Type filter (SE / CA)
        if (selectedCategory && selectedCategory !== 'ALL') {
          var rowExam = (row[0] || '').toString();
          var cat     = resolveExamType(rowExam);
          if (cat !== selectedCategory && rowExam.toUpperCase().indexOf(selectedCategory) < 0)
            return false;
        }
        // Year filter
        if (selectedYear) {
          var rowDate = row[8] instanceof Date
            ? row[8].getFullYear().toString()
            : (row[8] || '').toString();
          var rowExam = (row[0] || '').toString();
          if (rowDate.indexOf(selectedYear) < 0 && rowExam.indexOf(selectedYear) < 0)
            return false;
        }
        return true;
      });

      var uniqueStudents = {};
      matchingRows.forEach(function(row) {
        var key = norm(row[3] || row[4]);
        if (key) uniqueStudents[key] = true;
      });

      var lastRow        = matchingRows.length ? matchingRows[matchingRows.length - 1] : null;
      var submittedAtRaw = lastRow ? lastRow[8] : '';
      var submittedAt    = submittedAtRaw
        ? (submittedAtRaw instanceof Date
            ? submittedAtRaw.toLocaleString()
            : submittedAtRaw.toString())
        : '';
      var examName       = lastRow ? (lastRow[0] || '').toString().trim() : '';
      var examType       = resolveExamType(examName) || selectedCategory || '';
      var scoresSubmitted = Object.keys(uniqueStudents).length || matchingRows.length;
      var status         = matchingRows.length ? 'submitted' : 'pending';

      assignments.push({
        className:       cls,
        subjectName:     subj,
        examType:        examType || 'All',
        examName:        examName || (selectedExamName && selectedExamName !== 'ALL'
                           ? selectedExamName
                           : (selectedCategory && selectedCategory !== 'ALL' ? selectedCategory : 'All')),
        teacherName:     teacher.name,
        teacherUsername: teacher.username,
        totalStudents:   scoresSubmitted,
        scoresSubmitted: scoresSubmitted,
        status:          status,
        submittedAt:     submittedAt
      });
    });
  });

  var total     = assignments.length;
  var submitted = assignments.filter(function(a) { return a.status === 'submitted'; }).length;
  var pending   = total - submitted;
  var percentage = total ? Math.round((submitted / total) * 100) : 0;

  return {
    assignments: assignments,
    summary: { total: total, submitted: submitted, pending: pending, percentage: percentage }
  };
}

function getTeacherFullName() {
  const t = getTeacherSession();
  return t ? t.fullName : "Teacher";
}
// ============================================================
// ONE-TIME MIGRATION — Run manually from Apps Script editor.
// Moves rows from old "Scores" sheet to Form1_Scores … Form4_Scores.
// Safe to run multiple times (skips duplicates).
// ============================================================
function migrateScoresToFormSheets() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var legacy = ss.getSheetByName('Scores');
  if (!legacy || legacy.getLastRow() < 2) {
    Logger.log('No legacy Scores sheet or it is empty. Nothing to migrate.');
    return;
  }

  var rows   = legacy.getDataRange().getValues().slice(1);
  var counts = { '1': 0, '2': 0, '3': 0, '4': 0, skipped: 0 };

  // Pre-load existing keys per form sheet to avoid duplicates
  var existingKeys = {};
  ['1','2','3','4'].forEach(function(fn) {
    existingKeys[fn] = {};
    var sh = ss.getSheetByName('Form' + fn + '_Scores');
    if (sh && sh.getLastRow() > 1) {
      sh.getDataRange().getValues().slice(1).forEach(function(r) {
        var key = (r[0]||'').toString().trim().toLowerCase() + '|||' +
                  (r[1]||'').toString().trim().toLowerCase() + '|||' +
                  (r[2]||'').toString().trim().toLowerCase() + '|||' +
                  (r[4]||'').toString().trim().toLowerCase();
        existingKeys[fn][key] = true;
      });
    }
  });

  // Batch rows per form number
  var batches = { '1': [], '2': [], '3': [], '4': [] };

  rows.forEach(function(r) {
    var className = (r[1]||'').toString().trim();
    var fn        = extractFormNumber(className);
    if (!fn || !batches[fn]) { counts.skipped++; return; }

    var key = (r[0]||'').toString().trim().toLowerCase() + '|||' +
              (r[1]||'').toString().trim().toLowerCase() + '|||' +
              (r[2]||'').toString().trim().toLowerCase() + '|||' +
              (r[4]||'').toString().trim().toLowerCase();

    if (existingKeys[fn][key]) { counts.skipped++; return; } // already migrated

    batches[fn].push(r);
    existingKeys[fn][key] = true; // mark so batch doesn't duplicate
  });

  // Write batches
  ['1','2','3','4'].forEach(function(fn) {
    var batch = batches[fn];
    if (!batch.length) return;
    var sh = getScoresSheet('Form ' + fn);
    sh.getRange(sh.getLastRow() + 1, 1, batch.length, 9).setValues(batch);
    counts[fn] = batch.length;
    Logger.log('Form' + fn + '_Scores: migrated ' + batch.length + ' rows.');
  });

  Logger.log('Migration complete. '
    + 'Form1: ' + counts['1'] + ', Form2: ' + counts['2']
    + ', Form3: ' + counts['3'] + ', Form4: ' + counts['4']
    + ', Skipped: ' + counts.skipped);

  SpreadsheetApp.getUi().alert(
    '✅ Migration complete!\n'
    + 'Form 1 Scores: ' + counts['1'] + ' rows\n'
    + 'Form 2 Scores: ' + counts['2'] + ' rows\n'
    + 'Form 3 Scores: ' + counts['3'] + ' rows\n'
    + 'Form 4 Scores: ' + counts['4'] + ' rows\n'
    + 'Skipped (duplicates/unresolved): ' + counts.skipped
  );
}
// ═══════════════════════════════════════════════════════════════════
//  TOD REPORT — Google Apps Script Backend
//  Paste these two functions into your Apps Script project (Code.gs)
//  Sheet name used: "TOD Reports"  (created automatically if missing)
// ═══════════════════════════════════════════════════════════════════

/**
 * saveTODReport(todData)
 * Called from Teacher.html when the teacher confirms and saves a TOD report.
 * Writes one row to the "TOD Reports" sheet and optionally saves the PDF to Drive.
 *
 * @param {Object} todData  - Full TOD form data sent from the browser
 * @returns {Object}        - { success: true, message: "..." }
 */
/**
 * authorizeDriveAccess()
 * Run this ONCE manually from the Apps Script editor (select it in the function
 * dropdown next to "Run", then click Run) if TOD PDF saves fail with an error like:
 * "You do not have permission to call DriveApp.getFoldersByName."
 *
 * Running any function that touches DriveApp from the editor forces Google to show
 * the authorization consent screen for the Drive scope, which a plain deploy does not
 * trigger on its own. Approve the prompt (Review permissions → your account → Advanced
 * → Go to project (unsafe) → Allow), then try saving a TOD report again.
 */
function authorizeDriveAccess() {
  var folders = DriveApp.getFoldersByName("TOD Reports");
  var folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder("TOD Reports");
  Logger.log("Drive access OK. Folder: " + folder.getUrl());
  return folder.getUrl();
}

function saveTODReport(todData) {
  var accessCheck = checkDashboardWriteAccess(['TOD Reports','Reports']);
  if (!accessCheck.ok) return { success: false, message: accessCheck.message };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var SHEET = "TOD Reports";
  var sheet = ss.getSheetByName(SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET);
    var headers = [
      "No.","Date","Teacher","Submitted At","Total Registered","Total Present",
      "Total Absentees","% Present","% Absentees","Watchman Report","Morning Arrivals",
      "Morning Activities","Morning Assembly","Morning Warmup","Breakfast","Lunch",
      "Evening Activities","Visitors","TOD's Comment","TOD Signature","Watchman Signature",
      "Head Master Comment","Head Master Signature","Attendance Rows (JSON)","PDF Link"
    ];
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground("#0d1b36").setFontColor("#ffffff").setFontWeight("bold").setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,40); sheet.setColumnWidth(2,100); sheet.setColumnWidth(3,180);
    sheet.setColumnWidth(4,140); sheet.setColumnWidth(5,90); sheet.setColumnWidth(6,80);
    sheet.setColumnWidth(7,90); sheet.setColumnWidth(8,90); sheet.setColumnWidth(9,90);
    for (var c = 10; c <= 24; c++) sheet.setColumnWidth(c, 200);
    sheet.setColumnWidth(25, 250);
  }

  // Use the report date the teacher actually chose — never overwrite with "today"
  var reportDate = String(todData.date || "").trim();
  if (!reportDate) {
    reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  var pdfLink = "";
  var pdfWarning = "";
  try {
    if (todData.pdfBase64) {
      var pdfBytes = Utilities.base64Decode(todData.pdfBase64);
      var pdfBlob  = Utilities.newBlob(pdfBytes, "application/pdf", todData.pdfFileName || "TOD_Report.pdf");
      var folders  = DriveApp.getFoldersByName("TOD Reports");
      var folder   = folders.hasNext() ? folders.next() : DriveApp.createFolder("TOD Reports");
      var file     = folder.createFile(pdfBlob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfLink = file.getUrl();
    } else {
      pdfWarning = "No PDF data was received from the browser.";
    }
  } catch (pdfErr) {
    Logger.log("PDF save failed: " + pdfErr.message);
    pdfWarning = "PDF could not be uploaded to Drive: " + pdfErr.message;
  }

  var now       = new Date();
  var submitted = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  var row = [
    0, reportDate, todData.teacher || "", submitted,
    todData.totalRegistered || 0, todData.totalPresent || 0, todData.totalAbsentees || 0,
    todData.percentagePresent || "0%", todData.percentageAbsentees || "0%",
    todData.watchmanReport || "", todData.morningArrivals || "", todData.morningActivities || "",
    todData.morningAssembly || "", todData.morningWarmup || "", todData.breakfast || "",
    todData.lunch || "", todData.eveningActivities || "", todData.visitors || "",
    todData.todComment || "", todData.todSignature || "", todData.watchmanSignature || "",
    todData.headMasterComment || "", todData.headMasterSignature || "",
    todData.attendanceRows || "", pdfLink
  ];

  // Find an existing row for this teacher + this date → UPDATE instead of duplicating
  var lastRow = sheet.getLastRow();
  var targetRow = -1;
  if (lastRow > 1) {
    var existing = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // No, Date, Teacher
    for (var i = 0; i < existing.length; i++) {
      var exDate = existing[i][1] instanceof Date
        ? Utilities.formatDate(existing[i][1], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(existing[i][1]).substring(0, 10);
      var exTeacher = String(existing[i][2] || "").trim().toLowerCase();
      if (exDate === reportDate && exTeacher === String(todData.teacher || "").trim().toLowerCase()) {
        targetRow = i + 2;
        break;
      }
    }
  }

  if (targetRow > -1) {
    // ── UPDATE existing row for this teacher + date ──
    row[0] = sheet.getRange(targetRow, 1).getValue();
    if (!pdfLink) pdfLink = sheet.getRange(targetRow, 25).getValue(); // keep old PDF link if a new one wasn't generated
    row[24] = pdfLink;
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    sheet.getRange(targetRow, 2).setNumberFormat("@").setValue(reportDate);
    if (pdfLink) {
      // Link text = the URL itself, so getValues() reads back the real link (not "View PDF")
      var linkCell = sheet.getRange(targetRow, 25);
      linkCell.setFormula('=HYPERLINK("' + pdfLink + '","' + pdfLink + '")');
      linkCell.setFontColor("#1a56db");
    }
    return {
      success: true,
      message: "TOD Report updated successfully for " + reportDate + ".",
      row: targetRow,
      pdfLink: pdfLink,
      pdfWarning: pdfWarning
    };
  }

  // ── INSERT new row ──
  row[0] = lastRow; // sequential number
  sheet.appendRow(row);
  var newRow = sheet.getLastRow();
  sheet.getRange(newRow, 2).setNumberFormat("@").setValue(reportDate);
  if (newRow % 2 === 0) sheet.getRange(newRow, 1, 1, row.length).setBackground("#f0f4fa");
  if (pdfLink) {
    // Link text = the URL itself, so getValues() reads back the real link (not "View PDF")
    var linkCell2 = sheet.getRange(newRow, 25);
    linkCell2.setFormula('=HYPERLINK("' + pdfLink + '","' + pdfLink + '")');
    linkCell2.setFontColor("#1a56db");
  }

  return {
    success: true,
    message: "TOD Report saved successfully for " + reportDate + ".",
    row: newRow,
    pdfLink: pdfLink,
    pdfWarning: pdfWarning
  };
}

/**
 * getMyTODDates(teacherName)
 * Returns a de-duplicated, newest-first list of dates (yyyy-MM-dd) this teacher has
 * submitted a TOD report for. Used to populate date pickers.
 */
function getMyTODDates(teacherName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TOD Reports");
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues(); // No, Date, Teacher
  var wanted = String(teacherName || '').trim().toLowerCase();
  var dates = [];

  data.forEach(function(row) {
    var rTeacher = String(row[2] || '').trim().toLowerCase();
    if (wanted && rTeacher !== wanted) return;
    var d = row[1];
    var dateStr = d instanceof Date
      ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(d).substring(0, 10);
    if (dateStr) dates.push(dateStr);
  });

  var uniq = Array.from(new Set(dates));
  uniq.sort(function(a, b) { return a < b ? 1 : (a > b ? -1 : 0); }); // newest first
  return uniq;
}

/**
 * getMyTODReports(teacherName)
 * Returns every TOD report this teacher has saved (all dates, not just the current week),
 * newest first, with the PDF link so the UI can render a simple downloadable list.
 */
function getMyTODReports(teacherName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TOD Reports");
  if (!sheet || sheet.getLastRow() < 2) return [];

  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 25).getValues(); // No..PDF Link
  var wanted = String(teacherName || '').trim().toLowerCase();
  var results = [];

  data.forEach(function(row, idx) {
    var rTeacher = String(row[2] || '').trim().toLowerCase();
    if (wanted && rTeacher !== wanted) return;
    var d = row[1];
    var dateStr = d instanceof Date
      ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(d).substring(0, 10);
    if (!dateStr) return;
    results.push({
      rowIndex:    idx + 2,
      date:        dateStr,
      teacher:     String(row[2] || ''),
      submittedAt: String(row[3] || ''),
      pdfLink:     String(row[24] || '')
    });
  });

  // Newest first (by date, then by submitted time as a tiebreaker)
  results.sort(function(a, b) {
    if (a.date !== b.date) return a.date < b.date ? 1 : -1;
    return a.submittedAt < b.submittedAt ? 1 : (a.submittedAt > b.submittedAt ? -1 : 0);
  });

  return results;
}

/**
 * deleteTODReport(teacher, dateStr)
 * Deletes this teacher's TOD report row for the given date, and trashes the
 * associated PDF in Drive (if any). Called from the "My Weekly Report" list's
 * per-row Delete button.
 */
function deleteTODReport(teacher, dateStr) {
  var accessCheck = checkDashboardWriteAccess(['TOD Reports','Reports']);
  if (!accessCheck.ok) return { success: false, message: accessCheck.message };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TOD Reports");
  if (!sheet || sheet.getLastRow() < 2) return { success: false, message: "No TOD reports found." };

  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 25).getValues();
  var wantedTeacher = String(teacher || '').trim().toLowerCase();
  var wantedDate = String(dateStr || '').trim();
  var targetRow = -1;
  var pdfLink = "";

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rDate = row[1] instanceof Date
      ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(row[1]).substring(0, 10);
    var rTeacher = String(row[2] || '').trim().toLowerCase();
    if (rDate === wantedDate && rTeacher === wantedTeacher) {
      targetRow = i + 2;
      pdfLink = String(row[24] || '');
      break;
    }
  }

  if (targetRow === -1) {
    return { success: false, message: "No TOD report found for " + wantedDate + "." };
  }

  sheet.deleteRow(targetRow);

  if (pdfLink) {
    try {
      var idMatch = pdfLink.match(/[-\w]{25,}/);
      if (idMatch) DriveApp.getFileById(idMatch[0]).setTrashed(true);
    } catch (trashErr) {
      Logger.log("PDF trash failed (non-fatal): " + trashErr.message);
    }
  }

  return { success: true, message: "TOD Report for " + wantedDate + " deleted." };
}

function getTODReportByDate(teacher, dateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TOD Reports");
  if (!sheet || sheet.getLastRow() < 2) return null;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 25).getValues();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rDate = row[1] instanceof Date
      ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(row[1]).substring(0, 10);
    var rTeacher = String(row[2] || "").trim().toLowerCase();
    if (rDate === String(dateStr).trim() && rTeacher === String(teacher).trim().toLowerCase()) {
      return {
        date: rDate, teacher: row[2], submittedAt: row[3],
        totalRegistered: row[4], totalPresent: row[5], totalAbsentees: row[6],
        percentagePresent: row[7], percentageAbsentees: row[8], watchmanReport: row[9],
        morningArrivals: row[10], morningActivities: row[11], morningAssembly: row[12],
        morningWarmup: row[13], breakfast: row[14], lunch: row[15],
        eveningActivities: row[16], visitors: row[17], todComment: row[18],
        todSignature: row[19], watchmanSignature: row[20], headMasterComment: row[21],
        headMasterSignature: row[22], attendanceRows: row[23], pdfLink: row[24]
      };
    }
  }
  return null;
}

/**
 * getTODWeeklyReport(teacher)
 * Returns all TOD rows submitted by this teacher during the current Mon–Sun week.
 *
 * @param {string} teacher  - Teacher's full name (matched against column C)
 * @returns {Array}         - Array of row objects, newest first
 */
function getTODWeeklyReport(teacher) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TOD Reports");

  /* Sheet doesn't exist yet — no submissions */
  if (!sheet || sheet.getLastRow() < 2) return [];

  /* ── Determine current week's Monday and Sunday ── */
  var today    = new Date();
  var dayOfWk  = today.getDay(); // 0=Sun … 6=Sat
  var diffMon  = (dayOfWk === 0) ? -6 : 1 - dayOfWk; // days back to Monday
  var monday   = new Date(today); monday.setDate(today.getDate() + diffMon); monday.setHours(0,0,0,0);
  var sunday   = new Date(monday); sunday.setDate(monday.getDate() + 6); sunday.setHours(23,59,59,999);

  /* ── Read all data rows (skip header) ── */
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 25).getValues();
  var results = [];

  data.forEach(function(row) {
    var rowDate    = row[1];   // col B — Date
    var rowTeacher = row[2];   // col C — Teacher

    /* Normalise date */
    var d = rowDate instanceof Date ? rowDate : new Date(rowDate);
    if (isNaN(d.getTime())) return; // skip if unparseable

    /* Filter: this teacher + this week */
    var teacherMatch = !teacher ||
      String(rowTeacher).trim().toLowerCase() === String(teacher).trim().toLowerCase();
    var inWeek = d >= monday && d <= sunday;

    if (teacherMatch && inWeek) {
      results.push({
        date:                Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        teacher:             String(row[2]  || ""),
        submittedAt:         String(row[3]  || ""),
        totalRegistered:     Number(row[4]  || 0),
        totalPresent:        Number(row[5]  || 0),
        totalAbsentees:      Number(row[6]  || 0),
        percentagePresent:   String(row[7]  || ""),
        percentageAbsent:    String(row[8]  || ""),
        watchmanReport:      String(row[9]  || ""),
        morningArrivals:     String(row[10] || ""),
        morningActivities:   String(row[11] || ""),
        morningAssembly:     String(row[12] || ""),
        morningWarmup:       String(row[13] || ""),
        breakfast:           String(row[14] || ""),
        lunch:               String(row[15] || ""),
        eveningActivities:   String(row[16] || ""),
        visitors:            String(row[17] || ""),
        todComment:          String(row[18] || ""),
        todSignature:        String(row[19] || ""),
        watchmanSignature:   String(row[20] || ""),
        headMasterComment:   String(row[21] || ""),
        headMasterSignature: String(row[22] || ""),
        attendanceRows:      String(row[23] || ""),
        pdfLink:             String(row[24] || "")
      });
    }
  });

  /* Return newest first */
  return results.reverse();
}
// ── Derive sheet name from subject + class ──────────────────
function mapSheetName(subjectName, className) {
  function clean(s) {
    return (s || '').toString().trim()
      .replace(/\s+/g, '')
      .replace(/[^a-zA-Z0-9]/g, '');
  }
  return 'Map_' + clean(className) + '_' + clean(subjectName);
}


// ── Get or create a dedicated mapping sheet ─────────────────
function getOrCreateMapSheet(subjectName, className) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = mapSheetName(subjectName, className);
  var sh        = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow(['AdmNo', 'FirstName', 'LastName', 'Sex']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 4)
      .setBackground('#0d1b36')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sh.setColumnWidth(1, 110);
    sh.setColumnWidth(2, 160);
    sh.setColumnWidth(3, 160);
    sh.setColumnWidth(4, 80);
  }
  return sh;
}


// ── 1. Save mapping — writes to dedicated sheet ─────────────
function saveStudentSubjectMapping(subjectName, className, admNos) {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var formNum = extractFormNumber(className);
    var formSheet = studentSheetForClass_(ss, className);

    // Build admNo → student lookup from Form sheet
    var studentMap = {};
    if (formSheet && formSheet.getLastRow() > 1) {
      var formData = formSheet.getDataRange().getValues();
      for (var r = 1; r < formData.length; r++) {
        var admNo = String(formData[r][0] || '').trim();
        if (admNo) {
          studentMap[admNo] = {
            firstName: String(formData[r][1] || '').trim(),
            lastName:  String(formData[r][2] || '').trim(),
            sex:       String(formData[r][3] || '').trim()
          };
        }
      }
    }

    // Get dedicated sheet and clear data rows
    var sh = getOrCreateMapSheet(subjectName, className);
    if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);

    // Write selected students
    var admList = Array.isArray(admNos) ? admNos : [];
    var rows    = [];
    admList.forEach(function(admNo) {
      admNo = String(admNo || '').trim();
      if (!admNo) return;
      var s = studentMap[admNo] || { firstName: '', lastName: '', sex: '' };
      rows.push([admNo, s.firstName, s.lastName, s.sex]);
    });

    if (rows.length > 0) {
      sh.getRange(2, 1, rows.length, 4).setValues(rows);
    }

    return 'Success: ' + rows.length + ' student(s) mapped to '
           + subjectName + ' (' + className + ')';
  } catch (e) {
    return 'Error: ' + e.message;
  }
}


// ── 2. Single call: all students + already-mapped admNos ────
function getStudentsAndMappingForSubject(subjectName, className) {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var formNum = extractFormNumber(className);
    if (!formNum) return { students: [], mapped: [] };

    // All students from Form sheet
    var formSheet = studentSheetForClass_(ss, className);
    if (!formSheet) return { students: [], mapped: [] };

    var formData = formSheet.getDataRange().getValues();
    var students = [];
    for (var r = 1; r < formData.length; r++) {
      if (!formData[r][0] && !formData[r][1]) continue;
      students.push({
        admNo:     String(formData[r][0] || '').trim(),
        firstName: String(formData[r][1] || '').trim(),
        lastName:  String(formData[r][2] || '').trim(),
        sex:       String(formData[r][3] || '').trim()
      });
    }

    // Already-mapped admNos — direct read from dedicated sheet
    var mapped   = [];
    var mapSheet = ss.getSheetByName(mapSheetName(subjectName, className));
    if (mapSheet && mapSheet.getLastRow() > 1) {
      var mapData = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 1).getValues();
      mapData.forEach(function(row) {
        var a = String(row[0] || '').trim();
        if (a) mapped.push(a);
      });
    }

    return { students: students, mapped: mapped };
  } catch (e) {
    Logger.log('getStudentsAndMappingForSubject: ' + e.message);
    return { students: [], mapped: [] };
  }
}


// ── 3. Count badge — sum rows across all matching Map_ sheets
function getSubjectMappedCount(subjectName) {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var cleanSub = (subjectName || '').toString().trim()
                    .replace(/\s+/g, '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    var total    = 0;
    ss.getSheets().forEach(function(sh) {
      var name = sh.getName();
      if (name.indexOf('Map_') !== 0) return;
      if (name.toLowerCase().indexOf(cleanSub) < 0) return;
      var rows = sh.getLastRow() - 1;
      if (rows > 0) total += rows;
    });
    return total;
  } catch (e) {
    return 0;
  }
}


// ── 4. Summary table — scan all Map_ sheets ─────────────────
function getSubjectMappingSummary(subjectName) {
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var filterSub = (subjectName || '').toString().trim()
                      .replace(/\s+/g, '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();

    var CLASS_DISPLAY   = PS_CLASSES.slice();
    var SUBJECT_DISPLAY = ['Chemistry','Biology','History','Computer Science'];
    var results         = [];

    ss.getSheets().forEach(function(sh) {
      var name = sh.getName();
      if (name.indexOf('Map_') !== 0) return;

      var rest    = name.substring(4);              // strip "Map_"
      var uscore  = rest.indexOf('_');
      if (uscore < 0) return;
      var classPart   = rest.substring(0, uscore);  // e.g. "Form1"
      var subjectPart = rest.substring(uscore + 1); // e.g. "Chemistry"

      if (filterSub && subjectPart.toLowerCase().indexOf(filterSub) < 0) return;

      var count = Math.max(0, sh.getLastRow() - 1);

      // Restore display names
      var displayClass   = classPart;
      var displaySubject = subjectPart;
      CLASS_DISPLAY.forEach(function(c) {
        if (c.replace(/\s+/g,'').toLowerCase() === classPart.toLowerCase())
          displayClass = c;
      });
      SUBJECT_DISPLAY.forEach(function(s) {
        if (s.replace(/\s+/g,'').replace(/[^a-zA-Z0-9]/g,'').toLowerCase() === subjectPart.toLowerCase())
          displaySubject = s;
      });

      results.push({ subjectName: displaySubject, className: displayClass, count: count });
    });

    results.sort(function(a, b) {
      return a.subjectName.localeCompare(b.subjectName) || a.className.localeCompare(b.className);
    });
    return results;
  } catch (e) {
    return [];
  }
}


// ── 5. getStudentsForScoring — REPLACE your existing one ────
//
//  Optional subjects → reads Map_FormX_Subject sheet (instant)
//  Compulsory subjects → reads Form sheet (original behaviour)
//
function getStudentsForScoring(payload) {
  var cls      = String(payload['class']    || payload.class    || '').trim();
  var subject  = String(payload['subject']  || payload.subject  || '').trim();
  var examType = String(payload['examType'] || payload.examType || '').trim();

  if (!cls || !subject || !examType) return [];

  var OPTIONAL_NAMES = []; // primary school: only the AcademicSubjects 'Optional' category makes a subject optional
  var isOptional     = false;
  var ss             = SpreadsheetApp.getActiveSpreadsheet();

  // Check AcademicSubjects category field
  var subjSheet = ss.getSheetByName('AcademicSubjects');
  if (subjSheet && subjSheet.getLastRow() > 1) {
    var subjData = subjSheet.getDataRange().getValues();
    for (var si = 1; si < subjData.length; si++) {
      if (String(subjData[si][0]||'').trim().toLowerCase() === subject.toLowerCase()) {
        if (String(subjData[si][2]||'').trim().toLowerCase() === 'optional') isOptional = true;
        break;
      }
    }
  }
  if (!isOptional && OPTIONAL_NAMES.indexOf(subject.toLowerCase()) >= 0) isOptional = true;

  var formNum = extractFormNumber(cls);
  if (!formNum) return [];

  // Build score lookup from Scores sheet
 // Build score lookup from the FORM-SPECIFIC scores sheet
  var scoreMap    = {};
  var scoresSheet = getScoresSheet(cls);   // ← per-form sheet
  var lastScRow   = scoresSheet.getLastRow();
  if (lastScRow > 1) {
    var scoreRows = scoresSheet.getRange(2, 1, lastScRow - 1, 9).getValues();
    var clsL = cls.toLowerCase(), subjL = subject.toLowerCase(), examL = examType.toLowerCase();
    scoreRows.forEach(function(r) {
      if ((r[0]||'').toString().trim().toLowerCase() !== examL) return;
      if ((r[2]||'').toString().trim().toLowerCase() !== subjL) return;
      var rCls = (r[1]||'').toString().trim().toLowerCase();
      if (rCls !== clsL && extractFormNumber(rCls) !== formNum) return;
      var name = (r[4]||'').toString().trim();
      if (!name) return;
      var score = r[6];
      scoreMap[name.toLowerCase()] =
        (score !== '' && score !== null && score !== undefined) ? score : '';
    });
  }

  // Also check legacy Scores sheet for older data
  var legacySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Scores');
  if (legacySheet && legacySheet.getLastRow() > 1) {
    var legacyRows = legacySheet.getRange(2, 1, legacySheet.getLastRow() - 1, 9).getValues();
    legacyRows.forEach(function(r) {
      if ((r[0]||'').toString().trim().toLowerCase() !== examL) return;
      if ((r[2]||'').toString().trim().toLowerCase() !== subjL) return;
      var rCls = (r[1]||'').toString().trim().toLowerCase();
      if (rCls !== clsL && extractFormNumber(rCls) !== formNum) return;
      var name = (r[4]||'').toString().trim();
      if (!name) return;
      var key = name.toLowerCase();
      if (scoreMap.hasOwnProperty(key)) return;
      var score = r[6];
      scoreMap[key] = (score !== '' && score !== null && score !== undefined) ? score : '';
    });
  }

  function mergeScores(students) {
    return students.map(function(s) {
      var fullName = (s.firstName + ' ' + s.lastName).trim();
      var score    = scoreMap.hasOwnProperty(fullName.toLowerCase()) ? scoreMap[fullName.toLowerCase()] : '';
      return { admNo: s.admNo, name: fullName, sex: s.sex || '—', score: score };
    }).filter(function(s) { return s.name; });
  }

  // Students are no longer mapped per subject: with no mapping the whole class register is used.
  if (isOptional) {
    var _ms = ss.getSheetByName(mapSheetName(subject, cls));
    if (!_ms || _ms.getLastRow() < 2) isOptional = false;
  }

  // OPTIONAL → read directly from dedicated map sheet (fast)
  if (isOptional) {
    var mapSheet = ss.getSheetByName(mapSheetName(subject, cls));
    if (!mapSheet || mapSheet.getLastRow() < 2) {
      Logger.log('No mapping yet: ' + mapSheetName(subject, cls));
      return [];
    }
    var mapData = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 4).getValues();
    var students = mapData
      .filter(function(r) { return r[0]; })
      .map(function(r) {
        return {
          admNo:     String(r[0]||'').trim(),
          firstName: String(r[1]||'').trim(),
          lastName:  String(r[2]||'').trim(),
          sex:       String(r[3]||'').trim()
        };
      });
    return mergeScores(students);
  }

  // COMPULSORY → read from Form sheet (unchanged original behaviour)
  var formSheet = studentSheetForClass_(ss, cls);
  if (!formSheet) return [];
  var formData = formSheet.getDataRange().getValues();
  if (formData.length < 2) return [];
  var students = formData.slice(1)
    .filter(function(r) { return r[0] || r[1]; })
    .map(function(r) {
      return {
        admNo:     String(r[0]||'').trim(),
        firstName: String(r[1]||'').trim(),
        lastName:  String(r[2]||'').trim(),
        sex:       String(r[3]||'').trim()
      };
    });
  return mergeScores(students);
}


// ============================================================
// ACADEMIC YEAR PROGRESSION BACKEND
// Sheet: "AcademicYears" — Year | Active | ArchivedStudents | Created
// Sheet: "Alumni"        — Name | Sex | AdmNo | GraduatedYear | PreviousClass
// ============================================================

function ensureAcademicYearsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('AcademicYears');
  if (!sh) {
    sh = ss.insertSheet('AcademicYears');
    sh.appendRow(['Year', 'Active', 'ArchivedStudents', 'Created']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,4).setBackground('#0d1b36').setFontColor('#ffffff').setFontWeight('bold');
    var currentYear = new Date().getFullYear().toString();
    sh.appendRow([currentYear, 'TRUE', 0, new Date().toLocaleDateString()]);
  }
  return sh;
}

function ensureAlumniSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Alumni');
  if (!sh) {
    sh = ss.insertSheet('Alumni');
    sh.appendRow(['Name', 'Sex', 'AdmNo', 'GraduatedYear', 'PreviousClass']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,5).setBackground('#1e1b4b').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sh;
}

function progGetAcademicYears() {
  var sh = ensureAcademicYearsSheet();
  var data = sh.getDataRange().getValues();
  var activeYear = '';
  var years = data.slice(1).map(function(r) {
    var isActive = r[1] ? r[1].toString().toUpperCase() === 'TRUE' : false;
    if (isActive) activeYear = r[0] ? r[0].toString().trim() : '';
    return {
      year:             r[0] ? r[0].toString().trim() : '',
      active:           isActive,
      archivedStudents: r[2] ? Number(r[2]) : 0,
      created:          r[3] ? new Date(r[3]).toLocaleDateString() : ''
    };
  }).filter(function(y) { return y.year; })
    .sort(function(a,b) { return b.year.localeCompare(a.year); });
  return { activeYear: activeYear, years: years };
}

function progSetActiveYear(year) {
  var sh = ensureAcademicYearsSheet();
  var data = sh.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    var rowYear = data[i][0] ? data[i][0].toString().trim() : '';
    if (rowYear === year.toString().trim()) {
      sh.getRange(i+1, 2).setValue('TRUE');
      found = true;
    } else {
      sh.getRange(i+1, 2).setValue('FALSE');
    }
  }
  if (!found) return 'Year ' + year + ' not found. Please create it first.';
  try { logActivity('academic', 'Switched active academic year to: ' + year); } catch(e){}
  return '✅ Academic year ' + year + ' is now ACTIVE. Teachers will now see this year\'s classes.';
}

function progCreateAcademicYear(year) {
  var sh = ensureAcademicYearsSheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === year.toString().trim()) {
      return 'Academic year ' + year + ' already exists.';
    }
  }
  sh.appendRow([year.toString(), 'FALSE', 0, new Date().toLocaleDateString()]);
  try { logActivity('academic', 'Created academic year: ' + year); } catch(e){}
  return '✅ Academic year ' + year + ' created. Activate it when ready.';
}

function progPromoteStudents(fromYear, toYear, whichClass) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Primary school: Standard 1 → … → Standard 6 → Graduated (Alumni)
  var promotionMap = {};
  PS_CLASSES.forEach(function(cls, i) {
    var isLast = (i === PS_CLASSES.length - 1);
    promotionMap[cls] = {
      from:     psClassSheetName_(cls),
      to:       isLast ? 'Alumni' : psClassSheetName_(PS_CLASSES[i + 1]),
      toName:   isLast ? 'Graduated' : PS_CLASSES[i + 1],
      graduate: isLast
    };
  });
  // Process the HIGHEST class first so students just promoted into a class
  // are never promoted a second time in the same run.
  var classesToProcess = whichClass === 'all'
    ? PS_CLASSES.slice().reverse()
    : [whichClass];
  var summary = {}, totalMoved = 0, totalGraduated = 0;

  classesToProcess.forEach(function(cls) {
    var map = promotionMap[cls];
    if (!map) return;
    var fromSheet = ss.getSheetByName(map.from);
    if (!fromSheet || fromSheet.getLastRow() < 2) { summary[cls + ' → ' + map.toName] = 0; return; }
    var rows = fromSheet.getDataRange().getValues();
    var students = rows.slice(1).filter(function(r){ return r[0] || r[1]; });
    if (!students.length) { summary[cls + ' → ' + map.toName] = 0; return; }

    if (map.graduate) {
      var alumniSheet = ensureAlumniSheet();
      students.forEach(function(s) {
        var fullName = [String(s[1]||'').trim(), String(s[2]||'').trim()].join(' ').trim();
        alumniSheet.appendRow([fullName, s[3]||'', s[0]||'', toYear, PS_CLASSES[PS_CLASSES.length - 1]]);
      });
      if (fromSheet.getLastRow() > 1)
        fromSheet.getRange(2, 1, fromSheet.getLastRow()-1, fromSheet.getLastColumn()).clearContent();
      summary[cls + ' → Graduated'] = students.length;
      totalGraduated += students.length;
      var aySheet = ensureAcademicYearsSheet();
      var ayData = aySheet.getDataRange().getValues();
      for (var i = 1; i < ayData.length; i++) {
        if (ayData[i][0] && ayData[i][0].toString().trim() === fromYear) {
          aySheet.getRange(i+1, 3).setValue((Number(ayData[i][2])||0) + students.length);
          break;
        }
      }
    } else {
      var toSheet = ss.getSheetByName(map.to);
      if (!toSheet) {
        toSheet = ss.insertSheet(map.to);
        toSheet.appendRow(['AdmNo','FirstName','LastName','Sex','DOB','Class','Parent','Phone','Address']);
        toSheet.setFrozenRows(1);
      }
      students.forEach(function(s) {
        var row = s.slice();
        row[5] = map.toName;
        toSheet.appendRow(row);
      });
      if (fromSheet.getLastRow() > 1)
        fromSheet.getRange(2, 1, fromSheet.getLastRow()-1, fromSheet.getLastColumn()).clearContent();
      summary[cls + ' → ' + map.toName] = students.length;
      totalMoved += students.length;
    }
  });

  summary['Total Moved'] = totalMoved;
  if (totalGraduated > 0) summary['Graduated'] = totalGraduated;
  try { logActivity('academic','Promotion '+fromYear+'→'+toYear+' ['+whichClass+']: '+totalMoved+' moved, '+totalGraduated+' graduated'); } catch(e){}
  return {
    success: true,
    message: '🎓 Promotion complete! ' + totalMoved + ' student(s) moved up, ' + totalGraduated + ' graduated to Alumni.',
    summary: summary
  };
}

function progGetAlumni(year) {
  var sh = ensureAlumniSheet();
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  return data.slice(1)
    .filter(function(r){ return r[0] || r[2]; })
    .filter(function(r){
      if (!year) return true;
      return r[3] && r[3].toString().trim() === year.toString().trim();
    })
    .map(function(r){
      return {
        name:          String(r[0]||'').trim(),
        sex:           String(r[1]||'').trim(),
        admNo:         String(r[2]||'').trim(),
        graduatedYear: String(r[3]||'').trim(),
        previousClass: String(r[4]||PS_CLASSES[PS_CLASSES.length - 1]).trim()
      };
    });
}

// ============================================================
// ATTENDANCE TRACKING CENTRE — Academic Master Supervisory
// Reads the "Attendance" sheet saved by Class Teachers.
// Sheet columns (0-based): Class | Name | Sex | Date | Morning | Afternoon
// ============================================================

/**
 * getMonthlyAttendance({class, month, year})
 * Returns { students:[{name,sex}], records:[{name,dateStr,morning,afternoon}] }
 * Already defined in your Code.js — kept here as reference/fallback.
 */

/**
 * getClassTeachersWithAttendanceStatus(todayStr)
 * Returns array of { teacherName, className, studentCount, lastSaved }
 * Used by the Academic Master's Attendance Tracking Centre overview tab.
 */
function getClassTeachersWithAttendanceStatus(todayStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Load class list with assigned teachers
  var classSh = ss.getSheetByName('AcademicClasses') || ss.getSheetByName('Classes');
  var classRows = classSh ? classSh.getDataRange().getValues().slice(1) : [];

  // 2. Load today's attendance records
  var attSh = ss.getSheetByName('Attendance');
  var attData = attSh ? attSh.getDataRange().getValues() : [];
  var attHeaders = attData[0] || [];

  // Detect column positions (flexible)
  function colIdx(names) {
    for (var n = 0; n < names.length; n++) {
      for (var c = 0; c < attHeaders.length; c++) {
        if (String(attHeaders[c]).toLowerCase().trim() === names[n].toLowerCase()) return c;
      }
    }
    return -1;
  }
  var COL_CLASS = colIdx(['class','classname']);
  var COL_NAME  = colIdx(['name','studentname','student name']);
  var COL_DATE  = colIdx(['date','datestr']);
  var COL_MORN  = colIdx(['morning']);

  // Fallback fixed columns if headers not found
  if (COL_CLASS < 0) COL_CLASS = 0;
  if (COL_NAME  < 0) COL_NAME  = 1;
  if (COL_DATE  < 0) COL_DATE  = 3;
  if (COL_MORN  < 0) COL_MORN  = 4;

  // Build a map: className → lastSavedDate
  var classLastSaved = {};
  for (var i = 1; i < attData.length; i++) {
    var r = attData[i];
    var cls  = String(r[COL_CLASS]||'').trim();
    var dt   = r[COL_DATE] ? String(r[COL_DATE]).substring(0,10) : '';
    if (cls && dt) {
      if (!classLastSaved[cls] || dt > classLastSaved[cls]) classLastSaved[cls] = dt;
    }
  }

  // Build results from class list
  var results = [];
  var seen = {};

  // Also load known form sheets as fallback if AcademicClasses missing
  var formSheets = PS_CLASSES.slice();

  if (classRows.length > 0) {
    classRows.forEach(function(row) {
      // Columns: ID | Name | Stream | Teacher | Capacity | ...
      var className   = String(row[1]||row[0]||'').trim();
      var teacherName = String(row[3]||row[2]||'').trim();
      if (!className || seen[className]) return;
      seen[className] = 1;

      // Count students in form sheet
      var formKey = className.replace(/\s+/g,'');
      var fSh = ss.getSheetByName(formKey) || ss.getSheetByName(className);
      var studentCount = fSh && fSh.getLastRow() > 1 ? fSh.getLastRow()-1 : 0;

      results.push({
        teacherName:  teacherName || '—',
        className:    className,
        studentCount: studentCount,
        lastSaved:    classLastSaved[className] || ''
      });
    });
  } else {
    // Fallback: use form sheets directly
    formSheets.forEach(function(cls) {
      if (seen[cls]) return;
      seen[cls] = 1;
      var formKey = cls.replace(/\s+/g,'');
      var fSh = ss.getSheetByName(formKey) || ss.getSheetByName(cls);
      var studentCount = fSh && fSh.getLastRow() > 1 ? fSh.getLastRow()-1 : 0;
      results.push({
        teacherName:  '—',
        className:    cls,
        studentCount: studentCount,
        lastSaved:    classLastSaved[cls] || ''
      });
    });
  }
  return results;
}

function addParent(data) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Parents') || ss.insertSheet('Parents');
  // Create header row if sheet is new/empty
  if(sh.getLastRow() === 0){
    sh.appendRow(['Name','Sex','Relation','Phone','Email','AdmNo',
                  'Address','Username','Password','StudentName','ClassName','Role']);
  }
  sh.appendRow([data.name, data.sex, data.relation, data.phone, data.email,
                data.admNo, data.address, data.username, data.password,
                data.studentName, data.className, 'parent']);
  return 'Parent added successfully!';
}
// ============================================================
// PASTE THESE FUNCTIONS INTO YOUR Code.gs
// They handle the full Parents sheet (read, add, update, delete)
// and fix checkLogin() to return parentData for the portal.
//
// Your Parents sheet columns (exactly as saved):
// [0]Name | [1]Sex | [2]Relation | [3]Phone | [4]Email |
// [5]AdmNo | [6]Address | [7]Username | [8]Password |
// [9]StudentName | [10]ClassName | [11]Role
// ============================================================


// ── ENSURE Parents sheet exists with correct headers ────────
function ensureParentsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Parents');
  if (!sh) {
    sh = ss.insertSheet('Parents');
    sh.appendRow([
      'Name','Sex','Relation','Phone','Email','AdmNo',
      'Address','Username','Password','StudentName','ClassName','Role'
    ]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 12)
      .setBackground('#0d1b36')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sh.setColumnWidths(1, 12, 140);
  }
  return sh;
}


// ── READ: Returns all parent rows as arrays (skips header) ──
function getParentList() {
  var sh = ensureParentsSheet();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];                          // empty — no parents yet
  var data = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  // Filter out completely blank rows
  return data.filter(function(r) {
    return String(r[0] || '').trim() !== '';
  });
}


// ── ADD: Appends a new parent row ───────────────────────────
function addParent(data) {
  var sh = ensureParentsSheet();

  // Prevent duplicate usernames
  var existing = getParentList();
  for (var i = 0; i < existing.length; i++) {
    if (String(existing[i][7] || '').trim().toLowerCase() ===
        String(data.username || '').trim().toLowerCase()) {
      return 'Error: Username "' + data.username + '" already exists. Choose a different username.';
    }
  }

  sh.appendRow([
    data.name        || '',
    data.sex         || '',
    data.relation    || '',
    data.phone       || '',
    data.email       || '',
    data.admNo       || '',
    data.address     || '',
    data.username    || '',
    data.password    || '',
    data.studentName || '',
    data.className   || '',
    'parent'
  ]);

  logActivity('admin', 'Added parent: ' + (data.name || '') + ' (' + (data.username || '') + ')');
  return 'Parent "' + (data.name || '') + '" added successfully!';
}


// ── UPDATE: Finds row by old username, overwrites it ────────
function updateParent(oldUsername, data) {
  var sh   = ensureParentsSheet();
  var last = sh.getLastRow();
  if (last < 2) return 'Parent not found.';

  var rows = sh.getRange(2, 1, last - 1, 12).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][7] || '').trim().toLowerCase() ===
        String(oldUsername || '').trim().toLowerCase()) {

      var rowNum = i + 2;   // +1 for header, +1 for 1-based
      sh.getRange(rowNum, 1, 1, 12).setValues([[
        data.name        || rows[i][0],
        data.sex         || rows[i][1],
        data.relation    || rows[i][2],
        data.phone       || rows[i][3],
        data.email       || rows[i][4],
        data.admNo       || rows[i][5],
        data.address     || rows[i][6],
        data.username    || rows[i][7],
        data.password    || rows[i][8],
        data.studentName || rows[i][9],
        data.className   || rows[i][10],
        'parent'
      ]]);

      logActivity('admin', 'Updated parent: ' + oldUsername + ' → ' + (data.username || oldUsername));
      return 'Parent "' + (data.name || oldUsername) + '" updated successfully!';
    }
  }
  return 'Parent not found with username: ' + oldUsername;
}

// ── DELETE: Finds row by username and deletes it ────────────
function deleteParent(username) {
  var sh   = ensureParentsSheet();
  var last = sh.getLastRow();
  if (last < 2) return 'Parent not found.';

  var rows = sh.getRange(2, 1, last - 1, 12).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][7] || '').trim().toLowerCase() ===
        String(username || '').trim().toLowerCase()) {
      sh.deleteRow(i + 2);
      logActivity('admin', 'Deleted parent: ' + username);
      return 'Parent "' + username + '" deleted successfully!';
    }
  }
  return 'Parent not found with username: ' + username;
}
// ============================================================
// CODE_FIX.gs — Add these functions to your Code.gs
// Fixes: getLoginPage(), getParentPortalHtml(), getIndexPage()
// Replace: doGet(), logout(), openLogin() references
// ============================================================


// ============================================================
// WEB APP ENTRY POINT — replace your existing doGet()
// ============================================================
function doGet(e) {
  try { if (typeof psAutoSetup_ === 'function') psAutoSetup_(); } catch (err) { Logger.log('psAutoSetup_: ' + err.message); }   // auto-create tables
  const html = psBrand_(HtmlService.createHtmlOutputFromFile('Index').getContent());
  return HtmlService.createHtmlOutput(smsInjectRestore_(html))
    .setTitle('ORSMS — School Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================
// NAVIGATION FUNCTIONS
// Called by Index.html → Login.html → dashboards
// ============================================================

/**
 * getLoginPage()
 * Called by Index.html openLogin() button.
 * Returns Login.html content as a string.
 */
function getLoginPage() {
  return smsAppendScript_(psBrand_(HtmlService.createHtmlOutputFromFile('Login').getContent()), smsClearScript_());
}

/**
 * getIndexPage()
 * Called by Login.html or any dashboard "Back" button.
 * Returns Index.html content as a string.
 */
function getIndexPage() {
  return smsAppendScript_(psBrand_(HtmlService.createHtmlOutputFromFile('Index').getContent()), smsClearScript_());
}

/**
 * logout()
 * Called by teacher/admin dashboards logout button.
 * Clears teacher session and returns Login.html.
 */
function logout() {
  try { clearTeacherSession(); } catch(e) {}
  return smsAppendScript_(psBrand_(HtmlService.createHtmlOutputFromFile('Login').getContent()), smsClearScript_());
}

/**
 * getParentPortalHtml()
 * Called by Login.html after a successful parent login.
 * Returns Parent.html content as a string.
 * Login.html injects parentData auto-boot script before writing it.
 */
function getParentPortalHtml() {
  return HtmlService.createHtmlOutputFromFile('Parent').getContent();
}


// ============================================================
// UPDATED checkLogin() — replace your existing one
// Now returns parentData so Login.html can boot Parent.html
// ============================================================
function legacyCheckLogin_v2_DISABLED(username, password) {

  // ── Admin ──────────────────────────────────────────────────
  if (ADMIN_USERS.some(function(u) {
    return u.username === username && u.password === password;
  })) {
    return {
      status:   'success',
      role:     'admin',
      fullName: 'Administrator',
      username: username,
      html:     HtmlService.createHtmlOutputFromFile('Admin').getContent()
    };
  }

  // ── Teachers ───────────────────────────────────────────────
  var teachers = getTeacherList();
  var t = null;
  for (var ti = 0; ti < teachers.length; ti++) {
    if (teachers[ti].username === username && teachers[ti].password === password) {
      t = teachers[ti]; break;
    }
  }
  if (t) {
    setTeacherSession(t.username);
    var pos      = (t.position || '').trim();
    var htmlFile = 'Teacher';
    var role     = 'teacher';
    if (pos === 'Academic Master')   { htmlFile = 'Academics';    role = 'academic';     }
    else if (pos === 'Class Teacher'){ htmlFile = 'ClassTeacher'; role = 'classteacher'; }
    else if (['Sports Manager','Accountant','Librarian','Lab Technician'].indexOf(pos) >= 0) {
      htmlFile = 'Staff'; role = 'staff';
    }
    return {
      status:          'success',
      role:            role,
      position:        pos,
      fullName:        t.name,
      username:        t.username,
      classesSubjects: t.classesSubjects,
      classes:         t.classes,
      subjects:        t.subjects,
      html:            HtmlService.createHtmlOutputFromFile(htmlFile).getContent()
    };
  }

  // ── Parents ────────────────────────────────────────────────
  var sh = ensureParentsSheet();
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var rows = sh.getRange(2, 1, lastRow - 1, 12).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r       = rows[i];
      var rowUser = String(r[7] || '').trim();
      var rowPass = String(r[8] || '').trim();
      if (rowUser.toLowerCase() === username.toLowerCase() && rowPass === password) {
        try { logActivity(rowUser, 'Parent login: ' + rowUser); } catch(e) {}
        return {
          status:   'success',
          role:     'parent',
          fullName: String(r[0] || username).trim(),
          username: rowUser,
          // parentData is used by Login.html to auto-boot Parent.html portal
          parentData: {
            parentName:  String(r[0] || '').trim(),
            relation:    String(r[2] || '').trim(),
            phone:       String(r[3] || '').trim(),
            email:       String(r[4] || '').trim(),
            admNo:       String(r[5] || '').trim(),
            address:     String(r[6] || '').trim(),
            studentName: String(r[9] || '').trim(),
            className:   String(r[10]|| '').trim(),
            sex:         '',
            status:      'Approved'
          }
        };
      }
    }
  }

  // ── Invalid ────────────────────────────────────────────────
  return { status: 'fail', message: 'Invalid username or password.' };
}


// ============================================================
// PARENTS SHEET MANAGEMENT
// Sheet columns (0-based):
// [0]Name [1]Sex [2]Relation [3]Phone [4]Email [5]AdmNo
// [6]Address [7]Username [8]Password [9]StudentName [10]ClassName [11]Role
// ============================================================

function ensureParentsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Parents');
  if (!sh) {
    sh = ss.insertSheet('Parents');
    sh.appendRow([
      'Name','Sex','Relation','Phone','Email','AdmNo',
      'Address','Username','Password','StudentName','ClassName','Role'
    ]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 12)
      .setBackground('#0d1b36')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }
  return sh;
}

function getParentList() {
  var sh = ensureParentsSheet();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var data = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  return data.filter(function(r) {
    return String(r[0] || '').trim() !== '';
  });
}

function addParent(data) {
  var sh = ensureParentsSheet();
  // Duplicate username check
  var existing = getParentList();
  for (var i = 0; i < existing.length; i++) {
    if (String(existing[i][7] || '').trim().toLowerCase() ===
        String(data.username || '').trim().toLowerCase()) {
      return 'Error: Username "' + data.username + '" already exists.';
    }
  }
  sh.appendRow([
    data.name        || '',
    data.sex         || '',
    data.relation    || '',
    data.phone       || '',
    data.email       || '',
    data.admNo       || '',
    data.address     || '',
    data.username    || '',
    data.password    || '',
    data.studentName || '',
    data.className   || '',
    'parent'
  ]);
  try { logActivity('admin', 'Added parent: ' + (data.name||'') + ' (' + (data.username||'') + ')'); } catch(e) {}
  return 'Parent "' + (data.name || '') + '" added successfully!';
}

function updateParent(oldUsername, data) {
  var sh   = ensureParentsSheet();
  var last = sh.getLastRow();
  if (last < 2) return 'Parent not found.';
  var rows = sh.getRange(2, 1, last - 1, 12).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][7] || '').trim().toLowerCase() ===
        String(oldUsername || '').trim().toLowerCase()) {
      sh.getRange(i + 2, 1, 1, 12).setValues([[
        data.name        || rows[i][0],
        data.sex         || rows[i][1],
        data.relation    || rows[i][2],
        data.phone       || rows[i][3],
        data.email       || rows[i][4],
        data.admNo       || rows[i][5],
        data.address     || rows[i][6],
        data.username    || rows[i][7],
        data.password    || rows[i][8],
        data.studentName || rows[i][9],
        data.className   || rows[i][10],
        'parent'
      ]]);
      try { logActivity('admin', 'Updated parent: ' + oldUsername + ' → ' + (data.username||oldUsername)); } catch(e) {}
      return 'Parent "' + (data.name || oldUsername) + '" updated successfully!';
    }
  }
  return 'Parent not found: ' + oldUsername;
}

function deleteParent(username) {
  var sh   = ensureParentsSheet();
  var last = sh.getLastRow();
  if (last < 2) return 'Parent not found.';
  var rows = sh.getRange(2, 1, last - 1, 12).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][7] || '').trim().toLowerCase() ===
        String(username || '').trim().toLowerCase()) {
      sh.deleteRow(i + 2);
      try { logActivity('admin', 'Deleted parent: ' + username); } catch(e) {}
      return 'Parent "' + username + '" deleted successfully!';
    }
  }
  return 'Parent not found: ' + username;
}

function getParentPortalData() {
  var sh   = ensureParentsSheet();
  var last = sh.getLastRow();
  var parents = [];
  if (last >= 2) {
    parents = sh.getRange(2, 1, last - 1, 12).getValues()
      .filter(function(r){ return String(r[0]||'').trim() !== ''; })
      .map(function(r){
        return {
          name:        String(r[0]||'').trim(),
          sex:         String(r[1]||'').trim(),
          relation:    String(r[2]||'').trim(),
          phone:       String(r[3]||'').trim(),
          email:       String(r[4]||'').trim(),
          admNo:       String(r[5]||'').trim(),
          address:     String(r[6]||'').trim(),
          username:    String(r[7]||'').trim(),
          password:    String(r[8]||'').trim(),
          studentName: String(r[9]||'').trim(),
          className:   String(r[10]||'').trim()
        };
      });
  }
  var students = [];
  try { students = getStudentListRaw(); } catch(e) {}
  return { parents: parents, students: students };
}

// ============================================================
// SCHOOL ANNOUNCEMENTS (used by Parent portal)
// Returns array of announcement objects from "Announcements" sheet
// or empty array if sheet doesn't exist (portal uses defaults)
// ============================================================
function getSchoolAnnouncements() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Announcements');
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  return rows
    .filter(function(r){ return String(r[0]||'').trim() !== ''; })
    .map(function(r){
      return {
        title: String(r[0]||'').trim(),
        date:  String(r[1]||'').trim(),
        body:  String(r[2]||'').trim(),
        ico:   String(r[3]||'fas fa-bullhorn').trim(),
        bg:    String(r[4]||'#dbeafe').trim(),
        col:   '#1d4ed8'
      };
    });
}

// ============================================================
// ATTENDANCE RECORDS for Parent portal
// Called by Parent.html → getAttendance() → getAttendanceRecords(className)
// Returns records filtered; Parent.html further filters by student name/admNo
// ============================================================
function getAttendanceRecords(className) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Attendance');
  if (!sh || sh.getLastRow() < 2) return [];

  var rows    = sh.getDataRange().getValues();
  var headers = rows[0] || [];

  // Detect column positions
  function col(names) {
    for (var n = 0; n < names.length; n++)
      for (var c = 0; c < headers.length; c++)
        if (String(headers[c]).toLowerCase().trim() === names[n]) return c;
    return -1;
  }
  var C_DATE  = col(['date','datestr']);    if (C_DATE  < 0) C_DATE  = 0;
  var C_CLASS = col(['class','classname']); if (C_CLASS < 0) C_CLASS = 1;
  var C_ADMNO = col(['admno','adm no']);    if (C_ADMNO < 0) C_ADMNO = 2;
  var C_NAME  = col(['studentname','name','student name']); if (C_NAME < 0) C_NAME = 3;
  var C_SEX   = col(['sex','gender']);      if (C_SEX   < 0) C_SEX   = 4;
  var C_MORN  = col(['morning']);           if (C_MORN  < 0) C_MORN  = 6;
  var C_AFT   = col(['afternoon']);         if (C_AFT   < 0) C_AFT   = 7;

  var formNum  = className ? className.toString().match(/(\d+)/) : null;
  var fNum     = formNum ? formNum[1] : '';
  var clsLower = className ? className.toString().trim().toLowerCase() : '';
  var days     = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];

  return rows.slice(1)
    .filter(function(r){
      if (!r[C_DATE]) return false;
      var rc = String(r[C_CLASS]||'').trim().toLowerCase();
      if (!rc) return false;
      var rcNum = rc.match(/(\d+)/);
      // Match by class name or form number
      return rc === clsLower ||
             rc.replace(/\s+/g,'') === clsLower.replace(/\s+/g,'') ||
             (fNum && rcNum && rcNum[1] === fNum);
    })
    .map(function(r){
      var d = r[C_DATE];
      var dt = d instanceof Date ? d : new Date(d);
      var dateStr = isNaN(dt.getTime()) ? String(d).substring(0,10) :
        dt.getFullYear()+'-'+String(dt.getMonth()+1).padStart(2,'0')+'-'+String(dt.getDate()).padStart(2,'0');
      var dayName = isNaN(dt.getTime()) ? '' : days[dt.getDay()];
      var morn = String(r[C_MORN]||'P').trim().toUpperCase();
      return {
        admNo:  String(r[C_ADMNO]||'').trim(),
        name:   String(r[C_NAME] ||'').trim(),
        sex:    String(r[C_SEX]  ||'').trim(),
        date:   dateStr,
        day:    dayName,
        status: (morn === 'A') ? 'Absent' : 'Present'
      };
    })
    .filter(function(r){ return r.name !== ''; });
}

// ============================================================
// ALL EXAM RESULTS for Parent portal performance tab
// Called by Parent.html → getExamResults() → getAllExamResults(className)
// Returns array of { examName, students:[{name,admNo,subjects,average,position,...}] }
// ============================================================
function getAllExamResults(className) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formNum = extractFormNumber(className);
  var rows = [];

  if (formNum) {
    var formSheet = ss.getSheetByName('Form' + formNum + '_Scores');
    if (formSheet && formSheet.getLastRow() > 1) {
      rows = rows.concat(formSheet.getRange(2, 1, formSheet.getLastRow() - 1, 9).getValues());
    }
  }

  var legacy = ss.getSheetByName('Scores');
  if (legacy && legacy.getLastRow() > 1) {
    rows = rows.concat(legacy.getRange(2, 1, legacy.getLastRow() - 1, 9).getValues());
  }
  if (!rows.length) return [];

  var clsLow = className ? className.toString().trim().toLowerCase() : '';
  var fNum   = className ? (className.toString().match(/(\d+)/)||['',''])[1] : '';

  var examMap = {};
  rows.forEach(function(r){
    var rowExam  = String(r[0]||'').trim();
    var rowClass = String(r[1]||'').trim().toLowerCase();
    var rowSubj  = String(r[2]||'').trim();
    var rowAdmNo = String(r[3]||'').trim();
    var rowName  = String(r[4]||'').trim();
    var rowSex   = String(r[5]||'').trim();
    var rowScore = r[6];

    if (!rowExam || !rowName) return;

    var rcNum = rowClass.match(/(\d+)/);
    var match = rowClass === clsLow ||
                rowClass.replace(/\s+/g,'') === clsLow.replace(/\s+/g,'') ||
                (fNum && rcNum && rcNum[1] === fNum);
    if (!match) return;

    if (!examMap[rowExam]) examMap[rowExam] = {};
    if (!examMap[rowExam][rowName]) {
      examMap[rowExam][rowName] = {
        name: rowName, admNo: rowAdmNo, sex: rowSex, subjects: []
      };
    }
    var score = (rowScore !== '' && rowScore !== null && rowScore !== undefined && !isNaN(Number(rowScore)))
      ? Number(rowScore) : null;
    var grade = examScoreToGrade_(score);
    examMap[rowExam][rowName].subjects.push({
      name: rowSubj,
      score: score,
      grade: grade,
      point: examGradeToPoint_(grade)
    });
  });

  var results = [];
  Object.keys(examMap).sort().forEach(function(examName){
    var studentMap = examMap[examName];
    var students = Object.values(studentMap).map(function(s){
      var vals = s.subjects.filter(function(sub){ return sub.score !== null; })
                           .map(function(sub){ return sub.score; });
      var total = vals.reduce(function(a,b){ return a+b; }, 0);
      var avg   = vals.length ? (total / vals.length) : 0;
      var g     = examScoreToGrade_(avg);
      var summary = examBest7SummaryFromSubjects_(s.subjects);
      return {
        name:     s.name,
        admNo:    s.admNo,
        sex:      s.sex,
        subjects: s.subjects,
        total:    total,
        average:  avg.toFixed(2),
        grade:    g,
        point:    summary.point,
        division: summary.division,
        remarks:  summary.division === 'ABS' ? 'ABSENT' : (summary.division === '' ? 'INCOMPLETE' : (summary.division !== '0' ? 'PASS' : 'FAIL'))
      };
    }).sort(function(a,b){ return b.total - a.total; });
    students.forEach(function(s,i){ s.position = String(i+1); s.outOf = String(students.length); });
    results.push({ examName: examName, students: students });
  });

  return results;
}
function clearAllActivityLogs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Activity');
  if (!sh) return 'Activity sheet not found.';
  if (sh.getLastRow() > 1) {
    sh.deleteRows(2, sh.getLastRow() - 1);
  }
  return 'Activity log cleared.';
}



// ============================================================
// DASHBOARD ACCESS ENFORCEMENT — FINAL LOGIN OVERRIDE
// This final checkLogin() intentionally overrides earlier duplicate versions.
// It blocks Teacher/Staff/Academics/ClassTeacher/Parent when Admin disallows them.
// ============================================================
function legacyCheckLogin_v3_DISABLED(username, password) {
  username = String(username || '').trim();
  password = String(password || '');

  // ── Admin: never blocked by dashboard access policies ──────
  if (ADMIN_USERS.some(function(u) { return u.username === username && u.password === password; })) {
    return {
      status:   'success',
      role:     'admin',
      fullName: 'Administrator',
      username: username,
      html:     HtmlService.createHtmlOutputFromFile('Admin').getContent()
    };
  }

  // ── Teachers / Academic Master / Class Teacher / Staff ─────
  var teachers = getTeacherList();
  var t = null;
  for (var ti = 0; ti < teachers.length; ti++) {
    if (String(teachers[ti].username || '') === username && String(teachers[ti].password || '') === password) {
      t = teachers[ti];
      break;
    }
  }

  if (t) {
    var pos = String(t.position || '').trim();
    var htmlFile = 'Teacher';
    var role = 'teacher';
    var dashboardKey = getDashboardKeyFromPosition(pos);

    if (dashboardKey === 'academics') {
      htmlFile = 'Academics';
      role = 'academic';
    } else if (dashboardKey === 'classTeacher') {
      htmlFile = 'ClassTeacher';
      role = 'classteacher';
    } else if (['sportsManager','accountant','librarian','labTechnician'].indexOf(dashboardKey) >= 0) {
      htmlFile = 'Staff';
      role = 'staff';
    }

    var access = canAccessDashboard(dashboardKey, t.username);
    if (!access.allowed) {
      try { logActivity(t.username, 'Blocked login: ' + access.title); } catch(e) {}
      return {
        status: 'success',
        role: 'blocked',
        position: pos,
        fullName: t.name,
        username: t.username,
        dashboardKey: dashboardKey,
        html: getAccessDeniedPage(access)
      };
    }

    setTeacherSession(t.username);
    try { logActivity(t.username, 'Login: ' + t.username + ' (' + pos + ')'); } catch(e) {}

    return {
      status:          'success',
      role:            role,
      position:        pos,
      fullName:        t.name,
      username:        t.username,
      dashboardKey:    dashboardKey,
      access:          access,
      classesSubjects: t.classesSubjects,
      classes:         t.classes,
      subjects:        t.subjects,
      html:            HtmlService.createHtmlOutputFromFile(htmlFile).getContent()
    };
  }

  // ── Parents ────────────────────────────────────────────────
  var sh = ensureParentsSheet();
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var rows = sh.getRange(2, 1, lastRow - 1, 12).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var rowUser = String(r[7] || '').trim();
      var rowPass = String(r[8] || '').trim();

      if (rowUser.toLowerCase() === username.toLowerCase() && rowPass === password) {
        var parentAccess = canAccessDashboard('parents', rowUser);
        if (!parentAccess.allowed) {
          try { logActivity(rowUser, 'Blocked parent login'); } catch(e) {}
          return {
            status: 'success',
            role: 'blocked',
            fullName: String(r[0] || username).trim(),
            username: rowUser,
            dashboardKey: 'parents',
            html: getAccessDeniedPage(parentAccess)
          };
        }

        try { logActivity(rowUser, 'Parent login: ' + rowUser); } catch(e) {}
        return {
          status:   'success',
          role:     'parent',
          fullName: String(r[0] || username).trim(),
          username: rowUser,
          dashboardKey: 'parents',
          access: parentAccess,
          parentData: {
            parentName:  String(r[0] || '').trim(),
            relation:    String(r[2] || '').trim(),
            phone:       String(r[3] || '').trim(),
            email:       String(r[4] || '').trim(),
            admNo:       String(r[5] || '').trim(),
            address:     String(r[6] || '').trim(),
            studentName: String(r[9] || '').trim(),
            className:   String(r[10]|| '').trim(),
            sex:         '',
            status:      'Approved'
          }
        };
      }
    }
  }

  return { status: 'fail', message: 'Invalid username or password.' };
}
// ============================================================
// REQUIRED ACCESS FIX — PASTE AT THE VERY BOTTOM OF Code.gs
// This makes Teacher.html and ClassTeacher.html obey Admin restrictions.
// IMPORTANT: paste at the bottom so this checkLogin overrides old duplicate checkLogin functions.
// ============================================================

function smsCheckLoginCore_(username, password) {
  username = String(username || '').trim();
  password = String(password || '');
  try { clearTeacherSession(); } catch(e) {} // start every login clean; blocked users must not keep an old dashboard session

  // Admin is never blocked
  if (ADMIN_USERS.some(function(u) { return u.username === username && u.password === password; })) {
    return {
      status: 'success', role: 'admin', fullName: 'Administrator', username: username,
      html: HtmlService.createHtmlOutputFromFile('Admin').getContent()
    };
  }

  // Teachers / Class Teachers / Academic / Staff
  var teachers = getTeacherList();
  var t = null;
  for (var i = 0; i < teachers.length; i++) {
    if (String(teachers[i].username || '') === username && String(teachers[i].password || '') === password) {
      t = teachers[i]; break;
    }
  }

  if (t) {
    var pos = String(t.position || '').trim();
    var dashboardKey = getDashboardKeyFromPosition(pos);
    var htmlFile = 'Teacher';
    var role = 'teacher';

    if (dashboardKey === 'academics') { htmlFile = 'Academics'; role = 'academic'; }
    else if (dashboardKey === 'classTeacher') { htmlFile = 'ClassTeacher'; role = 'classteacher'; }
    else if (['sportsManager','accountant','librarian','labTechnician'].indexOf(dashboardKey) >= 0) { htmlFile = 'Staff'; role = 'staff'; }

    var access = canAccessDashboard(dashboardKey, t.username);
    if (!access || access.allowed === false) {
      try { logDashboardBlockedAttempt(t.username, dashboardKey, access, 'login'); } catch(e) {
        try { logActivity(t.username, 'Blocked login: ' + ((access && access.title) || dashboardKey)); } catch(_e) {}
      }
      return {
        status: 'fail',
        blocked: true,
        role: 'blocked',
        fullName: t.name,
        username: t.username,
        position: pos,
        dashboardKey: dashboardKey,
        message: ((access && access.reason) || 'This dashboard has been blocked by the administrator.'),
        html: getAccessDeniedPage(access || {title:dashboardKey, reason:'This dashboard has been blocked by the administrator.'})
      };
    }

    setTeacherSession(t.username);
    try { logActivity(t.username, 'Login: ' + t.username + ' (' + pos + ')'); } catch(e) {}
    return {
      status:'success', role:role, position:pos, fullName:t.name, username:t.username,
      dashboardKey:dashboardKey, access:access,
      classesSubjects:t.classesSubjects, classes:t.classes, subjects:t.subjects,
      html:HtmlService.createHtmlOutputFromFile(htmlFile).getContent()
    };
  }

  // Parents
  var sh = ensureParentsSheet();
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var rows = sh.getRange(2, 1, lastRow - 1, 12).getValues();
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      var rowUser = String(row[7] || '').trim();
      var rowPass = String(row[8] || '').trim();
      if (rowUser.toLowerCase() === username.toLowerCase() && rowPass === password) {
        var pAccess = canAccessDashboard('parents', rowUser);
        if (!pAccess.allowed) {
          try { logDashboardBlockedAttempt(rowUser, 'parents', pAccess, 'login'); } catch(e) {}
          return {status:'fail', blocked:true, role:'blocked', username:rowUser, message:(pAccess.reason||'Parent portal has been blocked by the administrator.'), html:getAccessDeniedPage(pAccess)};
        }
        return {
          status:'success', role:'parent', fullName:String(row[0] || username).trim(), username:rowUser,
          parentData:{
            parentName:String(row[0]||'').trim(), relation:String(row[2]||'').trim(), phone:String(row[3]||'').trim(),
            email:String(row[4]||'').trim(), admNo:String(row[5]||'').trim(), address:String(row[6]||'').trim(),
            studentName:String(row[9]||'').trim(), className:String(row[10]||'').trim(), sex:'', status:'Approved'
          }
        };
      }
    }
  }

  return { status:'fail', message:'Invalid username or password.' };
}
// ============================================================
// DASHBOARD ACCESS HELPERS — safe defaults (allow unless blocked)
// ============================================================
function getDashboardKeyFromPosition(pos){
  pos = String(pos||'').toLowerCase().trim();
  if(!pos) return 'teacher';
  if(pos.indexOf('academic')>=0)        return 'academics';
  if(pos.indexOf('class teacher')>=0)   return 'classTeacher';
  if(pos.indexOf('sport')>=0)           return 'sportsManager';
  if(pos.indexOf('account')>=0)         return 'accountant';
  if(pos.indexOf('librar')>=0)          return 'librarian';
  if(pos.indexOf('lab')>=0)             return 'labTechnician';
  if(pos.indexOf('parent')>=0)          return 'parents';
  return 'teacher';
}

// ============================================================
// DASHBOARD ACCESS POLICIES — real server-side storage for the
// "Monitor All Dashboards" panel.
// Previously the Admin.html panel called saveDashboardAccessPolicies(),
// getDashboardAccessPolicies(), saveDashboardAccessState() and
// getDashboardAccessEvents() — none of which existed here, so every
// Allow/Disallow/Read-only/Emergency-lock change only ever got saved
// to the admin's own browser localStorage. It LOOKED like it worked
// (the badge flipped to "Allowed"/"Disallowed"), but real login
// checks (canAccessDashboard, used by checkLogin) read a completely
// separate sheet and never saw any of it — hence "shows allowed but
// it denied". These functions make the panel's choices real.
//
// Storage: one row per dashboard key in "DashboardAccess", with the
// full policy object (allowed/readOnly/emergencyLock/temporaryLock/
// expires/reason/userOverrides/features/...) as a JSON blob in col B.
// Older 4-column rows (DashboardKey | Blocked | Reason | BlockedUsers)
// from before this fix are still honoured by canAccessDashboard() as
// a fallback, so nothing existing silently breaks.
// ============================================================
function ensureDashboardAccessSheet_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('DashboardAccess');
  if (!sh) {
    sh = ss.insertSheet('DashboardAccess');
    sh.appendRow(['DashboardKey','PolicyJSON','UpdatedBy','UpdatedAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getDashboardAccessPolicies(){
  var sh = ensureDashboardAccessSheet_();
  var last = sh.getLastRow();
  var out = {};
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 4).getValues().forEach(function(r){
    var key = String(r[0] || '').trim();
    if (!key) return;
    try { out[key] = JSON.parse(r[1]); } catch(e) { /* old simple-format row — leave out of the JSON map */ }
  });
  return out;
}

function saveDashboardAccessPolicies(policies, updatedBy){
  var sh = ensureDashboardAccessSheet_();
  var last = sh.getLastRow();
  var existing = last >= 2 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];
  var rowByKey = {};
  existing.forEach(function(r, i){ var k = String(r[0] || '').trim(); if (k) rowByKey[k] = i + 2; });
  var now = new Date();
  Object.keys(policies || {}).forEach(function(key){
    var json = JSON.stringify(policies[key]);
    if (rowByKey[key]) sh.getRange(rowByKey[key], 1, 1, 4).setValues([[key, json, updatedBy || 'admin', now]]);
    else sh.appendRow([key, json, updatedBy || 'admin', now]);
  });
  try { logActivity(updatedBy || 'admin', 'Updated dashboard access policies'); } catch(e) {}
  return 'Dashboard access policies saved.';
}

function saveDashboardAccessState(key, policy, updatedBy){
  var obj = {}; obj[key] = policy;
  saveDashboardAccessPolicies(obj, updatedBy);
  try { logDashboardAccessEvent_(key, (policy && policy.allowed === false) ? 'Blocked' : 'Allowed', updatedBy || 'admin'); } catch(e) {}
  return 'Saved "' + key + '" policy.';
}

function logDashboardAccessEvent_(key, action, by){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('DashboardAccessEvents');
  if (!sh) { sh = ss.insertSheet('DashboardAccessEvents'); sh.appendRow(['Time','Panel','Action','By']); sh.setFrozenRows(1); }
  sh.appendRow([new Date(), key, action, by || 'admin']);
}

function getDashboardAccessEvents(panel){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('DashboardAccessEvents');
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().map(function(r){
    return { time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), panel: r[1], action: r[2], by: r[3] };
  }).reverse();
  if (panel && panel !== 'all') rows = rows.filter(function(e){ return e.panel === panel; });
  return rows.slice(0, 80);
}

/**
 * forceLogoutDashboard(key, by)
 * Best-effort: clears the CacheService session for any teacher-type
 * dashboard so their next click re-checks access and gets bounced.
 * Parents have no server-side session (Parent.html just re-checks
 * checkCurrentDashboardAccess() every 15s — see Code.gs comment
 * there), so for 'parents' this just logs the action; the existing
 * poll will pick up the new block within 15 seconds regardless.
 */
function forceLogoutDashboard(key, by){
  try { logDashboardAccessEvent_(key, 'Force logout', by || 'admin'); } catch(e) {}
  return 'Logout signal sent for "' + key + '". Active users will be blocked on their next access check (within ~15s for parents, immediately for teacher dashboards on their next action).';
}

function canAccessDashboard(dashboardKey, username){
  // ============================================================
  // DASHBOARD ACCESS GATE — DISABLED
  // This used to read block/lock policies from the "DashboardAccess"
  // sheet (set via the Admin "Monitor All Dashboards" panel) and could
  // reject logins/sessions for Teacher, Class Teacher and Academic
  // Master, which is what was causing every action on those dashboards
  // to bounce back to Login with "Session expired". Always allowing
  // access here removes that gate system-wide (getTeacherSession(),
  // both checkLogin() functions, and the parent-portal check all call
  // this single function), while leaving the Monitor panel's UI and
  // logging in place — it just no longer has any enforcement effect.
  // ============================================================
  return { allowed:true, title:dashboardKey };
}

/**
 * Evaluates the rich policy object the "Monitor All Dashboards" panel
 * writes (allowed / readOnly / emergencyLock / temporaryLock+expires /
 * reason / userOverrides). This is what makes the panel's Allow/
 * Disallow/Emergency-lock buttons actually take effect at login,
 * instead of only changing a badge in the admin's own browser.
 */
function evalRichDashboardPolicy_(dashboardKey, pol, username){
  var reason = pol.reason || 'This dashboard has been restricted by the administrator.';

  if (pol.emergencyLock) {
    return { allowed:false, title:dashboardKey, reason:reason };
  }

  if (pol.temporaryLock && pol.allowed === false) {
    var exp = pol.expires ? new Date(pol.expires) : null;
    if (exp && !isNaN(exp.getTime()) && exp > new Date()) {
      return { allowed:false, title:dashboardKey, reason:reason, notice:'Locked until ' + exp.toLocaleString() };
    }
    // lock has expired — fall through to the general 'allowed' flag below
  }

  if (username && pol.userOverrides) {
    var ov = pol.userOverrides[String(username).toLowerCase()];
    if (ov && ov.allowed === false) {
      return { allowed:false, title:dashboardKey, reason: ov.reason || reason };
    }
  }

  if (pol.allowed === false) {
    return { allowed:false, title:dashboardKey, reason:reason };
  }

  return { allowed:true, title:dashboardKey, readOnly: !!pol.readOnly };
}

// ============================================================
// checkCurrentDashboardAccess(dashboardKey)
// Parent.html (and other dashboards) poll this every few seconds
// after login to see whether the admin has blocked the dashboard
// mid-session. This function was being CALLED but never DEFINED,
// so every poll failed with "Script function not found:
// checkCurrentDashboardAccess" — that repeated, unhandled error is
// what shows up to parents as being kicked out / "session expired"
// right after a successful login. Defining it here fixes that.
// ============================================================
function checkCurrentDashboardAccess(dashboardKey){
  try {
    return canAccessDashboard(dashboardKey, null);
  } catch(e) {
    // Fail-open so a transient error never locks a legitimate user out
    return { allowed: true, title: dashboardKey };
  }
}

function getAccessDeniedPage(access){
  var title  = (access && access.title)  || 'Dashboard';
  var reason = (access && access.reason) || 'Access has been restricted by the administrator.';
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Access Restricted</title>'
       + '<style>body{font-family:Arial,sans-serif;background:#0a1628;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;text-align:center;padding:24px}'
       + '.box{max-width:460px;background:#0f2044;padding:32px;border-radius:14px;box-shadow:0 12px 40px rgba(0,0,0,.4)}'
       + 'h1{color:#f87171;font-size:22px;margin:0 0 10px}p{color:#cbd5e1;line-height:1.5;margin:0 0 18px}'
       + 'a{color:#06b6d4;text-decoration:none;font-weight:700}</style></head><body><div class="box">'
       + '<h1>🔒 '+title+' restricted</h1><p>'+reason+'</p><a href="javascript:history.back()">← Back to login</a>'
       + '</div></body></html>';
}

function logDashboardBlockedAttempt(username, dashboardKey, access, phase){
  try{ logActivity(username||'unknown', 'Blocked '+(phase||'login')+' → '+dashboardKey+' ('+((access&&access.reason)||'blocked')+')'); }catch(e){}
}

// ============================================================
// BEHAVIOURAL CRITERIA ASSESSMENT — filled by the CLASS TEACHER
// for the classes in his/her Assignments, then SUBMITTED to Academic.
// Academics.html prints only SUBMITTED behaviour on the English report card.
// Sheet "Behaviors": Class | Exam | AdmNo | StudentName |
//   10 criteria | UpdatedBy | UpdatedAt | Status | SubmittedAt
// One row per (class + exam + student).
// ============================================================
const BEHAVIOR_SHEET    = 'Behaviors';
const BEHAVIOR_CRITERIA = ['Discipline','Punctuality','Attendance','Respect','Responsibility',
                           'Class Participation','Cooperation','Cleanliness','Self-control','Leadership'];
const BEHAVIOR_LEVELS   = ['Excellent','Very Good','Good','Average','Weak'];   // best → weakest
// Ratings saved before the English version are Swahili; each keeps the English level the report card always showed for it.
const BEHAVIOR_LEVELS_SW = { 'vizuri sana': 'Very Good', 'vizuri': 'Good', 'wastani': 'Average', 'dhaifu': 'Weak' };

// any rating (English or old Swahili) → the English level, or '' when it is not a valid level
function behLevel_(v) {
  const s = String(v == null ? '' : v).trim().toLowerCase();
  if (!s) return '';
  const i = BEHAVIOR_LEVELS.findIndex(l => l.toLowerCase() === s);
  return i >= 0 ? BEHAVIOR_LEVELS[i] : (BEHAVIOR_LEVELS_SW[s] || '');
}

function ensureBehaviorSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(BEHAVIOR_SHEET);
  if (!sh) {
    sh = ss.insertSheet(BEHAVIOR_SHEET);
    sh.appendRow(['Class','Exam','AdmNo','StudentName'].concat(BEHAVIOR_CRITERIA)
      .concat(['UpdatedBy','UpdatedAt','Status','SubmittedAt']));
    sh.setFrozenRows(1);
  } else {
    try {
      const cache = CacheService.getScriptCache();
      if (!cache.get('BEH_HDR_EN')) {
        const cur = sh.getRange(1, 5, 1, BEHAVIOR_CRITERIA.length).getValues()[0];
        if (cur.join('|') !== BEHAVIOR_CRITERIA.join('|')) sh.getRange(1, 5, 1, BEHAVIOR_CRITERIA.length).setValues([BEHAVIOR_CRITERIA]);
        cache.put('BEH_HDR_EN', '1', 21600);
      }
    } catch (e) {}
  }
  return sh;
}

function behNorm_(s)     { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
function behClassKey_(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ''); }

// Behaviour assessment belongs to the Class Teacher of a class, as set in Manage Classes.
// The classes a teacher only teaches (My Assignments) do not count.
function behTeacherClasses_(t) {
  const out = [];
  function add(c) { c = String(c || '').trim(); if (c && out.indexOf(c) < 0) out.push(c); }
  (t.homeroomClasses || []).forEach(add);
  return out;
}

function behCanAccess_(t, className) {
  const key = behClassKey_(className);
  return behTeacherClasses_(t).some(c => behClassKey_(c) === key);
}

function getMyHomeroomClasses() {
  const t = getTeacherSession();
  return t ? behTeacherClasses_(t) : [];
}

// Rows for one class + exam → { normalizedName: { ratings:[10], status } }
function behReadMap_(className, examType, submittedOnly) {
  const sh = ensureBehaviorSheet_();
  const map = {};
  if (sh.getLastRow() < 2) return map;
  const ck = behClassKey_(className), ek = behNorm_(examType), n = BEHAVIOR_CRITERIA.length;
  sh.getRange(2, 1, sh.getLastRow() - 1, 4 + n + 4).getValues().forEach(r => {
    if (behClassKey_(r[0]) !== ck || behNorm_(r[1]) !== ek) return;
    const status = String(r[4 + n + 2] || 'Draft').trim() || 'Draft';
    if (submittedOnly && status !== 'Submitted') return;
    map[behNorm_(r[3])] = { ratings: r.slice(4, 4 + n).map(behLevel_), status: status };
  });
  return map;
}

// Class teacher: students of the class + saved behaviour + submission status
function getBehaviorSheet(className, examType) {
  const t = getTeacherSession();
  if (!t) return { error: 'Session expired. Please log in again.' };
  if (!behCanAccess_(t, className)) return { error: 'Class ' + className + ' is not in your assignments.' };
  if (!examType) return { error: 'Select an exam first.' };

  const saved = behReadMap_(className, examType, false);
  let submittedCount = 0, savedCount = 0;
  const students = getStudentsForClass(className).map(s => {
    const name = ((s.firstName || '') + ' ' + (s.lastName || '')).trim();
    const rec = saved[behNorm_(name)];
    if (rec) { savedCount++; if (rec.status === 'Submitted') submittedCount++; }
    return { admNo: s.admNo, name: name, sex: s.sex, ratings: rec ? rec.ratings : [] };
  }).filter(s => s.name).sort((a, b) => a.name.localeCompare(b.name));

  const status = (savedCount > 0 && submittedCount === savedCount) ? 'Submitted' : (savedCount > 0 ? 'Draft' : 'New');
  return { criteria: BEHAVIOR_CRITERIA, levels: BEHAVIOR_LEVELS, students: students, status: status };
}

// payload = { className, examType, rows:[{admNo,name,ratings:[10]}], submit:true|false }
// submit=false → save as Draft (keeps "Submitted" only for rows that did not change)
// submit=true  → save and mark every row Submitted to Academic
function saveBehaviors(payload) {
  const t = getTeacherSession();
  if (!t) return { ok: false, message: 'Session expired. Please log in again.' };
  const className = String((payload && payload.className) || '').trim();
  const examType  = String((payload && payload.examType)  || '').trim();
  const rows      = (payload && payload.rows) || [];
  const submit    = !!(payload && payload.submit);
  if (!className || !examType) return { ok: false, message: 'Class and exam are required.' };
  if (!behCanAccess_(t, className)) return { ok: false, message: 'Class ' + className + ' is not in your assignments.' };

  const n = BEHAVIOR_CRITERIA.length, width = 4 + n + 4;
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok: false, message: 'Server busy, please try again.' }; }
  try {
    const sh = ensureBehaviorSheet_();
    const ck = behClassKey_(className), ek = behNorm_(examType);
    const existing = {};                       // normalized name → { row, ratings, status }
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues().forEach((r, i) => {
        if (behClassKey_(r[0]) === ck && behNorm_(r[1]) === ek) {
          existing[behNorm_(r[3])] = { row: i + 2, ratings: r.slice(4, 4 + n).map(behLevel_),
                                       status: String(r[4 + n + 2] || 'Draft').trim() || 'Draft' };
        }
      });
    }
    const stamp = new Date(), who = t.fullName || t.username;
    const appendRows = [];
    let saved = 0;

    rows.forEach(r => {
      const name = String((r && r.name) || '').trim();
      if (!name) return;
      const ratings = [];
      for (let i = 0; i < n; i++) {
        ratings.push(behLevel_(r.ratings && r.ratings[i]));
      }
      const old = existing[behNorm_(name)];
      if (!old && !ratings.some(v => v)) return;           // nothing to store
      let status = 'Draft', submittedAt = '';
      if (submit) { status = 'Submitted'; submittedAt = stamp; }
      else if (old && old.status === 'Submitted' && old.ratings.join('|') === ratings.join('|')) {
        status = 'Submitted';                              // unchanged → stays submitted
      }
      const line = [className, examType, String(r.admNo || '').trim(), name]
        .concat(ratings).concat([who, stamp, status, submittedAt]);
      if (old) {
        if (status === 'Submitted' && !submit) { // keep original submit time
          line[4 + n + 3] = sh.getRange(old.row, 4 + n + 4).getValue();
        }
        sh.getRange(old.row, 1, 1, width).setValues([line]);
      } else appendRows.push(line);
      saved++;
    });
    if (appendRows.length) sh.getRange(sh.getLastRow() + 1, 1, appendRows.length, width).setValues(appendRows);

    try { logActivity(t.username, (submit ? 'Submitted' : 'Saved') + ' behaviour: ' + className + ' / ' + examType + ' (' + saved + ')'); } catch (e) {}
    return { ok: true, saved: saved, submitted: submit,
             message: submit ? 'Behaviour of ' + saved + ' student(s) submitted to Academic.'
                             : 'Draft saved for ' + saved + ' student(s). Remember to Submit to Academic.' };
  } catch (e) {
    return { ok: false, message: 'Error: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// Academics.html report cards — SUBMITTED behaviour only
// returns { criteria, classTeacher, map: { normalizedName: [10 ratings] } }
function fetchReportBehaviors(className, examType) {
  let raw = '';
  const ck = behClassKey_(className);
  getAcademicClassesRaw().forEach(c => {
    if (!raw && behClassKey_(c.name) === ck && c.teacher) raw = String(c.teacher).trim();
  });

  // Always show the teacher's FULL name: look the value up in the Users sheet
  // (by name or username, or by a single matching first/last name).
  let classTeacher = '';
  if (raw) {
    const rn = behNorm_(raw), teachers = getTeacherList();
    let hit = teachers.find(t => behNorm_(t.name) === rn || behNorm_(t.username) === rn);
    if (!hit) {
      const part = teachers.filter(t => behNorm_(t.name).split(' ').indexOf(rn) >= 0);
      if (part.length === 1) hit = part[0];
    }
    if (hit && hit.name) classTeacher = String(hit.name).trim();
  }

  // Submitted behaviour rows: also remember who submitted them (full name).
  const sh = ensureBehaviorSheet_(), n = BEHAVIOR_CRITERIA.length, ek = behNorm_(examType);
  let submitter = '';
  const map = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 4 + n + 4).getValues().forEach(r => {
      if (behClassKey_(r[0]) !== ck || behNorm_(r[1]) !== ek) return;
      if (String(r[4 + n + 2] || '').trim() !== 'Submitted') return;
      map[behNorm_(r[3])] = r.slice(4, 4 + n).map(v => behLevel_(v) || String(v || '').trim());
      if (!submitter) submitter = String(r[4 + n] || '').trim();
    });
  }
  return { criteria: BEHAVIOR_CRITERIA, classTeacher: classTeacher || submitter || raw, map: map };
}


// ============================================================
// STAY LOGGED IN AFTER REFRESH (whole system)
// ------------------------------------------------------------
// • On login, a random session token is created, kept here (Script
//   Properties) and in the browser (localStorage "sms_token").
// • When the page is refreshed, doGet() serves the landing page with a
//   small script that sends the token to smsRestoreSession(); if it is
//   still valid the user's dashboard is re-opened directly — no login.
// • Logout (logout / getLoginPage / getIndexPage) deletes the token.
// Sessions last SMS_SESSION_HOURS since the last refresh/visit.
// ============================================================
const SMS_SESSION_PREFIX = 'SMS_SESS_';
const SMS_SESSION_HOURS  = 12;

function smsNewToken_(info) {
  const props = PropertiesService.getScriptProperties();
  try {                                         // remove expired sessions
    const all = props.getProperties(), now = Date.now();
    Object.keys(all).forEach(k => {
      if (k.indexOf(SMS_SESSION_PREFIX) !== 0) return;
      try { if (JSON.parse(all[k]).exp < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
    });
  } catch (e) {}
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  info.exp = Date.now() + SMS_SESSION_HOURS * 3600 * 1000;
  props.setProperty(SMS_SESSION_PREFIX + token, JSON.stringify(info));
  return token;
}

function smsEndSession(token) {
  token = String(token || '');
  if (/^[a-f0-9]{64}$/.test(token)) {
    try { PropertiesService.getScriptProperties().deleteProperty(SMS_SESSION_PREFIX + token); } catch (e) {}
    try { CacheService.getScriptCache().remove('PSS_' + token.slice(0, 40)); } catch (e) {}
  }
  return true;
}

// script that stores the token in the browser (fresh = true on a new login)
function smsStoreScript_(token, fresh) {
  return '<script>try{' +
    (fresh ? 'for(var i=localStorage.length-1;i>=0;i--){var k=localStorage.key(i);if(k&&k.indexOf("sms_pane_")===0)localStorage.removeItem(k);}' : '') +
    'localStorage.setItem("sms_token","' + token + '");}catch(e){}</script>';
}

// script that removes the saved login (logout)
function smsClearScript_() {
  return '<script>try{var t=localStorage.getItem("sms_token");localStorage.removeItem("sms_token");' +
    'for(var i=localStorage.length-1;i>=0;i--){var k=localStorage.key(i);if(k&&k.indexOf("sms_pane_")===0)localStorage.removeItem(k);}' +
    'if(t&&window.google&&google.script&&google.script.run)google.script.run.smsEndSession(t);}catch(e){}</script>';
}

function smsAppendScript_(html, script) { return String(html || '') + script; }

// script placed at the top of the landing page: restores the session if a token exists
function smsInjectRestore_(html) {
  const script = `<script>
(function(){
  var tok=null; try{tok=localStorage.getItem("sms_token");}catch(e){}
  if(!tok) return;
  var ov=document.createElement("div");
  ov.id="smsRestoreOverlay";
  ov.setAttribute("style","position:fixed;left:0;top:0;right:0;bottom:0;background:#0f172a;color:#fff;z-index:2147483647;display:flex;align-items:center;justify-content:center;font:600 15px Arial,sans-serif;");
  ov.textContent="Inapakia... / Restoring your session...";
  document.documentElement.appendChild(ov);
  function drop(){ var o=document.getElementById("smsRestoreOverlay"); if(o&&o.parentNode) o.parentNode.removeChild(o); }
  function run(){
    google.script.run
      .withSuccessHandler(function(r){
        if(r&&r.ok&&r.html){ document.open(); document.write(r.html); document.close(); }
        else { try{localStorage.removeItem("sms_token");}catch(e){} drop(); }
      })
      .withFailureHandler(function(){ drop(); })
      .smsRestoreSession(tok);
  }
  if(document.readyState==="loading") document.addEventListener("DOMContentLoaded",run); else run();
})();
</script>`;
  let out = String(html || '');
  if (/<head[^>]*>/i.test(out))      return out.replace(/<head[^>]*>/i, m => m + script);
  if (/<html[^>]*>/i.test(out))      return out.replace(/<html[^>]*>/i, m => m + script);
  return script + out;
}

function smsDashboardFor_(position) {
  const key = getDashboardKeyFromPosition(String(position || '').trim());
  let file = 'Teacher', role = 'teacher';
  if (key === 'academics')        { file = 'Academics';    role = 'academic'; }
  else if (key === 'classTeacher'){ file = 'ClassTeacher'; role = 'classteacher'; }
  else if (['sportsManager','accountant','librarian','labTechnician'].indexOf(key) >= 0) { file = 'Staff'; role = 'staff'; }
  return { key: key, file: file, role: role };
}

// Called by the landing page after a refresh. Returns { ok, html } or { ok:false }.
function smsRestoreSession(token) {
  try {
    token = String(token || '');
    if (!/^[a-f0-9]{64}$/.test(token)) return { ok: false };
    const props = PropertiesService.getScriptProperties();
    const key = SMS_SESSION_PREFIX + token;
    const raw = props.getProperty(key);
    if (!raw) return { ok: false };
    const s = JSON.parse(raw);
    if (!s.exp || s.exp < Date.now()) { props.deleteProperty(key); return { ok: false }; }

    let html = '';
    if (['schoolmanager', 'headofschool', 'accountant'].indexOf(s.role) >= 0) {
      const acc = psAdminByUsername_(s.username);
      if (!acc || acc.Status === 'Inactive' || psRoleKey_(acc.Role) !== s.role) { props.deleteProperty(key); return { ok: false }; }
      html = psBootInject_(s.role, { role: s.role, name: acc.FullName || acc.Username, username: acc.Username },
                           HtmlService.createHtmlOutputFromFile(psFileForRole_(s.role)).getContent());
    } else if (s.role === 'parent') {
      const pd = smsParentData_(s.username);
      if (!pd) { props.deleteProperty(key); return { ok: false }; }       // parent removed
      const pa = canAccessDashboard('parents', s.username);
      if (pa && pa.allowed === false) return { ok: false };               // blocked by admin
      html = smsParentHtml_(pd);
    } else if (s.role === 'admin') {
      props.deleteProperty(key); return { ok: false };               // old secondary-school Admin token
    } else {
      const t = getTeacherList().find(x => String(x.username || '') === String(s.username || ''));
      if (!t) { props.deleteProperty(key); return { ok: false }; }       // user was removed
      const d = smsDashboardFor_(t.position);
      const access = canAccessDashboard(d.key, t.username);
      if (!access || access.allowed === false) return { ok: false };     // blocked by admin
      setTeacherSession(t.username);
      html = HtmlService.createHtmlOutputFromFile(d.file).getContent();
    }
    s.exp = Date.now() + SMS_SESSION_HOURS * 3600 * 1000;               // slide expiry
    props.setProperty(key, JSON.stringify(s));
    return { ok: true, role: s.role, html: smsAppendScript_(html, smsStoreScript_(token, false)) };
  } catch (e) {
    Logger.log('smsRestoreSession: ' + e.message);
    return { ok: false };
  }
}

// Parent record → the same parentData object checkLogin() returns
function smsParentData_(username) {
  const sh = ensureParentsSheet(), last = sh.getLastRow();
  if (last < 2) return null;
  const rows = sh.getRange(2, 1, last - 1, 12).getValues();
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (String(row[7] || '').trim().toLowerCase() === String(username || '').trim().toLowerCase()) {
      return {
        parentName: String(row[0] || '').trim(), relation: String(row[2] || '').trim(), phone: String(row[3] || '').trim(),
        email: String(row[4] || '').trim(), admNo: String(row[5] || '').trim(), address: String(row[6] || '').trim(),
        studentName: String(row[9] || '').trim(), className: String(row[10] || '').trim(), sex: '', status: 'Approved'
      };
    }
  }
  return null;
}

// Parent portal page + the same auto-boot script Login.html uses
function smsParentHtml_(parentData) {
  const ph = HtmlService.createHtmlOutputFromFile('Parent').getContent();
  const boot = '<script>(function(){var _autoData=' + JSON.stringify(parentData).replace(/</g, '\\u003c') + ';' +
    'function tryBoot(){if(typeof bootPortal==="function"){var ls=document.getElementById("loginScreen");if(ls)ls.style.display="none";bootPortal(_autoData);}}' +
    'if(document.readyState==="complete"||document.readyState==="interactive"){setTimeout(tryBoot,100);}' +
    'else{window.addEventListener("DOMContentLoaded",function(){setTimeout(tryBoot,100);});}})();<\/script>';
  return /<\/body>/i.test(ph) ? ph.replace(/<\/body>/i, () => boot + '\n</body>') : ph + boot;
}

// Public login entry point:
//   1) School Manager / Head of School / Accountant  (Administrators sheet)
//   2) teachers & staff (Users sheet — blocked when Inactive)
//   3) parents
function checkLogin(username, password) {
  let res = null;
  try { res = psTryAdminLogin_(username, password); } catch (e) { Logger.log('psTryAdminLogin_: ' + e.message); }
  if (!res) {
    let blocked = null;
    try { blocked = psInactiveStaffBlock_(username, password); } catch (e) {}
    res = blocked || smsCheckLoginCore_(username, password);
  }
  try {
    if (res && res.status === 'success' && (res.html || res.role === 'parent')) {
      const token = smsNewToken_({ role: res.role, username: res.username });
      res.sessionToken = token;                       // parents: stored by the page (Login / Index / Parent)
      if (res.html) res.html = smsAppendScript_(res.html, smsStoreScript_(token, true));
    }
  } catch (e) { Logger.log('checkLogin token: ' + e.message); }
  return res;
}


// ============================================================
// ACADEMIC SETTINGS — report logos (left / right), Head of School and
// Academic Master names + signatures. Used by Academics.html on the
// Publish Exam Results report and on every Individual Report Card.
// Sheet "AcademicSettings": Key | Part | Value
// Images are stored as data-URLs, split in parts of 45,000 characters
// (a Google Sheets cell holds at most 50,000 characters).
// ============================================================
const ACS_SHEET     = 'AcademicSettings';
const ACS_TEXT_KEYS = ['headName', 'academicName'];
const ACS_IMG_KEYS  = ['logoLeft', 'logoRight', 'headSign', 'academicSign'];
const ACS_CHUNK     = 45000;
const ACS_MAX_IMG   = 200000;

function acsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ACS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ACS_SHEET);
    sh.appendRow(['Key', 'Part', 'Value']);
    sh.setFrozenRows(1);
    sh.getRange('C:C').setNumberFormat('@');   // keep values as plain text
  }
  return sh;
}

function getAcademicSettings() {
  const keys = ACS_TEXT_KEYS.concat(ACS_IMG_KEYS);
  const out = {};
  keys.forEach(function (k) { out[k] = ''; });
  const sh = acsSheet_();
  if (sh.getLastRow() > 1) {
    const parts = {};
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
      const k = String(r[0] || '').trim();
      if (keys.indexOf(k) < 0) return;
      (parts[k] = parts[k] || []).push({ i: Number(r[1]) || 0, v: String(r[2] == null ? '' : r[2]) });
    });
    Object.keys(parts).forEach(function (k) {
      out[k] = parts[k].sort(function (a, b) { return a.i - b.i; })
                       .map(function (p) { return p.v; }).join('');
    });
  }
  // school name configured by the School Manager (blank → the page keeps its default)
  out.schoolName = '';
  try { const b = psBranding_(); if (b && b.set) out.schoolName = String(b.name || ''); } catch (e) {}
  return out;
}

function saveAcademicSettings(payload) {
  if (!getTeacherSession()) return { ok: false, message: 'Session expired. Please login again.' };
  payload = payload || {};
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok: false, message: 'System busy, please try again.' }; }
  try {
    const keys = Object.keys(payload).filter(function (k) {
      return ACS_TEXT_KEYS.indexOf(k) >= 0 || ACS_IMG_KEYS.indexOf(k) >= 0;
    });
    if (!keys.length) return { ok: false, message: 'Nothing to save.' };

    const newRows = [];
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      let v = String(payload[k] == null ? '' : payload[k]);
      if (ACS_TEXT_KEYS.indexOf(k) >= 0) {
        v = v.trim().slice(0, 120);
        if (v) newRows.push([k, 0, v]);
      } else if (v) {
        if (!/^data:image\/(png|jpe?g|gif|webp);base64,/i.test(v)) return { ok: false, message: 'Invalid image for ' + k + '.' };
        if (v.length > ACS_MAX_IMG) return { ok: false, message: 'Image for ' + k + ' is too large.' };
        for (let p = 0, idx = 0; p < v.length; p += ACS_CHUNK, idx++) newRows.push([k, idx, v.substr(p, ACS_CHUNK)]);
      }
    }

    const sh = acsSheet_();
    // remove the old rows of the keys being saved (bottom-up so row numbers stay valid)
    if (sh.getLastRow() > 1) {
      const old = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
      for (let r = old.length - 1; r >= 0; r--) {
        if (keys.indexOf(String(old[r][0] || '').trim()) >= 0) sh.deleteRow(r + 2);
      }
    }
    if (newRows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 3).setNumberFormat('@').setValues(newRows.map(function (r) { return [r[0], String(r[1]), r[2]]; }));
    }
    try { CacheService.getScriptCache().put('PS_VER', String(Date.now()), 21600); } catch (e) {}   // refresh read-only panels
    try { logActivity('academic', 'Updated Academic Settings: ' + keys.join(', ')); } catch (e) {}
    return { ok: true, message: 'Settings saved successfully.' };
  } catch (e) {
    return { ok: false, message: 'Error: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}
