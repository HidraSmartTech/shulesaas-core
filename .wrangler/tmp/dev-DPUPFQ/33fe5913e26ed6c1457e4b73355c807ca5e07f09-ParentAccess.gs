// ============================================================
// PARENT ACCESS (v3) — managed by the CLASS TEACHER (ClassTeacher.html)
// ------------------------------------------------------------
// ONE LOGIN FOR EVERYONE:
// • Parents no longer have their own login page / link / username-only login.
// • Parents sign in on the SAME login (Index.html modal / Login.html) with
//   USERNAME + PASSWORD. checkLogin() detects the parent account
//   (role = 'parent') and opens the Parent Portal automatically.
// • The Class Teacher → "Manage Parents" panel:
//     - sees ONLY the students / parents of the class(es) assigned to him
//     - creates / edits / deletes parent accounts
//     - SETS the parent's Username and Password
//       (Username is suggested: Student first name + last 3 phone digits,
//        e.g. Asha + 0712 345 345 → Asha345, but he can change it)
// Uses the existing "Parents" sheet (see ensureParentsSheet() in Code.gs):
// [0]Name [1]Sex [2]Relation [3]Phone [4]Email [5]AdmNo
// [6]Address [7]Username [8]Password [9]StudentName [10]ClassName [11]Role
// ============================================================

// ── helpers ─────────────────────────────────────────────────
function ctpNorm_(v) { return String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' '); }

// Builds the username: first name (letters only, Capitalised) + last 3 phone digits
function ctpMakeUsername_(firstName, phone) {
  var fn = String(firstName || '').trim().split(/\s+/)[0].replace(/[^A-Za-z]/g, '');
  var digits = String(phone || '').replace(/\D/g, '');
  if (!fn || digits.length < 3) return '';
  return fn.charAt(0).toUpperCase() + fn.slice(1).toLowerCase() + digits.slice(-3);
}

// Classes the logged-in Class Teacher is responsible for: the class(es) given in Manage Classes
function ctpMyClasses_() {
  var t = getTeacherSession();
  if (!t) return null;
  var list = (t.homeroomClasses || []).map(function (c) { return String(c || '').trim(); }).filter(String);
  return { teacher: t, classes: list };
}

function ctpClassAllowed_(cls, classes) {
  return classes.some(function (c) { return ctpNorm_(c) === ctpNorm_(cls) || classLevelMatch(cls, c); });
}

// ── 1) data for the Manage Parents panel ────────────────────
function ctpGetParentsPanel() {
  var me = ctpMyClasses_();
  if (!me) return { ok: false, message: 'Session expired. Please login again.' };

  var students = [];
  me.classes.forEach(function (cls) {
    try {
      (getStudentsForClass(cls) || []).forEach(function (s) {
        students.push({
          admNo: s.admNo, firstName: s.firstName, lastName: s.lastName,
          name: (s.firstName + ' ' + s.lastName).trim(), sex: s.sex, className: cls
        });
      });
    } catch (e) {}
  });

  var parents = [];
  getParentList().forEach(function (r) {
    var cls = String(r[10] || '').trim();
    if (!cls || !ctpClassAllowed_(cls, me.classes)) return;
    var pw = String(r[8] == null ? '' : r[8]).trim();
    parents.push({
      name: String(r[0] || '').trim(), sex: String(r[1] || '').trim(), relation: String(r[2] || '').trim(),
      phone: String(r[3] || '').trim(), email: String(r[4] || '').trim(), admNo: String(r[5] || '').trim(),
      address: String(r[6] || '').trim(), username: String(r[7] || '').trim(),
      password: pw,                // class teacher can see/share the credentials he set
      hasPassword: !!pw,           // false → old account created before passwords: set one!
      studentName: String(r[9] || '').trim(), className: cls
    });
  });

  return { ok: true, classes: me.classes, students: students, parents: parents, link: ctpLoginLink_() };
}

// Suggested username for the form (teacher may change it)
function ctpSuggestUsername(firstName, phone) {
  return ctpMakeUsername_(firstName, phone);
}

// ── username must be unique across the WHOLE system (one login) ──
// exceptParentUser = the parent's current username when editing
function ctpUsernameTaken_(username, exceptParentUser) {
  var u = ctpNorm_(username), ex = ctpNorm_(exceptParentUser);
  if (!u) return false;
  // other parents
  var taken = getParentList().some(function (r) {
    var pu = ctpNorm_(r[7]);
    return pu && pu === u && pu !== ex;
  });
  if (taken) return true;
  // teachers / staff
  try { if (getTeacherList().some(function (t) { return ctpNorm_(t.username) === u; })) return true; } catch (e) {}
  // administrators
  try { if (typeof ADMIN_USERS !== 'undefined' && ADMIN_USERS.some(function (a) { return ctpNorm_(a.username) === u; })) return true; } catch (e) {}
  try { if (typeof psAdminByUsername_ === 'function' && psAdminByUsername_(username)) return true; } catch (e) {}
  return false;
}

// ── 2) create / update a parent account ─────────────────────
// data: { oldUsername?, admNo, name, sex, relation, phone, email, address, username, password }
//  • username  : blank on NEW → suggested one is used; blank on EDIT → unchanged
//  • password  : REQUIRED on NEW (min 4 chars); blank on EDIT → unchanged
function ctpSaveParent(data) {
  data = data || {};
  var me = ctpMyClasses_();
  if (!me) return { ok: false, message: 'Session expired. Please login again.' };

  var admNo = String(data.admNo || '').trim();
  var phone = String(data.phone || '').trim();
  var name  = String(data.name || '').trim();
  if (!admNo) return { ok: false, message: 'Please select a student.' };
  if (!name)  return { ok: false, message: 'Please enter the parent name.' };
  if (phone.replace(/\D/g, '').length < 9) return { ok: false, message: 'Please enter a valid phone number.' };

  // The student must belong to the teacher's class
  // (the same Adm No. can exist in different classes, so count the matches)
  var stu = null, found = 0;
  me.classes.forEach(function (cls) {
    var m = (getStudentsForClass(cls) || []).filter(function (s) { return String(s.admNo) === admNo; })[0];
    if (m && (!data.className || ctpNorm_(data.className) === ctpNorm_(cls))) { found++; if (!stu) { stu = m; stu.className = cls; } }
  });
  if (!stu) return { ok: false, message: 'This student is not in your class.' };
  if (found > 1) return { ok: false, message: 'Adm No. ' + admNo + ' exists in more than one of your classes. Contact the administrator.' };

  var oldU = ctpNorm_(data.oldUsername);
  var isEdit = !!oldU;

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { return { ok: false, message: 'System busy, please try again.' }; }
  try {
    var sh = ensureParentsSheet(), last = sh.getLastRow();
    var rows = last >= 2 ? sh.getRange(2, 1, last - 1, 12).getValues() : [];
    var editRow = -1, existingUser = '', existingPass = '';

    if (isEdit) {
      for (var i = 0; i < rows.length; i++) {
        if (ctpNorm_(rows[i][7]) === oldU) {
          if (!ctpClassAllowed_(rows[i][10], me.classes)) return { ok: false, message: 'You can only edit parents of your class.' };
          editRow = i + 2;
          existingUser = String(rows[i][7] || '').trim();
          existingPass = String(rows[i][8] == null ? '' : rows[i][8]).trim();
          break;
        }
      }
      if (editRow < 0) return { ok: false, message: 'Parent account not found.' };
    }

    // ── username ──
    var username = String(data.username || '').trim();
    if (!username) username = isEdit ? existingUser : ctpMakeUsername_(stu.firstName, phone);
    if (!username) return { ok: false, message: 'Please enter a username.' };
    if (!/^[A-Za-z0-9._-]{3,30}$/.test(username))
      return { ok: false, message: 'Username must be 3–30 characters: letters, numbers, dot, dash or underscore (no spaces).' };
    if (ctpUsernameTaken_(username, isEdit ? existingUser : ''))
      return { ok: false, message: 'Username "' + username + '" is already used by another account. Please choose a different username.' };

    // ── password ──
    var password = String(data.password == null ? '' : data.password).trim();
    if (!password) {
      if (isEdit && existingPass) password = existingPass;           // keep current password
      else return { ok: false, message: 'Please set a password for the parent.' };
    }
    if (password.length < 4) return { ok: false, message: 'Password must have at least 4 characters.' };
    if (/\s/.test(password)) return { ok: false, message: 'Password must not contain spaces.' };

    var row = [
      name, String(data.sex || '').trim(), String(data.relation || '').trim(), phone,
      String(data.email || '').trim(), admNo, String(data.address || '').trim(),
      username, password, (stu.firstName + ' ' + stu.lastName).trim(), stu.className, 'parent'
    ];
    var target = editRow > 0 ? editRow : last + 1;
    var rng = sh.getRange(target, 1, 1, 12);
    rng.setNumberFormat('@');            // keep phone / numeric passwords as text (no lost leading zeros)
    rng.setValues([row]);

    try { logActivity(me.teacher.username, (editRow > 0 ? 'Updated' : 'Added') + ' parent ' + name + ' (' + username + ')'); } catch (e) {}
    return { ok: true, username: username, password: password,
             message: 'Saved. Parent login → Username: ' + username + '  |  Password: ' + password };
  } finally {
    lock.releaseLock();
  }
}

// ── 3) delete a parent account ──────────────────────────────
function ctpDeleteParent(username) {
  var me = ctpMyClasses_();
  if (!me) return { ok: false, message: 'Session expired. Please login again.' };
  var sh = ensureParentsSheet(), last = sh.getLastRow();
  if (last < 2) return { ok: false, message: 'Not found.' };
  var rows = sh.getRange(2, 1, last - 1, 12).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (ctpNorm_(rows[i][7]) === ctpNorm_(username)) {
      if (!ctpClassAllowed_(rows[i][10], me.classes)) return { ok: false, message: 'You can only delete parents of your class.' };
      sh.deleteRow(i + 2);
      try { logActivity(me.teacher.username, 'Deleted parent ' + username); } catch (e) {}
      return { ok: true, message: 'Parent account deleted.' };
    }
  }
  return { ok: false, message: 'Not found.' };
}

// ── 4) system link (the SAME single login for everybody) ────
// There is no separate parent page any more. This only returns the normal
// web-app link, which the class teacher can send with the parent's
// username + password.
function ctpLoginLink_() {
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  return url;
}
function ctpGetLoginLink() {
  if (!ctpMyClasses_()) return { ok: false, message: 'Session expired. Please login again.' };
  var link = ctpLoginLink_();
  return link ? { ok: true, link: link } : { ok: false, message: 'Deploy the Web App first (Deploy → New deployment).' };
}

// REMOVED in v3 (single login):
//  • ctpParentPage_()        – the separate Parent page served by ?page=parent
//  • parentUsernameLogin()   – the username-only login (it had no password)
// Parents now log in through checkLogin(username, password) like everyone else.

// Parent-only reads: derive identity from the existing login token, never client class/name.
function parentIdentity_(token) {
  token=String(token||'');
  if(!/^[a-f0-9]{64}$/.test(token)) throw new Error('Please sign in with your parent username.');
  var props=PropertiesService.getScriptProperties(),key=SMS_SESSION_PREFIX+token;
  var raw=props.getProperty(key),session=raw?JSON.parse(raw):null;
  if(!session||session.role!=='parent'||!session.exp||session.exp<Date.now())
    throw new Error('Please sign in again to load your child\'s records.');
  var pd=smsParentData_(session.username);
  if(!pd||!String(pd.admNo||'').trim()) throw new Error('Your student account is not linked. Contact your class teacher.');
  var access=canAccessDashboard('parents',session.username);
  if(access&&access.allowed===false) throw new Error(access.reason||'Parent portal is disabled.');
  session.exp=Date.now()+SMS_SESSION_HOURS*3600*1000;
  props.setProperty(key,JSON.stringify(session));
  return pd;
}
function parentGetReportCard(token) {
  try {
    var pd=parentIdentity_(token),adm=String(pd.admNo).trim(),exams=[],examCount=0,withScores=0;
    (fetchExamTypes(pd.className)||[]).forEach(function(examName){
      // Same ranked results used by Academics, narrowed before returning to the browser.
      var all=getExamResults(pd.className,examName)||[];
      examCount++;if(all.length)withScores++;
      var matches=all.filter(function(s){return String(s.admNo||'').trim()===adm;});
      if(!matches.length){
        // Scores saved without an admission number are linked by full name, but only when that name
        // is unique in the class register and the score row carries no admission number of its own.
        var nm=ctpNorm_(pd.studentName),same=(getStudentsForClass(pd.className)||[]).filter(function(st){
          return ctpNorm_(String(st.firstName||'')+' '+String(st.lastName||''))===nm;});
        if(nm&&same.length===1)matches=all.filter(function(s){return !String(s.admNo||'').trim()&&ctpNorm_(s.name)===nm;});
      }
      if(matches.length!==1)return; // fail closed on missing or ambiguous identity
      var source=matches[0],report={};
      ['name','admNo','sex','subjects','total','average','grade','point','division','position','outOf','remarks'].forEach(function(k){report[k]=source[k];});
      report.admNo=adm;report.outOf=all.length;report.tabia=[];report.classTeacher='';
      // Teacher name is the same as Academics; ratings join by admission number, not name.
      var b=fetchReportBehaviors(pd.className,examName)||{};
      report.classTeacher=b.classTeacher||'';
      var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BEHAVIOR_SHEET),n=BEHAVIOR_CRITERIA.length;
      if(sh&&sh.getLastRow()>1){sh.getRange(2,1,sh.getLastRow()-1,4+n+4).getValues().forEach(function(r){
        if(behClassKey_(r[0])===behClassKey_(pd.className)&&behNorm_(r[1])===behNorm_(examName)&&String(r[2]||'').trim()===adm&&String(r[4+n+2]||'').trim()==='Submitted')
          report.tabia=r.slice(4,4+n).map(function(v){return String(v||'').trim();});
      });}
      exams.push({examName:examName,students:[report]});
    });
    var settings={},a=getAcademicSettings()||{};
    ['schoolName','logoLeft','logoRight','academicName','academicSign','headName','headSign'].forEach(function(k){settings[k]=a[k]||'';});
    // when nothing can be shown, say why (so the school knows what to fix)
    var hint='';
    if(!exams.length){
      if(!examCount)hint='No examinations have been set up for your child\'s class yet.';
      else if(!withScores)hint='Scores for your child\'s class have not been entered yet.';
      else hint='Scores have been entered, but none could be linked to the admission number on your account. Please ask the class teacher to check it.';
    }
    return {ok:true,exams:exams,settings:settings,hint:hint};
  } catch(e){return {ok:false,message:e.message||'Could not load report card.'};}
}
function parentAttendanceMark_(v){var s=String(v==null?'':v).trim().toUpperCase();return ['P','A','L'].indexOf(s)>=0?s:'';}
function parentGetAttendance(token) {
  try {
    var pd=parentIdentity_(token),adm=String(pd.admNo).trim();
    var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName('Attendance');
    if(!sh||sh.getLastRow()<2)return {ok:true,records:[]};
    var rows=sh.getDataRange().getValues(),headers=rows.shift();
    function col(names){for(var i=0;i<headers.length;i++){var h=String(headers[i]).toLowerCase().replace(/[^a-z0-9]/g,'');if(names.indexOf(h)>=0)return i;}return -1;}
    var ad=col(['admno','admissionnumber']),dt=col(['date','datestr']),mo=col(['morning']),af=col(['afternoon']),cl=col(['class','classname']);
    if(ad<0||dt<0||mo<0||af<0)throw new Error('Attendance columns are missing. Contact the class teacher.');
    var days=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'],byDate={};
    rows.forEach(function(r){
      if(String(r[ad]||'').trim()!==adm||!r[dt])return;
      // the same Adm No. may exist in another class: only this child's own class counts
      if(cl>=0&&pd.className&&String(r[cl]||'').trim()&&ctpNorm_(r[cl]).replace(/\s+/g,'')!==ctpNorm_(pd.className).replace(/\s+/g,''))return;
      var date=r[dt] instanceof Date?Utilities.formatDate(r[dt],ss.getSpreadsheetTimeZone(),'yyyy-MM-dd'):String(r[dt]).trim().slice(0,10);
      if(!/^\d{4}-\d{2}-\d{2}$/.test(date))return;
      var dayDate=new Date(date+'T12:00:00Z');if(isNaN(dayDate.getTime()))return;
      var morning=parentAttendanceMark_(r[mo]),afternoon=parentAttendanceMark_(r[af]);
      // Daily counters use morning registration; both teacher-entered sessions remain visible.
      byDate[date]={date:date,day:days[dayDate.getUTCDay()],morning:morning,afternoon:afternoon,status:({P:'Present',A:'Absent',L:'Late'})[morning]||'Not marked'};
    });
    return {ok:true,records:Object.keys(byDate).sort().map(function(d){return byDate[d];})};
  }catch(e){return {ok:false,message:e.message||'Could not load attendance.'};}
}
