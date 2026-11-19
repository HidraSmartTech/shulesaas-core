/**
 * ParentPortal.gs — real data for the Parent portal (no sample data)
 * The Parent page sends its session token; the server returns ONLY that parent's child:
 * the same Individual Report Card data the Academic account uses, real attendance, and the
 * Academic Settings (logos / names / signatures) so the card looks exactly like Academic's.
 */
var rbacStudentCache_ = null;

function rbacNormAdm_(v) { return String(v || '').replace(/\s+/g, '').toLowerCase(); }

/** "Amina John Mangube John Mangube" -> "Amina John Mangube" (removes an immediately repeated word group) */
function rbacCleanName_(n) {
  n = String(n || '').replace(/\s+/g, ' ').trim();
  var prev;
  do { prev = n; n = n.replace(/\b(\S+(?:\s\S+)*?)\s\1(?=\s|$)/i, '$1'); } while (n !== prev);
  return n;
}

/** Current name / class / sex of a student from the registration sheets (not the stale copy stored with the parent). */
function rbacLiveStudent_(admNo, storedName, storedClass) {
  var fallback = { name: rbacCleanName_(storedName), className: String(storedClass || '').trim(), sex: '' };
  try {
    if (!rbacStudentCache_) rbacStudentCache_ = getStudentListRaw();
    var key = rbacNormAdm_(admNo); if (!key) return fallback;
    var hits = rbacStudentCache_.filter(function (s) { return rbacNormAdm_(s.admNo) === key; });
    if (!hits.length) return fallback;
    var pick = hits.filter(function (s) { return String(s.className || '').trim() === fallback.className; })[0] || hits[0];
    var nm = rbacCleanName_(String(pick.firstName || '').trim() + ' ' + String(pick.lastName || '').trim());
    return { name: nm || fallback.name, className: String(pick.className || fallback.className).trim(), sex: String(pick.sex || '').trim() };
  } catch (e) { return fallback; }
}

function parentFromToken_(token) {
  token = String(token || '');
  if (!/^[a-f0-9]{64}$/.test(token)) throw new Error('Session expired. Please log in again.');
  var raw = PropertiesService.getScriptProperties().getProperty(SMS_SESSION_PREFIX + token);
  if (!raw) throw new Error('Session expired. Please log in again.');
  var s = JSON.parse(raw);
  if (!s.exp || s.exp < Date.now() || s.role !== 'parent') throw new Error('Session expired. Please log in again.');
  var pd = smsParentData_(s.username);
  if (!pd) throw new Error('Session expired. Please log in again.');
  if (rbacIsInactiveParent_(s.username)) throw new Error('This account has been deactivated. Contact the school office.');
  var pa = canAccessDashboard('parents', s.username);
  if (pa && pa.allowed === false) throw new Error(pa.reason || 'This portal is currently disabled by the administrator.');
  return { username: s.username, data: pd };
}

function parentPortalData(token) {
  var p = parentFromToken_(token), pd = p.data;
  var live = rbacLiveStudent_(pd.admNo, pd.studentName, pd.className);
  var className = live.className, myKey = rbacNormAdm_(pd.admNo), myName = behNorm_(live.name);
  var feat = rbacFeatures_('parents');
  var wantReports = feat['Report Card'] !== false || feat['Performance'] !== false;
  var wantAtt = feat['Attendance'] !== false;

  var reports = [];
  if (wantReports && className) {
    var exams = []; try { exams = fetchExamTypes(className) || []; } catch (e) { exams = []; }
    exams.forEach(function (exam) {
      var list = []; try { list = getExamResults(className, exam) || []; } catch (e) { list = []; }
      if (!list.length) return;
      var me = list.filter(function (r) { return myKey && rbacNormAdm_(r.admNo) === myKey; })[0]
            || list.filter(function (r) { return behNorm_(r.name) === myName; })[0];
      if (!me) return;                                     // this child has no results in this exam
      var beh = { map: {}, classTeacher: '' };
      try { beh = fetchReportBehaviors(className, exam) || beh; } catch (e) {}
      me.outOf = list.length;
      me.tabia = (beh.map || {})[behNorm_(me.name)] || null;
      me.classTeacher = beh.classTeacher || '';
      reports.push({ examName: exam, report: me });
    });
  }

  var attendance = [];
  if (wantAtt && className) {
    try {
      attendance = getAttendanceRecords(className).filter(function (r) {
        return (myKey && rbacNormAdm_(r.admNo) === myKey) || behNorm_(r.name) === myName;
      }).map(function (r) { return { date: r.date, day: r.day, status: r.status }; });
    } catch (e) { attendance = []; }
  }

  var a = acsReadAll_();
  return {
    student: { name: live.name, admNo: pd.admNo, className: className, sex: live.sex },
    reports: reports,
    attendance: attendance,
    settings: { logoLeft: a.logoLeft || '', logoRight: a.logoRight || '', headName: a.headName || '', headSign: a.headSign || '',
                academicName: a.academicName || '', academicSign: a.academicSign || '', schoolName: a.schoolName || '', schoolLogo: a.schoolLogo || '', schoolInfo: a.schoolInfo || {} },
    features: { report: feat['Report Card'] !== false, performance: feat['Performance'] !== false, attendance: wantAtt }
  };
}