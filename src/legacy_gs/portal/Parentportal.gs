/**
 * ParentPortal.gs — real data for the Parent portal (no sample data)
 * Supports both Cloudflare D1 host sessions (__HOST__.session) and legacy script property tokens.
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
  var username = '';

  // 1. Cloudflare D1 authenticated host session
  if (typeof __HOST__ !== 'undefined' && __HOST__ && __HOST__.session && __HOST__.session.username) {
    if (__HOST__.session.role !== 'parent' && __HOST__.session.role !== 'student') {
      throw new Error('Access denied. Parent account required.');
    }
    username = String(__HOST__.session.username || '').trim();
  } else {
    // 2. Legacy GAS token fallback
    token = String(token || '');
    if (!/^[a-f0-9]{64}$/.test(token)) throw new Error('Session expired. Please log in again.');
    var raw = PropertiesService.getScriptProperties().getProperty(SMS_SESSION_PREFIX + token);
    if (!raw) throw new Error('Session expired. Please log in again.');
    var s = JSON.parse(raw);
    if (!s.exp || s.exp < Date.now() || s.role !== 'parent') throw new Error('Session expired. Please log in again.');
    username = s.username;
  }

  var pd = smsParentData_(username);
  if (!pd) throw new Error('Parent profile not found for this student. Contact the school office.');
  if (rbacIsInactiveParent_(username)) throw new Error('This account has been deactivated. Contact the school office.');
  var pa = canAccessDashboard('parents', username);
  if (pa && pa.allowed === false) throw new Error(pa.reason || 'This portal is currently disabled by the administrator.');
  return { username: username, data: pd };
}

function parentBuildFeeSummary_(admNo, className) {
  var key = rbacNormAdm_(admNo);
  var d1Pays = (typeof __HOST__ !== 'undefined' && __HOST__ && Array.isArray(__HOST__.d1Payments)) ? __HOST__.d1Payments : [];
  var d1Fees = (typeof __HOST__ !== 'undefined' && __HOST__ && Array.isArray(__HOST__.d1FeeStructures)) ? __HOST__.d1FeeStructures : [];
  var currentYear = String(new Date().getFullYear());

  var matchingPays = d1Pays.filter(function (p) {
    return rbacNormAdm_(p.student_id) === key && String(p.status || 'completed').toLowerCase() === 'completed';
  });

  var totalPaid = 0;
  var history = matchingPays.map(function (p) {
    var amt = Number(p.amount_paid || 0);
    totalPaid += amt;
    return {
      receiptNo: p.receipt_no || '',
      date: String(p.payment_date || '').slice(0, 10),
      amount: amt,
      currency: p.currency || 'TZS',
      method: p.payment_method || 'cash',
      category: p.fee_category || 'Tuition',
      reference: p.reference_no || '',
      academicYear: p.academic_year || currentYear
    };
  });

  var normClass = String(className || '').replace(/\s+/g, '').toLowerCase();
  var expectedFee = 0;
  d1Fees.forEach(function (f) {
    var fClass = String(f.class_level || '').replace(/\s+/g, '').toLowerCase();
    var fYear = String(f.academic_year || currentYear);
    if (fYear === currentYear && (fClass === 'all' || (normClass && fClass === normClass))) {
      expectedFee += Number(f.amount || 0);
    }
  });

  return {
    academicYear: currentYear,
    totalFee: Math.round(expectedFee * 100) / 100,
    totalPaid: Math.round(totalPaid * 100) / 100,
    balance: Math.max(0, Math.round((expectedFee - totalPaid) * 100) / 100),
    payments: history
  };
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
      if (!me) return;
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

  var fees = parentBuildFeeSummary_(pd.admNo, className);
  var a = acsReadAll_();
  return {
    student: { name: live.name, admNo: pd.admNo, className: className, sex: live.sex },
    reports: reports,
    attendance: attendance,
    fees: fees,
    settings: { logoLeft: a.logoLeft || '', logoRight: a.logoRight || '', headName: a.headName || '', headSign: a.headSign || '',
                academicName: a.academicName || '', academicSign: a.academicSign || '', schoolName: a.schoolName || '', schoolLogo: a.schoolLogo || '', schoolInfo: a.schoolInfo || {} },
    features: { report: feat['Report Card'] !== false, performance: feat['Performance'] !== false, attendance: wantAtt }
  };
}
