/**
 * ============================================================================
 * DashboardAccess.gs — Unified Dashboard Access Control & Parent Management
 * ============================================================================
 * Multi-Tenant SaaS Engine (Cloudflare Workers + D1 + GAS Runtime)
 *
 * Enforces:
 *   1) Role-based and individual access policies across 8 system roles:
 *      'schoolManager', 'headOfSchool', 'secondMaster', 'academics',
 *      'classTeacher', 'teacher', 'accountant', 'parents'
 *   2) Plan-level restrictions (Basic, Pro, Advanced)
 *   3) Live dashboard access validation and emergency locks
 *   4) Read-only modes & feature switches per dashboard panel
 *   5) Parent Management (Class Teacher -> Manage Parents panel)
 *   6) Parent Portal records access (Report Cards & Attendance)
 * ============================================================================
 */

/* ═════════════════════════════════════════════════════════════════════════
   SECTION 1: CONSTANTS & DASHBOARD DEFINITIONS (ALL 8 ROLES)
   ═════════════════════════════════════════════════════════════════════════ */

var DASHBOARD_POLICY_PROP = 'DASHBOARD_ACCESS_POLICIES_V2';
var DASHBOARD_EVENT_PROP  = 'DASHBOARD_ACCESS_EVENTS_V2';

var DASHBOARD_KEYS = [
  'schoolManager', 'headOfSchool', 'secondMaster', 'academics',
  'classTeacher', 'teacher', 'accountant', 'parents'
];

var DASHBOARD_TITLES = {
  schoolManager: 'School Manager',
  headOfSchool:  'Head of School',
  secondMaster:  'Second Master / Depute',
  academics:     'Academic Master',
  classTeacher:  'Class Teacher',
  teacher:       'Subject Teacher',
  accountant:    'Bursar / Accountant',
  parents:       'Parent Portal'
};

var DASHBOARD_FEATURES = {
  schoolManager: [
    'School Settings', 'Manage Users', 'Manage Backups', 'Audit Logs',
    'Financial Overview', 'System Lock', 'Subscription & Billing'
  ],
  headOfSchool: [
    'School Overview', 'Approve Reports', 'Staff Management', 'TOD Roster Review',
    'Academic Performance', 'Publish Results', 'Emergency Lock'
  ],
  secondMaster: [
    'My Assignments', 'Submit Scores', 'Score Tracking', 'Submission Report',
    'Attendance Tracking', 'TOD Tracking', 'Lesson Plan', 'Discipline Records'
  ],
  academics: [
    'Student Admissions', 'Manage Classes', 'Map Students', 'Year Progression',
    'Attendance Tracking', 'Manage Subjects', 'Exam Setup', 'Score Deadline',
    'Publish Results', 'Report Cards', 'Submit Scores', 'Tracking Scores',
    'Timetable Settings', 'TOD Report', 'Messages', 'Academic Settings',
    'Manage Staff Accounts', 'Manage Parent Accounts', 'Create Class Teacher Account',
    'Create Subject Teacher Account', 'Create Depute Headmaster Account',
    'Create Parent Account', 'Lesson Plan'
  ],
  classTeacher: [
    'My Assignments', 'My Timetable', 'Behavioural Criteria', 'Submit Scores',
    'Score Tracking', 'Submission Report', 'Attendance', 'TOD Report',
    'Messages', 'Lesson Plan', 'Manage Parents'
  ],
  teacher: [
    'My Assignments', 'My Timetable', 'Submit Scores', 'Score Tracking',
    'TOD Report', 'Submission Report', 'Messages', 'Lesson Plan'
  ],
  accountant: [
    'Fee Structures', 'Record Payments', 'Fee Statements', 'Defaulters List',
    'Financial Reports', 'Expenses', 'SMS Reminders'
  ],
  parents: [
    'Report Card', 'Performance', 'Attendance', 'Announcements', 'Student Info', 'Fee Status'
  ]
};

/* ═════════════════════════════════════════════════════════════════════════
   SECTION 2: NORMALISATION & POLICY ENGINE HELPERS
   ═════════════════════════════════════════════════════════════════════════ */

function _now_() { return new Date().toISOString(); }

function _dashNorm_(v) {
  return String(v || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function _userNorm_(v) {
  return String(v || '').trim().toLowerCase();
}

function _bool_(v, def) {
  if (v === undefined || v === null || v === '') return !!def;
  if (v === true || v === false) return v;
  if (typeof v === 'number') return v !== 0;
  var s = String(v).trim().toLowerCase();
  if (['false','no','0','off','disabled','deny','denied','blocked','closed','locked'].indexOf(s) >= 0) return false;
  if (['true','yes','1','on','enabled','allow','allowed','open','active'].indexOf(s) >= 0) return true;
  return !!def;
}

function _isBlockingValue_(v) {
  if (v === true) return true;
  if (v === false || v === undefined || v === null || v === '') return false;
  var s = String(v).trim().toLowerCase();
  return ['true','yes','1','on','blocked','block','disabled','disable','denied','deny','locked','lock','closed','inactive','off'].indexOf(s) >= 0;
}

function normalizeDashboardKey_(key) {
  if (!key) return 'teacher';
  var raw = String(key || '').trim();
  if (DASHBOARD_KEYS.indexOf(raw) >= 0) return raw;
  var k = _dashNorm_(raw);

  if (k.indexOf('schoolmanager') >= 0 || k === 'manager') return 'schoolManager';
  if (k.indexOf('headofschool') >= 0 || k.indexOf('headmaster') >= 0 || k === 'hos') return 'headOfSchool';
  if (k.indexOf('secondmaster') >= 0 || k.indexOf('depute') >= 0 || k === 'second') return 'secondMaster';
  if (k.indexOf('academic') >= 0 || k === 'academics' || k === 'academicmaster') return 'academics';
  if (k.indexOf('classteacher') >= 0 || k === 'classteachers' || k === 'classmaster') return 'classTeacher';
  if (k.indexOf('accountant') >= 0 || k === 'bursar' || k === 'finance') return 'accountant';
  if (k.indexOf('parent') >= 0 || k === 'parents') return 'parents';
  if (k.indexOf('teacher') >= 0 || k === 'teachers') return 'teacher';
  return raw;
}

function _dashboardAliasKeys_(key) {
  key = normalizeDashboardKey_(key);
  var common = [key, _dashNorm_(key), DASHBOARD_TITLES[key], _dashNorm_(DASHBOARD_TITLES[key])];
  if (key === 'schoolManager') common = common.concat(['schoolManager', 'schoolmanager', 'School Manager']);
  if (key === 'headOfSchool')  common = common.concat(['headOfSchool', 'headofschool', 'Head of School']);
  if (key === 'secondMaster')  common = common.concat(['secondMaster', 'secondmaster', 'Second Master', 'second-master']);
  if (key === 'academics')     common = common.concat(['academic', 'academics', 'Academic Master']);
  if (key === 'classTeacher')  common = common.concat(['classTeacher', 'classteacher', 'Class Teacher']);
  if (key === 'teacher')       common = common.concat(['teacher', 'teachers', 'Subject Teacher']);
  if (key === 'accountant')    common = common.concat(['accountant', 'bursar', 'finance']);
  if (key === 'parents')       common = common.concat(['parent', 'parents', 'Parent Portal']);
  return common.filter(function(v, i, a){ return v && a.indexOf(v) === i; });
}

function _defaultDashboardPolicy_(key) {
  key = normalizeDashboardKey_(key);
  var features = {};
  (DASHBOARD_FEATURES[key] || []).forEach(function(f){
    features[f] = !/^(Manage (Staff|Parent) Accounts|Create )/.test(f);
  });
  return {
    key: key, allowed: true, readOnly: false, emergencyLock: false, temporaryLock: false,
    expires: '', reason: '', notice: '', forceLogoutAt: '',
    features: features, userOverrides: {}, updated: 'Never', updatedBy: ''
  };
}

function _mergePolicy_(base, incoming) {
  incoming = incoming || {};
  var out = Object.assign({}, base, incoming);
  out.features = Object.assign({}, base.features || {}, incoming.features || {});
  out.userOverrides = Object.assign({}, base.userOverrides || {}, incoming.userOverrides || {}, incoming.users || {}, incoming.userAccess || {});
  
  var cleanedOverrides = {};
  Object.keys(out.userOverrides || {}).forEach(function(u){
    if (!u) return;
    var ov = out.userOverrides[u] || {};
    var nu = _userNorm_(ov.username || u);
    if (!nu) return;
    ov.username = ov.username || u;
    if (Object.prototype.hasOwnProperty.call(ov, 'allowed')) ov.allowed = _bool_(ov.allowed, true);
    ov.readOnly = _bool_(ov.readOnly, false);
    ov.temporaryLock = _bool_(ov.temporaryLock, false) || _bool_(ov.timedLock, false);
    cleanedOverrides[nu] = ov;
  });
  out.userOverrides = cleanedOverrides;

  var allowed = _bool_(out.allowed, true);
  if (Object.prototype.hasOwnProperty.call(out, 'enabled')) allowed = _bool_(out.enabled, allowed);
  if (Object.prototype.hasOwnProperty.call(out, 'active')) allowed = _bool_(out.active, allowed);
  if (_isBlockingValue_(out.blocked) || _isBlockingValue_(out.disabled) || _isBlockingValue_(out.locked)) allowed = false;
  if (out.status && ['blocked','disabled','denied','closed','locked','inactive','off'].indexOf(String(out.status).trim().toLowerCase()) >= 0) allowed = false;
  if (out.access && ['blocked','disabled','denied','closed','locked','inactive','off'].indexOf(String(out.access).trim().toLowerCase()) >= 0) allowed = false;

  out.allowed = allowed;
  out.readOnly = _bool_(out.readOnly, false) || _isBlockingValue_(out.readonly);
  out.emergencyLock = _bool_(out.emergencyLock, false) || _isBlockingValue_(out.emergency) || _isBlockingValue_(out.locked);
  out.temporaryLock = _bool_(out.temporaryLock, false) || _bool_(out.timedLock, false) || _bool_(out.autoUnlock, false);

  Object.keys(out.features || {}).forEach(function(k){ out.features[k] = _bool_(out.features[k], true); });
  return out;
}

function _logDashboardAccessEvent_(action, key, adminUser, details) {
  try {
    var props = PropertiesService.getScriptProperties();
    var events = [];
    try { events = JSON.parse(props.getProperty(DASHBOARD_EVENT_PROP) || '[]') || []; } catch(e){ events = []; }
    events.unshift({ time: _now_(), action: action, panel: key, admin: adminUser || 'admin', details: details || '' });
    props.setProperty(DASHBOARD_EVENT_PROP, JSON.stringify(events.slice(0, 200)));
  } catch(e) {}
}

/* ═════════════════════════════════════════════════════════════════════════
   SECTION 3: POLICY STORAGE & MANAGEMENT
   ═════════════════════════════════════════════════════════════════════════ */

function _readDashboardPolicyStores_() {
  var props = PropertiesService.getScriptProperties();
  var names = [
    'DASHBOARD_ACCESS_POLICIES',
    'DASHBOARD_ACCESS_POLICY',
    'DASHBOARD_RESTRICTIONS',
    'DASHBOARD_ACCESS',
    'DASHBOARD_POLICIES',
    'dashboardAccessPolicies',
    'dashboardPolicies',
    DASHBOARD_POLICY_PROP
  ];
  var stores = [];
  names.forEach(function(name){
    try {
      var raw = props.getProperty(name);
      if (!raw) return;
      var obj = JSON.parse(raw);
      if (obj) stores.push(obj);
    } catch(e) {}
  });
  return stores;
}

function _normalisePolicyStore_(store) {
  if (!store) return {};
  if (Array.isArray(store)) {
    var o = {};
    store.forEach(function(row){
      if (!row) return;
      var k = row.key || row.panel || row.dashboard || row.name || row.id || row.title;
      if (k) o[k] = row;
    });
    return o;
  }
  if (store.policies) return _normalisePolicyStore_(store.policies);
  if (store.dashboards) return _normalisePolicyStore_(store.dashboards);
  if (store.access) return _normalisePolicyStore_(store.access);
  if (store.dashboardAccess) return _normalisePolicyStore_(store.dashboardAccess);
  return store;
}

function getDashboardAccessPolicies() {
  var stores = _readDashboardPolicyStores_().map(_normalisePolicyStore_);
  var out = {};
  DASHBOARD_KEYS.forEach(function(key){
    var policy = _defaultDashboardPolicy_(key);
    stores.forEach(function(saved){
      _dashboardAliasKeys_(key).forEach(function(alias){
        if (saved[alias]) policy = _mergePolicy_(policy, saved[alias]);
        var n = _dashNorm_(alias);
        if (saved[n]) policy = _mergePolicy_(policy, saved[n]);
      });
      Object.keys(saved || {}).forEach(function(savedKey){
        if (normalizeDashboardKey_(savedKey) === key || _dashNorm_(savedKey) === _dashNorm_(key)) {
          policy = _mergePolicy_(policy, saved[savedKey]);
        }
      });
    });
    policy.key = key;
    out[key] = policy;
  });
  return out;
}

function getDashboardPolicy(key) {
  key = normalizeDashboardKey_(key);
  return getDashboardAccessPolicies()[key] || _defaultDashboardPolicy_(key);
}

function saveDashboardAccessState(key, policy, adminUser) {
  key = normalizeDashboardKey_(key);
  if (DASHBOARD_KEYS.indexOf(key) === -1) throw new Error('Unknown dashboard key: ' + key);
  var all = getDashboardAccessPolicies();
  var def = _defaultDashboardPolicy_(key);
  all[key] = _mergePolicy_(def, Object.assign({}, policy || {}, { key: key, updated: _now_(), updatedBy: adminUser || 'admin' }));
  var props = PropertiesService.getScriptProperties();
  props.setProperty(DASHBOARD_POLICY_PROP, JSON.stringify(all));
  try { props.setProperty('DASHBOARD_ACCESS_POLICIES', JSON.stringify(all)); } catch(e) {}
  _logDashboardAccessEvent_('POLICY_SAVED', key, adminUser || 'admin', all[key].reason || '');
  return 'Dashboard access policy saved for ' + (DASHBOARD_TITLES[key] || key);
}

function saveDashboardAccessPolicies(policies, adminUser) {
  policies = policies || {};
  var current = getDashboardAccessPolicies();

  Object.keys(policies).forEach(function(inputKey){
    var key = normalizeDashboardKey_(inputKey);
    if (DASHBOARD_KEYS.indexOf(key) === -1) return;
    var def = _defaultDashboardPolicy_(key);
    current[key] = _mergePolicy_(def, Object.assign({}, policies[inputKey] || {}, { key: key, updated: _now_(), updatedBy: adminUser || 'admin' }));
  });

  var props = PropertiesService.getScriptProperties();
  props.setProperty(DASHBOARD_POLICY_PROP, JSON.stringify(current));
  try { props.setProperty('DASHBOARD_ACCESS_POLICIES', JSON.stringify(current)); } catch(e) {}
  _logDashboardAccessEvent_('BULK_POLICY_SAVED', 'all', adminUser || 'admin', 'Bulk dashboard policy update');
  return 'All dashboard access policies saved';
}

function saveDashboardUserAccess(key, username, override, adminUser) {
  key = normalizeDashboardKey_(key || 'teacher');
  var userKey = _userNorm_(username);
  if (DASHBOARD_KEYS.indexOf(key) === -1) throw new Error('Unknown dashboard key: ' + key);
  if (!userKey) throw new Error('Missing username');

  var all = getDashboardAccessPolicies();
  var pol = all[key] || _defaultDashboardPolicy_(key);
  pol.userOverrides = pol.userOverrides || {};

  override = override || {};
  var ov = {
    username: String(username || '').trim(),
    allowed: _bool_(override.allowed, true),
    readOnly: _bool_(override.readOnly, false),
    temporaryLock: _bool_(override.temporaryLock, false),
    expires: override.expires || '',
    reason: override.reason || (override.allowed === false ? 'Restricted by administrator' : ''),
    notice: override.notice || '',
    updated: _now_(),
    updatedBy: adminUser || 'admin'
  };
  pol.userOverrides[userKey] = ov;
  pol.updated = _now_();
  pol.updatedBy = adminUser || 'admin';
  all[key] = pol;

  var props = PropertiesService.getScriptProperties();
  props.setProperty(DASHBOARD_POLICY_PROP, JSON.stringify(all));
  try { props.setProperty('DASHBOARD_ACCESS_POLICIES', JSON.stringify(all)); } catch(e) {}
  _logDashboardAccessEvent_('USER_POLICY_SAVED', key, adminUser || 'admin', userKey + ' → ' + (ov.allowed ? 'allowed' : 'blocked') + (ov.reason ? ' — ' + ov.reason : ''));
  try { if (typeof logActivity === 'function') logActivity(adminUser || 'admin', 'Set individual dashboard access: ' + userKey + ' / ' + key + ' → ' + (ov.allowed ? 'allowed' : 'blocked')); } catch(e) {}
  return 'Individual access saved for ' + username;
}

function clearDashboardUserAccess(key, username, adminUser) {
  key = normalizeDashboardKey_(key || 'teacher');
  var userKey = _userNorm_(username);
  if (!userKey) throw new Error('Missing username');
  var all = getDashboardAccessPolicies();
  var pol = all[key] || _defaultDashboardPolicy_(key);
  pol.userOverrides = pol.userOverrides || {};
  delete pol.userOverrides[userKey];
  pol.updated = _now_();
  pol.updatedBy = adminUser || 'admin';
  all[key] = pol;
  var props = PropertiesService.getScriptProperties();
  props.setProperty(DASHBOARD_POLICY_PROP, JSON.stringify(all));
  try { props.setProperty('DASHBOARD_ACCESS_POLICIES', JSON.stringify(all)); } catch(e) {}
  _logDashboardAccessEvent_('USER_POLICY_CLEARED', key, adminUser || 'admin', userKey + ' returned to dashboard default');
  return 'Individual access cleared for ' + username;
}

function getDashboardKeyFromPosition(position) {
  var pos = String(position || '').trim();
  return normalizeDashboardKey_(pos || 'teacher');
}

/* ═════════════════════════════════════════════════════════════════════════
   SECTION 4: CORE ACCESS VALIDATION & LIVE CHECKS
   ═════════════════════════════════════════════════════════════════════════ */

function canAccessDashboard(key, username) {
  key = normalizeDashboardKey_(key || 'teacher');
  if (DASHBOARD_KEYS.indexOf(key) === -1) key = getDashboardKeyFromPosition(key);
  var policy = getDashboardPolicy(key);
  var userKey = _userNorm_(username || '');
  var override = userKey ? ((policy.userOverrides || {})[userKey] || null) : null;

  var allowed = !!policy.allowed && !policy.emergencyLock;
  var readOnly = !!policy.readOnly;
  var reason = policy.reason || (allowed ? '' : 'This dashboard has been disabled by the administrator.');
  var notice = policy.notice || '';

  if (!allowed && policy.temporaryLock === true && policy.expires) {
    var expiry = new Date(policy.expires);
    if (!isNaN(expiry.getTime()) && new Date() > expiry) {
      policy.allowed = true;
      policy.emergencyLock = false;
      policy.temporaryLock = false;
      policy.reason = '';
      saveDashboardAccessState(key, policy, 'system (temporary lock expired)');
      allowed = true;
      reason = '';
    }
  }

  if (override && !policy.emergencyLock) {
    var ov = override;
    if (ov.temporaryLock === true && ov.expires && ov.allowed === false) {
      var uExpiry = new Date(ov.expires);
      if (!isNaN(uExpiry.getTime()) && new Date() > uExpiry) {
        try { clearDashboardUserAccess(key, userKey, 'system (individual temporary lock expired)'); } catch(e) {}
        ov = null;
      }
    }

    if (ov) {
      if (ov.allowed === false) {
        allowed = false;
        reason = ov.reason || 'Your individual dashboard access has been restricted by the administrator.';
      } else if (ov.allowed === true) {
        allowed = true;
        reason = '';
      }
      if (ov.readOnly === true) readOnly = true;
      if (ov.notice) notice = ov.notice;
    }
  }

  if (policy.emergencyLock) {
    allowed = false;
    reason = policy.reason || 'This dashboard is under emergency lock by the administrator.';
  }

  return {
    key:           key,
    title:         DASHBOARD_TITLES[key] || key,
    allowed:       allowed,
    readOnly:      readOnly,
    emergencyLock: !!policy.emergencyLock,
    temporaryLock: !!policy.temporaryLock,
    reason:        reason,
    expires:       (override && override.expires) || policy.expires || '',
    notice:        notice,
    forceLogoutAt: policy.forceLogoutAt || '',
    features:      policy.features || {},
    userOverride:  override || null,
    policy:        policy
  };
}

function _denyDashboardAccess_(key, reason, notice) {
  key = normalizeDashboardKey_(key || 'teacher');
  return {
    key: key,
    title: DASHBOARD_TITLES[key] || key,
    allowed: false,
    readOnly: false,
    emergencyLock: true,
    temporaryLock: false,
    reason: reason || 'Access denied by administrator.',
    expires: '',
    notice: notice || '',
    forceLogoutAt: '',
    features: {},
    policy: {}
  };
}

function logDashboardBlockedAttempt(username, key, access, source) {
  try {
    key = normalizeDashboardKey_(key || (access && access.key) || 'teacher');
    access = access || {};
    var title = access.title || DASHBOARD_TITLES[key] || key;
    var reason = access.reason || 'Blocked by administrator';
    var who = username || 'unknown user';
    try { if (typeof logActivity === 'function') logActivity(who, 'Blocked dashboard access: ' + title + ' (' + (source || 'unknown') + ') — ' + reason); } catch(e1) {}
    _logDashboardAccessEvent_('BLOCKED_ATTEMPT', key, who, (source || 'unknown') + ' — ' + reason);
  } catch(e) {}
}

function checkCurrentDashboardAccess(key) {
  try {
    var requestedKey = normalizeDashboardKey_(key || 'teacher');

    if (requestedKey === 'parents') return canAccessDashboard('parents', '');

    var sess = null;
    try { sess = getTeacherSession(); } catch(sessionErr) { sess = null; }
    if (!sess || !sess.username) {
      return _denyDashboardAccess_(requestedKey, 'Session expired. Please log in again.');
    }

    var actualKey = getDashboardKeyFromPosition(sess.position || sess.role || 'teacher');
    if (requestedKey !== actualKey && sess.role !== 'schoolmanager' && sess.role !== 'headofschool') {
      var mismatch = _denyDashboardAccess_(requestedKey, 'This account is not allowed to use the ' + (DASHBOARD_TITLES[requestedKey] || requestedKey) + ' dashboard.');
      logDashboardBlockedAttempt(sess.username, requestedKey, mismatch, 'dashboard-mismatch');
      return mismatch;
    }

    var access = canAccessDashboard(requestedKey, sess.username);
    if (access && access.allowed === false) logDashboardBlockedAttempt(sess.username, requestedKey, access, 'live-dashboard-check');
    return access;
  } catch(e) {
    return _denyDashboardAccess_(key || 'teacher', 'Access check failed: ' + e.message);
  }
}

function _dashboardFeaturesPermit_(access, featureNames) {
  featureNames = Array.isArray(featureNames) ? featureNames : (featureNames ? [featureNames] : []);
  if (!featureNames.length) return true;
  var features = (access && access.features) || {};
  var sawKnownFeature = false;
  var sawAllowedFeature = false;

  featureNames.forEach(function(name){
    if (Object.prototype.hasOwnProperty.call(features, name)) {
      sawKnownFeature = true;
      if (features[name] !== false) sawAllowedFeature = true;
    }
  });

  return !sawKnownFeature || sawAllowedFeature;
}

function checkDashboardWriteAccess(featureNames) {
  try {
    var sess = getTeacherSession();
    if (!sess || !sess.username) return { ok:false, message:'Session expired. Please log in again.' };

    var key = getDashboardKeyFromPosition(sess.position || sess.role || 'teacher');
    var access = canAccessDashboard(key, sess.username);

    if (!access.allowed) {
      return { ok:false, key:key, access:access, message: access.reason || 'This dashboard has been disabled by the administrator.' };
    }
    if (access.readOnly) {
      return { ok:false, key:key, access:access, message:'This dashboard is currently in read-only mode by administrator.' };
    }
    if (!_dashboardFeaturesPermit_(access, featureNames)) {
      return { ok:false, key:key, access:access, message:'This feature has been disabled by administrator.' };
    }
    return { ok:true, key:key, access:access, message:'' };
  } catch(e) {
    return { ok:false, message:'Access verification failed: ' + e.message };
  }
}

function forceLogoutDashboard(key, adminUser) {
  key = normalizeDashboardKey_(key || 'all');
  var all = getDashboardAccessPolicies();
  if (key !== 'all' && DASHBOARD_KEYS.indexOf(key) === -1) throw new Error('Unknown key: ' + key);
  var stamp = _now_();
  var keys = (key === 'all') ? DASHBOARD_KEYS : [key];
  keys.forEach(function(k){ all[k].forceLogoutAt = stamp; all[k].updated = stamp; all[k].updatedBy = adminUser || 'admin'; });
  var props = PropertiesService.getScriptProperties();
  props.setProperty(DASHBOARD_POLICY_PROP, JSON.stringify(all));
  try { props.setProperty('DASHBOARD_ACCESS_POLICIES', JSON.stringify(all)); } catch(e) {}
  _logDashboardAccessEvent_('FORCE_LOGOUT', key, adminUser || 'admin', 'Force logout sent');
  return 'Force logout sent to ' + key + ' dashboard users';
}

function sendDashboardNotice(key, message, adminUser) {
  key = normalizeDashboardKey_(key);
  var policy = getDashboardPolicy(key);
  policy.notice = message || '';
  saveDashboardAccessState(key, policy, adminUser || 'admin');
  _logDashboardAccessEvent_('NOTICE_SENT', key, adminUser || 'admin', message || '');
  return 'Notice sent to ' + key + ' dashboard';
}

function getDashboardAccessEvents(key) {
  key = key ? normalizeDashboardKey_(key) : key;
  var events = [];
  try { events = JSON.parse(PropertiesService.getScriptProperties().getProperty(DASHBOARD_EVENT_PROP) || '[]') || []; } catch(e){ events = []; }
  if (key && key !== 'all') events = events.filter(function(e){ return normalizeDashboardKey_(e.panel) === key || e.panel === 'all'; });
  return events;
}

function _escapeHtml_(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function getAccessDeniedPage(access) {
  access = access || {};
  var title  = _escapeHtml_(access.title  || 'Dashboard');
  var reason = _escapeHtml_(access.reason || 'This dashboard has been disabled by the administrator.');
  var notice = _escapeHtml_(access.notice || '');
  var mode   = access.emergencyLock ? 'Emergency Lock' : (access.temporaryLock ? 'Temporary Lock' : 'Restricted');
  var until  = access.expires ? _escapeHtml_(new Date(access.expires).toLocaleString()) : 'Until Admin allows access';
  var now    = _escapeHtml_(new Date().toLocaleString());
  return '<!DOCTYPE html><html><head><base target="_top">'
    + '<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>Access Restricted</title>'
    + '<style>'
    + ':root{--navy:#07152f;--blue:#2563eb;--cyan:#06b6d4;--red:#ef4444;--amber:#f59e0b;--text:#0f172a;--muted:#64748b;--line:#e2e8f0;}'
    + '*{box-sizing:border-box}body{margin:0;min-height:100vh;font-family:Inter,Segoe UI,Arial,sans-serif;color:var(--text);background:radial-gradient(circle at 15% 10%,rgba(37,99,235,.30),transparent 34%),radial-gradient(circle at 85% 15%,rgba(6,182,212,.24),transparent 32%),linear-gradient(135deg,#eef5ff,#f8fafc 52%,#eef2ff);display:flex;align-items:center;justify-content:center;padding:22px;}'
    + '.wrap{width:min(980px,100%);display:grid;grid-template-columns:1.05fr .95fr;gap:0;background:rgba(255,255,255,.86);border:1px solid rgba(226,232,240,.9);border-radius:30px;overflow:hidden;box-shadow:0 30px 90px rgba(15,23,42,.22);backdrop-filter:blur(18px);}'
    + '.hero{position:relative;padding:42px;background:linear-gradient(145deg,#081a3a,#0f3b67 60%,#0e7490);color:white;overflow:hidden}.hero:before{content:"";position:absolute;width:260px;height:260px;border-radius:50%;right:-80px;top:-80px;background:rgba(255,255,255,.12)}.hero:after{content:"";position:absolute;width:180px;height:180px;border-radius:50%;left:-70px;bottom:-70px;background:rgba(6,182,212,.18)}'
    + '.shield{position:relative;width:86px;height:86px;border-radius:26px;background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.22);display:flex;align-items:center;justify-content:center;font-size:46px;margin-bottom:24px;box-shadow:inset 0 1px 0 rgba(255,255,255,.2)}'
    + '.pill{position:relative;display:inline-flex;gap:8px;align-items:center;background:rgba(239,68,68,.18);border:1px solid rgba(252,165,165,.45);color:#fee2e2;border-radius:99px;padding:7px 12px;font-size:12px;font-weight:900;text-transform:uppercase;letter-spacing:.08em}.hero h1{position:relative;font-size:34px;line-height:1.05;margin:18px 0 12px}.hero p{position:relative;color:rgba(255,255,255,.78);line-height:1.7;margin:0;font-size:14px}.content{padding:42px;background:linear-gradient(180deg,#fff,#f8fafc)}'
    + '.dash{display:inline-flex;align-items:center;gap:8px;background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe;border-radius:99px;padding:7px 13px;font-size:12px;font-weight:900}.reason{margin-top:18px;background:#fff7ed;border:1px solid #fed7aa;border-left:5px solid #f97316;border-radius:16px;padding:16px;color:#9a3412;font-weight:750;line-height:1.55}.notice{margin-top:12px;background:#fffbeb;border:1px solid #fde68a;border-radius:16px;padding:14px;color:#92400e;font-weight:750;line-height:1.55}.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:20px 0}.item{background:#f8fafc;border:1px solid var(--line);border-radius:14px;padding:13px}.item span{display:block;color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.08em;font-weight:900;margin-bottom:5px}.item b{font-size:13px}.actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:18px}.btn{border:0;border-radius:14px;padding:12px 16px;font-weight:900;cursor:pointer;font-size:14px}.primary{background:linear-gradient(135deg,#1d4ed8,#06b6d4);color:white;box-shadow:0 12px 28px rgba(37,99,235,.28)}.ghost{background:#f1f5f9;color:#334155;border:1px solid #e2e8f0}.foot{margin-top:18px;color:#64748b;font-size:12px;line-height:1.6}@media(max-width:820px){.wrap{grid-template-columns:1fr}.hero,.content{padding:28px}.hero h1{font-size:28px}.grid{grid-template-columns:1fr}}'
    + '</style></head><body><main class="wrap">'
    + '<section class="hero"><div class="shield">🛡️</div><div class="pill">● Access Control Active</div><h1>Access Restricted</h1><p>Your request was stopped by the school administration access policy. This protects sensitive school records and keeps the system secure.</p></section>'
    + '<section class="content"><div class="dash">🔒 ' + title + '</div><div class="reason">' + reason + '</div>'
    + (notice ? '<div class="notice">📢 Admin notice: ' + notice + '</div>' : '')
    + '<div class="grid"><div class="item"><span>Status</span><b>' + mode + '</b></div><div class="item"><span>Restricted until</span><b>' + until + '</b></div><div class="item"><span>Checked at</span><b>' + now + '</b></div><div class="item"><span>Action</span><b>Contact Administrator</b></div></div>'
    + '<div class="actions"><button class="btn primary" onclick="google.script.run.withSuccessHandler(function(h){document.open();document.write(h);document.close();}).logout();">Back to Login</button><button class="btn ghost" onclick="location.reload()">Check Again</button></div>'
    + '<div class="foot">If you believe this is a mistake, contact the Academic/Admin office and mention the dashboard name shown above.</div>'
    + '</section></main></body></html>';
}

/* ═════════════════════════════════════════════════════════════════════════
   SECTION 5: PARENT MANAGEMENT (CLASS TEACHER PANEL & PORTAL ACCESS)
   ═════════════════════════════════════════════════════════════════════════ */

function ctpNorm_(v) { return String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' '); }

function ctpMakeUsername_(firstName, phone) {
  var fn = String(firstName || '').trim().split(/\s+/)[0].replace(/[^A-Za-z]/g, '');
  var digits = String(phone || '').replace(/\D/g, '');
  if (!fn || digits.length < 3) return '';
  return fn.charAt(0).toUpperCase() + fn.slice(1).toLowerCase() + digits.slice(-3);
}

function ctpMyClasses_() {
  var t = getTeacherSession();
  if (!t) return null;
  var list = (t.classes || []).map(function (c) { return String(c || '').trim(); }).filter(String);
  try {
    var nm = ctpNorm_(t.fullName), un = ctpNorm_(t.username);
    getAcademicClassesRaw().forEach(function (c) {
      var ct = ctpNorm_(c.teacher);
      if (ct && (ct === nm || ct === un) && list.indexOf(String(c.name).trim()) < 0) list.push(String(c.name).trim());
    });
  } catch (e) {}
  return { teacher: t, classes: list };
}

function ctpClassAllowed_(cls, classes) {
  return classes.some(function (c) { return ctpNorm_(c) === ctpNorm_(cls) || (typeof classLevelMatch === 'function' && classLevelMatch(cls, c)); });
}

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
      password: pw,
      hasPassword: !!pw,
      studentName: String(r[9] || '').trim(), className: cls
    });
  });

  return { ok: true, classes: me.classes, students: students, parents: parents, link: ctpLoginLink_() };
}

function ctpSuggestUsername(firstName, phone) {
  return ctpMakeUsername_(firstName, phone);
}

function ctpUsernameTaken_(username, exceptParentUser) {
  var u = ctpNorm_(username), ex = ctpNorm_(exceptParentUser);
  if (!u) return false;
  var taken = getParentList().some(function (r) {
    var pu = ctpNorm_(r[7]);
    return pu && pu === u && pu !== ex;
  });
  if (taken) return true;
  try { if (getTeacherList().some(function (t) { return ctpNorm_(t.username) === u; })) return true; } catch (e) {}
  try { if (typeof ADMIN_USERS !== 'undefined' && ADMIN_USERS.some(function (a) { return ctpNorm_(a.username) === u; })) return true; } catch (e) {}
  try { if (typeof psAdminByUsername_ === 'function' && psAdminByUsername_(username)) return true; } catch (e) {}
  return false;
}

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

    var username = String(data.username || '').trim();
    if (!username) username = isEdit ? existingUser : ctpMakeUsername_(stu.firstName, phone);
    if (!username) return { ok: false, message: 'Please enter a username.' };
    if (!/^[A-Za-z0-9._-]{3,30}$/.test(username))
      return { ok: false, message: 'Username must be 3–30 characters: letters, numbers, dot, dash or underscore (no spaces).' };
    if (ctpUsernameTaken_(username, isEdit ? existingUser : ''))
      return { ok: false, message: 'Username "' + username + '" is already used by another account. Please choose a different username.' };

    var password = String(data.password == null ? '' : data.password).trim();
    if (!password) {
      if (isEdit && existingPass) password = existingPass;
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
    rng.setNumberFormat('@');
    rng.setValues([row]);

    try { logActivity(me.teacher.username, (editRow > 0 ? 'Updated' : 'Added') + ' parent ' + name + ' (' + username + ')'); } catch (e) {}
    return { ok: true, username: username, password: password,
             message: 'Saved. Parent login → Username: ' + username + '  |  Password: ' + password };
  } finally {
    lock.releaseLock();
  }
}

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

function ctpLoginLink_() {
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  return url;
}

function ctpGetLoginLink() {
  if (!ctpMyClasses_()) return { ok: false, message: 'Session expired. Please login again.' };
  var link = ctpLoginLink_();
  return link ? { ok: true, link: link } : { ok: false, message: 'Deploy the Web App first.' };
}

function parentIdentity_(token) {
  token = String(token || '');
  if (!/^[a-f0-9]{64}$/.test(token)) throw new Error('Please sign in with your parent username.');
  var props = PropertiesService.getScriptProperties(), key = SMS_SESSION_PREFIX + token;
  var raw = props.getProperty(key), session = raw ? JSON.parse(raw) : null;
  if (!session || session.role !== 'parent' || !session.exp || session.exp < Date.now())
    throw new Error('Please sign in again to load your child\'s records.');
  var pd = smsParentData_(session.username);
  if (!pd || !String(pd.admNo || '').trim()) throw new Error('Your student account is not linked. Contact your class teacher.');
  var access = canAccessDashboard('parents', session.username);
  if (access && access.allowed === false) throw new Error(access.reason || 'Parent portal is disabled.');
  session.exp = Date.now() + SMS_SESSION_HOURS * 3600 * 1000;
  props.setProperty(key, JSON.stringify(session));
  return pd;
}

function parentGetReportCard(token) {
  try {
    var pd = parentIdentity_(token), adm = String(pd.admNo).trim(), exams = [], examCount = 0, withScores = 0;
    (fetchExamTypes(pd.className) || []).forEach(function(examName){
      var all = getExamResults(pd.className, examName) || [];
      examCount++; if (all.length) withScores++;
      var matches = all.filter(function(s){ return String(s.admNo || '').trim() === adm; });
      if (!matches.length) {
        var nm = ctpNorm_(pd.studentName), same = (getStudentsForClass(pd.className) || []).filter(function(st){
          return ctpNorm_(String(st.firstName || '') + ' ' + String(st.lastName || '')) === nm;
        });
        if (nm && same.length === 1) matches = all.filter(function(s){ return !String(s.admNo || '').trim() && ctpNorm_(s.name) === nm; });
      }
      if (matches.length !== 1) return;
      var source = matches[0], report = {};
      ['name','admNo','sex','subjects','total','average','grade','point','division','position','outOf','remarks'].forEach(function(k){ report[k] = source[k]; });
      report.admNo = adm; report.outOf = all.length; report.tabia = []; report.classTeacher = '';
      var b = fetchReportBehaviors(pd.className, examName) || {};
      report.classTeacher = b.classTeacher || '';
      var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BEHAVIOR_SHEET), n = (typeof BEHAVIOR_CRITERIA !== 'undefined' ? BEHAVIOR_CRITERIA.length : 6);
      if (sh && sh.getLastRow() > 1) {
        sh.getRange(2, 1, sh.getLastRow() - 1, 4 + n + 4).getValues().forEach(function(r){
          if (typeof behClassKey_ === 'function' && behClassKey_(r[0]) === behClassKey_(pd.className) && behNorm_(r[1]) === behNorm_(examName) && String(r[2] || '').trim() === adm && String(r[4 + n + 2] || '').trim() === 'Submitted')
            report.tabia = r.slice(4, 4 + n).map(function(v){ return String(v || '').trim(); });
        });
      }
      exams.push({ examName: examName, students: [report] });
    });
    var settings = {}, a = (typeof getAcademicSettings === 'function' ? getAcademicSettings() : {}) || {};
    ['schoolName','logoLeft','logoRight','academicName','academicSign','headName','headSign'].forEach(function(k){ settings[k] = a[k] || ''; });
    var hint = '';
    if (!exams.length) {
      if (!examCount) hint = 'No examinations have been set up for your child\'s class yet.';
      else if (!withScores) hint = 'Scores for your child\'s class have not been entered yet.';
      else hint = 'Scores have been entered, but none could be linked to the admission number on your account. Please ask the class teacher to check it.';
    }
    return { ok: true, exams: exams, settings: settings, hint: hint };
  } catch(e) {
    return { ok: false, message: e.message || 'Could not load report card.' };
  }
}

function parentAttendanceMark_(v) {
  var s = String(v == null ? '' : v).trim().toUpperCase();
  return ['P','A','L'].indexOf(s) >= 0 ? s : '';
}

function parentGetAttendance(token) {
  try {
    var pd = parentIdentity_(token), adm = String(pd.admNo).trim();
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('Attendance');
    if (!sh || sh.getLastRow() < 2) return { ok: true, records: [] };
    var rows = sh.getDataRange().getValues(), headers = rows.shift();
    function col(names){
      for (var i = 0; i < headers.length; i++) {
        var h = String(headers[i]).toLowerCase().replace(/[^a-z0-9]/g, '');
        if (names.indexOf(h) >= 0) return i;
      }
      return -1;
    }
    var ad = col(['admno','admissionnumber']), dt = col(['date','datestr']), mo = col(['morning']), af = col(['afternoon']), cl = col(['class','classname']);
    if (ad < 0 || dt < 0 || mo < 0 || af < 0) throw new Error('Attendance columns are missing. Contact the class teacher.');
    var days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'], byDate = {};
    rows.forEach(function(r){
      if (String(r[ad] || '').trim() !== adm || !r[dt]) return;
      if (cl >= 0 && pd.className && String(r[cl] || '').trim() && ctpNorm_(r[cl]).replace(/\s+/g,'') !== ctpNorm_(pd.className).replace(/\s+/g,'')) return;
      var date = r[dt] instanceof Date ? Utilities.formatDate(r[dt], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(r[dt]).trim().slice(0, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
      var dayDate = new Date(date + 'T12:00:00Z'); if (isNaN(dayDate.getTime())) return;
      var morning = parentAttendanceMark_(r[mo]), afternoon = parentAttendanceMark_(r[af]);
      byDate[date] = { date: date, day: days[dayDate.getUTCDay()], morning: morning, afternoon: afternoon, status: ({P:'Present', A:'Absent', L:'Late'})[morning] || 'Not marked' };
    });
    return { ok: true, records: Object.keys(byDate).sort().map(function(d){ return byDate[d]; }) };
  } catch(e) {
    return { ok: false, message: e.message || 'Could not load attendance.' };
  }
}