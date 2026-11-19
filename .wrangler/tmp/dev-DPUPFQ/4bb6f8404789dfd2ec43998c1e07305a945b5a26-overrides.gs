// ============================================================
// HOST OVERRIDES — loaded after every original file, so these win.
// The session itself (token → account) is verified by the Node host against the
// `app_sessions` table; the legacy code only needs to know WHO is calling.
// ============================================================

// School Manager / Head of School / Accountant calls (psCall, apiRequest)
function psSessionFromToken_(token) {
  var s = __HOST__.session;
  if (!s || ['schoolmanager', 'headofschool', 'accountant'].indexOf(s.role) < 0) return null;
  var acc = psAdminByUsername_(s.username);
  if (!acc || acc.Status === 'Inactive' || psRoleKey_(acc.Role) !== s.role) return null;
  return { role: s.role, username: acc.Username, name: acc.FullName || acc.Username, adminId: acc.AdminID };
}

// Parent portal reads
function parentIdentity_(token) {
  var s = __HOST__.session;
  if (!s || s.role !== 'parent') throw new Error('Please sign in with your parent username.');
  var pd = smsParentData_(s.username);
  if (!pd || !String(pd.admNo || '').trim()) throw new Error('Your student account is not linked. Contact your class teacher.');
  var access = canAccessDashboard('parents', s.username);
  if (access && access.allowed === false) throw new Error(access.reason || 'Parent portal is disabled.');
  return pd;
}

// Sessions live in the database and are ended by the host.
function smsEndSession(token) { return true; }

// Parent passwords are stored only as hashes now. Never send the placeholder to the browser.
ctpGetParentsPanel = (function (original) {
  return function () {
    var res = original.apply(null, arguments);
    if (res && res.parents) {
      res.parents.forEach(function (p) {
        if (p.password === '__keep__') { p.password = ''; p.hasPassword = true; }
      });
    }
    return res;
  };
})(ctpGetParentsPanel);

// ───────────── login / restore payloads (password already verified by the host) ─────────────
function nativeBlockedPayload_(name, username, position, dashboardKey, access) {
  var acc = access || { title: dashboardKey, reason: 'This dashboard has been blocked by the administrator.' };
  return {
    status: 'fail', blocked: true, role: 'blocked', fullName: name, username: username, position: position,
    dashboardKey: dashboardKey, message: (acc.reason || 'This dashboard has been blocked by the administrator.'),
    html: getAccessDeniedPage(acc)
  };
}

function nativeFindStaff_(username) {
  var u = String(username || '').trim().toLowerCase(), list = getTeacherList();
  for (var i = 0; i < list.length; i++) if (String(list[i].username || '').trim().toLowerCase() === u) return list[i];
  return null;
}

function nativeLoginPayload_(username) {
  username = String(username || '').trim();

  var adm = psAdminByUsername_(username);
  if (adm) {
    var role = PS_ROLE_MAP[adm.Role];
    if (!role) return { status: 'fail', message: 'Account role is not valid.' };
    try { psTable_('Administrators').update(adm._row, { LastLogin: psNow_() }); } catch (e) {}
    try { logActivity(adm.Username, 'Login (' + adm.Role + ')'); } catch (e) {}
    return {
      status: 'success', role: role.key, position: adm.Role, fullName: adm.FullName, username: adm.Username,
      html: psBootInject_(role.key, { role: role.key, name: adm.FullName || adm.Username, username: adm.Username },
                          HtmlService.createHtmlOutputFromFile(role.file).getContent())
    };
  }

  var t = nativeFindStaff_(username);
  if (t) {
    var pos = String(t.position || '').trim();
    var d = smsDashboardFor_(pos);
    var access = canAccessDashboard(d.key, t.username);
    if (!access || access.allowed === false) {
      try { logDashboardBlockedAttempt(t.username, d.key, access, 'login'); } catch (e) {}
      return nativeBlockedPayload_(t.name, t.username, pos, d.key, access);
    }
    try { logActivity(t.username, 'Login: ' + t.username + ' (' + pos + ')'); } catch (e) {}
    return {
      status: 'success', role: d.role, position: pos, fullName: t.name, username: t.username,
      dashboardKey: d.key, access: access,
      classesSubjects: t.classesSubjects, classes: t.classes, homeroomClasses: t.homeroomClasses, subjects: t.subjects,
      html: HtmlService.createHtmlOutputFromFile(d.file).getContent()
    };
  }

  var pd = smsParentData_(username);
  if (pd) {
    var pAccess = canAccessDashboard('parents', username);
    if (!pAccess || pAccess.allowed === false) {
      try { logDashboardBlockedAttempt(username, 'parents', pAccess, 'login'); } catch (e) {}
      return nativeBlockedPayload_(pd.parentName, username, '', 'parents', pAccess);
    }
    return { status: 'success', role: 'parent', fullName: pd.parentName || username, username: username, parentData: pd };
  }
  return { status: 'fail', message: 'Invalid username or password.' };
}

function nativeRestorePayload_(username, role) {
  if (['schoolmanager', 'headofschool', 'accountant'].indexOf(role) >= 0) {
    var acc = psAdminByUsername_(username);
    if (!acc || acc.Status === 'Inactive' || psRoleKey_(acc.Role) !== role) return { ok: false };
    return { ok: true, role: role, html: psBootInject_(role, { role: role, name: acc.FullName || acc.Username, username: acc.Username },
                                                       HtmlService.createHtmlOutputFromFile(psFileForRole_(role)).getContent()) };
  }
  if (role === 'parent') {
    var pd = smsParentData_(username);
    if (!pd) return { ok: false };
    var pa = canAccessDashboard('parents', username);
    if (pa && pa.allowed === false) return { ok: false };
    return { ok: true, role: role, parentData: pd, html: smsParentHtml_(pd) };
  }
  var t = nativeFindStaff_(username);
  if (!t) return { ok: false };
  var d = smsDashboardFor_(t.position);
  var access = canAccessDashboard(d.key, t.username);
  if (!access || access.allowed === false) return { ok: false };
  return { ok: true, role: d.role, html: HtmlService.createHtmlOutputFromFile(d.file).getContent() };
}
