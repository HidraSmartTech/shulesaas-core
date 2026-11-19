// ============================================================
// PRIMARY SCHOOL SYSTEM — SERVER API
// School Manager  →  creates Head of School + Accountant
// Head of School  →  creates all teachers / staff accounts
// Accountant      →  fees, payments, income, expenditure, cash & bank,
//                    procurement & stock, reports, audit
//
// Every call from the new dashboards goes through psCall(token, fn, args):
//   • the token must belong to a logged-in, active account
//   • the account's role must be allowed to run that function
// ============================================================

// ───────────────────────── small helpers ─────────────────────────
function psNum_(v) { const n = Number(String(v === undefined || v === null ? '' : v).replace(/,/g, '')); return isNaN(n) ? 0 : n; }
function psRound_(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }
function psFmtDate_(d) {
  const h = d.getHours() || d.getMinutes() || d.getSeconds();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), h ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
}
function psPlain_(v) { return v instanceof Date ? psFmtDate_(v) : (v === null || v === undefined ? '' : v); }
function psToday_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function psNow_()   { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }
function psId_(prefix) { return prefix + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase().padStart(2, '0'); }
function psNormClass_(c) { return String(c || '').toLowerCase().replace(/\s+/g, ''); }
function psNormName_(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
function psErr_(msg) { throw new Error(msg); }

// Generic table access (rows come back as objects, dates as text)
// SPEED: sheet headers and values are remembered for the duration of ONE server call,
// so a page that reads the same sheet 5 times only hits the spreadsheet once.
// Every write below (append/update/remove) forgets the remembered values of that sheet.
let PS_MEMO_ = {};
function psMemoReset_() { PS_MEMO_ = {}; }
function psTable_(name) {
  let m = PS_MEMO_[name];
  if (!m) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(name);
    if (!sh) { psRunSetup_(); sh = ss.getSheetByName(name); }
    if (!sh) psErr_('Table "' + name + '" is missing.');
    const lastCol = Math.max(1, sh.getLastColumn());
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
    m = PS_MEMO_[name] = { sh: sh, headers: headers, values: null };
  }
  const sh = m.sh, headers = m.headers;
  function forget() { m.values = null; }
  return {
    sh: sh, headers: headers,
    rows: function () {
      if (m.values === null) {
        const n = sh.getLastRow() - 1;
        m.values = n < 1 ? [] : sh.getRange(2, 1, n, headers.length).getValues();
      }
      return m.values.map((r, i) => {
        const o = { _row: i + 2 };
        headers.forEach((h, j) => { if (h) o[h] = psPlain_(r[j]); });
        return o;
      });
    },
    append: function (obj) { forget(); sh.appendRow(headers.map(h => obj[h] === undefined ? '' : obj[h])); },
    appendMany: function (list) {
      if (!list.length) return;
      forget();
      const vals = list.map(o => headers.map(h => o[h] === undefined ? '' : o[h]));
      sh.getRange(sh.getLastRow() + 1, 1, vals.length, headers.length).setValues(vals);
    },
    update: function (rowNum, obj) {
      forget();
      const range = sh.getRange(rowNum, 1, 1, headers.length);
      const cur = range.getValues()[0];
      headers.forEach((h, j) => { if (obj[h] !== undefined) cur[j] = obj[h]; });
      range.setValues([cur]);
    },
    remove: function (rowNum) { forget(); sh.deleteRow(rowNum); }
  };
}

// numbers with a prefix (RCP-2026-00001). Locks itself unless psCall already holds the lock.
let PS_LOCK_HELD_ = false;
function psNextNumbers_(counterKey, prefixKey, count) {
  if (count <= 0) return [];
  const lock = LockService.getScriptLock();
  if (!PS_LOCK_HELD_) lock.waitLock(20000);
  try {
    const t = psTable_('FinanceSettings');
    const rows = t.rows();
    const c = rows.find(r => String(r.Key).trim() === counterKey);
    const p = rows.find(r => String(r.Key).trim() === prefixKey);
    if (!c) psErr_('Counter ' + counterKey + ' not found in FinanceSettings');
    const prefix = p ? (String(p.Value).trim() || 'NO') : 'NO';
    const start = Math.max(1, psNum_(c.Value));
    t.update(c._row, { Value: start + count });
    const out = [], y = new Date().getFullYear();
    for (let i = 0; i < count; i++) out.push(prefix + '-' + y + '-' + String(start + i).padStart(5, '0'));
    return out;
  } finally { if (!PS_LOCK_HELD_) lock.releaseLock(); }
}

function psHash_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + ':' + String(password), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function psSalt_() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); }

// ───────────────────────── accounts & login ─────────────────────────
const PS_ROLE_MAP = {
  'School Manager': { key: 'schoolmanager', file: 'SchoolManager' },
  'Head of School': { key: 'headofschool',  file: 'HeadOfSchool' },
  'Accountant':     { key: 'accountant',    file: 'Accountant' }
};
function psRoleKey_(label) { return (PS_ROLE_MAP[label] || {}).key || ''; }
function psFileForRole_(key) {
  const k = Object.keys(PS_ROLE_MAP).find(l => PS_ROLE_MAP[l].key === key);
  return k ? PS_ROLE_MAP[k].file : '';
}
function psAdminByUsername_(username) {
  const u = String(username || '').trim().toLowerCase();
  return psTable_('Administrators').rows().find(r => String(r.Username).trim().toLowerCase() === u) || null;
}

// Is this username already used by any account in the system?
function psUsernameTaken_(username, exceptAdminId, exceptUsersRow) {
  const u = String(username || '').trim().toLowerCase();
  if (!u) return false;
  if (ADMIN_USERS.some(a => String(a.username).toLowerCase() === u)) return true;
  if (psTable_('Administrators').rows().some(a => String(a.Username).trim().toLowerCase() === u && a.AdminID !== exceptAdminId)) return true;
  if (psTable_('Users').rows().some(a => String(a.Username).trim().toLowerCase() === u && a._row !== exceptUsersRow)) return true;
  if (psTable_('Parents').rows().some(a => String(a.Username).trim().toLowerCase() === u)) return true;
  return false;
}

// ───────────────────────── first-screen data, delivered WITH the page ─────────────────────────
// SPEED: after login the dashboard used to start empty and then make two sequential server calls
// (meta → first page) before showing anything. The data of the first screen is now computed here,
// while the login request is already running, and embedded in the HTML as window.__PS_BOOT.
// It never breaks login: any error just skips the embedding and the page loads data the normal way.
function psBootCached_(fn, role, args, ttl, compute) {
  const ck = psCacheKey_(fn, role, args);
  const hit = psCacheGet_(ck);
  if (hit !== undefined) return hit;
  const data = compute();
  psCachePut_(ck, data === undefined ? null : data, ttl);
  return data;
}
function psBootInject_(roleKey, sess, html) {
  try {
    psMemoReset_();
    const boot = { role: roleKey };
    if (roleKey === 'accountant') {
      AB_MEMO_ = {};
      boot.pages = { dashboard: psBootCached_('ab.dashboard', 'accountant', [{ action: 'dashboard', page: 'dashboard' }], AB_CACHE_TTL_,
        function () { return ab_page_(sess, { page: 'dashboard' }); }) };
    } else {
      boot.meta = api_meta_(sess);
      boot.overview = psBootCached_('overview', roleKey, [], PS_CACHE_TTL_['overview'] || 30, function () { return api_overview_(sess); });
    }
    const json = JSON.stringify(boot).replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
    const tag = '<script>window.__PS_BOOT=' + json + ';</' + 'script>';
    html = String(html || '');
    return /<head[^>]*>/i.test(html) ? html.replace(/<head[^>]*>/i, function (m) { return m + tag; }) : tag + html;
  } catch (e) {
    Logger.log('psBootInject_: ' + e.message);
    return html;
  }
}

// Called by checkLogin() first. Returns null when the username is not an administrator account.
function psTryAdminLogin_(username, password) {
  username = String(username || '').trim();
  password = String(password || '');
  const t = psTable_('Administrators');
  const rows = t.rows();
  const hasManager = rows.some(r => r.Role === 'School Manager');
  const row = rows.find(r => String(r.Username).trim().toLowerCase() === username.toLowerCase());

  let acc = null;
  if (row) {
    if (psHash_(password, row.Salt) !== row.PasswordHash) return { status: 'fail', message: 'Invalid username or password.' };
    if (row.Status === 'Inactive') return { status: 'fail', message: 'This account is inactive. Contact the School Manager.' };
    acc = row;
  } else if (!hasManager && ADMIN_USERS.some(a => a.username === username && a.password === password)) {
    // first ever login of the School Manager → keep the password only as a hash from now on
    const salt = psSalt_();
    const rec = { AdminID: psId_('ADM'), FullName: 'School Manager', Username: username, PasswordHash: psHash_(password, salt), Salt: salt,
                  Role: 'School Manager', Sex: '', Phone: '', Email: '', Status: 'Active', PicUrl: '', CreatedBy: 'system', CreatedAt: psNow_(), LastLogin: '' };
    t.append(rec);
    acc = psAdminByUsername_(username);
  }
  if (!acc) return null;

  const role = PS_ROLE_MAP[acc.Role];
  if (!role) return { status: 'fail', message: 'Account role is not valid.' };
  try { clearTeacherSession(); } catch (e) {}
  try { psTable_('Administrators').update(acc._row, { LastLogin: psNow_() }); } catch (e) {}
  try { logActivity(acc.Username, 'Login (' + acc.Role + ')'); } catch (e) {}
  return {
    status: 'success', role: role.key, position: acc.Role, fullName: acc.FullName, username: acc.Username,
    html: psBootInject_(role.key, { role: role.key, name: acc.FullName || acc.Username, username: acc.Username },
                        HtmlService.createHtmlOutputFromFile(role.file).getContent())
  };
}

// Staff whose account was set to Inactive by the Head of School cannot log in
function psInactiveStaffBlock_(username, password) {
  const u = String(username || '').trim().toLowerCase();
  const r = psTable_('Users').rows().find(x => String(x.Username).trim().toLowerCase() === u && String(x.Password) === String(password));
  if (r && String(r.Status || '').toLowerCase() === 'inactive') return { status: 'fail', message: 'This account is inactive. Contact the Head of School.' };
  return null;
}

// Editor-only recovery: run once if the School Manager forgets the password.
// After running, the default ADMIN_USERS password works again at the next login.
function psResetSchoolManager() {
  const t = psTable_('Administrators');
  t.rows().filter(r => r.Role === 'School Manager').reverse().forEach(r => t.remove(r._row));
  Logger.log('School Manager account removed. Log in with the ADMIN_USERS username/password from Code.gs.');
}

// session from the persistent login token (see smsNewToken_)
function psSessionFromToken_(token) {
  token = String(token || '');
  if (!/^[a-f0-9]{64}$/.test(token)) return null;
  const cache = CacheService.getScriptCache(), ck = 'PSS_' + token.slice(0, 40);
  const hit = cache.get(ck);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const props = PropertiesService.getScriptProperties();
  const key = SMS_SESSION_PREFIX + token;
  const raw = props.getProperty(key);
  if (!raw) return null;
  let s; try { s = JSON.parse(raw); } catch (e) { return null; }
  if (!s.exp || s.exp < Date.now()) return null;
  const acc = psAdminByUsername_(s.username);
  if (!acc || acc.Status === 'Inactive' || psRoleKey_(acc.Role) !== s.role) return null;
  s.exp = Date.now() + SMS_SESSION_HOURS * 3600 * 1000;
  props.setProperty(key, JSON.stringify(s));
  const out = { role: s.role, username: acc.Username, name: acc.FullName || acc.Username, adminId: acc.AdminID };
  cache.put(ck, JSON.stringify(out), 120);
  return out;
}

// ───────────────────────── server-side read cache ─────────────────────────
// Read-only lists are kept for a short time in CacheService so repeated opens
// (by any user) skip the slow spreadsheet reads. Any write made through psCall
// bumps the version, which invalidates everything at once.
const PS_CACHE_TTL_ = {
  'acad.classes': 60, 'acad.subjects': 60, 'acad.exams': 60, 'acad.examTypes': 60,
  'acad.deadline': 30, 'acad.settings': 120, 'acad.students': 30, 'acad.results': 30,
  'acad.fullResults': 30, 'acad.attendance': 30, 'overview': 30, 'terms.list': 120,
  'staff.list': 60
};
const PS_CACHE_CHUNK_ = 30000;
function psCacheVer_() { try { return CacheService.getScriptCache().get('PS_VER') || '0'; } catch (e) { return '0'; } }
function psCacheBump_() { try { CacheService.getScriptCache().put('PS_VER', String(Date.now()), 21600); } catch (e) {} }
function psCacheGet_(key) {
  try {
    const c = CacheService.getScriptCache(), head = c.get(key);
    if (head === null || head === undefined) return undefined;
    if (head.charAt(0) !== '#') return JSON.parse(head);
    const n = Number(head.slice(1)), keys = [];
    for (let i = 0; i < n; i++) keys.push(key + '~' + i);
    const parts = c.getAll(keys), out = [];
    for (let i = 0; i < n; i++) { const p = parts[key + '~' + i]; if (p === null || p === undefined) return undefined; out.push(p); }
    return JSON.parse(out.join(''));
  } catch (e) { return undefined; }
}
function psCachePut_(key, val, ttl) {
  try {
    const c = CacheService.getScriptCache(), str = JSON.stringify(val === undefined ? null : val);
    if (str.length <= PS_CACHE_CHUNK_) { c.put(key, str, ttl); return; }
    const n = Math.ceil(str.length / PS_CACHE_CHUNK_);
    if (n > 40) return;                       // too big to cache — just skip
    const all = {};
    for (let i = 0; i < n; i++) all[key + '~' + i] = str.substr(i * PS_CACHE_CHUNK_, PS_CACHE_CHUNK_);
    all[key] = '#' + n;
    c.putAll(all, ttl);
  } catch (e) {}
}
function psCacheKey_(fn, role, args) {
  const h = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(args || []))).slice(0, 22);
  return 'PSC_' + psCacheVer_() + '_' + fn + '_' + role + '_' + h;
}

// ───────────────────────── dispatcher ─────────────────────────
function psCall(token, fn, args) {
  psMemoReset_();
  let locked = false, wrote = false;
  const lock = LockService.getScriptLock();
  try {
    const s = psSessionFromToken_(token);
    if (!s) return { ok: false, code: 'SESSION', error: 'Your session has expired. Please log in again.' };
    const def = psApi_()[fn];
    if (!def) return { ok: false, error: 'Unknown action: ' + fn };
    if (def.roles.indexOf(s.role) < 0) return { ok: false, error: 'You are not allowed to do this.' };
    const ttl = def.write ? 0 : (PS_CACHE_TTL_[fn] || 0);
    let ck = '';
    if (ttl) {
      ck = psCacheKey_(fn, s.role, args);
      const hit = psCacheGet_(ck);
      if (hit !== undefined) return { ok: true, data: hit };
    }
    if (def.write) { lock.waitLock(25000); locked = true; wrote = true; PS_LOCK_HELD_ = true; }
    const data = def.fn.apply(null, [s].concat(args || []));
    if (ttl) psCachePut_(ck, data === undefined ? null : data, ttl);
    return { ok: true, data: data === undefined ? null : data };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  } finally {
    PS_LOCK_HELD_ = false;
    if (locked) lock.releaseLock();
    if (wrote) psCacheBump_();
  }
}

const PS_R_ALL = ['schoolmanager', 'headofschool', 'accountant'];
const PS_R_SM = ['schoolmanager'];
const PS_R_HEAD = ['headofschool'];
const PS_R_FIN = ['accountant'];
const PS_R_FINREAD = ['schoolmanager', 'accountant'];
const PS_R_ACAD = ['schoolmanager', 'headofschool'];

function psApi_() {
  return {
    'meta':               { roles: PS_R_ALL, fn: api_meta_ },
    'overview':           { roles: PS_R_ALL, fn: api_overview_ },
    'me.changePassword':  { roles: PS_R_ALL, fn: api_changeMyPassword_, write: true },
    'manager.updateAccount': { roles: PS_R_SM, fn: api_managerUpdateAccount_, write: true },
    'manager.academicYear.set': { roles: PS_R_SM, fn: api_managerSetAcademicYear_, write: true },
    'manager.backups.list': { roles: PS_R_SM, fn: api_managerBackupsList_ },
    'manager.backups.create': { roles: PS_R_SM, fn: api_managerBackupCreate_, write: true },
    'manager.backups.restore': { roles: PS_R_SM, fn: api_managerBackupRestore_, write: true },
    'school.save':        { roles: PS_R_SM, fn: api_schoolSave_, write: true },
    'terms.list':         { roles: PS_R_ALL, fn: api_termsList_ },
    'terms.save':         { roles: PS_R_SM, fn: api_termsSave_, write: true },
    // academics (read only)
    'acad.classes':       { roles: PS_R_ALL, fn: api_classes_ },
    'acad.students':      { roles: PS_R_ALL, fn: api_students_ },
    'acad.subjects':      { roles: PS_R_ACAD, fn: api_subjects_ },
    'acad.exams':         { roles: PS_R_ACAD, fn: api_exams_ },
    'acad.results':       { roles: PS_R_ACAD, fn: api_results_ },
    'acad.examTypes':     { roles: PS_R_ACAD, fn: api_exam_types_ },
    'acad.deadline':      { roles: PS_R_ACAD, fn: api_exam_deadline_ },
    'acad.fullResults':   { roles: PS_R_ACAD, fn: api_full_results_ },
    'acad.attendance':    { roles: PS_R_ACAD, fn: api_attendance_ },
    'acad.settings':      { roles: PS_R_ACAD, fn: api_acad_settings_ },
    // School Manager → administrators
    'admins.list':        { roles: PS_R_SM, fn: api_adminsList_ },
    'admins.save':        { roles: PS_R_SM, fn: api_adminsSave_, write: true },
    'admins.delete':      { roles: PS_R_SM, fn: api_adminsDelete_, write: true },
    'admins.resetPassword': { roles: PS_R_SM, fn: api_adminsReset_, write: true },
    // Head of School → staff accounts
    'staff.list':         { roles: PS_R_HEAD.concat(PS_R_SM), fn: api_staffList_ },
    // Teacher On Duty tracking (Head of School) — see TODReportAPI.gs
    'tod.teachers':       { roles: PS_R_HEAD, fn: api_todTeachers_ },
    'tod.roster':         { roles: PS_R_HEAD, fn: api_todRosterList_ },
    'tod.rosterSave':     { roles: PS_R_HEAD, fn: api_todRosterSave_, write: true },
    'tod.rosterDelete':   { roles: PS_R_HEAD, fn: api_todRosterDelete_, write: true },
    'tod.reports':        { roles: PS_R_HEAD, fn: api_todReports_ },
    'tod.report':         { roles: PS_R_HEAD, fn: api_todReport_ },
    'staff.save':         { roles: PS_R_HEAD, fn: api_staffSave_, write: true },
    'staff.delete':       { roles: PS_R_HEAD, fn: api_staffDelete_, write: true },
    'staff.resetPassword': { roles: PS_R_HEAD, fn: api_staffReset_, write: true },
    // finance — read
    'fin.payments':       { roles: PS_R_FINREAD, fn: fin_payments_ },
    'fin.debtors':        { roles: PS_R_FINREAD, fn: fin_debtors_ },
    'fin.expenses':       { roles: PS_R_FINREAD, fn: fin_expenses_ },
    'fin.categories.list': { roles: PS_R_FINREAD, fn: fin_catList_ },
    'fin.categories.save': { roles: PS_R_FINREAD, fn: fin_catSave_, write: true },
    'fin.categories.delete': { roles: PS_R_FINREAD, fn: fin_catDelete_, write: true },
    // finance — accountant
    'fin.fees.list':      { roles: PS_R_FINREAD, fn: fin_feesList_ },
    'fin.fees.save':      { roles: PS_R_FIN, fn: fin_feesSave_, write: true },
    'fin.fees.delete':    { roles: PS_R_FIN, fn: fin_feesDelete_, write: true },
    'fin.invoices.list':  { roles: PS_R_FINREAD, fn: fin_invList_ },
    'fin.invoices.generate': { roles: PS_R_FIN, fn: fin_invGenerate_, write: true },
    'fin.invoices.addLine': { roles: PS_R_FIN, fn: fin_invAddLine_, write: true },
    'fin.invoices.setDiscount': { roles: PS_R_FIN, fn: fin_invDiscount_, write: true },
    'fin.invoices.cancel': { roles: PS_R_FIN, fn: fin_invCancel_, write: true },
    'fin.studentSearch':  { roles: PS_R_FIN, fn: fin_studentSearch_ },
    'fin.studentAccount': { roles: PS_R_FIN, fn: fin_studentAccount_ },
    'fin.pay':            { roles: PS_R_FIN, fn: fin_pay_, write: true },
    'fin.payment.void':   { roles: PS_R_FIN, fn: fin_voidPayment_, write: true },
    'fin.receipt':        { roles: PS_R_FINREAD, fn: fin_receipt_ },
    'fin.income.list':    { roles: PS_R_FINREAD, fn: fin_incomeList_ },
    'fin.income.save':    { roles: PS_R_FIN, fn: fin_incomeSave_, write: true },
    'fin.income.void':    { roles: PS_R_FIN, fn: fin_incomeVoid_, write: true },
    'fin.expense.save':   { roles: PS_R_FIN, fn: fin_expenseSave_, write: true },
    'fin.expense.pay':    { roles: PS_R_FIN, fn: fin_expensePay_, write: true },
    'fin.expense.reverse': { roles: PS_R_FIN, fn: fin_expenseReverse_, write: true },
    'fin.expense.delete': { roles: PS_R_FIN, fn: fin_expenseDelete_, write: true },
    'fin.accounts.list':  { roles: PS_R_FINREAD, fn: fin_accountsList_ },
    'fin.accounts.save':  { roles: PS_R_FIN, fn: fin_accountsSave_, write: true },
    'fin.txns':           { roles: PS_R_FINREAD, fn: fin_txns_ },
    'fin.transfer':       { roles: PS_R_FIN, fn: fin_transfer_, write: true },
    'fin.recon.list':     { roles: PS_R_FINREAD, fn: fin_reconList_ },
    'fin.recon.save':     { roles: PS_R_FIN, fn: fin_reconSave_, write: true },
    'fin.followups.list': { roles: PS_R_FIN, fn: fin_fuList_ },
    'fin.followups.save': { roles: PS_R_FIN, fn: fin_fuSave_, write: true },
    'fin.followups.delete': { roles: PS_R_FIN, fn: fin_fuDelete_, write: true },
    'fin.dashboard':      { roles: PS_R_FINREAD, fn: fin_dashboard_ },
    'fin.report':         { roles: PS_R_FINREAD, fn: fin_report_ },
    'fin.suppliers.list': { roles: PS_R_FIN, fn: fin_supList_ },
    'fin.suppliers.save': { roles: PS_R_FIN, fn: fin_supSave_, write: true },
    'fin.suppliers.delete': { roles: PS_R_FIN, fn: fin_supDelete_, write: true },
    'fin.requests.list':  { roles: PS_R_FIN, fn: fin_reqList_ },
    'fin.requests.save':  { roles: PS_R_FIN, fn: fin_reqSave_, write: true },
    'fin.requests.status': { roles: PS_R_FIN, fn: fin_reqStatus_, write: true },
    'fin.requests.receive': { roles: PS_R_FIN, fn: fin_reqReceive_, write: true },
    'fin.stock.items':    { roles: PS_R_FIN, fn: fin_stockItems_ },
    'fin.stock.saveItem': { roles: PS_R_FIN, fn: fin_stockSaveItem_, write: true },
    'fin.stock.move':     { roles: PS_R_FIN, fn: fin_stockMove_, write: true },
    'fin.stock.movements': { roles: PS_R_FIN, fn: fin_stockMovements_ },
    'fin.audit.list':     { roles: PS_R_FINREAD, fn: fin_auditList_ },
    'fin.audit.checks':   { roles: PS_R_FINREAD, fn: fin_auditChecks_ },
    'fin.settings.list':  { roles: PS_R_FINREAD, fn: fin_settingsList_ },
    'fin.settings.save':  { roles: PS_R_FIN, fn: fin_settingsSave_, write: true }
  };
}

function psAuditS_(s, module, action, recordId, oldV, newV, details) {
  psAudit_(s.username, s.role, module, action, recordId, oldV, newV, details);
}

// ───────────────────────── shared data readers ─────────────────────────
function psAllStudents_(onlyClass) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), out = [];
  PS_CLASSES.forEach(c => {
    if (onlyClass && psNormClass_(onlyClass) !== psNormClass_(c)) return;
    const sh = ss.getSheetByName(psClassSheetName_(c));
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().forEach(r => {
      if (!String(r[0]).trim() && !String(r[1]).trim()) return;
      out.push({ AdmNo: String(r[0]).trim(), FirstName: String(r[1]).trim(), LastName: String(r[2]).trim(),
                 Name: (String(r[1]) + ' ' + String(r[2])).trim(), Sex: String(r[3]).trim(), DOB: psPlain_(r[4]),
                 Class: c, Parent: String(r[6]).trim(), Phone: String(r[7]).trim(), Address: String(r[8]).trim() });
    });
  });
  return out;
}
function psSettings_() {
  const o = {};
  psTable_('FinanceSettings').rows().forEach(r => { o[String(r.Key).trim()] = r.Value; });
  return o;
}
function psActiveYear_() {
  const y = psTable_('AcademicYears').rows().find(r => String(r.Active).toUpperCase() === 'TRUE');
  return y ? String(y.Year) : String(new Date().getFullYear());
}
function psActiveTerm_() {
  const year = psActiveYear_();
  const terms = psTable_('AcademicTerms').rows();
  const t = terms.find(r => r.Status === 'Active' && String(r.AcademicYear) === year);
  return t ? t.Term : 'Term 1';
}
function psBalances_() {
  const acc = psTable_('CashBankAccounts').rows(), tx = psTable_('CashBankTransactions').rows(), m = {};
  acc.forEach(a => { m[a.AccountID] = psNum_(a.OpeningBalance); });
  tx.forEach(t => { if (m[t.AccountID] !== undefined) m[t.AccountID] += (t.Direction === 'In' ? 1 : -1) * psNum_(t.Amount); });
  return acc.map(a => Object.assign({}, a, { Balance: psRound_(m[a.AccountID]) }));
}
function psTxn_(accountId, direction, amount, refType, refId, description, date, by, transferTo) {
  psTable_('CashBankTransactions').append({
    TxnID: psId_('TX'), Date: date || psToday_(), AccountID: accountId, Direction: direction, Amount: psRound_(amount),
    RefType: refType, RefID: refId, Description: description || '', TransferToAccountID: transferTo || '',
    RecordedBy: by || '', CreatedAt: psNow_()
  });
}
function psPickAccount_(accountId, method) {
  const accs = psTable_('CashBankAccounts').rows().filter(a => a.Status !== 'Inactive');
  if (accountId && accs.some(a => a.AccountID === accountId)) return accountId;
  const type = method === 'Cash' ? 'Cash' : (method === 'Mobile Money' ? 'Mobile Money' : 'Bank');
  const hit = accs.find(a => a.Type === type) || accs[0];
  if (!hit) psErr_('No active cash/bank account. Add one under Cash & Bank.');
  return hit.AccountID;
}
function psInRange_(d, from, to) {
  const x = String(d || '').slice(0, 10);
  if (from && x < from) return false;
  if (to && x > to) return false;
  return true;
}

// ───────────────────────── common ─────────────────────────
function api_meta_(s) {
  const key = 'PSC_' + psCacheVer_() + '_metastatic';
  let m = psCacheGet_(key);
  if (m === undefined) { m = api_metaFresh_({ role: '', name: '', username: '' }); psCachePut_(key, m, 120); }
  return Object.assign({}, m, { role: s.role, name: s.name, username: s.username });
}
function api_metaFresh_(s) {
  const subjects = psTable_('AcademicSubjects').rows().map(r => r.Name);
  const st = psSettings_();
  return {
    role: s.role, name: s.name, username: s.username,
    classes: PS_CLASSES, subjects: subjects,
    positions: ['Academic Master', 'Class Teacher', 'Subject Teacher', 'Sports Manager', 'Librarian', 'Lab Technician'],
    year: psActiveYear_(), term: psActiveTerm_(),
    years: psTable_('AcademicYears').rows().map(r => String(r.Year)),
    terms: ['Term 1', 'Term 2'],
    methods: PS_METHODS,
    accounts: psTable_('CashBankAccounts').rows().filter(a => a.Status !== 'Inactive').map(a => ({ AccountID: a.AccountID, AccountName: a.AccountName, Type: a.Type })),
    incomeCats: psTable_('FinanceCategories').rows().filter(c => c.Type === 'Income' && c.Status !== 'Inactive').map(c => c.Name),
    expenseCats: psTable_('FinanceCategories').rows().filter(c => c.Type === 'Expense' && c.Status !== 'Inactive').map(c => c.Name),
    school: { name: st.SchoolName || '', nameSet: !!(st.SchoolName && st.SchoolName !== PS_PLACEHOLDER_NAME), code: st.SchoolCode || '', council: st.SchoolCouncil || '', address: st.SchoolAddress || '',
              phone: st.SchoolPhone || '', email: st.SchoolEmail || '', motto: st.SchoolMotto || '', logo: st.SchoolLogo || '',
              header1: st.ReportHeader1 || '', header2: st.ReportHeader2 || '',
              currency: st.Currency || 'TZS', receiptFooter: st.ReceiptFooter || '' }
  };
}

function api_changeMyPassword_(s, oldPw, newPw) {
  newPw = String(newPw || '');
  if (newPw.length < 6) psErr_('New password must have at least 6 characters.');
  const acc = psAdminByUsername_(s.username);
  if (!acc) psErr_('Account not found.');
  if (psHash_(oldPw, acc.Salt) !== acc.PasswordHash) psErr_('Current password is incorrect.');
  const salt = psSalt_();
  psTable_('Administrators').update(acc._row, { Salt: salt, PasswordHash: psHash_(newPw, salt) });
  psAuditS_(s, 'Account', 'Changed own password', acc.AdminID, '', '', '');
  return 'Password changed.';
}

function api_managerUpdateAccount_(s, p) {
  p = p || {};
  const t = psTable_('Administrators');
  const cur = t.rows().find(a => a.AdminID === s.adminId && a.Role === 'School Manager');
  if (!cur) psErr_('School Manager account not found.');
  const currentPassword = String(p.CurrentPassword || '');
  if (!currentPassword || psHash_(currentPassword, cur.Salt) !== cur.PasswordHash) psErr_('Current password is incorrect.');
  const fullName = String(p.FullName || '').trim();
  const username = String(p.Username || '').trim();
  if (!fullName) psErr_('Full name is required.');
  if (username.length < 3 || /\s/.test(username)) psErr_('Username must have at least 3 characters and no spaces.');
  if (psUsernameTaken_(username, cur.AdminID, null)) psErr_('Username "' + username + '" is already used.');

  const newPassword = String(p.NewPassword || '');
  if (newPassword && newPassword.length < 6) psErr_('New password must have at least 6 characters.');
  const usernameChanged = username !== cur.Username;
  const passwordChanged = !!newPassword;
  const update = { FullName: fullName, Username: username };
  if (passwordChanged) {
    const salt = psSalt_();
    update.Salt = salt;
    update.PasswordHash = psHash_(newPassword, salt);
  }
  t.update(cur._row, update);
  psAuditS_(s, 'Administrators', 'Updated own account', cur.AdminID, cur.Username, username,
    fullName + (passwordChanged ? ' (password changed)' : ''));
  return {
    fullName: fullName, username: username, usernameChanged: usernameChanged, passwordChanged: passwordChanged,
    message: 'Account settings updated.'
  };
}

function api_managerSetAcademicYear_(s, value) {
  const input = String(value || '').trim();
  const match = input.match(/^(\d{4})(?:\/(\d{4}))?$/);
  if (!match || (match[2] && Number(match[2]) !== Number(match[1]) + 1)) {
    psErr_('Enter an academic year such as 2026/2027.');
  }
  const year = match[1];
  const t = psTable_('AcademicYears');
  const rows = t.rows();
  const existing = rows.find(r => String(r.Year).trim() === year);
  if (!existing) t.append({ Year: year, Active: 'TRUE', ArchivedStudents: 0, Created: psToday_() });
  rows.forEach(r => t.update(r._row, { Active: String(r.Year).trim() === year ? 'TRUE' : 'FALSE' }));
  psAuditS_(s, 'AcademicYears', existing ? 'Activated' : 'Created and activated', year, '', year, 'Student records were not promoted or changed.');
  return { year: year, label: year + '/' + (Number(year) + 1), created: !existing };
}

const PS_BACKUP_FOLDER_NAME = 'PrimarySMS Backups';
const PS_BACKUP_MARKER = 'PrimarySMS system backup; sourceSpreadsheetId=';

function psBackupFolder_() {
  const folders = DriveApp.getFoldersByName(PS_BACKUP_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(PS_BACKUP_FOLDER_NAME);
}
function psCreateSystemBackup_(s, ss) {
  const folder = psBackupFolder_();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const name = 'PrimarySMS Backup - ' + ss.getName() + ' - ' + stamp;
  const file = DriveApp.getFileById(ss.getId()).makeCopy(name, folder);
  file.setDescription(PS_BACKUP_MARKER + ss.getId() + '; createdBy=' + s.username);
  return { id: file.getId(), name: file.getName(), createdAt: psPlain_(file.getDateCreated()) };
}
function api_managerBackupsList_(s) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folders = DriveApp.getFoldersByName(PS_BACKUP_FOLDER_NAME);
  if (!folders.hasNext()) return [];
  const files = folders.next().getFiles();
  const backups = [];
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS ||
        String(file.getDescription() || '').indexOf(PS_BACKUP_MARKER + ss.getId() + ';') !== 0) continue;
    backups.push({ id: file.getId(), name: file.getName(), createdAt: psPlain_(file.getDateCreated()) });
  }
  return backups.sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
}
function api_managerBackupCreate_(s) {
  const backup = psCreateSystemBackup_(s, SpreadsheetApp.getActiveSpreadsheet());
  psAuditS_(s, 'System', 'Created backup', backup.id, '', backup.name, '');
  return { backup: backup };
}
function api_managerBackupRestore_(s, backupId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(String(backupId || ''));
  if (file.getMimeType() !== MimeType.GOOGLE_SHEETS ||
      String(file.getDescription() || '').indexOf(PS_BACKUP_MARKER + ss.getId() + ';') !== 0) {
    psErr_('This is not a valid backup of the current school spreadsheet.');
  }
  const backup = SpreadsheetApp.openById(file.getId());
  const liveSheets = ss.getSheets();
  const backupSheets = backup.getSheets();
  const liveNames = liveSheets.map(sh => sh.getName()).sort();
  const backupNames = backupSheets.map(sh => sh.getName()).sort();
  if (liveNames.length !== backupNames.length || liveNames.some((name, i) => name !== backupNames[i])) {
    psErr_('The backup has a different sheet structure. Restore was cancelled to prevent a partial restore.');
  }

  const safetyBackup = psCreateSystemBackup_(s, ss);
  try {
    const sourceByName = {};
    backupSheets.forEach(sh => { sourceByName[sh.getName()] = sh; });
    liveSheets.forEach(target => {
      const source = sourceByName[target.getName()];
      const sourceRange = source.getDataRange();
      const values = sourceRange.getValues();
      const rows = values.length, cols = values[0] ? values[0].length : 0;
      if (rows > target.getMaxRows()) target.insertRowsAfter(target.getMaxRows(), rows - target.getMaxRows());
      if (cols > target.getMaxColumns()) target.insertColumnsAfter(target.getMaxColumns(), cols - target.getMaxColumns());
      target.getDataRange().clearContent();
      if (rows && cols) target.getRange(1, 1, rows, cols).setValues(values);
    });
    SpreadsheetApp.flush();
  } catch (e) {
    psErr_('Restore did not complete. A safety backup was created ("' + safetyBackup.name + '"). Details: ' + (e.message || e));
  }
  psAuditS_(s, 'System', 'Restored backup', file.getId(), '', file.getName(), 'Safety backup: ' + safetyBackup.name);
  return { backup: file.getName(), safetyBackup: safetyBackup.name };
}

function api_classes_(s) {
  const students = psAllStudents_();
  const counts = {};
  students.forEach(x => { counts[x.Class] = (counts[x.Class] || 0) + 1; });
  return psTable_('AcademicClasses').rows().map(c => ({
    ID: c.ID, Name: c.Name, Stream: c.Stream, Teacher: c.Teacher, Capacity: psNum_(c.Capacity), Students: counts[c.Name] || 0
  }));
}
function api_students_(s, className) { return psAllStudents_(className || ''); }
function api_subjects_(s) { return psTable_('AcademicSubjects').rows().map(r => ({ Name: r.Name, Code: r.Code, Category: r.Category, Department: r.Department, Classes: r.Classes })); }
function api_exams_(s) { return psTable_('Examinations').rows().map(r => ({ ID: r.ID, Name: r.Name, Type: r.Type, Year: r.Year, StartDate: r.StartDate, EndDate: r.EndDate, Classes: r.Classes, Status: r.Status, Remarks: r.Remarks })); }

function psGradeFor_(avg, scale) {
  const g = scale.find(r => avg >= psNum_(r.MinScore) && avg <= psNum_(r.MaxScore));
  return g ? { grade: g.Grade, points: psNum_(g.Points), remark: g.RemarkSW || g.RemarkEN || '' } : { grade: '-', points: 0, remark: '' };
}
function api_results_(s, className, examName) {
  if (!className || !examName) return { subjects: [], students: [], exams: [] };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const own = ss.getSheetByName(psClassSheetName_(className) + '_Scores');
  const rows = (own ? psTable_(psClassSheetName_(className) + '_Scores') : psTable_('Scores')).rows()
    .filter(r => psNormClass_(r.Class) === psNormClass_(className) && psNormName_(r.ExamName) === psNormName_(examName));
  const scale = psTable_('GradingScale').rows();
  const subjects = [];
  const by = {};
  rows.forEach(r => {
    if (subjects.indexOf(r.Subject) < 0) subjects.push(r.Subject);
    const k = r.AdmNo || r.StudentName;
    by[k] = by[k] || { AdmNo: r.AdmNo, Name: r.StudentName, Sex: r.Sex, scores: {} };
    by[k].scores[r.Subject] = psNum_(r.Score);
  });
  subjects.sort();
  const students = Object.keys(by).map(k => {
    const st = by[k], vals = subjects.map(x => st.scores[x]).filter(v => v !== undefined);
    st.total = psRound_(vals.reduce((a, b) => a + b, 0));
    st.average = vals.length ? psRound_(st.total / vals.length) : 0;
    const g = psGradeFor_(st.average, scale);
    st.grade = g.grade; st.remark = g.remark;
    return st;
  }).sort((a, b) => b.total - a.total);
  students.forEach((st, i) => { st.position = (i > 0 && st.total === students[i - 1].total) ? students[i - 1].position : i + 1; });
  const avg = students.length ? psRound_(students.reduce((a, b) => a + b.average, 0) / students.length) : 0;
  return { subjects: subjects, students: students, classAverage: avg };
}
function api_exam_types_(s, className) {
  return fetchExamTypes(className);
}
// Academic Settings (report logos, names, signatures) — read only for School Manager / Head of School
function api_acad_settings_(s) {
  return getAcademicSettings();
}
function api_exam_deadline_(s) {
  return getDeadlineData();
}
function api_full_results_(s, className, examName) {
  const report = fetchTopPanelData(className, examName);
  if (!report || !report.students || !report.students.length) return { report: report || {}, behavior: { map: {} } };
  return { report: report, behavior: fetchReportBehaviors(className, examName) || { map: {} } };
}

function api_attendance_(s, p) {
  p = p || {};
  return psTable_('AttendanceSummary').rows()
    .filter(r => psInRange_(r.Date, p.from, p.to) && (!p.className || psNormClass_(r.Class) === psNormClass_(p.className)))
    .map(r => {
      const total = psNum_(r.Total), absent = psNum_(r.Absent);
      return { Date: r.Date, Class: r.Class, Total: total, Absent: absent, Present: total - absent,
               PresentPct: total ? psRound_((total - absent) * 100 / total) : 0, Reason: r.Reason, Teacher: r.Teacher };
    }).sort((a, b) => String(b.Date).localeCompare(String(a.Date)));
}

function api_overview_(s) {
  const students = psAllStudents_();
  const out = { role: s.role, students: students.length, classes: PS_CLASSES.length };
  const users = psTable_('Users').rows();
  out.staff = users.filter(u => String(u.Status || '').toLowerCase() !== 'inactive').length;
  out.parents = psTable_('Parents').rows().length;
  const today = psToday_();
  const att = psTable_('AttendanceSummary').rows().filter(r => String(r.Date).slice(0, 10) === today);
  const tot = att.reduce((a, r) => a + psNum_(r.Total), 0), abs = att.reduce((a, r) => a + psNum_(r.Absent), 0);
  out.attendanceToday = tot ? psRound_((tot - abs) * 100 / tot) : null;
  out.exams = psTable_('Examinations').rows().length;
  if (s.role !== 'headofschool') {
    const f = fin_totals_();
    out.billed = f.billed; out.collected = f.collected; out.outstanding = f.outstanding;
    out.income = f.otherIncome; out.expenses = f.expenses;
    out.cash = psBalances_().reduce((a, b) => a + b.Balance, 0);
  }
  if (s.role === 'schoolmanager') out.admins = api_adminsList_(s);
  if (s.role === 'headofschool') out.recentStaff = users.slice(-5).reverse().map(u => ({ Name: u.Name, Position: u.Position }));
  return out;
}

// ───────────────────────── School Manager: administrators ─────────────────────────
function api_adminsList_(s) {
  return psTable_('Administrators').rows().filter(a => a.Role === 'Head of School' || a.Role === 'Accountant').map(a => ({
    AdminID: a.AdminID, FullName: a.FullName, Username: a.Username, Role: a.Role, Sex: a.Sex, Phone: a.Phone,
    Email: a.Email, Status: a.Status, CreatedAt: a.CreatedAt, LastLogin: a.LastLogin
  }));
}
function api_adminsSave_(s, p) {
  p = p || {};
  const t = psTable_('Administrators');
  const role = String(p.Role || '');
  if (['Head of School', 'Accountant'].indexOf(role) < 0) psErr_('Role must be Head of School or Accountant.');
  const name = String(p.FullName || '').trim(), username = String(p.Username || '').trim();
  if (!name) psErr_('Full name is required.');
  if (username.length < 3 || /\s/.test(username)) psErr_('Username must have at least 3 characters and no spaces.');

  if (p.AdminID) {
    const cur = t.rows().find(a => a.AdminID === p.AdminID);
    if (!cur) psErr_('Administrator not found.');
    if (cur.Role !== 'Head of School' && cur.Role !== 'Accountant') psErr_('Only Head of School and Accountant accounts can be managed here.');
    if (psUsernameTaken_(username, cur.AdminID, null)) psErr_('Username "' + username + '" is already used.');
    const upd = { FullName: name, Username: username, Sex: p.Sex || '', Phone: p.Phone || '', Email: p.Email || '', Status: p.Status || 'Active' };
    upd.Role = role;
    t.update(cur._row, upd);
    psAuditS_(s, 'Administrators', 'Updated', cur.AdminID, cur.Username, username, name + ' (' + (upd.Role || cur.Role) + ')');
    return 'Administrator updated.';
  }
  const pw = String(p.Password || '');
  if (pw.length < 6) psErr_('Password must have at least 6 characters.');
  if (psUsernameTaken_(username, null, null)) psErr_('Username "' + username + '" is already used.');
  const salt = psSalt_(), id = psId_('ADM');
  t.append({ AdminID: id, FullName: name, Username: username, PasswordHash: psHash_(pw, salt), Salt: salt, Role: role,
             Sex: p.Sex || '', Phone: p.Phone || '', Email: p.Email || '', Status: p.Status || 'Active', PicUrl: '',
             CreatedBy: s.username, CreatedAt: psNow_(), LastLogin: '' });
  psAuditS_(s, 'Administrators', 'Created', id, '', username, name + ' (' + role + ')');
  return role + ' account created for ' + name + '.';
}
function api_adminsDelete_(s, adminId) {
  const t = psTable_('Administrators');
  const cur = t.rows().find(a => a.AdminID === adminId);
  if (!cur) psErr_('Administrator not found.');
  if (cur.Role !== 'Head of School' && cur.Role !== 'Accountant') psErr_('Only Head of School and Accountant accounts can be deleted here.');
  t.remove(cur._row);
  psAuditS_(s, 'Administrators', 'Deleted', adminId, cur.Username, '', cur.FullName + ' (' + cur.Role + ')');
  return 'Administrator deleted.';
}
function api_adminsReset_(s, adminId, newPassword) {
  newPassword = String(newPassword || '');
  if (newPassword.length < 6) psErr_('Password must have at least 6 characters.');
  const t = psTable_('Administrators');
  const cur = t.rows().find(a => a.AdminID === adminId);
  if (!cur) psErr_('Administrator not found.');
  if (cur.Role !== 'Head of School' && cur.Role !== 'Accountant') psErr_('Only Head of School and Accountant passwords can be reset here.');
  const salt = psSalt_();
  t.update(cur._row, { Salt: salt, PasswordHash: psHash_(newPassword, salt) });
  psAuditS_(s, 'Administrators', 'Reset password', adminId, '', '', cur.Username);
  return 'Password reset for ' + cur.FullName + '.';
}

// ───────────────────────── Head of School: staff accounts ─────────────────────────
const PS_POSITIONS = ['Academic Master', 'Class Teacher', 'Subject Teacher', 'Sports Manager', 'Librarian', 'Lab Technician'];

function api_staffList_(s) {
  return psTable_('Users').rows().map(u => ({
    Name: u.Name, Username: u.Username, Sex: u.Sex, Position: u.Position,
    Classes: u.Classes, Subjects: u.Subjects, Phone: u.Phone, Email: u.Email, Status: u.Status || 'Active', _row: u._row
  }));
}
// Class Teachers are assigned in Academic > Manage Classes. When a staff account is removed or its
// login name changes, keep that list pointing at the right person.
function psRelinkClassTeacher_(oldName, oldUsername, newUsername) {
  const t = psTable_('AcademicClasses');
  const a = psNormName_(oldName), b = psNormName_(oldUsername);
  t.rows().forEach(c => {
    const cur = psNormName_(c.Teacher);
    if (!cur || (cur !== a && cur !== b)) return;
    t.update(c._row, { Teacher: newUsername || '' });
  });
}
function api_staffSave_(s, p) {
  p = p || {};
  const t = psTable_('Users');
  const name = String(p.Name || '').trim(), username = String(p.Username || '').trim();
  if (!name) psErr_('Full name is required.');
  if (username.length < 3 || /\s/.test(username)) psErr_('Username must have at least 3 characters and no spaces.');
  if (PS_POSITIONS.indexOf(p.Position) < 0) psErr_('Choose a valid position.');
  const classes = (Array.isArray(p.Classes) ? p.Classes : String(p.Classes || '').split(',')).map(x => String(x).trim()).filter(Boolean);
  const subjects = (Array.isArray(p.Subjects) ? p.Subjects : String(p.Subjects || '').split(',')).map(x => String(x).trim()).filter(Boolean);
  classes.forEach(c => { if (PS_CLASSES.indexOf(c) < 0) psErr_('Unknown class: ' + c); });
  // A teacher may teach any number of classes and subjects. The Class Teacher of a class is chosen by the
  // Academic Master in Manage Classes, so there is no class rule here.
  const base = { Name: name, Username: username, Classes: classes.join(', '), Subjects: subjects.join(', '), Sex: p.Sex || '',
                 Position: p.Position, Phone: p.Phone || '', Email: p.Email || '', Status: p.Status || 'Active' };
  if (p._row) {
    const cur = t.rows().find(u => u._row === p._row);
    if (!cur) psErr_('Staff member not found.');
    if (psUsernameTaken_(username, null, cur._row)) psErr_('Username "' + username + '" is already used.');
    if (p.Password) {
      if (String(p.Password).length < 6) psErr_('Password must have at least 6 characters.');
      base.Password = String(p.Password);
    }
    t.update(cur._row, base);
    if (cur.Position === 'Class Teacher' && p.Position !== 'Class Teacher') psRelinkClassTeacher_(cur.Name, cur.Username, '');   // no longer a Class Teacher: the class is free again
    else if (psNormName_(cur.Username) !== psNormName_(username) || psNormName_(cur.Name) !== psNormName_(name)) psRelinkClassTeacher_(cur.Name, cur.Username, username);
    try { logActivity(s.username, 'Updated staff ' + username); } catch (e) {}
    return 'Staff account updated.';
  }
  const pw = String(p.Password || '');
  if (pw.length < 6) psErr_('Password must have at least 6 characters.');
  if (psUsernameTaken_(username, null, null)) psErr_('Username "' + username + '" is already used.');
  base.Password = pw; base.PicUrl = '';
  t.append(base);
  try { logActivity(s.username, 'Created staff ' + username + ' (' + p.Position + ')'); } catch (e) {}
  return p.Position + ' account created for ' + name + '.';
}
function api_staffDelete_(s, rowNum) {
  const t = psTable_('Users');
  const cur = t.rows().find(u => u._row === rowNum);
  if (!cur) psErr_('Staff member not found.');
  psRelinkClassTeacher_(cur.Name, cur.Username, '');
  t.remove(cur._row);
  try { logActivity(s.username, 'Deleted staff ' + cur.Username); } catch (e) {}
  return 'Staff account deleted.';
}
function api_staffReset_(s, rowNum, newPassword) {
  newPassword = String(newPassword || '');
  if (newPassword.length < 6) psErr_('Password must have at least 6 characters.');
  const t = psTable_('Users');
  const cur = t.rows().find(u => u._row === rowNum);
  if (!cur) psErr_('Staff member not found.');
  t.update(cur._row, { Password: newPassword });
  try { logActivity(s.username, 'Reset password of ' + cur.Username); } catch (e) {}
  return 'Password reset for ' + cur.Name + '.';
}

// ═════════════════════════ FINANCE ═════════════════════════
function psLineStatus_(net, paid) {
  if (paid <= 0) return 'Unpaid';
  return paid + 0.0001 >= net ? 'Paid' : 'Partial';
}
function fin_totals_() {
  const inv = psTable_('FeeInvoices').rows().filter(r => r.Status !== 'Cancelled');
  const billed = inv.reduce((a, r) => a + psNum_(r.NetAmount), 0);
  const paid = inv.reduce((a, r) => a + psNum_(r.Paid), 0);
  const inc = psTable_('Income').rows().filter(r => r.Status !== 'Voided').reduce((a, r) => a + psNum_(r.Amount), 0);
  const exp = psTable_('Expenditure').rows().filter(r => r.Status === 'Paid').reduce((a, r) => a + psNum_(r.Amount), 0);
  return { billed: psRound_(billed), collected: psRound_(paid), outstanding: psRound_(billed - paid), otherIncome: psRound_(inc), expenses: psRound_(exp) };
}

// ── fee structure ──
function fin_feesList_(s, p) {
  p = p || {};
  return psTable_('FeeStructure').rows().filter(r => (!p.year || String(r.AcademicYear) === String(p.year)))
    .map(r => ({ FeeID: r.FeeID, AcademicYear: String(r.AcademicYear), Class: r.Class, Term: r.Term, FeeName: r.FeeName, Amount: psNum_(r.Amount), Mandatory: r.Mandatory, Status: r.Status, Notes: r.Notes }));
}
function fin_feesSave_(s, p) {
  p = p || {};
  const name = String(p.FeeName || '').trim();
  if (!name) psErr_('Fee name is required.');
  if (psNum_(p.Amount) <= 0) psErr_('Amount must be greater than zero.');
  const t = psTable_('FeeStructure');
  const rec = { AcademicYear: String(p.AcademicYear || psActiveYear_()), Class: p.Class || 'All', Term: p.Term || 'Term 1', FeeName: name,
                Amount: psNum_(p.Amount), Mandatory: p.Mandatory || 'Yes', Status: p.Status || 'Active', Notes: p.Notes || '' };
  if (p.FeeID) {
    const cur = t.rows().find(r => r.FeeID === p.FeeID);
    if (!cur) psErr_('Fee not found.');
    t.update(cur._row, rec);
    psAuditS_(s, 'FeeStructure', 'Updated', p.FeeID, cur.Amount, rec.Amount, name);
    return 'Fee updated. (Existing invoices are not changed.)';
  }
  rec.FeeID = psId_('FEE');
  t.append(rec);
  psAuditS_(s, 'FeeStructure', 'Created', rec.FeeID, '', rec.Amount, name + ' / ' + rec.Class + ' / ' + rec.Term);
  return 'Fee added.';
}
function fin_feesDelete_(s, feeId) {
  const t = psTable_('FeeStructure');
  const cur = t.rows().find(r => r.FeeID === feeId);
  if (!cur) psErr_('Fee not found.');
  t.remove(cur._row);
  psAuditS_(s, 'FeeStructure', 'Deleted', feeId, cur.Amount, '', cur.FeeName);
  return 'Fee deleted.';
}

// ── invoices ──
function fin_invList_(s, p) {
  p = p || {};
  const q = psNormName_(p.q);
  return psTable_('FeeInvoices').rows().filter(r =>
    (!p.year || String(r.AcademicYear) === String(p.year)) && (!p.term || r.Term === p.term) &&
    (!p.className || psNormClass_(r.Class) === psNormClass_(p.className)) &&
    (!p.status || r.Status === p.status) &&
    (!q || psNormName_(r.StudentName).indexOf(q) >= 0 || psNormName_(r.AdmNo).indexOf(q) >= 0 || psNormName_(r.InvoiceNo).indexOf(q) >= 0))
    .map(r => ({ InvoiceNo: r.InvoiceNo, IssueDate: r.IssueDate, DueDate: r.DueDate, AdmNo: r.AdmNo, StudentName: r.StudentName, Class: r.Class,
                 AcademicYear: String(r.AcademicYear), Term: r.Term, FeeName: r.FeeName, Amount: psNum_(r.Amount), Discount: psNum_(r.Discount),
                 NetAmount: psNum_(r.NetAmount), Paid: psNum_(r.Paid), Balance: psNum_(r.Balance), Status: r.Status }))
    .slice(0, 3000);
}
function fin_invGenerate_(s, p) {
  p = p || {};
  const year = String(p.year || psActiveYear_()), term = String(p.term || psActiveTerm_());
  const fees = psTable_('FeeStructure').rows().filter(f => f.Status !== 'Inactive' && String(f.AcademicYear) === year &&
    f.Mandatory !== 'No' && (f.Term === term || (f.Term === 'Annual' && term === 'Term 1')));
  if (!fees.length) psErr_('No active mandatory fees for ' + year + ' / ' + term + '. Add them under School Fees → Fee Structure.');
  const cls = p.className && p.className !== 'All' ? p.className : '';
  const students = psAllStudents_(cls);
  if (!students.length) psErr_('No students found' + (cls ? ' in ' + cls : '') + '.');
  const inv = psTable_('FeeInvoices');
  const existing = inv.rows().filter(r => r.Status !== 'Cancelled' && String(r.AcademicYear) === year && r.Term === term);
  const have = {}, noBy = {};
  existing.forEach(r => { have[r.AdmNo + '|' + r.FeeName] = true; noBy[r.AdmNo] = r.InvoiceNo; });

  const plan = [], needNo = [];
  students.forEach(st => {
    const mine = fees.filter(f => f.Class === 'All' || psNormClass_(f.Class) === psNormClass_(st.Class));
    const missing = mine.filter(f => !have[st.AdmNo + '|' + f.FeeName]);
    if (!missing.length) return;
    if (!noBy[st.AdmNo]) needNo.push(st.AdmNo);
    plan.push({ st: st, fees: missing });
  });
  const nos = psNextNumbers_('NextInvoiceNo', 'InvoicePrefix', needNo.length);
  needNo.forEach((a, i) => { noBy[a] = nos[i]; });
  const rows = [], issue = psToday_(), due = p.dueDate || '';
  plan.forEach(x => x.fees.forEach(f => {
    const amt = psNum_(f.Amount);
    rows.push({ InvoiceNo: noBy[x.st.AdmNo], IssueDate: issue, DueDate: due, AdmNo: x.st.AdmNo, StudentName: x.st.Name, Class: x.st.Class,
                AcademicYear: year, Term: term, FeeName: f.FeeName, Amount: amt, Discount: 0, NetAmount: amt, Paid: 0, Balance: amt,
                Status: 'Unpaid', CreatedBy: s.username, CreatedAt: psNow_() });
  }));
  inv.appendMany(rows);
  psAuditS_(s, 'Invoices', 'Generated', year + '/' + term, '', rows.length, (cls || 'All classes') + ': ' + plan.length + ' students');
  return { students: plan.length, lines: rows.length, message: rows.length ? ('Billed ' + plan.length + ' student(s): ' + rows.length + ' fee line(s) created.') : 'Everything was already billed — nothing new to create.' };
}
function fin_invAddLine_(s, p) {
  p = p || {};
  const st = psAllStudents_().find(x => x.AdmNo === String(p.admNo || '').trim());
  if (!st) psErr_('Student not found.');
  const amt = psNum_(p.amount);
  if (!String(p.feeName || '').trim() || amt <= 0) psErr_('Fee name and amount are required.');
  const year = String(p.year || psActiveYear_()), term = p.term || psActiveTerm_();
  const inv = psTable_('FeeInvoices');
  const mine = inv.rows().filter(r => r.AdmNo === st.AdmNo && String(r.AcademicYear) === year && r.Term === term && r.Status !== 'Cancelled');
  if (mine.some(r => r.FeeName === p.feeName)) psErr_('This fee is already on the invoice.');
  const no = mine.length ? mine[0].InvoiceNo : psNextNumbers_('NextInvoiceNo', 'InvoicePrefix', 1)[0];
  inv.append({ InvoiceNo: no, IssueDate: psToday_(), DueDate: p.dueDate || '', AdmNo: st.AdmNo, StudentName: st.Name, Class: st.Class, AcademicYear: year,
               Term: term, FeeName: String(p.feeName).trim(), Amount: amt, Discount: 0, NetAmount: amt, Paid: 0, Balance: amt, Status: 'Unpaid',
               CreatedBy: s.username, CreatedAt: psNow_() });
  psAuditS_(s, 'Invoices', 'Added line', no, '', amt, st.Name + ' — ' + p.feeName);
  return 'Fee added to invoice ' + no + '.';
}
function fin_invDiscount_(s, invoiceNo, feeName, discount) {
  discount = psNum_(discount);
  const t = psTable_('FeeInvoices');
  const line = t.rows().find(r => r.InvoiceNo === invoiceNo && r.FeeName === feeName && r.Status !== 'Cancelled');
  if (!line) psErr_('Invoice line not found.');
  if (discount < 0 || discount > psNum_(line.Amount)) psErr_('Discount must be between 0 and ' + psNum_(line.Amount) + '.');
  const net = psRound_(psNum_(line.Amount) - discount), paid = psNum_(line.Paid);
  if (net + 0.0001 < paid) psErr_('Discount is too big: the student has already paid ' + paid + '.');
  t.update(line._row, { Discount: discount, NetAmount: net, Balance: psRound_(net - paid), Status: psLineStatus_(net, paid) });
  psAuditS_(s, 'Invoices', 'Discount', invoiceNo, line.Discount, discount, line.StudentName + ' — ' + feeName);
  return 'Discount saved.';
}
function fin_invCancel_(s, invoiceNo, feeName, reason) {
  const t = psTable_('FeeInvoices');
  const line = t.rows().find(r => r.InvoiceNo === invoiceNo && r.FeeName === feeName && r.Status !== 'Cancelled');
  if (!line) psErr_('Invoice line not found.');
  if (psNum_(line.Paid) > 0) psErr_('This fee already has payments. Void the payment first.');
  t.update(line._row, { Status: 'Cancelled', Balance: 0 });
  psAuditS_(s, 'Invoices', 'Cancelled line', invoiceNo, line.NetAmount, 0, line.StudentName + ' — ' + feeName + (reason ? ' (' + reason + ')' : ''));
  return 'Fee cancelled.';
}

// ── student lookup, account & payments ──
function fin_studentSearch_(s, q) {
  q = psNormName_(q);
  if (q.length < 2) return [];
  return psAllStudents_().filter(x => psNormName_(x.Name).indexOf(q) >= 0 || psNormName_(x.AdmNo).indexOf(q) >= 0)
    .slice(0, 15).map(x => ({ AdmNo: x.AdmNo, Name: x.Name, Class: x.Class, Parent: x.Parent, Phone: x.Phone }));
}
function fin_studentAccount_(s, admNo) {
  admNo = String(admNo || '').trim();
  const st = psAllStudents_().find(x => x.AdmNo === admNo);
  if (!st) psErr_('Student not found.');
  const lines = psTable_('FeeInvoices').rows().filter(r => r.AdmNo === admNo && r.Status !== 'Cancelled')
    .map(r => ({ InvoiceNo: r.InvoiceNo, AcademicYear: String(r.AcademicYear), Term: r.Term, FeeName: r.FeeName, NetAmount: psNum_(r.NetAmount),
                 Paid: psNum_(r.Paid), Balance: psNum_(r.Balance), Status: r.Status, IssueDate: r.IssueDate }));
  const pays = psTable_('Payments').rows().filter(r => r.AdmNo === admNo && r.Status !== 'Voided')
    .map(r => ({ ReceiptNo: r.ReceiptNo, PaymentDate: r.PaymentDate, FeeName: r.FeeName, Amount: psNum_(r.Amount), Method: r.Method }));
  return { student: st, lines: lines, payments: pays.slice(-30).reverse(),
           totals: { billed: psRound_(lines.reduce((a, b) => a + b.NetAmount, 0)), paid: psRound_(lines.reduce((a, b) => a + b.Paid, 0)),
                     balance: psRound_(lines.reduce((a, b) => a + b.Balance, 0)) } };
}
function fin_pay_(s, p) {
  p = p || {};
  const amount = psRound_(psNum_(p.amount));
  if (amount <= 0) psErr_('Amount must be greater than zero.');
  const adm = String(p.admNo || '').trim();
  const st = psAllStudents_().find(x => x.AdmNo === adm);
  if (!st) psErr_('Student not found.');
  const method = PS_METHODS.indexOf(p.method) >= 0 ? p.method : 'Cash';
  const accountId = psPickAccount_(p.accountId, method);
  const invT = psTable_('FeeInvoices');
  let lines = invT.rows().filter(r => r.AdmNo === adm && r.Status !== 'Cancelled' && r.Status !== 'Paid' && psNum_(r.Balance) > 0);
  if (p.invoiceNo) lines = lines.filter(r => r.InvoiceNo === p.invoiceNo);
  if (p.feeName) lines = lines.filter(r => r.FeeName === p.feeName);
  lines.sort((a, b) => String(a.IssueDate).localeCompare(String(b.IssueDate)) || a._row - b._row);
  const totalBal = psRound_(lines.reduce((a, r) => a + psNum_(r.Balance), 0));
  if (!lines.length) psErr_('This student has no unpaid fees. Generate invoices first.');
  if (amount > totalBal + 0.0001) psErr_('Amount (' + amount + ') is more than the outstanding balance (' + totalBal + ').');

  const receipt = psNextNumbers_('NextReceiptNo', 'ReceiptPrefix', 1)[0];
  let left = amount;
  const alloc = [], payRows = [], date = p.date || psToday_();
  lines.forEach(l => {
    if (left <= 0) return;
    const take = psRound_(Math.min(left, psNum_(l.Balance)));
    if (take <= 0) return;
    left = psRound_(left - take);
    const newPaid = psRound_(psNum_(l.Paid) + take), net = psNum_(l.NetAmount);
    invT.update(l._row, { Paid: newPaid, Balance: psRound_(net - newPaid), Status: psLineStatus_(net, newPaid) });
    alloc.push({ FeeName: l.FeeName, Term: l.Term, Amount: take });
    payRows.push({ ReceiptNo: receipt, PaymentDate: date, AdmNo: adm, StudentName: st.Name, Class: st.Class, AcademicYear: String(l.AcademicYear),
                   Term: l.Term, InvoiceNo: l.InvoiceNo, FeeName: l.FeeName, Amount: take, Method: method, AccountID: accountId,
                   Reference: p.reference || '', PaidBy: p.paidBy || '', ReceivedBy: s.name, Status: 'Valid', VoidReason: '', Notes: p.notes || '', CreatedAt: psNow_() });
  });
  psTable_('Payments').appendMany(payRows);
  psTxn_(accountId, 'In', amount, 'Payment', receipt, 'Fees — ' + st.Name + ' (' + adm + ')', date, s.name);
  psAuditS_(s, 'Payments', 'Received', receipt, '', amount, st.Name + ' / ' + method);
  const acct = fin_studentAccount_(s, adm);
  return { receiptNo: receipt, date: date, student: { AdmNo: adm, Name: st.Name, Class: st.Class }, lines: alloc, total: amount, method: method,
           reference: p.reference || '', paidBy: p.paidBy || '', receivedBy: s.name, balanceAfter: acct.totals.balance };
}
function fin_receipt_(s, receiptNo) {
  const rows = psTable_('Payments').rows().filter(r => r.ReceiptNo === receiptNo);
  if (!rows.length) psErr_('Receipt not found.');
  const f = rows[0];
  const acct = fin_studentAccount_(s, f.AdmNo);
  return { receiptNo: receiptNo, date: f.PaymentDate, student: { AdmNo: f.AdmNo, Name: f.StudentName, Class: f.Class },
           lines: rows.map(r => ({ FeeName: r.FeeName, Term: r.Term, Amount: psNum_(r.Amount) })),
           total: psRound_(rows.reduce((a, r) => a + psNum_(r.Amount), 0)), method: f.Method, reference: f.Reference, paidBy: f.PaidBy,
           receivedBy: f.ReceivedBy, status: f.Status, balanceAfter: acct.totals.balance };
}
function fin_voidPayment_(s, receiptNo, reason) {
  reason = String(reason || '').trim();
  if (reason.length < 3) psErr_('Give a reason for voiding the payment.');
  const pt = psTable_('Payments');
  const rows = pt.rows().filter(r => r.ReceiptNo === receiptNo);
  if (!rows.length) psErr_('Receipt not found.');
  if (rows.some(r => r.Status === 'Voided')) psErr_('This receipt is already voided.');
  const invT = psTable_('FeeInvoices'), inv = invT.rows();
  let total = 0;
  rows.forEach(r => {
    const line = inv.find(x => x.InvoiceNo === r.InvoiceNo && x.FeeName === r.FeeName);
    const amt = psNum_(r.Amount); total += amt;
    if (line) {
      const paid = Math.max(0, psRound_(psNum_(line.Paid) - amt)), net = psNum_(line.NetAmount);
      invT.update(line._row, { Paid: paid, Balance: psRound_(net - paid), Status: line.Status === 'Cancelled' ? 'Cancelled' : psLineStatus_(net, paid) });
      line.Paid = paid;
    }
    pt.update(r._row, { Status: 'Voided', VoidReason: reason });
  });
  psTxn_(rows[0].AccountID, 'Out', total, 'Adjustment', receiptNo, 'VOID receipt ' + receiptNo + ': ' + reason, psToday_(), s.name);
  psAuditS_(s, 'Payments', 'Voided', receiptNo, total, 0, rows[0].StudentName + ' — ' + reason);
  return 'Receipt ' + receiptNo + ' voided.';
}
function fin_payments_(s, p) {
  p = p || {};
  const q = psNormName_(p.q);
  return psTable_('Payments').rows().filter(r =>
    psInRange_(r.PaymentDate, p.from, p.to) && (!p.className || psNormClass_(r.Class) === psNormClass_(p.className)) &&
    (!p.method || r.Method === p.method) && (!p.status || r.Status === p.status) &&
    (!q || psNormName_(r.StudentName).indexOf(q) >= 0 || psNormName_(r.AdmNo).indexOf(q) >= 0 || psNormName_(r.ReceiptNo).indexOf(q) >= 0))
    .map(r => ({ ReceiptNo: r.ReceiptNo, PaymentDate: r.PaymentDate, AdmNo: r.AdmNo, StudentName: r.StudentName, Class: r.Class, Term: r.Term,
                 FeeName: r.FeeName, Amount: psNum_(r.Amount), Method: r.Method, Reference: r.Reference, ReceivedBy: r.ReceivedBy, Status: r.Status, VoidReason: r.VoidReason }))
    .sort((a, b) => String(b.PaymentDate).localeCompare(String(a.PaymentDate)) || String(b.ReceiptNo).localeCompare(String(a.ReceiptNo)))
    .slice(0, 3000);
}

// ── debtors ──
function fin_debtors_(s, p) {
  p = p || {};
  const min = psNum_(p.min), q = psNormName_(p.q);
  const students = {}; psAllStudents_().forEach(x => { students[x.AdmNo] = x; });
  const by = {};
  psTable_('FeeInvoices').rows().filter(r => r.Status !== 'Cancelled').forEach(r => {
    if (p.year && String(r.AcademicYear) !== String(p.year)) return;
    const o = by[r.AdmNo] = by[r.AdmNo] || { AdmNo: r.AdmNo, StudentName: r.StudentName, Class: r.Class, Billed: 0, Paid: 0, Balance: 0, Oldest: '' };
    o.Billed += psNum_(r.NetAmount); o.Paid += psNum_(r.Paid); o.Balance += psNum_(r.Balance);
    if (psNum_(r.Balance) > 0 && (!o.Oldest || String(r.IssueDate) < o.Oldest)) o.Oldest = String(r.IssueDate);
  });
  const fu = {}; psTable_('DebtorFollowups').rows().forEach(f => { if (!fu[f.AdmNo] || String(f.Date) > fu[f.AdmNo]) fu[f.AdmNo] = String(f.Date); });
  return Object.keys(by).map(k => by[k]).filter(o => o.Balance > 0.0001 && o.Balance >= min &&
      (!p.className || psNormClass_(o.Class) === psNormClass_(p.className)) &&
      (!q || psNormName_(o.StudentName).indexOf(q) >= 0 || psNormName_(o.AdmNo).indexOf(q) >= 0))
    .map(o => ({ AdmNo: o.AdmNo, StudentName: o.StudentName, Class: o.Class, Billed: psRound_(o.Billed), Paid: psRound_(o.Paid), Balance: psRound_(o.Balance),
                 Oldest: o.Oldest, Parent: (students[o.AdmNo] || {}).Parent || '', Phone: (students[o.AdmNo] || {}).Phone || '', LastFollowup: fu[o.AdmNo] || '' }))
    .sort((a, b) => b.Balance - a.Balance);
}
function fin_fuList_(s, admNo) {
  return psTable_('DebtorFollowups').rows().filter(f => !admNo || f.AdmNo === admNo).map(f => ({
    FollowupID: f.FollowupID, Date: f.Date, AdmNo: f.AdmNo, StudentName: f.StudentName, Class: f.Class, OutstandingAmount: psNum_(f.OutstandingAmount),
    Method: f.Method, Notes: f.Notes, PromisedDate: f.PromisedDate, PromisedAmount: psNum_(f.PromisedAmount), Status: f.Status, By: f.By
  })).sort((a, b) => String(b.Date).localeCompare(String(a.Date)));
}
function fin_fuSave_(s, p) {
  p = p || {};
  const st = psAllStudents_().find(x => x.AdmNo === String(p.AdmNo || '').trim());
  if (!st) psErr_('Student not found.');
  const bal = fin_studentAccount_(s, st.AdmNo).totals.balance;
  const rec = { Date: p.Date || psToday_(), AdmNo: st.AdmNo, StudentName: st.Name, Class: st.Class, OutstandingAmount: bal, Method: p.Method || 'Phone Call',
                Notes: p.Notes || '', PromisedDate: p.PromisedDate || '', PromisedAmount: psNum_(p.PromisedAmount), Status: p.Status || 'Open', By: s.name };
  const t = psTable_('DebtorFollowups');
  if (p.FollowupID) {
    const cur = t.rows().find(f => f.FollowupID === p.FollowupID);
    if (!cur) psErr_('Follow-up not found.');
    t.update(cur._row, rec);
    return 'Follow-up updated.';
  }
  rec.FollowupID = psId_('FU');
  t.append(rec);
  return 'Follow-up saved.';
}
function fin_fuDelete_(s, id) {
  const t = psTable_('DebtorFollowups');
  const cur = t.rows().find(f => f.FollowupID === id);
  if (!cur) psErr_('Follow-up not found.');
  t.remove(cur._row);
  return 'Follow-up deleted.';
}

// ── other income ──
function fin_incomeList_(s, p) {
  p = p || {};
  return psTable_('Income').rows().filter(r => psInRange_(r.Date, p.from, p.to) && (!p.category || r.Category === p.category) && (!p.status || r.Status === p.status))
    .map(r => ({ IncomeID: r.IncomeID, Date: r.Date, Category: r.Category, Description: r.Description, Amount: psNum_(r.Amount), Method: r.Method,
                 AccountID: r.AccountID, Reference: r.Reference, ReceivedFrom: r.ReceivedFrom, RecordedBy: r.RecordedBy, Status: r.Status }))
    .sort((a, b) => String(b.Date).localeCompare(String(a.Date)));
}
function fin_incomeSave_(s, p) {
  p = p || {};
  const amt = psRound_(psNum_(p.Amount));
  if (amt <= 0) psErr_('Amount must be greater than zero.');
  if (!String(p.Category || '').trim()) psErr_('Choose a category.');
  const method = PS_METHODS.indexOf(p.Method) >= 0 ? p.Method : 'Cash';
  const accountId = psPickAccount_(p.AccountID, method);
  const cat = psTable_('FinanceCategories').rows().find(c => c.Type === 'Income' && c.Name === p.Category);
  const id = psId_('INC'), date = p.Date || psToday_();
  psTable_('Income').append({ IncomeID: id, Date: date, CategoryID: cat ? cat.CategoryID : '', Category: p.Category, Description: p.Description || '', Amount: amt,
                              Method: method, AccountID: accountId, Reference: p.Reference || '', ReceivedFrom: p.ReceivedFrom || '', RecordedBy: s.name, Status: 'Valid', CreatedAt: psNow_() });
  psTxn_(accountId, 'In', amt, 'Income', id, p.Category + (p.Description ? ': ' + p.Description : ''), date, s.name);
  psAuditS_(s, 'Income', 'Recorded', id, '', amt, p.Category);
  return 'Income recorded.';
}
function fin_incomeVoid_(s, id, reason) {
  reason = String(reason || '').trim();
  if (reason.length < 3) psErr_('Give a reason.');
  const t = psTable_('Income');
  const cur = t.rows().find(r => r.IncomeID === id);
  if (!cur) psErr_('Income not found.');
  if (cur.Status === 'Voided') psErr_('Already voided.');
  t.update(cur._row, { Status: 'Voided', Description: (cur.Description ? cur.Description + ' | ' : '') + 'VOID: ' + reason });
  psTxn_(cur.AccountID, 'Out', psNum_(cur.Amount), 'Adjustment', id, 'VOID income ' + id + ': ' + reason, psToday_(), s.name);
  psAuditS_(s, 'Income', 'Voided', id, cur.Amount, 0, reason);
  return 'Income voided.';
}

// ── expenditure ──
function fin_expenses_(s, p) {
  p = p || {};
  return psTable_('Expenditure').rows().filter(r => psInRange_(r.Date, p.from, p.to) && (!p.category || r.Category === p.category) && (!p.status || r.Status === p.status))
    .map(r => ({ ExpenseID: r.ExpenseID, Date: r.Date, Category: r.Category, Description: r.Description, Amount: psNum_(r.Amount), Payee: r.Payee, SupplierID: r.SupplierID,
                 Method: r.Method, AccountID: r.AccountID, VoucherNo: r.VoucherNo, Reference: r.Reference, Status: r.Status, RequestedBy: r.RequestedBy, ApprovedBy: r.ApprovedBy }))
    .sort((a, b) => String(b.Date).localeCompare(String(a.Date)));
}
function fin_expenseSave_(s, p) {
  p = p || {};
  const amt = psRound_(psNum_(p.Amount));
  if (amt <= 0) psErr_('Amount must be greater than zero.');
  if (!String(p.Category || '').trim()) psErr_('Choose a category.');
  if (!String(p.Description || '').trim()) psErr_('Describe the expense.');
  const method = PS_METHODS.indexOf(p.Method) >= 0 ? p.Method : 'Cash';
  const cat = psTable_('FinanceCategories').rows().find(c => c.Type === 'Expense' && c.Name === p.Category);
  const status = p.Status === 'Pending' ? 'Pending' : 'Paid';
  const id = psId_('EXP'), date = p.Date || psToday_();
  const rec = { ExpenseID: id, Date: date, CategoryID: cat ? cat.CategoryID : '', Category: p.Category, Description: p.Description, Amount: amt, Payee: p.Payee || '',
                SupplierID: p.SupplierID || '', Method: method, AccountID: '', VoucherNo: '', Reference: p.Reference || '', Status: status, RequestedBy: s.name,
                ApprovedBy: '', ApprovedAt: '', PaidAt: '', Attachment: '', CreatedAt: psNow_() };
  if (status === 'Paid') {
    rec.AccountID = psPickAccount_(p.AccountID, method);
    rec.VoucherNo = psNextNumbers_('NextVoucherNo', 'VoucherPrefix', 1)[0];
    rec.ApprovedBy = s.name; rec.ApprovedAt = psNow_(); rec.PaidAt = psNow_();
    psTxn_(rec.AccountID, 'Out', amt, 'Expense', id, p.Category + ': ' + p.Description, date, s.name);
  }
  psTable_('Expenditure').append(rec);
  psAuditS_(s, 'Expenditure', status === 'Paid' ? 'Paid' : 'Requested', id, '', amt, p.Category + ' — ' + p.Description);
  return status === 'Paid' ? 'Expense recorded and paid (voucher ' + rec.VoucherNo + ').' : 'Expense saved as Pending.';
}
function fin_expensePay_(s, id, accountId, method) {
  const t = psTable_('Expenditure');
  const cur = t.rows().find(r => r.ExpenseID === id);
  if (!cur) psErr_('Expense not found.');
  if (cur.Status !== 'Pending') psErr_('Only Pending expenses can be paid.');
  method = PS_METHODS.indexOf(method) >= 0 ? method : (cur.Method || 'Cash');
  const acc = psPickAccount_(accountId, method), voucher = psNextNumbers_('NextVoucherNo', 'VoucherPrefix', 1)[0];
  t.update(cur._row, { Status: 'Paid', AccountID: acc, Method: method, VoucherNo: voucher, ApprovedBy: s.name, ApprovedAt: psNow_(), PaidAt: psNow_() });
  psTxn_(acc, 'Out', psNum_(cur.Amount), 'Expense', id, cur.Category + ': ' + cur.Description, psToday_(), s.name);
  psAuditS_(s, 'Expenditure', 'Paid', id, 'Pending', 'Paid', voucher);
  return 'Paid — voucher ' + voucher + '.';
}
function fin_expenseReverse_(s, id, reason) {
  reason = String(reason || '').trim();
  if (reason.length < 3) psErr_('Give a reason.');
  const t = psTable_('Expenditure');
  const cur = t.rows().find(r => r.ExpenseID === id);
  if (!cur) psErr_('Expense not found.');
  if (cur.Status === 'Paid') {
    psTxn_(cur.AccountID, 'In', psNum_(cur.Amount), 'Adjustment', id, 'REVERSE expense ' + id + ': ' + reason, psToday_(), s.name);
  } else if (cur.Status !== 'Pending') psErr_('This expense is already ' + cur.Status + '.');
  t.update(cur._row, { Status: 'Rejected', Reference: (cur.Reference ? cur.Reference + ' | ' : '') + 'REVERSED: ' + reason });
  psAuditS_(s, 'Expenditure', 'Reversed', id, cur.Status, 'Rejected', reason);
  return 'Expense reversed.';
}
function fin_expenseDelete_(s, id) {
  const t = psTable_('Expenditure');
  const cur = t.rows().find(r => r.ExpenseID === id);
  if (!cur) psErr_('Expense not found.');
  if (cur.Status === 'Paid') psErr_('A paid expense cannot be deleted — reverse it instead.');
  t.remove(cur._row);
  psAuditS_(s, 'Expenditure', 'Deleted', id, cur.Amount, '', cur.Description);
  return 'Expense deleted.';
}

// ── categories & settings ──
function fin_catList_(s, type) {
  return psTable_('FinanceCategories').rows().filter(c => !type || c.Type === type).map(c => ({ CategoryID: c.CategoryID, Type: c.Type, Name: c.Name, Description: c.Description, Status: c.Status }));
}
function fin_catSave_(s, p) {
  p = p || {};
  const name = String(p.Name || '').trim();
  if (!name) psErr_('Name is required.');
  if (['Income', 'Expense'].indexOf(p.Type) < 0) psErr_('Type must be Income or Expense.');
  if (s.role === 'schoolmanager' && p.Type !== 'Expense') psErr_('The School Manager manages expense categories only.');
  const t = psTable_('FinanceCategories');
  if (t.rows().some(c => c.Type === p.Type && psNormName_(c.Name) === psNormName_(name) && c.CategoryID !== p.CategoryID)) psErr_('This category already exists.');
  if (p.CategoryID) {
    const cur = t.rows().find(c => c.CategoryID === p.CategoryID);
    if (!cur) psErr_('Category not found.');
    t.update(cur._row, { Name: name, Description: p.Description || '', Status: p.Status || 'Active' });
    return 'Category updated.';
  }
  t.append({ CategoryID: (p.Type === 'Income' ? 'INC' : 'EXP') + Date.now().toString(36).toUpperCase(), Type: p.Type, Name: name, Description: p.Description || '', Status: p.Status || 'Active' });
  psAuditS_(s, 'Categories', 'Created', name, '', '', p.Type);
  return 'Category added.';
}
function fin_catDelete_(s, id) {
  const t = psTable_('FinanceCategories');
  const cur = t.rows().find(c => c.CategoryID === id);
  if (!cur) psErr_('Category not found.');
  if (s.role === 'schoolmanager' && cur.Type !== 'Expense') psErr_('The School Manager manages expense categories only.');
  const used = cur.Type === 'Income' ? psTable_('Income').rows().some(r => r.Category === cur.Name) : psTable_('Expenditure').rows().some(r => r.Category === cur.Name);
  if (used) psErr_('This category is already used in records — set it to Inactive instead.');
  t.remove(cur._row);
  return 'Category deleted.';
}
function fin_settingsList_(s) {
  return psTable_('FinanceSettings').rows().filter(r => String(r.Key).indexOf('School') !== 0).map(r => ({ Key: r.Key, Value: r.Value, Description: r.Description }));
}
function fin_settingsSave_(s, key, value) {
  if (String(key).indexOf('School') === 0) psErr_('School details are managed by the School Manager (School Settings).');
  const t = psTable_('FinanceSettings');
  const cur = t.rows().find(r => String(r.Key).trim() === key);
  if (!cur) psErr_('Setting not found.');
  t.update(cur._row, { Value: value });
  psAuditS_(s, 'Settings', 'Changed', key, cur.Value, value, '');
  return 'Setting saved.';
}

// ── cash & bank ──
function fin_accountsList_(s) { return psBalances_().map(a => ({ AccountID: a.AccountID, AccountName: a.AccountName, Type: a.Type, BankName: a.BankName, AccountNumber: a.AccountNumber, OpeningBalance: psNum_(a.OpeningBalance), Currency: a.Currency, Status: a.Status, Notes: a.Notes, Balance: a.Balance })); }
function fin_accountsSave_(s, p) {
  p = p || {};
  const name = String(p.AccountName || '').trim();
  if (!name) psErr_('Account name is required.');
  if (['Cash', 'Bank', 'Mobile Money'].indexOf(p.Type) < 0) psErr_('Choose the account type.');
  const t = psTable_('CashBankAccounts');
  const rec = { AccountName: name, Type: p.Type, BankName: p.BankName || '', AccountNumber: p.AccountNumber || '', OpeningBalance: psNum_(p.OpeningBalance), Currency: p.Currency || 'TZS', Status: p.Status || 'Active', Notes: p.Notes || '' };
  if (p.AccountID) {
    const cur = t.rows().find(a => a.AccountID === p.AccountID);
    if (!cur) psErr_('Account not found.');
    t.update(cur._row, rec);
    psAuditS_(s, 'CashBank', 'Updated account', p.AccountID, cur.OpeningBalance, rec.OpeningBalance, name);
    return 'Account updated.';
  }
  const n = t.rows().length + 1;
  rec.AccountID = 'ACC-' + String(n).padStart(3, '0');
  while (t.rows().some(a => a.AccountID === rec.AccountID)) rec.AccountID = 'ACC-' + String(Date.now()).slice(-6);
  t.append(rec);
  psAuditS_(s, 'CashBank', 'Created account', rec.AccountID, '', rec.OpeningBalance, name);
  return 'Account added.';
}
function fin_txns_(s, p) {
  p = p || {};
  return psTable_('CashBankTransactions').rows().filter(t => (!p.accountId || t.AccountID === p.accountId) && psInRange_(t.Date, p.from, p.to))
    .map(t => ({ TxnID: t.TxnID, Date: t.Date, AccountID: t.AccountID, Direction: t.Direction, Amount: psNum_(t.Amount), RefType: t.RefType, RefID: t.RefID, Description: t.Description, RecordedBy: t.RecordedBy }))
    .sort((a, b) => String(b.Date).localeCompare(String(a.Date))).slice(0, 2000);
}
function fin_transfer_(s, p) {
  p = p || {};
  const amt = psRound_(psNum_(p.amount));
  if (amt <= 0) psErr_('Amount must be greater than zero.');
  if (!p.from || !p.to || p.from === p.to) psErr_('Choose two different accounts.');
  const bal = psBalances_(), a = bal.find(x => x.AccountID === p.from), b = bal.find(x => x.AccountID === p.to);
  if (!a || !b) psErr_('Account not found.');
  if (a.Balance + 0.0001 < amt) psErr_(a.AccountName + ' only has ' + a.Balance + '.');
  const id = psId_('TRF'), date = p.date || psToday_();
  psTxn_(p.from, 'Out', amt, 'Transfer', id, 'Transfer to ' + b.AccountName + (p.note ? ': ' + p.note : ''), date, s.name, p.to);
  psTxn_(p.to, 'In', amt, 'Transfer', id, 'Transfer from ' + a.AccountName + (p.note ? ': ' + p.note : ''), date, s.name, p.from);
  psAuditS_(s, 'CashBank', 'Transfer', id, a.AccountName, b.AccountName, String(amt));
  return 'Transferred ' + amt + ' from ' + a.AccountName + ' to ' + b.AccountName + '.';
}
function fin_reconList_(s) {
  return psTable_('CashReconciliations').rows().map(r => ({ RecID: r.RecID, Date: r.Date, AccountID: r.AccountID, SystemBalance: psNum_(r.SystemBalance), ActualBalance: psNum_(r.ActualBalance), Difference: psNum_(r.Difference), Notes: r.Notes, ReconciledBy: r.ReconciledBy }))
    .sort((a, b) => String(b.Date).localeCompare(String(a.Date)));
}
function fin_reconSave_(s, p) {
  p = p || {};
  const acc = psBalances_().find(a => a.AccountID === p.AccountID);
  if (!acc) psErr_('Choose an account.');
  if (p.ActualBalance === '' || p.ActualBalance === undefined) psErr_('Enter the actual (counted / statement) balance.');
  const actual = psRound_(psNum_(p.ActualBalance)), diff = psRound_(actual - acc.Balance);
  psTable_('CashReconciliations').append({ RecID: psId_('REC'), Date: p.Date || psToday_(), AccountID: acc.AccountID, SystemBalance: acc.Balance, ActualBalance: actual, Difference: diff, Notes: p.Notes || '', ReconciledBy: s.name, CreatedAt: psNow_() });
  psAuditS_(s, 'CashBank', 'Reconciled', acc.AccountID, acc.Balance, actual, 'Difference ' + diff);
  return diff === 0 ? 'Balanced — no difference.' : 'Saved. Difference: ' + diff + '.';
}

// ── dashboard & reports ──
function fin_dashboard_(s, year) {
  year = String(year || new Date().getFullYear());
  const months = []; for (let m = 1; m <= 12; m++) months.push(year + '-' + String(m).padStart(2, '0'));
  const coll = {}, exp = {}; months.forEach(m => { coll[m] = 0; exp[m] = 0; });
  psTable_('Payments').rows().filter(r => r.Status !== 'Voided').forEach(r => { const k = String(r.PaymentDate).slice(0, 7); if (coll[k] !== undefined) coll[k] += psNum_(r.Amount); });
  psTable_('Income').rows().filter(r => r.Status !== 'Voided').forEach(r => { const k = String(r.Date).slice(0, 7); if (coll[k] !== undefined) coll[k] += psNum_(r.Amount); });
  const byCat = {};
  psTable_('Expenditure').rows().filter(r => r.Status === 'Paid').forEach(r => {
    const k = String(r.Date).slice(0, 7); if (exp[k] !== undefined) { exp[k] += psNum_(r.Amount); byCat[r.Category] = (byCat[r.Category] || 0) + psNum_(r.Amount); }
  });
  const classes = {};
  psTable_('FeeInvoices').rows().filter(r => r.Status !== 'Cancelled' && String(r.AcademicYear) === year).forEach(r => {
    const o = classes[r.Class] = classes[r.Class] || { Class: r.Class, Billed: 0, Paid: 0 };
    o.Billed += psNum_(r.NetAmount); o.Paid += psNum_(r.Paid);
  });
  const methods = {};
  psTable_('Payments').rows().filter(r => r.Status !== 'Voided' && String(r.PaymentDate).slice(0, 4) === year).forEach(r => { methods[r.Method] = (methods[r.Method] || 0) + psNum_(r.Amount); });
  const low = psTable_('StockItems').rows().filter(i => i.Status !== 'Inactive' && psNum_(i.CurrentQty) <= psNum_(i.ReorderLevel)).map(i => ({ ItemName: i.ItemName, CurrentQty: psNum_(i.CurrentQty), ReorderLevel: psNum_(i.ReorderLevel), Unit: i.Unit }));
  return {
    year: year, totals: fin_totals_(), balances: psBalances_().map(a => ({ AccountName: a.AccountName, Type: a.Type, Balance: a.Balance })),
    months: months.map(m => ({ Month: m, In: psRound_(coll[m]), Out: psRound_(exp[m]) })),
    expenseByCategory: Object.keys(byCat).map(k => ({ Category: k, Amount: psRound_(byCat[k]) })).sort((a, b) => b.Amount - a.Amount),
    byClass: Object.keys(classes).map(k => ({ Class: k, Billed: psRound_(classes[k].Billed), Paid: psRound_(classes[k].Paid) })),
    byMethod: Object.keys(methods).map(k => ({ Method: k, Amount: psRound_(methods[k]) })),
    lowStock: low,
    pendingExpenses: psTable_('Expenditure').rows().filter(r => r.Status === 'Pending').length
  };
}
function fin_report_(s, p) {
  p = p || {};
  const from = p.from || (new Date().getFullYear() + '-01-01'), to = p.to || psToday_();
  const pays = psTable_('Payments').rows().filter(r => r.Status !== 'Voided' && psInRange_(r.PaymentDate, from, to));
  const sum = (arr, k) => psRound_(arr.reduce((a, r) => a + psNum_(r[k]), 0));
  const group = (arr, key, amt) => { const o = {}; arr.forEach(r => { o[r[key] || '—'] = (o[r[key] || '—'] || 0) + psNum_(r[amt]); }); return Object.keys(o).map(k => ({ Name: k, Amount: psRound_(o[k]) })).sort((a, b) => b.Amount - a.Amount); };
  const inc = psTable_('Income').rows().filter(r => r.Status !== 'Voided' && psInRange_(r.Date, from, to));
  const exp = psTable_('Expenditure').rows().filter(r => r.Status === 'Paid' && psInRange_(r.Date, from, to));
  const feesTotal = sum(pays, 'Amount'), incTotal = sum(inc, 'Amount'), expTotal = sum(exp, 'Amount');
  return {
    from: from, to: to, feesCollected: feesTotal, otherIncome: incTotal, totalIncome: psRound_(feesTotal + incTotal), totalExpenses: expTotal, net: psRound_(feesTotal + incTotal - expTotal),
    feesByFee: group(pays, 'FeeName', 'Amount'), feesByClass: group(pays, 'Class', 'Amount'), feesByMethod: group(pays, 'Method', 'Amount'),
    incomeByCategory: group(inc, 'Category', 'Amount'), expenseByCategory: group(exp, 'Category', 'Amount'),
    outstanding: fin_totals_().outstanding, debtors: fin_debtors_(s, {}).length, balances: psBalances_().map(a => ({ AccountName: a.AccountName, Type: a.Type, Balance: a.Balance }))
  };
}

// ── procurement & stock ──
function fin_supList_(s) { return psTable_('Suppliers').rows().map(r => ({ SupplierID: r.SupplierID, Name: r.Name, ContactPerson: r.ContactPerson, Phone: r.Phone, Email: r.Email, Address: r.Address, TIN: r.TIN, Category: r.Category, Status: r.Status })); }
function fin_supSave_(s, p) {
  p = p || {};
  if (!String(p.Name || '').trim()) psErr_('Supplier name is required.');
  const t = psTable_('Suppliers');
  const rec = { Name: String(p.Name).trim(), ContactPerson: p.ContactPerson || '', Phone: p.Phone || '', Email: p.Email || '', Address: p.Address || '', TIN: p.TIN || '', Category: p.Category || '', Status: p.Status || 'Active' };
  if (p.SupplierID) { const cur = t.rows().find(r => r.SupplierID === p.SupplierID); if (!cur) psErr_('Supplier not found.'); t.update(cur._row, rec); return 'Supplier updated.'; }
  rec.SupplierID = psId_('SUP'); t.append(rec); return 'Supplier added.';
}
function fin_supDelete_(s, id) {
  const t = psTable_('Suppliers'); const cur = t.rows().find(r => r.SupplierID === id);
  if (!cur) psErr_('Supplier not found.');
  if (psTable_('ProcurementRequests').rows().some(r => r.SupplierID === id) || psTable_('Expenditure').rows().some(r => r.SupplierID === id)) psErr_('This supplier is used in records — set it to Inactive instead.');
  t.remove(cur._row); return 'Supplier deleted.';
}
function fin_reqList_(s, p) {
  p = p || {};
  const sup = {}; psTable_('Suppliers').rows().forEach(x => { sup[x.SupplierID] = x.Name; });
  return psTable_('ProcurementRequests').rows().filter(r => !p.status || r.Status === p.status).map(r => ({
    RequestID: r.RequestID, Date: r.Date, RequestedBy: r.RequestedBy, Department: r.Department, ItemName: r.ItemName, Quantity: psNum_(r.Quantity), Unit: r.Unit,
    EstUnitCost: psNum_(r.EstUnitCost), EstTotal: psNum_(r.EstTotal), Status: r.Status, ApprovedBy: r.ApprovedBy, SupplierID: r.SupplierID, Supplier: sup[r.SupplierID] || '',
    PONumber: r.PONumber, ReceivedDate: r.ReceivedDate, Notes: r.Notes })).sort((a, b) => String(b.Date).localeCompare(String(a.Date)));
}
function fin_reqSave_(s, p) {
  p = p || {};
  const qty = psNum_(p.Quantity), cost = psNum_(p.EstUnitCost);
  if (!String(p.ItemName || '').trim() || qty <= 0) psErr_('Item name and a quantity are required.');
  const t = psTable_('ProcurementRequests');
  const rec = { Department: p.Department || '', ItemName: String(p.ItemName).trim(), Quantity: qty, Unit: p.Unit || 'pcs', EstUnitCost: cost, EstTotal: psRound_(qty * cost), SupplierID: p.SupplierID || '', Notes: p.Notes || '' };
  if (p.RequestID) {
    const cur = t.rows().find(r => r.RequestID === p.RequestID);
    if (!cur) psErr_('Request not found.');
    if (['Received', 'Rejected'].indexOf(cur.Status) >= 0) psErr_('A ' + cur.Status + ' request cannot be edited.');
    t.update(cur._row, rec); return 'Request updated.';
  }
  rec.RequestID = psId_('REQ'); rec.Date = psToday_(); rec.RequestedBy = p.RequestedBy || s.name; rec.Status = 'Pending';
  t.append(rec);
  psAuditS_(s, 'Procurement', 'Request created', rec.RequestID, '', rec.EstTotal, rec.ItemName);
  return 'Request saved.';
}
function fin_reqStatus_(s, id, status, po) {
  if (['Approved', 'Ordered', 'Rejected'].indexOf(status) < 0) psErr_('Invalid status.');
  const t = psTable_('ProcurementRequests'); const cur = t.rows().find(r => r.RequestID === id);
  if (!cur) psErr_('Request not found.');
  if (['Received', 'Rejected'].indexOf(cur.Status) >= 0) psErr_('This request is already ' + cur.Status + '.');
  const upd = { Status: status };
  if (status === 'Approved') { upd.ApprovedBy = s.name; upd.ApprovedAt = psNow_(); }
  if (status === 'Ordered' && po) upd.PONumber = po;
  t.update(cur._row, upd);
  psAuditS_(s, 'Procurement', status, id, cur.Status, status, cur.ItemName);
  return 'Request ' + status.toLowerCase() + '.';
}
function fin_reqReceive_(s, id) {
  const t = psTable_('ProcurementRequests'); const cur = t.rows().find(r => r.RequestID === id);
  if (!cur) psErr_('Request not found.');
  if (cur.Status === 'Received') psErr_('Already received.');
  if (['Approved', 'Ordered'].indexOf(cur.Status) < 0) psErr_('Approve the request before receiving the goods.');
  const items = psTable_('StockItems');
  let item = items.rows().find(i => psNormName_(i.ItemName) === psNormName_(cur.ItemName));
  if (!item) {
    const rec = { ItemID: psId_('ITM'), ItemName: cur.ItemName, Category: cur.Department || 'General', Unit: cur.Unit || 'pcs', ReorderLevel: 0, CurrentQty: 0, UnitCost: psNum_(cur.EstUnitCost), Location: '', Status: 'Active' };
    items.append(rec); item = items.rows().find(i => i.ItemID === rec.ItemID);
  }
  fin_stockMove_(s, { itemId: item.ItemID, type: 'In', quantity: psNum_(cur.Quantity), unitCost: psNum_(cur.EstUnitCost), reference: cur.RequestID, notes: 'Received from request ' + cur.RequestID });
  t.update(cur._row, { Status: 'Received', ReceivedDate: psToday_() });
  return 'Goods received and added to stock. Record the payment under Expenditure if not paid yet.';
}
function fin_stockItems_(s) {
  return psTable_('StockItems').rows().map(i => ({ ItemID: i.ItemID, ItemName: i.ItemName, Category: i.Category, Unit: i.Unit, ReorderLevel: psNum_(i.ReorderLevel), CurrentQty: psNum_(i.CurrentQty),
    UnitCost: psNum_(i.UnitCost), Location: i.Location, Status: i.Status, Low: psNum_(i.CurrentQty) <= psNum_(i.ReorderLevel) }));
}
function fin_stockSaveItem_(s, p) {
  p = p || {};
  if (!String(p.ItemName || '').trim()) psErr_('Item name is required.');
  const t = psTable_('StockItems');
  const rec = { ItemName: String(p.ItemName).trim(), Category: p.Category || '', Unit: p.Unit || 'pcs', ReorderLevel: psNum_(p.ReorderLevel), UnitCost: psNum_(p.UnitCost), Location: p.Location || '', Status: p.Status || 'Active' };
  if (p.ItemID) { const cur = t.rows().find(i => i.ItemID === p.ItemID); if (!cur) psErr_('Item not found.'); t.update(cur._row, rec); return 'Item updated.'; }
  if (t.rows().some(i => psNormName_(i.ItemName) === psNormName_(rec.ItemName))) psErr_('This item already exists.');
  rec.ItemID = psId_('ITM'); rec.CurrentQty = psNum_(p.CurrentQty);
  t.append(rec);
  if (rec.CurrentQty > 0) psTable_('StockMovements').append({ MovementID: psId_('MV'), Date: psToday_(), ItemID: rec.ItemID, ItemName: rec.ItemName, Type: 'In', Quantity: rec.CurrentQty, UnitCost: rec.UnitCost, Reference: 'Opening stock', IssuedTo: '', RecordedBy: s.name, Notes: '' });
  return 'Item added.';
}
function fin_stockMove_(s, p) {
  p = p || {};
  const qty = psNum_(p.quantity);
  if (['In', 'Out', 'Adjustment'].indexOf(p.type) < 0) psErr_('Invalid movement type.');
  if (qty <= 0 && p.type !== 'Adjustment') psErr_('Quantity must be greater than zero.');
  const t = psTable_('StockItems'); const item = t.rows().find(i => i.ItemID === p.itemId);
  if (!item) psErr_('Item not found.');
  const cur = psNum_(item.CurrentQty);
  let next = cur;
  if (p.type === 'In') next = cur + qty;
  else if (p.type === 'Out') { if (qty > cur) psErr_('Only ' + cur + ' ' + item.Unit + ' in stock.'); next = cur - qty; }
  else next = qty;                                           // Adjustment = set counted quantity
  t.update(item._row, { CurrentQty: psRound_(next), UnitCost: p.type === 'In' && psNum_(p.unitCost) > 0 ? psNum_(p.unitCost) : psNum_(item.UnitCost) });
  psTable_('StockMovements').append({ MovementID: psId_('MV'), Date: p.date || psToday_(), ItemID: item.ItemID, ItemName: item.ItemName, Type: p.type,
    Quantity: p.type === 'Adjustment' ? psRound_(next - cur) : qty, UnitCost: psNum_(p.unitCost) || psNum_(item.UnitCost), Reference: p.reference || '', IssuedTo: p.issuedTo || '', RecordedBy: s.name, Notes: p.notes || '' });
  psAuditS_(s, 'Stock', p.type, item.ItemID, cur, next, item.ItemName);
  return 'Stock updated: ' + item.ItemName + ' = ' + psRound_(next) + ' ' + item.Unit + '.';
}
function fin_stockMovements_(s, p) {
  p = p || {};
  return psTable_('StockMovements').rows().filter(m => (!p.itemId || m.ItemID === p.itemId) && psInRange_(m.Date, p.from, p.to))
    .map(m => ({ MovementID: m.MovementID, Date: m.Date, ItemName: m.ItemName, Type: m.Type, Quantity: psNum_(m.Quantity), UnitCost: psNum_(m.UnitCost), Reference: m.Reference, IssuedTo: m.IssuedTo, RecordedBy: m.RecordedBy, Notes: m.Notes }))
    .sort((a, b) => String(b.Date).localeCompare(String(a.Date))).slice(0, 1500);
}

// ── control & audit ──
function fin_auditList_(s, p) {
  p = p || {};
  const q = psNormName_(p.q);
  return psTable_('AuditLog').rows().filter(r => (!p.module || r.Module === p.module) && psInRange_(r.Timestamp, p.from, p.to) &&
    (!q || psNormName_([r.Username, r.Action, r.RecordID, r.Details].join(' ')).indexOf(q) >= 0))
    .map(r => ({ Timestamp: r.Timestamp, Username: r.Username, Role: r.Role, Module: r.Module, Action: r.Action, RecordID: r.RecordID, OldValue: r.OldValue, NewValue: r.NewValue, Details: r.Details }))
    .reverse().slice(0, 1500);
}
function fin_auditChecks_(s) {
  const out = [];
  const inv = psTable_('FeeInvoices').rows().filter(r => r.Status !== 'Cancelled');
  const bad = inv.filter(r => Math.abs(psNum_(r.NetAmount) - psNum_(r.Paid) - psNum_(r.Balance)) > 0.01);
  out.push({ Check: 'Invoice lines where Net − Paid ≠ Balance', Count: bad.length, Ok: bad.length === 0, Detail: bad.slice(0, 5).map(r => r.InvoiceNo + ' ' + r.FeeName).join(', ') });
  const over = inv.filter(r => psNum_(r.Paid) > psNum_(r.NetAmount) + 0.01);
  out.push({ Check: 'Invoice lines paid more than billed', Count: over.length, Ok: over.length === 0, Detail: over.slice(0, 5).map(r => r.InvoiceNo + ' ' + r.FeeName).join(', ') });
  const pays = psTable_('Payments').rows().filter(r => r.Status !== 'Voided');
  const payTotal = psRound_(pays.reduce((a, r) => a + psNum_(r.Amount), 0));
  const invPaid = psRound_(inv.reduce((a, r) => a + psNum_(r.Paid), 0));
  out.push({ Check: 'Valid payments total equals amount paid on invoices', Count: Math.abs(payTotal - invPaid) > 0.01 ? 1 : 0, Ok: Math.abs(payTotal - invPaid) <= 0.01, Detail: 'Payments ' + payTotal + ' vs invoices ' + invPaid });
  const tx = psTable_('CashBankTransactions').rows();
  const txPay = psRound_(tx.filter(t => t.RefType === 'Payment').reduce((a, t) => a + psNum_(t.Amount), 0));
  const voidedTotal = psRound_(psTable_('Payments').rows().filter(r => r.Status === 'Voided').reduce((a, r) => a + psNum_(r.Amount), 0));
  out.push({ Check: 'Cash/bank receipts match payments (valid + voided)', Count: Math.abs(txPay - (payTotal + voidedTotal)) > 0.01 ? 1 : 0, Ok: Math.abs(txPay - (payTotal + voidedTotal)) <= 0.01, Detail: 'Cash/bank ' + txPay + ' vs payments ' + psRound_(payTotal + voidedTotal) });
  const noAcc = pays.filter(r => !r.AccountID);
  out.push({ Check: 'Payments without a cash/bank account', Count: noAcc.length, Ok: noAcc.length === 0, Detail: '' });
  const pend = psTable_('Expenditure').rows().filter(r => r.Status === 'Pending');
  out.push({ Check: 'Expenses still Pending (not paid)', Count: pend.length, Ok: pend.length === 0, Detail: psRound_(pend.reduce((a, r) => a + psNum_(r.Amount), 0)) + ' total' });
  const neg = psBalances_().filter(a => a.Balance < 0);
  out.push({ Check: 'Accounts with a negative balance', Count: neg.length, Ok: neg.length === 0, Detail: neg.map(a => a.AccountName).join(', ') });
  return out;
}


// ═════════════════════ SCHOOL PROFILE, BRANDING, TERMS ═════════════════════
const PS_PLACEHOLDER_NAME = 'PRIMARY SCHOOL NAME';
const PS_DEFAULT_LOGO = 'data:image/svg+xml;utf8,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><rect width="64" height="64" rx="14" fill="#1e3a8a"/><path d="M32 14 10 26l22 12 22-12z" fill="#fbbf24"/><path d="M18 33v10c0 4 7 8 14 8s14-4 14-8V33l-14 7z" fill="#fff"/></svg>');

function psSettingSet_(key, value, desc) {
  const t = psTable_('FinanceSettings');
  const cur = t.rows().find(r => String(r.Key).trim() === key);
  if (cur) t.update(cur._row, { Value: value }); else t.append({ Key: key, Value: value, Description: desc || '' });
}

// name / logo of the school (cached 5 min) — used to brand the public pages
function psBranding_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('PS_BRAND');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const st = psSettings_();
  const name = String(st.SchoolName || '').trim();
  const set = !!name && name !== PS_PLACEHOLDER_NAME;
  const b = { set: set, name: set ? name : 'School Management System',
              short: set ? (name.replace(/\s*(primary|secondary)?\s*school\s*$/i, '').trim() || name) : 'our school',
              logo: String(st.SchoolLogo || ''), motto: String(st.SchoolMotto || '') };
  try { cache.put('PS_BRAND', JSON.stringify(b), 300); } catch (e) {}
  return b;
}

// The landing and login pages fill in the school's name and logo themselves (getPublicBranding).
function psBrand_(html) { return String(html || ''); }

function api_schoolSave_(s, p) {
  p = p || {};
  const name = String(p.name || '').trim();
  if (name.length < 3) psErr_('Enter the full name of the school.');
  psSettingSet_('SchoolName', name, 'School name — printed on receipts, reports and login page');
  psSettingSet_('SchoolCode', String(p.code || '').trim(), 'School code / registration number');
  psSettingSet_('SchoolCouncil', String(p.council || '').trim(), 'District / Town Council');
  psSettingSet_('SchoolAddress', String(p.address || '').trim(), 'Postal address');
  psSettingSet_('SchoolPhone', String(p.phone || '').trim(), 'Phone');
  psSettingSet_('SchoolEmail', String(p.email || '').trim(), 'Email');
  psSettingSet_('SchoolMotto', String(p.motto || '').trim(), 'School motto');
  psSettingSet_('ReportHeader1', String(p.header1 || '').trim().slice(0, 80), 'Report heading, line 1 (e.g. country)');
  psSettingSet_('ReportHeader2', String(p.header2 || '').trim().slice(0, 80), 'Report heading, line 2 (e.g. ministry / authority)');
  if (p.removeLogo) psSettingSet_('SchoolLogo', '', 'School logo (image data)');
  else if (p.logo) {
    if (!/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+\/=]+$/.test(p.logo)) psErr_('The logo must be a JPEG, PNG or WebP image.');
    if (p.logo.length > 45000) psErr_('The logo is too large. Use a smaller image.');
    psSettingSet_('SchoolLogo', p.logo, 'School logo (image data)');
  }
  CacheService.getScriptCache().remove('PS_BRAND');
  psAuditS_(s, 'School', 'Profile saved', '', '', name, '');
  return 'School details saved.';
}

function api_termsList_(s) {
  return psTable_('AcademicTerms').rows().map(r => ({ TermID: r.TermID, AcademicYear: String(r.AcademicYear), Term: r.Term, StartDate: r.StartDate, EndDate: r.EndDate, Status: r.Status }))
    .sort((a, b) => String(b.AcademicYear).localeCompare(String(a.AcademicYear)) || String(a.Term).localeCompare(String(b.Term)));
}
function api_termsSave_(s, p) {
  p = p || {};
  const year = String(p.AcademicYear || '').trim();
  if (!/^\d{4}$/.test(year)) psErr_('Academic year must be 4 digits, e.g. 2026.');
  if (['Term 1', 'Term 2'].indexOf(p.Term) < 0) psErr_('Choose Term 1 or Term 2.');
  if (p.StartDate && p.EndDate && String(p.StartDate) > String(p.EndDate)) psErr_('The start date must be before the end date.');
  const status = ['Active', 'Closed', 'Upcoming'].indexOf(p.Status) >= 0 ? p.Status : 'Upcoming';
  const t = psTable_('AcademicTerms'), rows = t.rows();
  if (rows.some(r => String(r.AcademicYear) === year && r.Term === p.Term && r.TermID !== p.TermID)) psErr_(p.Term + ' ' + year + ' already exists.');
  const rec = { AcademicYear: year, Term: p.Term, StartDate: p.StartDate || '', EndDate: p.EndDate || '', Status: status };
  let id = p.TermID;
  if (id) {
    const cur = rows.find(r => r.TermID === id);
    if (!cur) psErr_('Term not found.');
    t.update(cur._row, rec);
  } else { id = 'T' + year + '-' + p.Term.slice(-1); rec.TermID = id; t.append(rec); }
  if (status === 'Active') t.rows().forEach(r => { if (r.TermID !== id && r.Status === 'Active') t.update(r._row, { Status: 'Closed' }); });
  psAuditS_(s, 'School', 'Term saved', id, '', status, p.Term + ' ' + year);
  return 'Term saved.' + (status === 'Active' ? ' It is now the current term for fees and invoices.' : '');
}
