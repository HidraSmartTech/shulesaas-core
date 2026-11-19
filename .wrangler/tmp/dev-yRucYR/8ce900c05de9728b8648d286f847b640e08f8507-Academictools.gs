/**
 * AcademicTools.gs
 *  1) Academic Settings  : report logos, Head of School / Academic Master names + signatures
 *                          (used by Publish Exam Results, Individual Report Cards and TOD Report)
 *  2) Student Excel import (uses the existing saveStudent(), so all existing rules still apply)
 */
var ACS_KEYS_ = { logoLeft: 'img', logoRight: 'img', schoolLogo: 'img', headSign: 'img', academicSign: 'img', headName: 'text', academicName: 'text',
                  admPrefix: 'adm', admSep: 'adm', admDigits: 'adm', admStart: 'adm' };
var ACS_CHUNK_ = 45000;      // a Sheet cell holds max 50,000 chars; images are stored in chunks
var ACS_MAX_IMG_ = 190000;   // same limit the Academic Settings panel enforces

function acsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('AcademicSettings');
  if (!sh) { sh = ss.insertSheet('AcademicSettings'); sh.appendRow(['Key', 'Part', 'Value']); sh.setFrozenRows(1); }
  return sh;
}

/** Internal reader (no session check) - used by getAcademicSettings() and by the Parent portal data. */
function acsReadAll_() {
  var sh = acsSheet_(), last = sh.getLastRow(), parts = {}, out = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 3).getValues().forEach(function (r) {
    var k = String(r[0]); if (!ACS_KEYS_[k]) return;
    (parts[k] = parts[k] || []).push([Number(r[1]) || 0, String(r[2])]);
  });
  Object.keys(parts).forEach(function (k) {
    out[k] = parts[k].sort(function (a, b) { return a[0] - b[0]; }).map(function (p) { return p[1]; }).join('');
  });
  var info = {}, cs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SchoolConfig');
  if (cs && cs.getLastRow() >= 2) cs.getRange(2, 1, cs.getLastRow() - 1, 2).getValues().forEach(function (r) { if (r[0]) info[String(r[0])] = String(r[1]); });
  out.schoolName = info.name || '';
  out.schoolInfo = info;
  out.schoolLogo = out.schoolLogo || info.logo || '';   // uploaded logo (older versions stored a link)
  return out;
}

/** Returns {logoLeft, logoRight, headSign, headName, academicSign, academicName, schoolName, schoolInfo} */
function getAcademicSettings() {
  var s = getTeacherSession();
  if (!s || !s.username) throw new Error('Session expired. Please log in again.');
  return acsReadAll_();
}

/** payload: any of the keys in ACS_KEYS_. Empty string removes a value (reports fall back to the default). */
function saveAcademicSettings(payload) {
  try { acadProfileSession_(); } catch (e) { return { ok: false, message: e.message }; }
  if (!payload || typeof payload !== 'object') return { ok: false, message: 'Nothing to save.' };
  var keys = Object.keys(payload).filter(function (k) { return ACS_KEYS_[k]; });
  if (!keys.length) return { ok: false, message: 'Nothing to save.' };
  var clean = {};
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i], v = String(payload[k] == null ? '' : payload[k]), kind = ACS_KEYS_[k];
    if (kind === 'img') {
      if (v && (v.length > ACS_MAX_IMG_ || !/^data:image\/(png|jpe?g|gif|webp);base64,[A-Za-z0-9+\/=]+$/.test(v)))
        return { ok: false, message: 'Invalid or too large image for ' + k + '.' };
    } else if (kind === 'adm') {
      v = v.trim();
      if (k === 'admPrefix' && !/^[A-Za-z0-9]{0,12}$/.test(v)) return { ok: false, message: 'Prefix can have letters and numbers only (max 12).' };
      if (k === 'admSep' && ['-', '/', '.', '_', 'none'].indexOf(v) < 0) return { ok: false, message: 'Invalid separator.' };
      if (k === 'admDigits' && !(/^\d+$/.test(v) && +v >= 1 && +v <= 8)) return { ok: false, message: 'Number of digits must be 1 to 8.' };
      if (k === 'admStart' && !(/^\d+$/.test(v) && +v >= 1 && +v <= 99999999)) return { ok: false, message: 'Start number must be 1 or more.' };
    } else { v = v.trim().slice(0, 120); }
    clean[k] = v;
  }
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = acsSheet_(), data = sh.getDataRange().getValues();
    for (var r = data.length - 1; r >= 1; r--) if (clean.hasOwnProperty(String(data[r][0]))) sh.deleteRow(r + 1);
    var rows = [];
    Object.keys(clean).forEach(function (k) {
      var v = clean[k]; if (!v) return;
      for (var p = 0, n = 0; p < v.length; p += ACS_CHUNK_, n++) rows.push([k, n, v.substr(p, ACS_CHUNK_)]);
    });
    if (rows.length) {
      var rng = sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3);
      rng.setNumberFormat('@'); rng.setValues(rows);
    }
  } finally { lock.releaseLock(); }
  rbacAdmFormatCache_ = null;
  try { logActivity('academic', 'Updated Academic Settings: ' + keys.join(', ')); } catch (e) {}
  return { ok: true, message: 'Saved successfully.' };
}

/** Excel import: rows = [{row, firstName, lastName, sex, dob, className, parent, phone, address, admNo}] (max 40 per call) */
function acadImportStudents(rows) {
  acadProfileSession_();
  if (rbacFeatures_('academics')['Student Admissions'] === false) throw new Error('Not enabled by Admin: Student Admissions');
  rows = rows || [];
  if (rows.length > 40) throw new Error('Send at most 40 rows per request.');
  var ok = 0, errors = [];
  rows.forEach(function (r) {
    var msg;
    try {
      msg = saveStudent({ firstName: String(r.firstName || '').trim(), lastName: String(r.lastName || '').trim(), sex: r.sex, dob: r.dob || '',
        className: r.className, parent: r.parent || '', phone: r.phone || '', address: r.address || '', admNo: String(r.admNo || '').trim(), oldId: '' });
    } catch (e) { msg = 'Error: ' + e.message; }
    if (/^Error/i.test(String(msg))) errors.push({ row: r.row, message: String(msg).replace(/^Error:\s*/i, '') }); else ok++;
  });
  return { ok: ok, errors: errors };
}


/* ───────── Dynamic school identity (comes ONLY from Settings, nothing is hardcoded) ───────── */
function rbacSchool_() {
  var a = {}; try { a = acsReadAll_(); } catch (e) {}
  var i = a.schoolInfo || {}, name = String(a.schoolName || '').trim(), district = String(i.district || '').trim();
  return {
    name: name, nameUpper: name.toUpperCase(), motto: i.motto || '', region: i.region || '', district: district,
    council: district ? district.toUpperCase() + ' DISTRICT COUNCIL' : '',
    address: i.address || '', phone: i.phone || '', email: i.email || '',
    logoLeft: a.logoLeft || '', logoRight: a.logoRight || '', logoUrl: a.logoRight || a.logoLeft || a.schoolLogo || '',
    headName: a.headName || '', academicName: a.academicName || ''
  };
}
/** School block used by the TOD report API. */
function rbacTodSchool_() {
  var s = rbacSchool_();
  return { councilName: s.council, schoolName: s.nameUpper, motto: s.motto, headName: s.headName, logoUrl: s.logoUrl };
}

/**
 * Replaces the school name / council / logo that the old pages had typed in with the school saved in Settings.
 * Used for every dashboard, the Parent portal, Login and Index, so each school sees its own identity.
 */
var RBAC_OLD_LOGO_ = 'https://i.ibb.co/pj376KSX/Logo-Right.jpg';
var RBAC_BLANK_PIXEL_ = 'data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==';
function rbacBrand_(html) {
  html = String(html || '');
  var sc = rbacSchool_();
  function esc(t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;'); }
  var name = sc.name ? esc(sc.name) : 'School Management System', upper = sc.name ? esc(sc.nameUpper) : 'SCHOOL MANAGEMENT SYSTEM';
  html = html.split('MANGI WINGIA SECONDARY SCHOOL').join(upper)
             .split('Mangi Wingia Secondary School').join(name)
             .split('Mangi Wingia Sec. School').join(name)
             .split('Mangi Wingia SS').join(name)
             .split('Mangi Wingia').join(name)
             .split('ROMBO DISTRICT COUNCIL').join(esc(sc.council))
             .split(RBAC_OLD_LOGO_).join(sc.logoUrl || RBAC_BLANK_PIXEL_);
  return html;
}


/* ───────── Admission No. format (every school has its own) ───────── */
var rbacAdmFormatCache_ = null;
function acsAdmFormat_() {
  if (rbacAdmFormatCache_) return rbacAdmFormatCache_;
  var a = {}; try { a = acsReadAll_(); } catch (e) {}
  var digits = parseInt(a.admDigits, 10), start = parseInt(a.admStart, 10), sep = String(a.admSep || '');
  rbacAdmFormatCache_ = { configured: !!String(a.admDigits || '').trim(), prefix: String(a.admPrefix || ''),
    sep: (sep === 'none') ? '' : sep, digits: digits > 0 ? digits : 3, start: start > 0 ? start : 1 };
  return rbacAdmFormatCache_;
}
function acsAdmBuild_(f, n) {
  var num = String(n); while (num.length < f.digits) num = '0' + num;
  return f.prefix + (f.prefix ? f.sep : '') + num;
}