// ============================================================
// FEATURES added after the move to Neon
//   • School profile      – the one place every report/page takes the school's identity from
//   • Student Excel import – bulk admission with the same rules as saveStudent()
//   • My profile          – real profile save + picture (stored as a small image)
// ============================================================

// ───────────────────────── School profile ─────────────────────────
function featSchoolProfile_() {
  var st = psSettings_();
  var name = String(st.SchoolName || '').trim();
  var set = !!name && name !== PS_PLACEHOLDER_NAME;
  return {
    name: set ? name : '', nameSet: set,
    code: String(st.SchoolCode || ''), council: String(st.SchoolCouncil || ''),
    address: String(st.SchoolAddress || ''), phone: String(st.SchoolPhone || ''),
    email: String(st.SchoolEmail || ''), motto: String(st.SchoolMotto || ''),
    logo: String(st.SchoolLogo || ''),
    header1: String(st.ReportHeader1 || ''), header2: String(st.ReportHeader2 || '')
  };
}

// For any signed-in user (report headers, print-outs, dashboards)
function getSchoolProfile() { return featSchoolProfile_(); }

// For the public landing / login pages: identity + a few headline numbers
function getPublicBranding() {
  var p = featSchoolProfile_();
  var out = { name: p.name, nameSet: p.nameSet, motto: p.motto, logo: p.logo, address: p.address,
              phone: p.phone, email: p.email, counts: {}, year: '' };
  try { out.counts.students = getStudentListRaw().length; } catch (e) {}
  try { out.counts.classes = getAcademicClassesRaw().length || PS_CLASSES.length; } catch (e) {}
  try {
    out.counts.teachers = getTeacherList().filter(function (t) {
      return ['Academic Master', 'Class Teacher', 'Subject Teacher'].indexOf(String(t.position || '').trim()) >= 0;
    }).length;
  } catch (e) {}
  try { out.counts.subjects = getAcademicSubjectsRaw().length; } catch (e) {}
  try { out.counts.parents = psTable_('Parents').rows().length; } catch (e) {}
  try { out.year = String(psActiveYear_() || ''); } catch (e) {}
  return out;
}

// ───────────────────────── Student Excel import ─────────────────────────
function featNormClass_(v) {
  var m = String(v == null ? '' : v).trim().toLowerCase().replace(/\./g, ' ').replace(/\s+/g, ' ').match(/^(?:standard|std|class|darasa(?: la)?)?\s*(\d+)$/);
  if (!m) return '';
  var n = Number(m[1]);
  return (n >= 1 && n <= PS_CLASSES.length) ? PS_CLASSES[n - 1] : '';
}
function featNormSex_(v) {
  var x = String(v == null ? '' : v).trim().toLowerCase();
  if (['m', 'male', 'me', 'boy', 'mvulana', 'kiume'].indexOf(x) >= 0) return 'Male';
  if (['f', 'female', 'ke', 'girl', 'msichana', 'kike'].indexOf(x) >= 0) return 'Female';
  return '';
}
// '' = no date, null = not understood, otherwise yyyy-MM-dd
function featNormDob_(v) {
  var x = String(v == null ? '' : v).trim();
  if (!x) return '';
  var m = x.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/) || x.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  var y, mo, d;
  if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; }
  else {
    m = x.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);   // day first (Tanzania)
    if (!m) return null;
    d = +m[1]; mo = +m[2]; y = +m[3];
  }
  var dt = new Date(Date.UTC(y, mo - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return null;
  var yr = new Date().getFullYear();
  if (y < 1990 || y > yr) return null;
  return y + '-' + String(mo).padStart(2, '0') + '-' + String(d).padStart(2, '0');
}

// rows: [{ firstName, lastName, sex, dob, className, parent, phone, address, premNo }]
// opts: { skipDuplicates: true }
function importStudents(rows, opts) {
  opts = opts || {};
  if (!Array.isArray(rows) || !rows.length) return { ok: false, message: 'There are no students to import.' };
  if (rows.length > 1000) return { ok: false, message: 'Import at most 1,000 students at a time.' };
  var skipDup = opts.skipDuplicates !== false;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = admSeriesConfig_();

  // what is already registered (Prem Nos and name+class pairs)
  var prems = {}, known = {};
  primaryStudentSheetNames_().forEach(function (nm) {
    var sh = ss.getSheetByName(nm);
    if (!sh || sh.getLastRow() < 2) return;
    var w = Math.max(sh.getLastColumn(), 11);
    sh.getRange(2, 1, sh.getLastRow() - 1, w).getValues().forEach(function (r) {
      var prem = String(r[10] || '').trim().toLowerCase();
      if (prem) prems[prem] = true;
      var cls = String(r[5] || '').trim() || '';
      known[(String(r[1] || '').trim() + '|' + String(r[2] || '').trim() + '|' + cls).toLowerCase()] = true;
    });
  });

  var counters = {};                 // class -> next admission number
  function nextAdm(cls) {
    if (!(cls in counters)) {
      var first = admSeriesNextFor_(ss, cls, cfg);
      counters[cls] = parseInt(String(first).split('-').pop(), 10);
    }
    return admSeriesFormat_(cfg, counters[cls]++);
  }

  var results = [], added = 0, skipped = 0, failed = 0;
  rows.forEach(function (raw, i) {
    var line = { row: i + 1 };
    var first = String(raw.firstName == null ? '' : raw.firstName).trim();
    var last  = String(raw.lastName  == null ? '' : raw.lastName).trim();
    var cls   = featNormClass_(raw.className);
    var sex   = featNormSex_(raw.sex);
    var dob   = featNormDob_(raw.dob);
    var prem  = String(raw.premNo == null ? '' : raw.premNo).trim();
    var err = '';
    if (!first || !last) err = 'First name and last name are required.';
    else if (!cls) err = 'Class "' + String(raw.className || '') + '" is not valid (use Standard 1 – Standard ' + PS_CLASSES.length + ').';
    else if (!sex) err = 'Sex must be Male or Female.';
    else if (dob === null) err = 'Date of birth "' + String(raw.dob) + '" is not understood (use dd/mm/yyyy).';
    if (err) { line.status = 'error'; line.message = err; failed++; results.push(line); return; }

    // a student who is already registered is skipped first, so importing the same file twice is harmless
    var key = (first + '|' + last + '|' + cls).toLowerCase();
    if (skipDup && known[key]) { line.status = 'skipped'; line.message = 'Already in the register (same name and class).'; skipped++; results.push(line); return; }
    if (prem && prems[prem.toLowerCase()]) { line.status = 'error'; line.message = 'Prem No ' + prem + ' is already used.'; failed++; results.push(line); return; }

    var sheetName = primaryClassSheetName_(cls);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['AdmNo', 'FirstName', 'LastName', 'Sex', 'DOB', 'Class', 'Parent', 'Phone', 'Address', 'Status', 'PremNo']);
      sheet.setFrozenRows(1);
    }
    admEnsureStudentCols_(sheet);
    var admNo = nextAdm(cls);
    sheet.appendRow([admNo, first, last, sex, dob || '', cls,
                     String(raw.parent == null ? '' : raw.parent).trim(), String(raw.phone == null ? '' : raw.phone).trim(),
                     String(raw.address == null ? '' : raw.address).trim(), '', prem]);
    if (prem) prems[prem.toLowerCase()] = true;
    known[key] = true;
    line.status = 'added'; line.admNo = admNo; line.className = cls; added++;
    results.push(line);
  });

  if (added) { try { logActivity('academic', 'Imported ' + added + ' student(s) from Excel'); } catch (e) {} }
  return { ok: true, added: added, skipped: skipped, failed: failed, results: results };
}

// ───────────────────────── My profile ─────────────────────────
function featMyUserRow_() {
  var t = getTeacherSession();
  if (!t) throw new Error('Please sign in again.');
  var u = psTable_('Users').rows().find(function (r) { return psNormName_(r.Username) === psNormName_(t.username); });
  if (!u) throw new Error('Your account was not found.');
  return u;
}

function getMyProfile() {
  var u = featMyUserRow_();
  return { name: u.Name, username: u.Username, position: u.Position, sex: u.Sex || '', phone: String(u.Phone || ''),
           email: u.Email || '', picUrl: u.PicUrl || '', classes: u.Classes || '', subjects: u.Subjects || '' };
}

function saveMyProfile(p) {
  p = p || {};
  var u = featMyUserRow_();
  var name = String(p.name == null ? u.Name : p.name).trim();
  if (name.length < 3) throw new Error('Enter your full name.');
  var sex = String(p.sex || '').trim();
  if (sex && ['M', 'F', 'Male', 'Female'].indexOf(sex) < 0) throw new Error('Choose Male or Female.');
  var email = String(p.email || '').trim();
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Enter a valid email address.');
  var rec = { Name: name, Sex: sex, Phone: String(p.phone || '').trim(), Email: email };
  if (p.removePicture) rec.PicUrl = '';
  else if (p.picUrl) {
    var pic = String(p.picUrl);
    if (!/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+\/=]+$/.test(pic)) throw new Error('The picture must be a JPEG, PNG or WebP image.');
    if (pic.length > 60000) throw new Error('The picture is too large. Choose a smaller image.');
    rec.PicUrl = pic;
  }
  psTable_('Users').update(u._row, rec);
  try { logActivity(u.Username, 'Updated profile'); } catch (e) {}
  return { message: 'Profile saved.', name: name, picUrl: rec.PicUrl !== undefined ? rec.PicUrl : (u.PicUrl || '') };
}
