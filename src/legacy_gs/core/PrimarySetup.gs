// ============================================================
// PRIMARY SCHOOL MANAGEMENT SYSTEM — TABLE AUTO-GENERATION
// ------------------------------------------------------------
// Creates every Google Sheet (table) the system needs, with headers,
// colours, drop-down validation and starter data.
//
//  • SAFE to run many times: it never deletes or overwrites your data.
//    - missing sheets are created
//    - missing columns are added at the END of an existing sheet
//    - starter data is written only when a sheet is empty
//  • Runs automatically when the web app opens (see psAutoSetup_),
//    and manually from the spreadsheet menu:  School System ▸ Set up / repair all tables
//  • To change the classes of the school edit PS_CLASSES below.
// ============================================================

const PS_SCHEMA_VERSION = '1.2.0';

// Primary school classes (edit if your school also has Pre-Primary / Awali)
const PS_CLASSES = ['Standard 1', 'Standard 2', 'Standard 3', 'Standard 4',
                    'Standard 5', 'Standard 6'];

const PS_METHODS = ['Cash', 'Bank', 'Mobile Money', 'Cheque'];

// header colours per group of tables
const PS_COLORS = {
  people:   '#0d1b36',   // accounts, students, parents
  academic: '#0d3d27',   // classes, subjects, exams, attendance
  finance:  '#92400e',   // fees, payments, income, expenditure, cash & bank
  stock:    '#4c1d95',   // procurement & stock
  system:   '#374151'    // settings, audit, logs
};

// "Standard 3" → "Standard3"  (name of the student sheet of that class)
function psClassSheetName_(className) {
  return String(className || '').replace(/\s+/g, '');
}

// ------------------------------------------------------------
// SCHEMA — one entry per table
//   headers : column names (existing system columns keep their order)
//   valid   : { Column : [allowed values] }   → drop-down in the sheet
//   fmt     : { Column : number format }
//   seed    : starter rows (same order as headers), written once
//   kv      : true → key/value table, only missing keys are added
// ------------------------------------------------------------
function psSchema_() {
  const defs = [];
  const year = new Date().getFullYear();
  const today = new Date();
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const MONEY = '#,##0';
  const DATE  = 'yyyy-mm-dd';
  const add = (name, group, headers, opt) => defs.push(Object.assign({ name: name, group: group, headers: headers }, opt || {}));

  // ═════════ PEOPLE & ACCESS ═════════
  add('Administrators', 'people',
    ['AdminID', 'FullName', 'Username', 'PasswordHash', 'Salt', 'Role', 'Sex', 'Phone', 'Email',
     'Status', 'PicUrl', 'CreatedBy', 'CreatedAt', 'LastLogin'],
    { valid: { Role: ['School Manager', 'Head of School', 'Accountant'], Sex: ['M', 'F'], Status: ['Active', 'Inactive'] } });

  // teachers & other staff (Academic, Class Teacher, Subject Teacher, Sports, Librarian, Lab)
  add('Users', 'people',
    ['Name', 'Username', 'Password', 'Classes', 'Subjects', 'Sex', 'Position', 'Department', 'PicUrl',
     'Phone', 'Email', 'Status'],
    { valid: {
        Position: ['Academic Master', 'Class Teacher', 'Subject Teacher', 'Sports Manager', 'Librarian', 'Lab Technician'],
        Status: ['Active', 'Inactive'] } });

  add('Parents', 'people',
    ['Name', 'Sex', 'Relation', 'Phone', 'Email', 'AdmNo', 'Address', 'Username', 'Password',
     'StudentName', 'ClassName', 'Role']);

  PS_CLASSES.forEach(c => add(psClassSheetName_(c), 'people',
    ['AdmNo', 'FirstName', 'LastName', 'Sex', 'DOB', 'Class', 'Parent', 'Phone', 'Address', 'Status', 'PremNo'],
    { valid: { Sex: ['M', 'F'] } }));

  add('Alumni', 'people', ['Name', 'Sex', 'AdmNo', 'GraduatedYear', 'PreviousClass']);
  add('DashboardAccess', 'system', ['DashboardKey', 'PolicyJSON', 'UpdatedBy', 'UpdatedAt']);
  add('DashboardAccessEvents', 'system', ['Time', 'Panel', 'Action', 'By']);

  // ═════════ ACADEMICS ═════════
  add('AcademicYears', 'academic', ['Year', 'Active', 'ArchivedStudents', 'Created'],
    { seed: [[String(year), 'TRUE', 0, todayStr]] });

  add('AcademicTerms', 'academic', ['TermID', 'AcademicYear', 'Term', 'StartDate', 'EndDate', 'Status'],
    { valid: { Status: ['Active', 'Closed', 'Upcoming'] }, fmt: { StartDate: DATE, EndDate: DATE },
      seed: [
        ['T' + year + '-1', String(year), 'Term 1', year + '-01-01', year + '-06-30', today.getMonth() < 6 ? 'Active' : 'Closed'],
        ['T' + year + '-2', String(year), 'Term 2', year + '-07-01', year + '-12-31', today.getMonth() >= 6 ? 'Active' : 'Upcoming']
      ] });

  add('AcademicClasses', 'academic', ['ID', 'Name', 'Stream', 'Teacher', 'Capacity', 'Students'],
    { seed: PS_CLASSES.map((c, i) => ['STD' + (i + 1), c, '', '', 45, 0]) });

  const upper = PS_CLASSES.slice(2).join(', ');           // Standard 3 … 6
  add('AcademicSubjects', 'academic', ['Name', 'Code', 'Category', 'Department', 'Classes'],
    { valid: { Category: ['Compulsory', 'Optional'] },
      seed: [
        ['Kiswahili',                'KIS', 'Compulsory', 'Languages', 'All'],
        ['English',                  'ENG', 'Compulsory', 'Languages', 'All'],
        ['Mathematics',              'MAT', 'Compulsory', 'Sciences',  'All'],
        ['Science and Technology',   'SCI', 'Compulsory', 'Sciences',  upper],
        ['Social Studies',           'SST', 'Compulsory', 'Humanities', upper],
        ['Civics and Moral Education', 'CME', 'Compulsory', 'Humanities', upper],
        ['Vocational Skills',        'VOC', 'Compulsory', 'Practical', upper]
      ] });

  add('Examinations', 'academic',
    ['ID', 'Name', 'Type', 'Year', 'StartDate', 'EndDate', 'Classes', 'Remarks', 'Status']);

  add('Scores', 'academic',
    ['ExamName', 'Class', 'Subject', 'AdmNo', 'StudentName', 'Sex', 'Score', 'SubmittedBy', 'SubmittedAt']);

  // PSLE style scale — edit freely
  add('GradingScale', 'academic', ['Grade', 'MinScore', 'MaxScore', 'Points', 'RemarkEN', 'RemarkSW'],
    { seed: [
        ['A', 81, 100, 1, 'Excellent',      'Bora Sana'],
        ['B', 61, 80,  2, 'Very Good',      'Vizuri Sana'],
        ['C', 41, 60,  3, 'Good',           'Vizuri'],
        ['D', 21, 40,  4, 'Satisfactory',   'Wastani'],
        ['E', 0,  20,  5, 'Needs Improvement', 'Dhaifu']
      ] });

  add('Attendance', 'academic',
    ['Date', 'Class', 'AdmNo', 'StudentName', 'Sex', 'RowIndex', 'Morning', 'Afternoon', 'Teacher', 'Reason']);
  add('AttendanceSummary', 'academic',
    ['Date', 'Class', 'Total', 'Absent', 'Percentage', 'Reason', 'Teacher']);

  const crit = (typeof BEHAVIOR_CRITERIA !== 'undefined') ? BEHAVIOR_CRITERIA :
    ['Nidhamu', 'Kuwahi', 'Mahudhurio', 'Heshima', 'Uwajibikaji', 'Ushiriki Darasani', 'Ushirikiano', 'Usafi', 'Kujidhibiti', 'Uongozi'];
  add('Behaviors', 'academic',
    ['Class', 'Exam', 'AdmNo', 'StudentName'].concat(crit).concat(['UpdatedBy', 'UpdatedAt', 'Status', 'SubmittedAt']));

  add('Messages', 'system', ['From', 'To', 'Subject', 'Body', 'Date', 'FromName', 'ToName']);

  // ═════════ FINANCE ═════════
  add('FinanceSettings', 'finance', ['Key', 'Value', 'Description'],
    { kv: true, seed: [
        ['SchoolName',          'PRIMARY SCHOOL NAME', 'Printed on receipts and reports'],
        ['SchoolAddress',       'P.O. Box ..., Town', 'Printed on receipts and reports'],
        ['SchoolCouncil',       '',     'District / Town Council'],
        ['SchoolMotto',         '',     'School motto'],
        ['SchoolLogo',          '',     'School logo (image data) — set under School Settings'],
        ['SchoolPhone',         '',     'Printed on receipts'],
        ['SchoolEmail',         '',     'Printed on receipts'],
        ['Currency',            'TZS',  'Currency used in all finance screens'],
        ['ReceiptPrefix',       'RCP',  'Prefix of receipt numbers'],
        ['NextReceiptNo',       1,      'Counter — next receipt number (do not edit by hand)'],
        ['InvoicePrefix',       'INV',  'Prefix of invoice numbers'],
        ['NextInvoiceNo',       1,      'Counter — next invoice number'],
        ['VoucherPrefix',       'PV',   'Prefix of payment vouchers (expenses)'],
        ['NextVoucherNo',       1,      'Counter — next voucher number'],
        ['AllowPartialPayment', 'TRUE', 'Allow parents to pay a fee in instalments'],
        ['ExpenseApprovalLimit', 0,     'Expenses above this amount need Head of School approval (0 = every expense)'],
        ['LateFeePercent',      0,      'Penalty % added to overdue balances (0 = none)'],
        ['LowStockAlert',       'TRUE', 'Warn when stock falls to its reorder level'],
        ['ReceiptFooter',       'Thank you for your payment', 'Text printed at the bottom of receipts']
      ] });

  add('FinanceCategories', 'finance', ['CategoryID', 'Type', 'Name', 'Description', 'Status'],
    { valid: { Type: ['Income', 'Expense'], Status: ['Active', 'Inactive'] },
      seed: (function () {
        const inc = ['Government Capitation Grant', 'Donations & Sponsorship', 'Fundraising / Harambee', 'School Shop / Canteen',
                     'Hire of School Facilities', 'Farm / Project Sales', 'Interest Earned', 'Other Income'];
        const exp = ['Salaries & Wages', 'Utilities (Electricity, Water)', 'Repairs & Maintenance', 'Stationery & Office',
                     'Teaching & Learning Materials', 'Food & Meals', 'Transport & Fuel', 'Sports & Games', 'Library',
                     'Laboratory', 'Security', 'Communication & Internet', 'Medical & First Aid',
                     'Furniture & Equipment', 'Examinations', 'Other Expenses'];
        return inc.map((n, i) => ['INC' + String(i + 1).padStart(2, '0'), 'Income', n, '', 'Active'])
          .concat(exp.map((n, i) => ['EXP' + String(i + 1).padStart(2, '0'), 'Expense', n, '', 'Active']));
      })() });

  add('FeeStructure', 'finance',
    ['FeeID', 'AcademicYear', 'Class', 'Term', 'FeeName', 'Amount', 'Mandatory', 'Status', 'Notes'],
    { valid: { Class: ['All'].concat(PS_CLASSES), Term: ['Term 1', 'Term 2', 'Annual'],
               Mandatory: ['Yes', 'No'], Status: ['Active', 'Inactive'] },
      fmt: { Amount: MONEY } });

  add('FeeInvoices', 'finance',
    ['InvoiceNo', 'IssueDate', 'DueDate', 'AdmNo', 'StudentName', 'Class', 'AcademicYear', 'Term', 'FeeName',
     'Amount', 'Discount', 'NetAmount', 'Paid', 'Balance', 'Status', 'CreatedBy', 'CreatedAt'],
    { valid: { Status: ['Unpaid', 'Partial', 'Paid', 'Cancelled'] },
      fmt: { IssueDate: DATE, DueDate: DATE, Amount: MONEY, Discount: MONEY, NetAmount: MONEY, Paid: MONEY, Balance: MONEY } });

  add('Payments', 'finance',
    ['ReceiptNo', 'PaymentDate', 'AdmNo', 'StudentName', 'Class', 'AcademicYear', 'Term', 'InvoiceNo', 'FeeName',
     'Amount', 'Method', 'AccountID', 'Reference', 'PaidBy', 'ReceivedBy', 'Status', 'VoidReason', 'Notes', 'CreatedAt'],
    { valid: { Method: PS_METHODS, Status: ['Valid', 'Voided'] },
      fmt: { PaymentDate: DATE, Amount: MONEY } });

  add('Income', 'finance',
    ['IncomeID', 'Date', 'CategoryID', 'Category', 'Description', 'Amount', 'Method', 'AccountID', 'Reference',
     'ReceivedFrom', 'RecordedBy', 'Status', 'CreatedAt'],
    { valid: { Method: PS_METHODS, Status: ['Valid', 'Voided'] }, fmt: { Date: DATE, Amount: MONEY } });

  add('Expenditure', 'finance',
    ['ExpenseID', 'Date', 'CategoryID', 'Category', 'Description', 'Amount', 'Payee', 'SupplierID', 'Method', 'AccountID',
     'VoucherNo', 'Reference', 'Status', 'RequestedBy', 'ApprovedBy', 'ApprovedAt', 'PaidAt', 'Attachment', 'CreatedAt'],
    { valid: { Method: PS_METHODS, Status: ['Pending', 'Approved', 'Paid', 'Rejected'] },
      fmt: { Date: DATE, Amount: MONEY } });

  add('CashBankAccounts', 'finance',
    ['AccountID', 'AccountName', 'Type', 'BankName', 'AccountNumber', 'OpeningBalance', 'Currency', 'Status', 'Notes'],
    { valid: { Type: ['Cash', 'Bank', 'Mobile Money'], Status: ['Active', 'Inactive'] }, fmt: { OpeningBalance: MONEY },
      seed: [
        ['ACC-001', 'Cash on Hand',       'Cash', '', '', 0, 'TZS', 'Active', 'Petty cash & fees received in cash'],
        ['ACC-002', 'Main Bank Account',  'Bank', '', '', 0, 'TZS', 'Active', 'Fill in bank name and account number']
      ] });

  add('CashBankTransactions', 'finance',
    ['TxnID', 'Date', 'AccountID', 'Direction', 'Amount', 'RefType', 'RefID', 'Description', 'TransferToAccountID',
     'RecordedBy', 'CreatedAt'],
    { valid: { Direction: ['In', 'Out'], RefType: ['Payment', 'Income', 'Expense', 'Transfer', 'Adjustment'] },
      fmt: { Date: DATE, Amount: MONEY } });

  add('CashReconciliations', 'finance',
    ['RecID', 'Date', 'AccountID', 'SystemBalance', 'ActualBalance', 'Difference', 'Notes', 'ReconciledBy', 'CreatedAt'],
    { fmt: { Date: DATE, SystemBalance: MONEY, ActualBalance: MONEY, Difference: MONEY } });

  add('DebtorFollowups', 'finance',
    ['FollowupID', 'Date', 'AdmNo', 'StudentName', 'Class', 'OutstandingAmount', 'Method', 'Notes',
     'PromisedDate', 'PromisedAmount', 'Status', 'By'],
    { valid: { Method: ['Phone Call', 'SMS', 'Letter', 'Visit', 'Meeting'], Status: ['Open', 'Promised', 'Settled'] },
      fmt: { Date: DATE, PromisedDate: DATE, OutstandingAmount: MONEY, PromisedAmount: MONEY } });

  // ═════════ PROCUREMENT & STOCK ═════════
  add('Suppliers', 'stock',
    ['SupplierID', 'Name', 'ContactPerson', 'Phone', 'Email', 'Address', 'TIN', 'Category', 'Status'],
    { valid: { Status: ['Active', 'Inactive'] } });

  add('ProcurementRequests', 'stock',
    ['RequestID', 'Date', 'RequestedBy', 'Department', 'ItemName', 'Quantity', 'Unit', 'EstUnitCost', 'EstTotal',
     'Status', 'ApprovedBy', 'ApprovedAt', 'SupplierID', 'PONumber', 'ReceivedDate', 'Notes'],
    { valid: { Status: ['Pending', 'Approved', 'Ordered', 'Received', 'Rejected'] },
      fmt: { Date: DATE, EstUnitCost: MONEY, EstTotal: MONEY } });

  add('StockItems', 'stock',
    ['ItemID', 'ItemName', 'Category', 'Unit', 'ReorderLevel', 'CurrentQty', 'UnitCost', 'Location', 'Status'],
    { valid: { Status: ['Active', 'Inactive'] }, fmt: { UnitCost: MONEY } });

  add('StockMovements', 'stock',
    ['MovementID', 'Date', 'ItemID', 'ItemName', 'Type', 'Quantity', 'UnitCost', 'Reference', 'IssuedTo', 'RecordedBy', 'Notes'],
    { valid: { Type: ['In', 'Out', 'Adjustment'] }, fmt: { Date: DATE, UnitCost: MONEY } });

  // ═════════ SYSTEM ═════════
  add('AuditLog', 'system',
    ['LogID', 'Timestamp', 'Username', 'Role', 'Module', 'Action', 'RecordID', 'OldValue', 'NewValue', 'Details']);
  add('Activity', 'system', ['Date', 'User', 'Action']);

  return defs;
}

// ------------------------------------------------------------
// ENGINE
// ------------------------------------------------------------
function psEnsureSheet_(ss, def, report) {
  let sh = ss.getSheetByName(def.name);
  const isNew = !sh;
  if (isNew) sh = ss.insertSheet(def.name);

  let lastCol = sh.getLastColumn();
  if (sh.getLastRow() === 0 || lastCol === 0) {
    sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
    lastCol = def.headers.length;
    if (isNew) report.created.push(def.name); else report.repaired.push(def.name + ' (header row written)');
  } else {
    const have = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
    const missing = def.headers.filter(h => have.indexOf(h) < 0);
    if (missing.length) {
      sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
      lastCol += missing.length;
      report.repaired.push(def.name + ' (+ columns: ' + missing.join(', ') + ')');
    } else {
      report.ok.push(def.name);
    }
  }

  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
  const colOf = name => head.indexOf(name) + 1;                       // 1-based, 0 if absent

  // header look
  sh.getRange(1, 1, 1, lastCol).setBackground(PS_COLORS[def.group] || '#0d1b36').setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);

  // number formats + drop-downs (rows 2…2000)
  const ROWS = 2000;
  Object.keys(def.fmt || {}).forEach(h => {
    const c = colOf(h); if (c) sh.getRange(2, c, ROWS, 1).setNumberFormat(def.fmt[h]);
  });
  Object.keys(def.valid || {}).forEach(h => {
    const c = colOf(h); if (!c) return;
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(def.valid[h], true).setAllowInvalid(true).build();
    sh.getRange(2, c, ROWS, 1).setDataValidation(rule);
  });

  // starter data
  if (def.seed && def.seed.length) {
    if (def.kv) {
      const existing = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
      const toAdd = def.seed.filter(r => existing.indexOf(String(r[0])) < 0);
      if (toAdd.length) {
        sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, def.headers.length).setValues(toAdd);
        report.seeded.push(def.name + ' (' + toAdd.length + ' settings)');
      }
    } else if (sh.getLastRow() <= 1) {
      const rows = def.seed.map(r => {
        const out = new Array(lastCol).fill('');
        def.headers.forEach((h, i) => { const c = colOf(h); if (c) out[c - 1] = r[i]; });
        return out;
      });
      sh.getRange(2, 1, rows.length, lastCol).setValues(rows);
      report.seeded.push(def.name + ' (' + rows.length + ' rows)');
    }
  }
  return sh;
}

// Creates / repairs everything. Returns a report object.
function psRunSetup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const report = { created: [], repaired: [], ok: [], seeded: [], version: PS_SCHEMA_VERSION };
  psSchema_().forEach(def => psEnsureSheet_(ss, def, report));

  // remove the empty default "Sheet1" that every new spreadsheet has
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) { try { ss.deleteSheet(s1); } catch (e) {} }

  PropertiesService.getScriptProperties().setProperty('PS_SCHEMA_VERSION', PS_SCHEMA_VERSION);
  CacheService.getScriptCache().put('PS_OK', '1', 600);
  Logger.log(JSON.stringify(report));
  return report;
}

// Manual entry point (menu or editor ▸ Run)
function psSetupAll() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let r;
  try { r = psRunSetup_(); } finally { lock.releaseLock(); }
  try {
    SpreadsheetApp.getUi().alert('School tables are ready',
      'Created: ' + r.created.length + ' table(s)\n' +
      'Repaired: ' + r.repaired.length + '\n' +
      'Already fine: ' + r.ok.length + '\n' +
      'Starter data added to: ' + r.seeded.length + ' table(s)',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { /* no UI when run from a trigger / web app */ }
  return r;
}

// Cheap automatic check — call from doGet(). Creates the tables the first time
// the web app opens, and again if someone deletes a sheet by mistake.
function psAutoSetup_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('PS_OK')) return false;
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const have = ss.getSheets().map(s => s.getName());
  const allThere = psSchema_().every(d => have.indexOf(d.name) >= 0);
  if (allThere && props.getProperty('PS_SCHEMA_VERSION') === PS_SCHEMA_VERSION) {
    cache.put('PS_OK', '1', 600);
    return false;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { psRunSetup_(); } finally { lock.releaseLock(); }
  return true;
}

// Table report (name, group, rows) — useful for an admin screen
function psSchemaReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return psSchema_().map(d => {
    const sh = ss.getSheetByName(d.name);
    return { name: d.name, group: d.group, exists: !!sh, rows: sh ? Math.max(0, sh.getLastRow() - 1) : 0, columns: d.headers.length };
  });
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('School System')
      .addItem('Set up / repair all tables', 'psSetupAll')
      .addItem('Show table report', 'psShowReport')
      .addToUi();
  } catch (e) {}
}

function psShowReport() {
  const rep = psSchemaReport();
  const lines = rep.map(r => (r.exists ? '✔ ' : '✘ ') + r.name + '  (' + r.rows + ' rows)');
  SpreadsheetApp.getUi().alert('Tables — ' + rep.filter(r => r.exists).length + ' of ' + rep.length + ' exist', lines.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ------------------------------------------------------------
// HELPERS the finance modules will use
// ------------------------------------------------------------

// Next number from FinanceSettings counter, e.g. psNextNumber_('NextReceiptNo','ReceiptPrefix') → "RCP-2026-00001"
function psNextNumber_(counterKey, prefixKey) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FinanceSettings');
    const data = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
    let counterRow = -1, prefix = 'NO';
    data.forEach((r, i) => {
      if (String(r[0]).trim() === counterKey) counterRow = i + 1;
      if (String(r[0]).trim() === prefixKey)  prefix = String(r[1]).trim() || prefix;
    });
    if (counterRow < 0) throw new Error('Counter ' + counterKey + ' not found in FinanceSettings');
    const n = Number(data[counterRow - 1][1]) || 1;
    sh.getRange(counterRow, 2).setValue(n + 1);
    return prefix + '-' + new Date().getFullYear() + '-' + String(n).padStart(5, '0');
  } finally { lock.releaseLock(); }
}

// One line in AuditLog (Control & Audit)
function psAudit_(username, role, module, action, recordId, oldValue, newValue, details) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AuditLog');
    if (!sh) return;
    sh.appendRow(['LOG' + Date.now() + Math.floor(Math.random() * 1000), new Date(), username || '', role || '', module || '',
                  action || '', recordId || '', oldValue == null ? '' : String(oldValue),
                  newValue == null ? '' : String(newValue), details || '']);
  } catch (e) { Logger.log('psAudit_: ' + e.message); }
}