// ============================================================
// ACCOUNTANT BRIDGE — connects Accountant_v2.html to PrimaryApi.gs
//
// Accountant_v2.html calls  google.script.run.apiRequest({ action, token, ... })
// This file receives those calls, checks the login token (same as psCall),
// and translates them to the existing fin_* functions and sheets.
//
// Install: add this file to the Apps Script project next to PrimaryApi.gs,
// save, then Deploy → Manage deployments → Edit → New version.
// ============================================================

// SPEED: read-only pages are kept in the server cache for AB_CACHE_TTL_ seconds.
// Every write made here (or through psCall) bumps the cache version, so the next open is always fresh.
var AB_CACHE_TTL_ = 60;
var AB_MEMO_ = {};

function apiRequest(payload) {
  payload = payload || {};
  AB_MEMO_ = {};
  psMemoReset_();
  var action = String(payload.action || '');
  if (action === 'logout') return { ok: true, data: null };

  var s = psSessionFromToken_(payload.token);
  if (!s) return { ok: false, code: 'SESSION', error: 'Your session has expired. Please log in again.' };
  if (s.role !== 'accountant') return { ok: false, error: 'You are not allowed to do this.' };

  var h = ab_handlers_()[action];
  if (!h) return { ok: false, error: 'Unknown action: ' + action };

  var cacheKey = '';
  if (!h.write && action === 'dashboard') {
    var args = {}; Object.keys(payload).forEach(function (k) { if (k !== 'token') args[k] = payload[k]; });
    cacheKey = psCacheKey_('ab.' + action, s.role, [args]);
    var hit = psCacheGet_(cacheKey);
    if (hit !== undefined) return { ok: true, data: hit };
  }

  var lock = LockService.getScriptLock(), locked = false, wrote = false;
  try {
    if (h.write) { lock.waitLock(25000); locked = true; wrote = true; PS_LOCK_HELD_ = true; }
    var data = h.fn(s, payload);
    if (cacheKey) psCachePut_(cacheKey, data === undefined ? null : data, AB_CACHE_TTL_);
    return { ok: true, data: data === undefined ? null : data };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  } finally {
    PS_LOCK_HELD_ = false;
    if (locked) lock.releaseLock();
    if (wrote) psCacheBump_();
  }
}

function ab_handlers_() {
  return {
    dashboard:            { fn: ab_page_ },
    savePayment:          { fn: ab_recordPayment_, write: true },
    recordPaymentFinance: { fn: ab_recordPayment_, write: true },
    saveFeeStructure:     { fn: ab_saveFee_, write: true },
    updateFinanceRecord:  { fn: ab_updateRecord_, write: true },
    deleteFinanceRecord:  { fn: ab_deleteRecord_, write: true },
    saveIncome:           { fn: ab_saveIncome_, write: true },
    saveExpense:          { fn: ab_saveExpense_, write: true },
    saveExpenseFinance:   { fn: ab_saveExpense_, write: true },
    expenseStatus:        { fn: ab_expenseStatus_, write: true },
    saveBankAccount:      { fn: ab_saveBank_, write: true },
    cashbookEntry:        { fn: ab_cashbook_, write: true },
    dailyClosing:         { fn: ab_dailyClosing_, write: true },
    savePurchase:         { fn: ab_savePurchase_, write: true },
    saveSupplier:         { fn: ab_saveSupplier_, write: true },
    adjustStock:          { fn: ab_adjustStock_, write: true },
    setPeriodLock:        { fn: ab_setLock_, write: true }
  };
}

// ───────────── helpers ─────────────
function ab_d_(v) { return String(v === null || v === undefined ? '' : v).slice(0, 10); }
function ab_method_(m) {
  m = String(m || '').toLowerCase();
  if (m.indexOf('mpesa') >= 0 || m.indexOf('m-pesa') >= 0 || m.indexOf('mobile') >= 0) return 'Mobile Money';
  if (m.indexOf('bank') >= 0) return 'Bank';
  if (m.indexOf('cheque') >= 0) return 'Cheque';
  return 'Cash';
}
function ab_lock_() {
  var raw = PropertiesService.getScriptProperties().getProperty('AB_PERIOD_LOCK');
  try { return raw ? JSON.parse(raw) : { locked: false, upTo: '' }; } catch (e) { return { locked: false, upTo: '' }; }
}
function ab_checkLock_(date) {
  var l = ab_lock_();
  if (l.locked && date && ab_d_(date) <= l.upTo) psErr_('This record falls in a locked period (up to ' + l.upTo + ').');
}
function ab_obj_(arr, key, amt) {
  var o = {};
  arr.forEach(function (r) { var k = r[key] || 'Other'; o[k] = psRound_((o[k] || 0) + psNum_(r[amt])); });
  return o;
}
function ab_school_() {
  var st = psSettings_();
  return { name: st.SchoolName || 'School', address: st.SchoolAddress || '', phone: st.SchoolPhone || '',
           email: st.SchoolEmail || '', motto: st.SchoolMotto || '', logoUrl: st.SchoolLogo || '', academicYear: psActiveYear_() };
}
function ab_students_() {
  if (!AB_MEMO_.students) AB_MEMO_.students = psAllStudents_().map(function (x) {
    return { id: x.AdmNo, firstName: x.FirstName, lastName: x.LastName, admissionNo: x.AdmNo,
             className: x.Class, classId: x.Class, parent: x.Parent, phone: x.Phone };
  });
  return AB_MEMO_.students;
}
function ab_feeStructures_() {
  return psTable_('FeeStructure').rows().filter(function (r) { return r.Status !== 'Inactive'; }).map(function (r) {
    return { id: r.FeeID, classId: r.Class, className: r.Class, term: r.Term, amount: psNum_(r.Amount),
             academicYear: String(r.AcademicYear), description: r.Notes || r.FeeName, feeName: r.FeeName };
  });
}
function ab_payments_() {
  return psTable_('Payments').rows().filter(function (r) { return r.Status !== 'Voided'; });
}

// Per-student statements: fee = fee structure for the class (active year), paid = valid payments that year
function ab_statements_() {
  var year = psActiveYear_(), fees = ab_feeStructures_().filter(function (f) { return f.academicYear === year; });
  var pays = ab_payments_().filter(function (p) { return String(p.AcademicYear) === year; });
  var byAdm = {};
  pays.forEach(function (p) { (byAdm[p.AdmNo] = byAdm[p.AdmNo] || []).push(p); });
  return ab_students_().map(function (st) {
    var mine = fees.filter(function (f) { return f.classId === 'All' || psNormClass_(f.classId) === psNormClass_(st.classId); });
    var fee = psRound_(mine.reduce(function (a, f) { return a + f.amount; }, 0));
    var ps = (byAdm[st.id] || []).slice().sort(function (a, b) { return ab_d_(a.PaymentDate).localeCompare(ab_d_(b.PaymentDate)); });
    var paid = psRound_(ps.reduce(function (a, p) { return a + psNum_(p.Amount); }, 0));
    var last = ps.length ? ps[ps.length - 1] : null;
    var terms = {};
    mine.forEach(function (f) { terms[f.term] = terms[f.term] || { term: f.term, fee: 0, paid: 0 }; terms[f.term].fee += f.amount; });
    ps.forEach(function (p) { terms[p.Term] = terms[p.Term] || { term: p.Term, fee: 0, paid: 0 }; terms[p.Term].paid += psNum_(p.Amount); });
    return {
      studentId: st.id, studentName: st.firstName + ' ' + st.lastName, admissionNo: st.admissionNo, className: st.className,
      fee: fee, paid: paid, balance: psRound_(fee - paid), payments: ps.length, lastPayment: last ? ab_d_(last.PaymentDate) : '',
      feeTerms: Object.keys(terms).map(function (k) { var t = terms[k]; return { term: t.term, fee: psRound_(t.fee), paid: psRound_(t.paid), balance: psRound_(t.fee - t.paid) }; }),
      lastPaymentRecord: last ? { id: last.ReceiptNo, receiptNo: last.ReceiptNo, paymentDate: ab_d_(last.PaymentDate), amount: psNum_(last.Amount),
        paymentMethod: last.Method, reference: last.Reference, term: last.Term, academicYear: String(last.AcademicYear),
        studentName: last.StudentName, accountName: last.ReceivedBy, accountDesignation: 'Accountant', paymentType: '' } : null
    };
  });
}

function ab_overview_() {
  var st = ab_statements_();
  var pays = ab_payments_();
  var inc = psTable_('Income').rows().filter(function (r) { return r.Status !== 'Voided'; });
  var expAll = psTable_('Expenditure').rows();
  var exp = expAll.filter(function (r) { return r.Status === 'Paid'; });
  var feeIncome = psRound_(pays.reduce(function (a, r) { return a + psNum_(r.Amount); }, 0));
  var otherIncome = psRound_(inc.reduce(function (a, r) { return a + psNum_(r.Amount); }, 0));
  var expenditure = psRound_(exp.reduce(function (a, r) { return a + psNum_(r.Amount); }, 0));
  var pending = psRound_(expAll.filter(function (r) { return r.Status === 'Pending' || r.Status === 'Approved'; })
    .reduce(function (a, r) { return a + psNum_(r.Amount); }, 0));
  var invoiced = psRound_(st.reduce(function (a, x) { return a + x.fee; }, 0));
  var collected = psRound_(st.reduce(function (a, x) { return a + Math.min(x.paid, x.fee || x.paid); }, 0));
  var debtors = st.filter(function (x) { return x.balance > 0; }).sort(function (a, b) { return b.balance - a.balance; })
    .map(function (x) { return { studentName: x.studentName, admissionNo: x.admissionNo, className: x.className, due: x.fee, paid: x.paid, balance: x.balance }; });
  var outstanding = psRound_(debtors.reduce(function (a, x) { return a + x.balance; }, 0));
  var bal = psBalances_(), cash = 0, bank = 0;
  bal.forEach(function (a) { if (a.Type === 'Cash') cash += a.Balance; else bank += a.Balance; });

  var months = [], now = new Date();
  for (var i = 5; i >= 0; i--) {
    var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    months.push({ key: Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM'), label: Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM yyyy'), income: 0, expenditure: 0 });
  }
  var mk = {}; months.forEach(function (m) { mk[m.key] = m; });
  pays.forEach(function (r) { var m = mk[ab_d_(r.PaymentDate).slice(0, 7)]; if (m) m.income += psNum_(r.Amount); });
  inc.forEach(function (r) { var m = mk[ab_d_(r.Date).slice(0, 7)]; if (m) m.income += psNum_(r.Amount); });
  exp.forEach(function (r) { var m = mk[ab_d_(r.Date).slice(0, 7)]; if (m) m.expenditure += psNum_(r.Amount); });
  months.forEach(function (m) { m.income = psRound_(m.income); m.expenditure = psRound_(m.expenditure); m.net = psRound_(m.income - m.expenditure); });

  var income = psRound_(feeIncome + otherIncome);
  return {
    totals: { income: income, expenditure: expenditure, balance: psRound_(income - expenditure), pendingPayments: pending,
              outstandingDebt: outstanding, cashInHand: psRound_(cash), bankTotal: psRound_(bank), otherIncome: otherIncome,
              feeIncome: feeIncome, invoiced: invoiced, collected: collected },
    collectionRate: invoiced > 0 ? Math.round(collected / invoiced * 100) : 0,
    monthly: months, debtors: debtors, debtorCount: debtors.length,
    _pays: pays, _inc: inc, _exp: exp, _statements: st
  };
}
function ab_clean_(o) { var c = {}; Object.keys(o).forEach(function (k) { if (k.charAt(0) !== '_') c[k] = o[k]; }); return c; }

// ───────────── pages ─────────────
function ab_page_(s, p) {
  var page = String(p.page || 'dashboard');
  var year = psActiveYear_();
  var o, t;
  switch (page) {
    case 'dashboard': {
      o = ab_overview_();
      return { stats: { revenue: o.totals.income, expenses: o.totals.expenditure, balance: o.totals.balance, feesOutstanding: o.totals.outstandingDebt },
               byMethod: ab_obj_(o._pays, 'Method', 'Amount'), byCategory: ab_obj_(o._exp, 'Category', 'Amount'), announcements: [] };
    }
    case 'finDashboard': {
      o = ab_overview_();
      return { overview: ab_clean_(o), expenseByCategory: ab_obj_(o._exp, 'Category', 'Amount'), incomeByCategory: ab_obj_(o._inc, 'Category', 'Amount') };
    }
    case 'finFees': {
      o = ab_overview_();
      var byClass = {};
      o._statements.forEach(function (x) {
        var c = byClass[x.className] = byClass[x.className] || { className: x.className, students: 0, feeAmount: x.fee, expected: 0, collected: 0, outstanding: 0, debtors: 0 };
        c.students++; c.expected += x.fee; c.collected += Math.min(x.paid, x.fee || x.paid);
        if (x.balance > 0) { c.outstanding += x.balance; c.debtors++; }
      });
      var cr = Object.keys(byClass).map(function (k) { var c = byClass[k];
        c.expected = psRound_(c.expected); c.collected = psRound_(c.collected); c.outstanding = psRound_(c.outstanding);
        c.rate = c.expected > 0 ? Math.round(c.collected / c.expected * 100) : 0; return c; });
      return { overview: ab_clean_(o), statements: o._statements, classReport: cr, students: ab_students_(),
               feeStructures: ab_feeStructures_(), academicYear: year, paymentMethods: PS_METHODS, schoolProfile: ab_school_() };
    }
    case 'finSettings': {
      var classes = psTable_('AcademicClasses').rows().filter(function (c) { return c.Name; }).map(function (c) { return { id: c.Name, name: c.Name }; });
      if (!classes.length) PS_CLASSES.forEach(function (c) { classes.push({ id: c, name: c }); });
      return { classes: classes, feeStructures: ab_feeStructures_(), academicYear: year };
    }
    case 'finIncome': {
      var inc = fin_incomeList_(s, {}).filter(function (r) { return r.Status !== 'Voided'; });
      o = ab_overview_();
      return { income: inc.map(function (r) { return { id: r.IncomeID, date: ab_d_(r.Date), category: r.Category, description: r.Description,
                 source: r.ReceivedFrom, method: r.Method, reference: r.Reference || r.IncomeID, amount: r.Amount }; }),
               total: psRound_(inc.reduce(function (a, r) { return a + r.Amount; }, 0)), byCategory: ab_obj_(inc, 'Category', 'Amount'),
               categories: fin_catList_(s, 'Income').filter(function (c) { return c.Status !== 'Inactive'; }).map(function (c) { return c.Name; }),
               paymentMethods: PS_METHODS, overview: ab_clean_(o) };
    }
    case 'finExpenses': {
      var ex = fin_expenses_(s, {}).filter(function (r) { return r.Category !== 'Purchases'; });
      var bs = { Pending: 0, Approved: 0, Paid: 0, Rejected: 0 };
      ex.forEach(function (r) { bs[r.Status] = (bs[r.Status] || 0) + 1; });
      var paid = ex.filter(function (r) { return r.Status === 'Paid'; });
      return { expenses: ex.map(function (r) { return { id: r.ExpenseID, date: ab_d_(r.Date), category: r.Category, description: r.Description,
                 method: r.Method, reference: r.VoucherNo || r.Reference || r.ExpenseID, evidenceUrl: '', amount: r.Amount, status: r.Status }; }),
               byStatus: bs, total: psRound_(paid.reduce(function (a, r) { return a + r.Amount; }, 0)), byCategory: ab_obj_(paid, 'Category', 'Amount'),
               categories: fin_catList_(s, 'Expense').filter(function (c) { return c.Status !== 'Inactive'; }).map(function (c) { return c.Name; }),
               paymentMethods: PS_METHODS };
    }
    case 'finCashBank': {
      var bal = psBalances_(), accName = {}, cashB = 0, bankB = 0;
      bal.forEach(function (a) { accName[a.AccountID] = a; if (a.Type === 'Cash') cashB += a.Balance; else bankB += a.Balance; });
      var tx = psTable_('CashBankTransactions').rows().slice().sort(function (a, b) {
        return ab_d_(a.Date).localeCompare(ab_d_(b.Date)) || a._row - b._row; });
      var run = bal.reduce(function (a, x) { return a + psNum_(x.OpeningBalance); }, 0), tIn = 0, tOut = 0, lock = ab_lock_();
      var cb = tx.map(function (r) {
        var amt = psNum_(r.Amount), isIn = r.Direction === 'In';
        if (isIn) { run += amt; tIn += amt; } else { run -= amt; tOut += amt; }
        var a = accName[r.AccountID] || {};
        return { date: ab_d_(r.Date), description: r.Description, category: r.RefType, account: a.AccountName || r.AccountID,
                 reference: r.RefID || r.TxnID, moneyIn: isIn ? amt : 0, moneyOut: isIn ? 0 : amt, balance: psRound_(run),
                 locked: r.RefType === 'Closing' || (lock.locked && ab_d_(r.Date) <= lock.upTo) };
      }).reverse();
      return { totals: { cashBalance: psRound_(cashB), bankBalance: psRound_(bankB), totalIn: psRound_(tIn), totalOut: psRound_(tOut) },
               cashbook: cb, banks: bal.filter(function (a) { return a.Type !== 'Cash' && a.Status !== 'Inactive'; }).map(function (a) {
                 return { id: a.AccountID, bankName: a.BankName || a.AccountName, accountName: a.AccountName, accountNumber: a.AccountNumber,
                          currency: a.Currency || 'TZS', balance: a.Balance }; }) };
    }
    case 'finReports': return ab_report_(s, p.period || 'Monthly');
    case 'finProcurement': {
      var items = fin_stockItems_(s).filter(function (i) { return i.Status !== 'Inactive'; });
      var sups = fin_supList_(s).filter(function (x) { return x.Status !== 'Inactive'; });
      var purch = psTable_('Expenditure').rows().filter(function (r) { return r.Category === 'Purchases' && r.Status !== 'Rejected'; })
        .map(function (r) { var m = String(r.Reference || '').match(/^PUR\|(.*)\|([\d.]+)\|([\d.]+)$/) || [];
          return { id: r.ExpenseID, date: ab_d_(r.Date), item: m[1] || r.Description, quantity: psNum_(m[2]) || 1, unitCost: psNum_(m[3]) || psNum_(r.Amount),
                   total: psNum_(r.Amount), supplier: r.Payee, reference: r.VoucherNo || r.ExpenseID }; })
        .sort(function (a, b) { return b.date.localeCompare(a.date); });
      return { stats: { purchaseTotal: psRound_(purch.reduce(function (a, x) { return a + x.total; }, 0)),
                        stockValue: psRound_(items.reduce(function (a, i) { return a + i.CurrentQty * i.UnitCost; }, 0)),
                        stockItems: items.length, supplierCount: sups.length },
               purchases: purch,
               stock: items.map(function (i) { return { id: i.ItemID, item: i.ItemName, quantity: i.CurrentQty, unitCost: i.UnitCost,
                        value: psRound_(i.CurrentQty * i.UnitCost), location: i.Location }; }),
               suppliers: sups.map(function (x) { return { id: x.SupplierID, name: x.Name, contactPerson: x.ContactPerson, phone: x.Phone, category: x.Category }; }) };
    }
    case 'finControl': {
      var all = fin_expenses_(s, {});
      var c = { pending: 0, approved: 0, rejected: 0, paid: 0 };
      all.forEach(function (r) { var k = String(r.Status || '').toLowerCase(); if (c[k] !== undefined) c[k]++; });
      var st2 = psSettings_();
      return { counts: c, periodLocked: ab_lock_(),
               approvalQueue: all.filter(function (r) { return r.Status === 'Pending'; }).map(function (r) {
                 return { id: r.ExpenseID, date: ab_d_(r.Date), category: r.Category, description: r.Description, reference: r.Reference || r.ExpenseID, amount: r.Amount }; }),
               auditTrail: fin_auditList_(s, {}).slice(0, 200).map(function (l) {
                 return { userName: l.Username, action: l.Module + ' · ' + l.Action, details: [l.RecordID, l.Details].filter(String).join(' — '), createdAt: String(l.Timestamp) }; }),
               referenceExamples: { payment: (st2.ReceiptPrefix || 'RCP-') + year + '-00001', income: 'INC…', expense: (st2.VoucherPrefix || 'PV-') + year + '-00001', purchase: 'PUR / ' + (st2.VoucherPrefix || 'PV-') + '…' } };
    }
    case 'payments': {
      return { payments: fin_payments_(s, {}).filter(function (r) { return r.Status !== 'Voided'; }).slice(0, 300).map(function (r) {
                 return { receiptNo: r.ReceiptNo, studentName: r.StudentName, className: r.Class, amount: r.Amount, paymentMethod: r.Method,
                          paymentDate: ab_d_(r.PaymentDate), term: r.Term, reference: r.Reference }; }),
               students: ab_students_() };
    }
    case 'debts': {
      o = ab_overview_();
      return { debtors: o.debtors.map(function (x) { return { studentName: x.studentName, admissionNo: x.admissionNo, className: x.className, fee: x.due, paid: x.paid, balance: x.balance }; }),
               academicYear: year, totalFeesDue: o.totals.invoiced,
               totalPaid: psRound_(o._statements.reduce(function (a, x) { return a + x.paid; }, 0)), totalDebt: o.totals.outstandingDebt };
    }
  }
  return {};
}

function ab_report_(s, period) {
  var tz = Session.getScriptTimeZone(), now = new Date(), from, to = psToday_(), fmt = function (d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); };
  if (period === 'Daily') from = to;
  else if (period === 'Weekly') from = fmt(new Date(now.getTime() - 6 * 86400000));
  else if (period === 'Termly') from = fmt(new Date(now.getFullYear(), now.getMonth() - 3, now.getDate()));
  else if (period === 'Annual') from = now.getFullYear() + '-01-01';
  else { period = 'Monthly'; from = fmt(new Date(now.getFullYear(), now.getMonth(), 1)); }
  var inRange = function (d) { return psInRange_(ab_d_(d), from, to); };
  var pays = ab_payments_(), inc = psTable_('Income').rows().filter(function (r) { return r.Status !== 'Voided'; });
  var expAll = psTable_('Expenditure').rows(), expPaid = expAll.filter(function (r) { return r.Status === 'Paid'; });
  var sum = function (a, k) { return psRound_(a.reduce(function (x, r) { return x + psNum_(r[k]); }, 0)); };
  var opening = psRound_(psBalances_().reduce(function (a, x) { return a + psNum_(x.OpeningBalance); }, 0) +
    sum(pays.filter(function (r) { return ab_d_(r.PaymentDate) < from; }), 'Amount') +
    sum(inc.filter(function (r) { return ab_d_(r.Date) < from; }), 'Amount') -
    sum(expPaid.filter(function (r) { return ab_d_(r.Date) < from; }), 'Amount'));
  var incomeRows = pays.filter(function (r) { return inRange(r.PaymentDate); }).map(function (r) {
      return { date: ab_d_(r.PaymentDate), category: 'School Fees', description: r.StudentName + ' (' + r.Class + ')', source: r.FeeName, method: r.Method, reference: r.ReceiptNo, amount: psNum_(r.Amount) }; })
    .concat(inc.filter(function (r) { return inRange(r.Date); }).map(function (r) {
      return { date: ab_d_(r.Date), category: r.Category, description: r.Description, source: r.ReceivedFrom, method: r.Method, reference: r.Reference || r.IncomeID, amount: psNum_(r.Amount) }; }))
    .sort(function (a, b) { return a.date.localeCompare(b.date); });
  var expenseRows = expAll.filter(function (r) { return inRange(r.Date); }).map(function (r) {
      return { date: ab_d_(r.Date), category: r.Category, description: r.Description, method: r.Method, reference: r.VoucherNo || r.ExpenseID, status: r.Status, amount: psNum_(r.Amount) }; })
    .sort(function (a, b) { return a.date.localeCompare(b.date); });
  var paidRows = expenseRows.filter(function (r) { return r.status === 'Paid'; });
  var income = sum(incomeRows, 'amount'), expenditure = sum(paidRows, 'amount'), sch = ab_school_();
  return {
    period: period, title: period + ' Financial Report', from: from, to: to, academicYear: psActiveYear_(),
    generatedBy: s.name, generatedByRole: 'Accountant', generatedAt: psNow_(),
    summary: { openingBalance: opening, income: income, incomeCount: incomeRows.length, expenditure: expenditure, expenseCount: paidRows.length,
               net: psRound_(income - expenditure), closingBalance: psRound_(opening + income - expenditure) },
    incomeRows: incomeRows, expenseRows: expenseRows,
    incomeByCategory: ab_obj_(incomeRows, 'category', 'amount'), expenseByCategory: ab_obj_(paidRows, 'category', 'amount'),
    schoolName: sch.name, schoolMotto: sch.motto, schoolLogoUrl: sch.logoUrl, schoolAddress: sch.address, schoolPhone: sch.phone, schoolEmail: sch.email
  };
}

// ───────────── writes ─────────────
function ab_recordPayment_(s, p) {
  var adm = String(p.studentId || '').trim(), amount = psRound_(psNum_(p.amount)), date = p.date || psToday_();
  if (!adm) psErr_('Please select a student.');
  if (amount <= 0) psErr_('Please enter a valid amount.');
  ab_checkLock_(date);
  var st = psAllStudents_().filter(function (x) { return x.AdmNo === adm; })[0];
  if (!st) psErr_('Student not found.');
  var year = psActiveYear_(), term = String(p.term || psActiveTerm_());

  if (!p.force) {
    var dup = ab_payments_().some(function (r) { return r.AdmNo === adm && ab_d_(r.PaymentDate) === ab_d_(date) && Math.abs(psNum_(r.Amount) - amount) < 0.01; });
    if (dup) psErr_('DUPLICATE: a payment of ' + amount + ' for this student on ' + ab_d_(date) + ' already exists.');
  }

  // make sure the student is billed for this term (from the fee structure)
  var invT = psTable_('FeeInvoices');
  var findInv = function () {
    return invT.rows().filter(function (r) { return r.AdmNo === adm && String(r.AcademicYear) === year && r.Term === term && r.Status !== 'Cancelled'; });
  };
  var lines = findInv();
  if (!lines.length) {
    try { fin_invGenerate_(s, { year: year, term: term, className: st.Class }); } catch (e) {}
    lines = findInv();
  }
  if (!lines.length) psErr_('No fee is set for ' + st.Class + ' / ' + term + '. Add it under Accountant Setting first.');

  var res = fin_pay_(s, { admNo: adm, amount: amount, method: ab_method_(p.paymentMethod), date: date,
                          reference: p.reference || '', invoiceNo: lines[0].InvoiceNo, notes: p.paymentType || '' });
  var sch = ab_school_();
  return { receiptNo: res.receiptNo, studentId: adm, studentName: st.Name, admissionNo: adm, className: st.Class,
           amount: amount, paymentMethod: res.method, date: res.date, reference: res.reference, paymentType: p.paymentType || '',
           term: term, academicYear: year, outstandingBalance: res.balanceAfter, accountName: s.name, accountDesignation: 'Accountant',
           schoolProfile: sch };
}

function ab_saveFee_(s, p) {
  var cls = String(p.classId || '').trim();
  if (!cls) psErr_('Choose a class.');
  var term = String(p.term || 'Term 1'), year = psActiveYear_();
  var exists = psTable_('FeeStructure').rows().filter(function (r) {
    return psNormClass_(r.Class) === psNormClass_(cls) && r.Term === term && String(r.AcademicYear) === year && r.Status !== 'Inactive'; })[0];
  var rec = { Class: cls, Term: term, AcademicYear: year, FeeName: exists ? exists.FeeName : 'School Fees', Amount: psNum_(p.amount), Notes: p.description || '' };
  if (exists) rec.FeeID = exists.FeeID;
  return fin_feesSave_(s, rec);
}

function ab_updateRecord_(s, p) {
  var f = p.fields || {};
  if (!Object.keys(f).length) return null;
  if (p.entity === 'feeStructures') {
    var cur = psTable_('FeeStructure').rows().filter(function (r) { return r.FeeID === p.id; })[0];
    if (!cur) psErr_('Fee not found.');
    return fin_feesSave_(s, { FeeID: cur.FeeID, Class: cur.Class, AcademicYear: String(cur.AcademicYear), FeeName: cur.FeeName,
      Term: f.term || cur.Term, Amount: f.amount !== undefined ? f.amount : cur.Amount, Notes: f.description !== undefined ? f.description : cur.Notes });
  }
  if (p.entity === 'payments') {
    psErr_('A recorded payment cannot be edited. Delete (void) it and record it again with the correct details.');
  }
  psErr_('Editing this record is not supported.');
}

function ab_deleteRecord_(s, p) {
  var id = String(p.id || '');
  ab_checkLock_(p.date);
  switch (p.entity) {
    case 'payments':      return fin_voidPayment_(s, id, 'Deleted by accountant');
    case 'feeStructures': return fin_feesDelete_(s, id);
    case 'income':        return fin_incomeVoid_(s, id, 'Deleted by accountant');
    case 'suppliers':     return fin_supDelete_(s, id);
    case 'purchases':
    case 'expenses': {
      if (/^ACC-/.test(id)) {
        var a = psTable_('CashBankAccounts').rows().filter(function (r) { return r.AccountID === id; })[0];
        if (!a) psErr_('Account not found.');
        return fin_accountsSave_(s, { AccountID: id, AccountName: a.AccountName, Type: a.Type, BankName: a.BankName, AccountNumber: a.AccountNumber,
                                      OpeningBalance: a.OpeningBalance, Currency: a.Currency, Status: 'Inactive', Notes: a.Notes });
      }
      var e = psTable_('Expenditure').rows().filter(function (r) { return r.ExpenseID === id; })[0];
      if (!e) psErr_('Record not found.');
      return e.Status === 'Paid' ? fin_expenseReverse_(s, id, 'Deleted by accountant') : fin_expenseDelete_(s, id);
    }
  }
  psErr_('Deleting this record is not supported.');
}

function ab_saveIncome_(s, p) {
  ab_checkLock_(p.date);
  return fin_incomeSave_(s, { Amount: p.amount, Category: p.category, Description: p.description, ReceivedFrom: p.source,
                              Method: ab_method_(p.method), Date: p.date, Reference: p.reference || '' });
}
function ab_saveExpense_(s, p) {
  ab_checkLock_(p.date);
  return fin_expenseSave_(s, { Amount: p.amount, Category: p.category, Description: p.description || p.category,
                               Method: ab_method_(p.method || p.paymentMethod), Date: p.date, Status: 'Pending', Reference: p.evidenceUrl || '' });
}
function ab_expenseStatus_(s, p) {
  if (p.status === 'Approved' || p.status === 'Paid') return fin_expensePay_(s, p.id, '', '');
  if (p.status === 'Rejected') return fin_expenseReverse_(s, p.id, 'Rejected by accountant');
  psErr_('Unknown status.');
}
function ab_saveBank_(s, p) {
  return fin_accountsSave_(s, { AccountName: p.accountName || p.bankName, Type: 'Bank', BankName: p.bankName, AccountNumber: p.accountNumber,
                                OpeningBalance: p.balance, Currency: 'TZS' });
}
function ab_cashbook_(s, p) {
  var amt = psRound_(psNum_(p.amount));
  if (amt <= 0) psErr_('Amount must be greater than zero.');
  ab_checkLock_(p.date);
  var acc = psPickAccount_('', p.account === 'Cash' ? 'Cash' : 'Bank');
  var dir = p.type === 'Withdrawal' ? 'Out' : 'In';
  if (dir === 'Out') {
    var b = psBalances_().filter(function (a) { return a.AccountID === acc; })[0];
    if (b && b.Balance + 0.0001 < amt) psErr_(b.AccountName + ' only has ' + b.Balance + '.');
  }
  var id = psId_('CB');
  psTxn_(acc, dir, amt, p.type === 'Withdrawal' ? 'Withdrawal' : 'Deposit', id, p.description || p.type, p.date || psToday_(), s.name);
  psAuditS_(s, 'CashBank', p.type, id, '', amt, p.description || '');
  return p.type + ' recorded.';
}
function ab_dailyClosing_(s) {
  var today = psToday_();
  var tx = psTable_('CashBankTransactions').rows().filter(function (t) { return ab_d_(t.Date) === today; });
  if (tx.some(function (t) { return t.RefType === 'Closing'; })) psErr_('Today has already been closed.');
  var tin = 0, tout = 0;
  tx.forEach(function (t) { if (t.Direction === 'In') tin += psNum_(t.Amount); else tout += psNum_(t.Amount); });
  var acc = psPickAccount_('', 'Cash');
  psTxn_(acc, 'In', 0, 'Closing', 'CLOSE-' + today, 'Daily closing — in ' + psRound_(tin) + ', out ' + psRound_(tout), today, s.name);
  psAuditS_(s, 'CashBank', 'Daily closing', today, '', psRound_(tin - tout), '');
  return { message: 'Day closed. In ' + psRound_(tin) + ', out ' + psRound_(tout) + '.', totalIn: psRound_(tin), totalOut: psRound_(tout) };
}
function ab_savePurchase_(s, p) {
  var item = String(p.item || '').trim(), qty = psNum_(p.quantity), cost = psNum_(p.unitCost);
  if (!item || qty <= 0 || cost <= 0) psErr_('Item, quantity and unit cost are required.');
  ab_checkLock_(p.date);
  var cats = psTable_('FinanceCategories');
  if (!cats.rows().some(function (c) { return c.Type === 'Expense' && c.Name === 'Purchases'; }))
    cats.append({ CategoryID: 'EXP' + Date.now().toString(36).toUpperCase(), Type: 'Expense', Name: 'Purchases', Description: 'Procurement & stock', Status: 'Active' });
  var msg = fin_expenseSave_(s, { Amount: psRound_(qty * cost), Category: 'Purchases', Description: item + ' × ' + qty, Payee: p.supplier || '',
                                  Method: 'Cash', Date: p.date, Status: 'Paid', Reference: 'PUR|' + item + '|' + qty + '|' + cost });
  var stock = psTable_('StockItems'), it = stock.rows().filter(function (i) { return psNormName_(i.ItemName) === psNormName_(item); })[0];
  if (!it) { fin_stockSaveItem_(s, { ItemName: item, UnitCost: cost, CurrentQty: 0 }); it = stock.rows().filter(function (i) { return psNormName_(i.ItemName) === psNormName_(item); })[0]; }
  fin_stockMove_(s, { itemId: it.ItemID, type: 'In', quantity: qty, unitCost: cost, date: p.date, reference: 'Purchase' });
  return msg;
}
function ab_saveSupplier_(s, p) {
  return fin_supSave_(s, { Name: p.name, ContactPerson: p.contactPerson, Phone: p.phone, Email: p.email, Category: p.category });
}
function ab_adjustStock_(s, p) {
  return fin_stockMove_(s, { itemId: p.id, type: 'Adjustment', quantity: psNum_(p.quantity) });
}
function ab_setLock_(s, p) {
  var l = p.locked ? { locked: true, upTo: ab_d_(p.upTo) || psToday_() } : { locked: false, upTo: '' };
  PropertiesService.getScriptProperties().setProperty('AB_PERIOD_LOCK', JSON.stringify(l));
  psAuditS_(s, 'Control', l.locked ? 'Period locked' : 'Period unlocked', l.upTo, '', '', '');
  return l;
}