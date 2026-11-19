-- Migration: 0002_roles_and_payments.sql
-- Engine: Cloudflare D1 (SQLite)
-- Purpose: Expand accounts.role CHECK constraint and add multi-tenant Fee Structures & Payments tables.

PRAGMA foreign_keys = OFF;

-- ============================================================================
-- 1. UPDATE ACCOUNTS TABLE (Expand role CHECK constraint safely)
-- ============================================================================
CREATE TABLE accounts_new (
  id TEXT PRIMARY KEY,
  school_id TEXT NOT NULL,
  seq INTEGER NOT NULL,
  username TEXT NOT NULL,
  full_name TEXT NOT NULL DEFAULT '',
  password_hash TEXT NOT NULL,
  password_salt TEXT NOT NULL,
  role TEXT NOT NULL CHECK (role IN (
    'superadmin', 'schoolmanager', 'headofschool', 'headteacher',
    'secondmaster', 'accountant', 'bursar', 'academic', 'academicmaster',
    'discipline', 'classteacher', 'teacher', 'librarian', 'storekeeper',
    'patron', 'matron', 'staff', 'parent', 'student'
  )),
  position TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT 'Active',
  profile TEXT NOT NULL DEFAULT '{}',
  created_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  updated_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  FOREIGN KEY (school_id) REFERENCES schools(id) ON DELETE CASCADE
);

INSERT INTO accounts_new (
  id, school_id, seq, username, full_name, password_hash, password_salt,
  role, position, status, profile, created_at, updated_at
)
SELECT
  id, school_id, seq, username, full_name, password_hash, password_salt,
  role, position, status, profile, created_at, updated_at
FROM accounts;

DROP TABLE accounts;
ALTER TABLE accounts_new RENAME TO accounts;

CREATE UNIQUE INDEX IF NOT EXISTS idx_accounts_tenant_username
  ON accounts (school_id, lower(trim(username)));

CREATE INDEX IF NOT EXISTS idx_accounts_tenant_role
  ON accounts (school_id, role);

-- ============================================================================
-- 2. FEE STRUCTURES TABLE (Multi-Tenant)
-- ============================================================================
CREATE TABLE IF NOT EXISTS fee_structures (
  id TEXT PRIMARY KEY,
  school_id TEXT NOT NULL,
  academic_year TEXT NOT NULL,                   -- e.g. '2026'
  class_level TEXT NOT NULL,                     -- e.g. 'Form 1', 'Standard 4', 'ALL'
  fee_name TEXT NOT NULL,                        -- e.g. 'Tuition Fee', 'Lunch', 'Exam Fee'
  amount REAL NOT NULL CHECK (amount >= 0),
  currency TEXT NOT NULL DEFAULT 'TZS',
  term TEXT NOT NULL DEFAULT 'Annual',           -- 'Term 1', 'Term 2', 'Annual'
  is_mandatory INTEGER NOT NULL DEFAULT 1,       -- 1 = True, 0 = False
  created_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  updated_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  FOREIGN KEY (school_id) REFERENCES schools(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_fee_structures_tenant
  ON fee_structures (school_id, academic_year, class_level);

-- ============================================================================
-- 3. PAYMENTS / TRANSACTIONS TABLE (Multi-Tenant)
-- ============================================================================
CREATE TABLE IF NOT EXISTS payments (
  id TEXT PRIMARY KEY,                           -- UUID or 'pay_...'
  school_id TEXT NOT NULL,                       -- Tenant key
  receipt_no TEXT NOT NULL,                      -- Unique receipt per school
  student_id TEXT NOT NULL,                      -- Admission No / Student ID
  student_name TEXT NOT NULL DEFAULT '',
  class_level TEXT NOT NULL DEFAULT '',
  academic_year TEXT NOT NULL DEFAULT '',
  fee_structure_id TEXT,                         -- Optional link to fee_structures
  fee_category TEXT NOT NULL DEFAULT 'Tuition',  -- e.g. 'Tuition', 'Transport', 'Uniform'
  amount_paid REAL NOT NULL CHECK (amount_paid > 0),
  currency TEXT NOT NULL DEFAULT 'TZS',
  payment_method TEXT NOT NULL CHECK (payment_method IN (
    'cash', 'bank', 'mobile_money', 'control_number', 'cheque', 'other'
  )) DEFAULT 'cash',
  reference_no TEXT NOT NULL DEFAULT '',         -- Bank slip / M-Pesa / TigoPesa / GePG Control No
  payment_date TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  status TEXT NOT NULL CHECK (status IN (
    'completed', 'pending', 'reversed', 'cancelled'
  )) DEFAULT 'completed',
  recorded_by TEXT NOT NULL DEFAULT '',          -- Username of accountant/bursar
  notes TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  updated_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  FOREIGN KEY (school_id) REFERENCES schools(id) ON DELETE CASCADE,
  FOREIGN KEY (fee_structure_id) REFERENCES fee_structures(id) ON DELETE SET NULL
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_payments_tenant_receipt
  ON payments (school_id, lower(trim(receipt_no)));

CREATE INDEX IF NOT EXISTS idx_payments_tenant_student
  ON payments (school_id, student_id, academic_year);

CREATE INDEX IF NOT EXISTS idx_payments_tenant_date
  ON payments (school_id, payment_date DESC);

PRAGMA foreign_keys = ON;
