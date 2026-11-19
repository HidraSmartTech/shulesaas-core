-- Migration: 0001_initial_schema.sql
-- Engine: Cloudflare D1 (SQLite)
-- SaaS Multi-tenant Schema for School Management System

PRAGMA foreign_keys = ON;

-- ============================================================================
-- 1. TENANTS / SCHOOLS TABLE
-- ============================================================================
CREATE TABLE IF NOT EXISTS schools (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  subdomain TEXT UNIQUE NOT NULL,
  education_level TEXT NOT NULL DEFAULT 'O-Level',
  plan TEXT NOT NULL CHECK (plan IN ('basic', 'pro', 'advanced')) DEFAULT 'basic',
  status TEXT NOT NULL CHECK (status IN ('active', 'suspended', 'expired')) DEFAULT 'active',
  max_students INTEGER NOT NULL DEFAULT 500,
  logo_url TEXT DEFAULT '',
  contact_email TEXT DEFAULT '',
  contact_phone TEXT DEFAULT '',
  subscription_expires_at TEXT,
  created_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  updated_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP)
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_schools_subdomain ON schools (lower(trim(subdomain)));

-- ============================================================================
-- 2. ACCOUNTS & SESSIONS (Multi-Tenant)
-- ============================================================================
CREATE TABLE IF NOT EXISTS accounts (
  id TEXT PRIMARY KEY,
  school_id TEXT NOT NULL,
  seq INTEGER NOT NULL,
  username TEXT NOT NULL,
  full_name TEXT NOT NULL DEFAULT '',
  password_hash TEXT NOT NULL,
  password_salt TEXT NOT NULL,
  role TEXT NOT NULL CHECK (role IN (
    'superadmin', 'schoolmanager', 'headofschool', 'secondmaster', 
    'accountant', 'academic', 'classteacher', 'teacher', 'staff', 'parent'
  )),
  position TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT 'Active',
  profile TEXT NOT NULL DEFAULT '{}',
  created_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  updated_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  FOREIGN KEY (school_id) REFERENCES schools(id) ON DELETE CASCADE
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_accounts_tenant_username 
  ON accounts (school_id, lower(trim(username)));

CREATE TABLE IF NOT EXISTS app_sessions (
  token_hash TEXT PRIMARY KEY,
  school_id TEXT NOT NULL,
  username TEXT NOT NULL,
  role TEXT NOT NULL,
  expires_at TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  FOREIGN KEY (school_id) REFERENCES schools(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_app_sessions_expiry ON app_sessions (expires_at);
CREATE INDEX IF NOT EXISTS idx_app_sessions_tenant ON app_sessions (school_id, token_hash);

-- ============================================================================
-- 3. SCHOOL DATA STORAGE (GAS Emulator Virtual Sheets)
-- ============================================================================
CREATE TABLE IF NOT EXISTS sheet_meta (
  school_id TEXT NOT NULL,
  name TEXT NOT NULL,
  version INTEGER NOT NULL DEFAULT 1,
  updated_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  PRIMARY KEY (school_id, name),
  FOREIGN KEY (school_id) REFERENCES schools(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS sheet_rows (
  school_id TEXT NOT NULL,
  sheet TEXT NOT NULL,
  row_no INTEGER NOT NULL CHECK (row_no >= 1),
  cells TEXT NOT NULL,
  PRIMARY KEY (school_id, sheet, row_no),
  FOREIGN KEY (school_id, sheet) REFERENCES sheet_meta(school_id, name) ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_sheet_rows_lookup ON sheet_rows (school_id, sheet);

-- ============================================================================
-- 4. BACKUPS & AUDIT (Multi-Tenant)
-- ============================================================================
CREATE TABLE IF NOT EXISTS system_backups (
  id TEXT PRIMARY KEY,
  school_id TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  created_by TEXT NOT NULL DEFAULT '',
  note TEXT NOT NULL DEFAULT '',
  data TEXT NOT NULL,
  FOREIGN KEY (school_id) REFERENCES schools(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_backups_tenant ON system_backups (school_id, created_at DESC);

-- ============================================================================
-- 5. SECURITY & THROTTLING
-- ============================================================================
CREATE TABLE IF NOT EXISTS login_failures (
  school_id TEXT NOT NULL,
  key TEXT NOT NULL,
  fails INTEGER NOT NULL DEFAULT 0,
  first_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
  PRIMARY KEY (school_id, key),
  FOREIGN KEY (school_id) REFERENCES schools(id) ON DELETE CASCADE
);
