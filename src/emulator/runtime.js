// Cloudflare Workers V8 GAS Engine Emulator (Multi-Tenant)
// Runs Google Apps Script (.gs) functions inside Cloudflare Worker Sandbox

import { Workbook } from "./sheets.js";
import { encodeChanges, ConflictError } from "./store.js";
import { runGas } from "../legacy_gs/compiled.js";

export const DEFAULT_TIMEZONE = "Africa/Dar_es_Salaam";
const PROPERTIES_SHEET = "__properties";
const MAX_ATTEMPTS = 30;
const STAFF_ROLES = new Set([
  "academic", "academicmaster", "classteacher", "teacher",
  "staff", "secondmaster", "discipline", "bursar", "accountant"
]);

export const VIRTUAL_HEADERS = {
  __accounts: ["seq", "id", "username", "full_name", "password_hash", "password_salt", "role", "position", "status", "profile"]
};

export function resolveEngineType(school) {
  const lvl = String(school?.education_level || "").toUpperCase();
  if (lvl.includes("PRIMARY") || lvl.includes("PRE") || lvl === "NURSERY") {
    return "PRIMARY";
  }
  return "SECONDARY";
}

export function accountsToRows(sheetName, accounts) {
  if (sheetName !== "__accounts") return [];
  return (accounts || []).map((a) => [
    a.seq,
    a.id,
    a.username,
    a.full_name || "",
    a.password_hash || "",
    a.password_salt || "",
    a.role,
    a.position || "",
    a.status || "Active",
    typeof a.profile === "object" ? JSON.stringify(a.profile) : a.profile || "{}"
  ]);
}

export function diffAccountSheet(sheetName, originalRows, currentRows, originalAccounts) {
  if (sheetName !== "__accounts") return [];
  const ops = [];
  const existingMap = new Map((originalAccounts || []).map((a) => [a.username.toLowerCase(), a]));

  for (const row of currentRows) {
    if (!row || !row[2]) continue;
    const username = String(row[2]).trim();
    const role = String(row[6] || "teacher");
    let profileObj = {};
    try {
      profileObj = typeof row[9] === "string" ? JSON.parse(row[9]) : row[9] || {};
    } catch (e) {
      profileObj = {};
    }

    const acc = {
      username,
      full_name: String(row[3] || ""),
      password_hash: String(row[4] || ""),
      password_salt: String(row[5] || ""),
      role,
      position: String(row[7] || ""),
      status: String(row[8] || "Active"),
      profile: profileObj,
    };

    const old = existingMap.get(username.toLowerCase());
    if (!old) {
      ops.push({ type: "insert", account: acc });
    } else {
      ops.push({ type: "update", match: old.username, account: acc });
      existingMap.delete(username.toLowerCase());
    }
  }

  for (const [uname] of existingMap) {
    ops.push({ type: "delete", match: uname });
  }

  return ops;
}

const sheetCache = new Map();
const accountsCache = new Map();
const prefetch = new Map();

export function resetRuntimeCaches() {
  sheetCache.clear();
  accountsCache.clear();
  prefetch.clear();
}

function virtualNames() {
  return Object.keys(VIRTUAL_HEADERS);
}

async function loadNames(store, schoolId, engineType, known, names, state, everything) {
  const wanted = new Set(names);
  if (everything) for (const n of known.keys()) wanted.add(n);

  const fetch = [];
  let needAccounts = false;
  for (const name of wanted) {
    if (state.loaded.has(name)) continue;
    if (VIRTUAL_HEADERS[name]) { needAccounts = true; continue; }
    if (!known.has(name)) continue;

    const cacheKey = `${engineType}:${schoolId}:${name}`;
    const cached = sheetCache.get(cacheKey);
    if (cached && cached.version === known.get(name)) {
      state.loaded.set(name, cached);
    } else {
      fetch.push(name);
    }
  }

  if (fetch.length > 0) {
    const got = await store.loadSheets(schoolId, fetch);
    for (const [name, data] of got) {
      sheetCache.set(`${engineType}:${schoolId}:${name}`, data);
      state.loaded.set(name, data);
    }
  }

  if (needAccounts || everything) {
    const version = known.get("__accounts") || 0;
    const accKey = `${engineType}:${schoolId}`;
    const accCached = accountsCache.get(accKey);
    if (!accCached || accCached.version !== version) {
      const fresh = await store.loadAccounts(schoolId);
      accountsCache.set(accKey, { version: fresh.version, accounts: fresh.accounts });
    }
    const currentAcc = accountsCache.get(accKey);
    state.accounts = currentAcc;
    for (const name of virtualNames()) {
      state.loaded.set(name, {
        version: currentAcc.version,
        rows: accountsToRows(name, currentAcc.accounts)
      });
    }
  }
}

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

function zonedParts(date, timeZone) {
  const fmt = new Intl.DateTimeFormat("en-US", {
    timeZone, hourCycle: "h23", year: "numeric", month: "numeric", day: "numeric",
    hour: "numeric", minute: "numeric", second: "numeric", weekday: "short",
  });
  const out = {};
  for (const p of fmt.formatToParts(date)) out[p.type] = p.value;
  const wd = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].indexOf(out.weekday);
  return {
    y: Number(out.year), M: Number(out.month), d: Number(out.day),
    H: Number(out.hour) % 24, m: Number(out.minute), s: Number(out.second), wd,
  };
}

export function formatDate(date, timeZone, pattern) {
  const d = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(d.getTime())) throw new Error("Invalid date");
  const p = zonedParts(d, timeZone || DEFAULT_TIMEZONE);
  const pad = (n, w = 2) => String(n).padStart(w, "0");
  return String(pattern).replace(/'([^']*)'|y{1,4}|M{1,4}|d{1,2}|E{3,4}|H{1,2}|h{1,2}|m{1,2}|s{1,2}|S{1,3}|a|Z/g, (token, quoted) => {
    if (quoted !== undefined) return quoted || "'";
    switch (token) {
      case "yyyy": return pad(p.y, 4);
      case "yy": return pad(p.y % 100);
      case "y": return String(p.y);
      case "MMMM": return MONTHS[p.M - 1];
      case "MMM": return MONTHS[p.M - 1].slice(0, 3);
      case "MM": return pad(p.M);
      case "M": return String(p.M);
      case "dd": return pad(p.d);
      case "d": return String(p.d);
      case "EEEE": return DAYS[p.wd];
      case "EEE": return DAYS[p.wd].slice(0, 3);
      case "HH": return pad(p.H);
      case "H": return String(p.H);
      case "hh": return pad(p.H % 12 || 12);
      case "h": return String(p.H % 12 || 12);
      case "mm": return pad(p.m);
      case "m": return String(p.m);
      case "ss": return pad(p.s);
      case "s": return String(p.s);
      case "SSS": return pad(d.getMilliseconds(), 3);
      case "a": return p.H < 12 ? "AM" : "PM";
      case "Z": return "+0000";
      default: return token;
    }
  });
}

function buildSandbox({ wb, school, session, readPage, appUrl, timezone, logs, engineType, d1Finance, d1Ops }) {
  const cacheStore = () => {
    const map = new Map();
    return {
      get: (k) => (map.has(k) ? map.get(k) : null),
      put: (k, v) => { map.set(String(k), String(v)); },
      putAll: (o) => { Object.entries(o).forEach(([k, v]) => map.set(k, String(v))); },
      getAll: (keys) => Object.fromEntries(keys.filter((k) => map.has(k)).map((k) => [k, map.get(k)])),
      remove: (k) => { map.delete(k); },
      removeAll: (keys) => keys.forEach((k) => map.delete(k)),
    };
  };

  const scriptCache = cacheStore();
  const userCache = cacheStore();
  if (session && STAFF_ROLES.has(session.role)) {
    userCache.put("currentTeacher", session.username);
  }

  const props = {
    _sheet(create) {
      let sh = wb.getSheetByName(PROPERTIES_SHEET);
      if (!sh && create) {
        sh = wb.insertSheet(PROPERTIES_SHEET);
        sh.appendRow(["Key", "Value"]);
      }
      return sh;
    },
    _all() {
      const sh = this._sheet(false);
      const out = {};
      if (sh) sh.getDataRange().getValues().slice(1).forEach((r) => { if (r[0] !== "") out[String(r[0])] = String(r[1]); });
      return out;
    },
    getProperty(key) {
      const all = this._all();
      return Object.prototype.hasOwnProperty.call(all, key) ? all[key] : null;
    },
    getProperties() { return this._all(); },
    setProperty(key, value) {
      const sh = this._sheet(true);
      const rows = sh.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(key)) { sh.getRange(i + 1, 2).setValue(String(value)); return this; }
      }
      sh.appendRow([String(key), String(value)]);
      return this;
    },
    setProperties(obj) { Object.entries(obj).forEach(([k, v]) => this.setProperty(k, v)); return this; },
    deleteProperty(key) {
      const sh = this._sheet(false);
      if (!sh) return this;
      const rows = sh.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) if (String(rows[i][0]) === String(key)) sh.deleteRow(i + 1);
      return this;
    },
    deleteAllProperties() { const sh = this._sheet(false); if (sh) sh.clearContents(); return this; },
  };

  const lock = { waitLock() {}, tryLock() { return true; }, releaseLock() {}, hasLock() { return true; } };

  const validation = {
    requireValueInList() { return validation; },
    setAllowInvalid() { return validation; },
    build() { return {}; },
  };

  const page = (name) => {
    const file = name === "Staff" ? "Teacher" : name;
    let content = null;
    return {
      getContent() { if (content === null) content = readPage(file); return content; },
      setTitle() { return this; },
      setXFrameOptionsMode() { return this; },
      addMetaTag() { return this; },
      setContent(c) { content = String(c); return this; },
    };
  };

  const schoolContext = school
    ? { ...school, engine_type: engineType }
    : { education_level: engineType === "PRIMARY" ? "Primary" : "O-Level", engine_type: engineType };

  return {
    Date,
    console: { log() {}, error() {}, warn() {} },
    __SCHOOL__: Object.freeze(schoolContext),
    __HOST__: {
      session: session ? { username: session.username, role: session.role, fullName: session.fullName || "" } : null,
      education_level: schoolContext.education_level,
      engine_type: engineType,
      appUrl,
      d1Payments: d1Finance ? d1Finance.payments : [],
      d1FeeStructures: d1Finance ? d1Finance.feeStructures : [],
      d1Ops,
    },
    Logger: { log: (...a) => { logs.push(a.join(" ")); } },
    Session: {
      getScriptTimeZone: () => timezone,
      getActiveUser: () => ({ getEmail: () => "" }),
      getEffectiveUser: () => ({ getEmail: () => "" }),
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => wb,
      getActive: () => wb,
      flush() {},
      newDataValidation: () => validation,
      getUi: () => ({
        alert() {},
        ButtonSet: { OK: "OK" },
        createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() {} }),
      }),
    },
    HtmlService: {
      createHtmlOutputFromFile: page,
      createHtmlOutput: (html) => page("").setContent(html),
      XFrameOptionsMode: { ALLOWALL: "ALLOWALL", DEFAULT: "DEFAULT" },
    },
    PropertiesService: { getScriptProperties: () => props, getUserProperties: () => props, getDocumentProperties: () => props },
    CacheService: { getScriptCache: () => scriptCache, getUserCache: () => userCache, getDocumentCache: () => scriptCache },
    LockService: { getScriptLock: () => lock, getUserLock: () => lock, getDocumentLock: () => lock },
    ScriptApp: { getService: () => ({ getUrl: () => appUrl }) },
    Utilities: {
      formatDate: (date, tz, pattern) => formatDate(date, tz, pattern),
      getUuid: () => crypto.randomUUID(),
      DigestAlgorithm: { SHA_256: "sha256", MD5: "md5", SHA_1: "sha1" },
      Charset: { UTF_8: "utf8" },
      base64Encode: (v) => btoa(String(v)),
      base64Decode: (s) => atob(String(s)),
      sleep() {},
    },
  };
}

export class LegacyError extends Error {
  constructor(message, { internal = false } = {}) {
    super(message);
    this.internal = internal;
  }
}

function executeInWorkerSandbox(codeString, sandbox, fnName, args) {
  return runGas(sandbox, fnName, args);
}

const FINANCE_PRELOAD_FNS = new Set([
  "parentPortalData", "getParentFeeStatus", "apiRequest", "getAccountantData", "getFinancialSummary"
]);

export async function runLegacy({
  store, schoolId, school = null, fn, args = [], session = null, readPage, combinedScriptCode, appUrl = "", timezone,
}) {
  if (!/^[A-Za-z_$][\w$]{0,99}$/.test(fn)) throw new Error("Invalid function name.");

  const resolvedSchool = school || (typeof store.getSchool === "function" ? await store.getSchool(schoolId) : null);
  const engineType = resolveEngineType(resolvedSchool);
  const prefetchKey = `${engineType}:${fn}`;
  const accCacheKey = `${engineType}:${schoolId}`;

  let d1Finance = null;
  if (FINANCE_PRELOAD_FNS.has(fn) && typeof store.loadFinanceData === "function") {
    d1Finance = await store.loadFinanceData(schoolId);
  }

  const tz = timezone || DEFAULT_TIMEZONE;
  const state = { loaded: new Map(), accounts: null };
  const wantedNames = new Set(prefetch.get(prefetchKey) || []);
  let conflicts = 0;
  let everything = false;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    const known = await store.listSheets(schoolId);
    for (const name of virtualNames()) {
      known.set(name, known.get("__accounts") || 0);
    }

    for (const [name, data] of [...state.loaded]) {
      if (known.get(name) !== data.version) state.loaded.delete(name);
    }
    if (state.accounts && state.accounts.version !== known.get("__accounts")) state.accounts = null;

    await loadNames(store, schoolId, engineType, known, wantedNames, state, everything);

    const logs = [];
    const d1Ops = [];
    const wb = new Workbook({
      known: [...known].filter(([n]) => n !== "__accounts"),
      loaded: state.loaded,
      timezone: tz,
      virtualSheets: new Map(virtualNames().map((n) => [n, true])),
    });

    const sandbox = buildSandbox({
      wb,
      school: resolvedSchool,
      session,
      readPage,
      appUrl,
      timezone: tz,
      logs,
      engineType,
      d1Finance,
      d1Ops,
    });

    let outcome;
    try {
      const result = executeInWorkerSandbox(combinedScriptCode, sandbox, fn, args);
      outcome = { ok: true, value: result };
    } catch (error) {
      outcome = { ok: false, error };
    }

    if (wb.missing.size > 0) {
      wb.missing.forEach((n) => wantedNames.add(n));
      if (attempt >= 5) everything = true;
      continue;
    }

    if (!outcome.ok) {
      const e = outcome.error;
      const detail = e && e.message ? `${e.name || "Error"}: ${e.message}` : String(e);
      throw new LegacyError(detail, { internal: false, cause: e });
    }

    prefetch.set(prefetchKey, new Set([...wantedNames, ...wb.states.keys()]));

    const payload = { sheets: encodeChanges(wb.collectChanges()) };
    if (d1Ops.length > 0) payload.d1Ops = d1Ops;

    if (state.accounts) {
      const ops = [];
      for (const name of virtualNames()) {
        const st = wb.states.get(name);
        if (!st || !st.dirty) continue;
        ops.push(...diffAccountSheet(name, st.original, st.rows, state.accounts.accounts));
      }
      if (ops.length > 0) payload.accounts = { expected: state.accounts.version, ops };
    }

    if (payload.sheets.length > 0 || payload.accounts || (payload.d1Ops && payload.d1Ops.length > 0)) {
      try {
        await store.commit(schoolId, payload);
      } catch (error) {
        if (error instanceof ConflictError && ++conflicts <= 4) {
          for (const c of payload.sheets) sheetCache.delete(`${engineType}:${schoolId}:${c.sheet}`);
          accountsCache.delete(accCacheKey);
          state.loaded = new Map();
          state.accounts = null;
          continue;
        }
        throw error;
      }

      for (const c of payload.sheets) sheetCache.delete(`${engineType}:${schoolId}:${c.sheet}`);
      if (payload.accounts) accountsCache.delete(accCacheKey);
    }

    return outcome.value;
  }

  throw new Error("The request needed too many data reloads.");
}
