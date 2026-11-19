// Cloudflare D1 (SQLite) Multi-Tenant Storage Layer for GAS Emulator
// All database operations are strictly isolated by school_id with atomic version checks.

export class ConflictError extends Error {
  constructor(sheet) {
    super(`The data changed while the request was running (${sheet}).`);
    this.sheet = sheet;
    this.isConflict = true;
  }
}

// Helper to encode JS values to JSON strings for SQLite text storage
function encodeCellValue(val) {
  if (val === undefined || val === null) return null;
  if (val instanceof Date) return { __type: "Date", iso: val.toISOString() };
  return val;
}

function decodeCellValue(val, DateCtor = Date) {
  if (val === null || val === undefined) return "";
  if (typeof val === "object" && val !== null && val.__type === "Date") {
    return new DateCtor(val.iso);
  }
  return val;
}

export function encodeRows(rows) {
  return rows.map((row) =>
    (row || []).map((cell) => encodeCellValue(cell))
  );
}

export function decodeRows(rows, DateCtor = Date) {
  return rows.map((row) =>
    (row || []).map((cell) => decodeCellValue(cell, DateCtor))
  );
}

function toRowArray(records, DateCtor) {
  const rows = [];
  for (const record of records) {
    while (rows.length < record.row_no - 1) rows.push([]);
    let parsed = [];
    try {
      parsed = JSON.parse(record.cells);
    } catch (e) {
      parsed = [];
    }
    rows[record.row_no - 1] = decodeRows([parsed], DateCtor)[0];
  }
  return rows;
}

export function createD1Store(db) {
  return {
    // 0. Soma taarifa za shule husika (kwa ajili ya kubaini Primary vs Secondary)
    async getSchool(schoolId) {
      const stmt = db.prepare(
        "SELECT id, name, subdomain, education_level, plan, status FROM schools WHERE id = ?"
      );
      return await stmt.bind(schoolId).first();
    },

    // 1. Orodhesha masheet yote ya shule husika pamoja na matoleo (versions) yake
    async listSheets(schoolId) {
      const stmt = db.prepare(
        "SELECT name, version FROM sheet_meta WHERE school_id = ?"
      );
      const { results } = await stmt.bind(schoolId).all();
      const map = new Map();
      for (const r of results || []) {
        map.set(r.name, Number(r.version));
      }
      return map;
    },

    // 2. Pakia data za masheet mahususi kwa ajili ya shule moja
    async loadSheets(schoolId, names, DateCtor) {
      const out = new Map();
      if (!names || names.length === 0) return out;

      const placeholders = names.map(() => "?").join(",");

      // Get metadata versions
      const metaStmt = db.prepare(
        `SELECT name, version FROM sheet_meta WHERE school_id = ? AND name IN (${placeholders})`
      );
      const { results: versions } = await metaStmt.bind(schoolId, ...names).all();

      // Get row records
      const rowsStmt = db.prepare(
        `SELECT sheet, row_no, cells FROM sheet_rows WHERE school_id = ? AND sheet IN (${placeholders}) ORDER BY sheet, row_no`
      );
      const { results: rows } = await rowsStmt.bind(schoolId, ...names).all();

      const bySheet = new Map();
      for (const r of rows || []) {
        if (!bySheet.has(r.sheet)) bySheet.set(r.sheet, []);
        bySheet.get(r.sheet).push(r);
      }

      for (const v of versions || []) {
        out.set(v.name, {
          version: Number(v.version),
          rows: toRowArray(bySheet.get(v.name) || [], DateCtor),
        });
      }
      return out;
    },

    // 3. Pakia akaunti za shule husika
    async loadAccounts(schoolId) {
      const stmt = db.prepare(
        `SELECT seq, id, username, full_name, password_hash, password_salt, role, position, status, profile 
         FROM accounts WHERE school_id = ? ORDER BY seq`
      );
      const { results: accounts } = await stmt.bind(schoolId).all();

      const formattedAccounts = (accounts || []).map((acc) => {
        let profileObj = {};
        try {
          profileObj = typeof acc.profile === "string" ? JSON.parse(acc.profile) : acc.profile;
        } catch (e) {
          profileObj = {};
        }
        return { ...acc, profile: profileObj };
      });

      const metaStmt = db.prepare(
        "SELECT version FROM sheet_meta WHERE school_id = ? AND name = '__accounts'"
      );
      const meta = await metaStmt.bind(schoolId).first();

      return {
        accounts: formattedAccounts,
        version: meta ? Number(meta.version) : 0,
      };
    },

    // 4. Hifadhi mabadiliko yote kwa mkupuo mmoja (Atomic Optimistic Concurrency via D1 batch)
    async commit(schoolId, payload) {
      const statements = [];
      // Track guarded version-check statements to verify rows updated atomically
      const guards = [];

      // Handle sheet changes
      for (const s of payload.sheets || []) {
        const nm = s.sheet;

        if (s.op === "create") {
          const idx = statements.length;
          statements.push(
            db.prepare(
              "INSERT OR IGNORE INTO sheet_meta (school_id, name, version) VALUES (?, ?, 1)"
            ).bind(schoolId, nm)
          );
          guards.push({ index: idx, sheet: nm });
        } else {
          const expectedVer = Number(s.expected);
          if (!Number.isFinite(expectedVer) || expectedVer < 1) {
            throw new ConflictError(nm);
          }

          if (s.op === "drop") {
            const idx = statements.length;
            statements.push(
              db.prepare(
                "DELETE FROM sheet_meta WHERE school_id = ? AND name = ? AND version = ?"
              ).bind(schoolId, nm, expectedVer)
            );
            guards.push({ index: idx, sheet: nm });

            statements.push(
              db.prepare(
                "DELETE FROM sheet_rows WHERE school_id = ? AND sheet = ?"
              ).bind(schoolId, nm)
            );
            continue;
          }

          // Atomic version increment: fails (changes = 0) if another writer updated version
          const idx = statements.length;
          statements.push(
            db.prepare(
              "UPDATE sheet_meta SET version = version + 1, updated_at = CURRENT_TIMESTAMP WHERE school_id = ? AND name = ? AND version = ?"
            ).bind(schoolId, nm, expectedVer)
          );
          guards.push({ index: idx, sheet: nm });

          // Truncate extra rows if length shortened
          if (s.length !== undefined) {
            statements.push(
              db.prepare(
                "DELETE FROM sheet_rows WHERE school_id = ? AND sheet = ? AND row_no > ?"
              ).bind(schoolId, nm, s.length)
            );
          }
        }

        // Insert or Update cell rows
        if (s.set && s.set.length > 0) {
          for (const [rowNo, cells] of s.set) {
            const encodedCells = JSON.stringify(encodeRows([cells])[0]);
            statements.push(
              db.prepare(
                `INSERT INTO sheet_rows (school_id, sheet, row_no, cells) VALUES (?, ?, ?, ?)
                 ON CONFLICT(school_id, sheet, row_no) DO UPDATE SET cells = excluded.cells`
              ).bind(schoolId, nm, rowNo, encodedCells)
            );
          }
        }
      }

      // Handle Accounts changes with atomic version check
      if (payload.accounts && payload.accounts.ops && payload.accounts.ops.length > 0) {
        const expectedAccVer = Number(payload.accounts.expected || 0);
        const idx = statements.length;

        if (expectedAccVer === 0) {
          statements.push(
            db.prepare(
              "INSERT OR IGNORE INTO sheet_meta (school_id, name, version) VALUES (?, '__accounts', 1)"
            ).bind(schoolId)
          );
        } else {
          statements.push(
            db.prepare(
              "UPDATE sheet_meta SET version = version + 1, updated_at = CURRENT_TIMESTAMP WHERE school_id = ? AND name = '__accounts' AND version = ?"
            ).bind(schoolId, expectedAccVer)
          );
        }
        guards.push({ index: idx, sheet: "__accounts" });

        for (const o of payload.accounts.ops) {
          const a = o.account || {};
          if (o.type === "insert") {
            const accId = a.id || `acc_${Date.now().toString(36)}_${Math.random().toString(36).substring(2, 6)}`;

            statements.push(
              db.prepare(
                `INSERT INTO accounts (id, school_id, seq, username, full_name, password_hash, password_salt, role, position, status, profile)
                 VALUES (
                   ?, ?,
                   (SELECT COALESCE(MAX(seq), 0) + 1 FROM accounts WHERE school_id = ?),
                   ?, ?, ?, ?, ?, ?, ?, ?
                 )`
              ).bind(
                accId,
                schoolId,
                schoolId,
                a.username,
                a.full_name || "",
                a.password_hash || "",
                a.password_salt || "",
                a.role,
                a.position || "",
                a.status || "Active",
                JSON.stringify(a.profile || {})
              )
            );
          } else if (o.type === "update") {
            const oldUsername = o.match;
            const profileStr = a.profile ? JSON.stringify(a.profile) : null;

            statements.push(
              db.prepare(
                `UPDATE accounts SET
                  username = COALESCE(?, username),
                  full_name = COALESCE(?, full_name),
                  password_hash = COALESCE(?, password_hash),
                  password_salt = COALESCE(?, password_salt),
                  role = COALESCE(?, role),
                  position = COALESCE(?, position),
                  status = COALESCE(?, status),
                  profile = COALESCE(?, profile),
                  updated_at = CURRENT_TIMESTAMP
                 WHERE school_id = ? AND lower(trim(username)) = lower(trim(?))`
              ).bind(
                a.username || null,
                a.full_name || null,
                a.password_hash || null,
                a.password_salt || null,
                a.role || null,
                a.position || null,
                a.status || null,
                profileStr,
                schoolId,
                oldUsername
              )
            );

            if (a.password_hash || (a.status && a.status.toLowerCase() !== "active")) {
              statements.push(
                db.prepare(
                  "DELETE FROM app_sessions WHERE school_id = ? AND lower(trim(username)) = lower(trim(?))"
                ).bind(schoolId, oldUsername)
              );
            }
          } else if (o.type === "delete") {
            statements.push(
              db.prepare(
                "DELETE FROM accounts WHERE school_id = ? AND lower(trim(username)) = lower(trim(?))"
              ).bind(schoolId, o.match)
            );
            statements.push(
              db.prepare(
                "DELETE FROM app_sessions WHERE school_id = ? AND lower(trim(username)) = lower(trim(?))"
              ).bind(schoolId, o.match)
            );
          }
        }
      }

      if (statements.length === 0) return;

      // D1 batch executes inside an implicit SQLite transaction.
      // If any version guard affected 0 rows, abort immediately so D1 rolls back the entire batch.
      const results = await db.batch(statements);
      for (const g of guards) {
        const res = results[g.index];
        const changes = res && res.meta ? Number(res.meta.changes || 0) : 0;
        if (changes === 0) {
          throw new ConflictError(g.sheet);
        }
      }
    },
  };
}

export function encodeChanges(changes) {
  return changes.map((c) => ({
    ...c,
    set: c.set ? c.set.map(([n, cells]) => [n, encodeRows([cells])[0]]) : undefined,
  }));
}
