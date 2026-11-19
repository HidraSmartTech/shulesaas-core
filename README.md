# ShuleSaaS Core (`shulesaas-core`)

Multi-Tenant School Management SaaS built on **Cloudflare Workers** and **Cloudflare D1 (SQLite)**.

## Architecture & Key Files
- **`src/index.js`**: Main Cloudflare Worker entry point, routing, authentication, and Super Admin endpoints.
- **`src/emulator/`**: Google Apps Script (GAS) V8 runtime emulator (`runtime.js`, `sheets.js`, `store.js`) isolated per `school_id` in D1.
- **`src/legacy_gs/`**: Legacy Google Apps Script modules (`core/`, `academics/`, `operations/`, `portal/`).
- **`build.js`**: Build script that compiles all `.gs` files into `src/legacy_gs/compiled.js` (avoids `new Function()` / `eval()` restrictions in Cloudflare Workers).
- **`migrations/`**: D1 database schema migrations.
- **`public/`**: Frontend HTML/JS assets (`index.html`, `register.html`, `/superadmin/`).

## Rules for Contributors & AI Assistants
1. **Do NOT manually edit `src/legacy_gs/compiled.js`**. Always edit `build.js` or the source `.gs` files inside `src/legacy_gs/`, then run `node build.js`.
2. **Strict Mode Compatibility**: All `.gs` scripts run inside an ES Module sandbox (`"use strict"`). Ensure any undeclared global variables or missing GAS globals are declared in `build.js`.
3. **Multi-Tenancy**: Every database query in `src/emulator/store.js` must remain strictly scoped by `school_id`.

## Deployment
Connected to Cloudflare Workers Auto-Deploy via GitHub:
- **Build command**: `node build.js`
- **Deploy command**: `npx wrangler deploy`
