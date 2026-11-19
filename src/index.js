import { resolveTenant } from "./middleware/tenant.js";
import { handleGasApiRequest } from "./api/gas-router.js";
import { COMBINED_SCRIPT } from "./legacy_gs/index.js";

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
    },
  });
}

function publicSchoolInfo(tenant) {
  return {
    name: tenant.name,
    subdomain: tenant.subdomain,
    plan: tenant.plan,
    educationLevel: tenant.educationLevel,
    logoUrl: tenant.logoUrl,
    contactEmail: tenant.contactEmail,
    contactPhone: tenant.contactPhone,
    subscriptionExpiresAt: tenant.subscriptionExpiresAt,
  };
}

const PAGE_NAMES = [
  "Index", "Login", "Parent", "HeadOfSchool", "SecondMaster",
  "DeputeHeadmaster", "Academics", "Teacher", "SubjectTeacher",
  "ClassTeacher", "Accountant", "SchoolManager", "Admin",
  "LessonPlanPanel", "Staff",
  "SecIndex", "SecLogin", "SecParent", "SecAcademics", "SecClassTeacher"
];


const pageCache = new Map();
let pagesLoaded = false;

async function ensurePagesLoaded(env, origin) {
  if (pagesLoaded) return;
  await Promise.all(
    PAGE_NAMES.map(async (name) => {
      for (const p of [`/${name}`, `/${name}.html`, `/app/${name}`, `/app/${name}.html`]) {
        try {
          const res = await env.ASSETS.fetch(new Request(new URL(p, origin), { redirect: "follow" }));
          if (res.ok) {
            pageCache.set(name, await res.text());
            return;
          }
        } catch (e) {}
      }
    })
  );
  pagesLoaded = true;
}

function readPage(name) {
  if (!/^[A-Za-z]+$/.test(String(name))) {
    throw new Error("Requested page is unavailable.");
  }
  const html = pageCache.get(name);
  if (html === undefined) {
    throw new Error(`Page "${name}" is not available on this deployment.`);
  }
  return html;
}

function injectSchoolContext(html, tenant) {
  const ctx = {
    id: tenant.id,
    name: tenant.name,
    subdomain: tenant.subdomain,
    plan: tenant.plan,
    educationLevel: tenant.educationLevel,
    logoUrl: tenant.logoUrl,
  };
  const script = `<script>window.__SCHOOL__=${JSON.stringify(ctx)};</script><script src="/gas-compat.js"></script>`;
  return /<head[^>]*>/i.test(html)
    ? html.replace(/<head[^>]*>/i, (m) => m + script)
    : script + html;
}

function tenantErrorHtml(message) {
  return `<!DOCTYPE html><html lang="sw"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Shule Haipatikani</title><style>body{font-family:system-ui,sans-serif;background:#0f172a;color:#e2e8f0;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;padding:1rem;}.card{max-width:520px;background:#1e293b;border:1px solid #334155;border-radius:12px;padding:2rem;text-align:center;}h1{font-size:1.25rem;margin:0 0 1rem;color:#f87171;}p{line-height:1.6;color:#cbd5e1;}</style></head><body><div class="card"><h1>Tatizo la Akaunti ya Shule</h1><p>${String(message || "Shule haipatikani.").replace(/[<>&"]/g, "")}</p></div></body></html>`;
}

function landingFallbackHtml(env) {
  const domain = (env.MAIN_DOMAIN || "domainyako.com").trim();
  return `<!DOCTYPE html><html lang="sw"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>School Management SaaS</title><style>body{font-family:system-ui,sans-serif;background:#0f172a;color:#e2e8f0;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;padding:1rem;}.card{max-width:560px;background:#1e293b;border:1px solid #334155;border-radius:12px;padding:2rem;}h1{margin-top:0;color:#38bdf8;}</style></head><body><div class="card"><h1>School Management SaaS</h1><p>Mfumo wa shule (Pre, Primary, O-Level &amp; A-Level) upo tayari kwenye ${domain}.</p></div></body></html>`;
}

const RESERVED_SUBDOMAINS = new Set([
  "www", "api", "app", "admin", "superadmin", "mail", "support",
  "help", "blog", "status", "cdn", "dashboard", "login", "register",
]);

const EDUCATION_LEVELS = ["Pre-Primary", "Primary", "O-Level", "A-Level", "Combined"];
const PLANS = ["basic", "pro", "advanced"];

async function sha256Hex(text) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

function randomHex(bytes) {
  const arr = new Uint8Array(bytes);
  crypto.getRandomValues(arr);
  return Array.from(arr).map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function handleRegisterSchool(request, env) {
  let body;
  try {
    body = await request.json();
  } catch (e) {
    return json({ ok: false, error: "Muundo wa JSON si sahihi." }, 400);
  }

  if (env.REGISTRATION_KEY && String(body.registration_key || "") !== String(env.REGISTRATION_KEY)) {
    return json({ ok: false, error: "Registration key si sahihi." }, 403);
  }

  const name = String(body.name || body.school_name || "").trim();
  const subdomain = String(body.subdomain || "").trim().toLowerCase();
  const educationLevel = String(body.education_level || "O-Level").trim();
  const plan = String(body.plan || body.package || "basic").trim().toLowerCase();
  const adminUsername = String(body.admin_username || body.contact_email || body.email || "admin").trim();
  const adminPassword = String(body.admin_password || "");
  const adminFullName = String(body.admin_full_name || body.admin_name || "School Manager").trim();
  const contactEmail = String(body.contact_email || body.email || "").trim();
  const contactPhone = String(body.contact_phone || "").trim();
  const logoUrl = String(body.logo_url || "").trim();
  const maxStudents = Number(body.max_students) > 0 ? Math.min(Number(body.max_students), 100000) : 500;
  const trialDays = Number(body.trial_days) > 0 ? Number(body.trial_days) : 30;

  const errors = [];
  if (name.length < 3) errors.push("Jina la shule liwe na angalau herufi 3.");
  if (!/^[a-z0-9]([a-z0-9-]{1,28}[a-z0-9])?$/.test(subdomain)) {
    errors.push("Subdomain si sahihi: tumia herufi ndogo, namba na '-' pekee (3-30 characters).");
  }
  if (RESERVED_SUBDOMAINS.has(subdomain)) errors.push(`Subdomain '${subdomain}' imehifadhiwa na mfumo.`);
  if (!EDUCATION_LEVELS.includes(educationLevel)) {
    errors.push(`education_level lazima iwe mojawapo ya: ${EDUCATION_LEVELS.join(", ")}`);
  }
  if (!PLANS.includes(plan)) {
    errors.push(`plan lazima iwe mojawapo ya: ${PLANS.join(", ")}`);
  }
  if (adminUsername.length < 3) errors.push("admin_username liwe na angalau herufi 3.");
  if (adminPassword.length < 6) errors.push("admin_password liwe na angalau herufi 6.");
  if (errors.length > 0) return json({ ok: false, errors, error: errors[0] }, 400);

  const existing = await env.DB.prepare(
    "SELECT id FROM schools WHERE lower(trim(subdomain)) = ? LIMIT 1"
  ).bind(subdomain).first();
  if (existing) {
    return json({ ok: false, error: `Subdomain '${subdomain}' tayari inatumika na shule nyingine.` }, 409);
  }

  const schoolId = "sch_" + crypto.randomUUID().replace(/-/g, "").slice(0, 16);
  const accountId = "acc_" + crypto.randomUUID().replace(/-/g, "").slice(0, 16);
  const salt = randomHex(32);
  const passwordHash = await sha256Hex(adminPassword + salt);
  const expiresAt = new Date(Date.now() + trialDays * 86400000).toISOString();
  const profile = JSON.stringify({
    AdminID: "ADM" + Date.now().toString(36).toUpperCase(),
    CreatedBy: "registration-api",
    CreatedAt: new Date().toISOString().slice(0, 16).replace("T", " "),
  });

  try {
    await env.DB.batch([
      env.DB.prepare(
        `INSERT INTO schools (id, name, subdomain, education_level, plan, status, max_students, logo_url, contact_email, contact_phone, subscription_expires_at)
         VALUES (?, ?, ?, ?, ?, 'active', ?, ?, ?, ?, ?)`
      ).bind(schoolId, name, subdomain, educationLevel, plan, maxStudents, logoUrl, contactEmail, contactPhone, expiresAt),

      env.DB.prepare(
        `INSERT INTO sheet_meta (school_id, name, version) VALUES (?, '__accounts', 1)`
      ).bind(schoolId),

      env.DB.prepare(
        `INSERT INTO accounts (id, school_id, seq, username, full_name, password_hash, password_salt, role, position, status, profile)
         VALUES (?, ?, 1, ?, ?, ?, ?, 'schoolmanager', 'School Manager', 'Active', ?)`
      ).bind(accountId, schoolId, adminUsername, adminFullName, passwordHash, salt, profile),
    ]);
  } catch (e) {
    return json({ ok: false, error: "Usajili umeshindwa: " + e.message }, 500);
  }

  const mainDomain = (env.MAIN_DOMAIN || "").trim();
  const loginUrl = mainDomain
    ? `https://${subdomain}.${mainDomain}/`
    : `?school=${subdomain}`;

  return json({
    ok: true,
    message: `Shule '${name}' imesajiliwa kikamilifu kwenye kifurushi cha '${plan.toUpperCase()}'.`,
    school: {
      id: schoolId,
      name,
      subdomain,
      education_level: educationLevel,
      plan,
      status: "active",
      subscription_expires_at: expiresAt,
      login_url: loginUrl,
    },
    admin: { username: adminUsername, role: "schoolmanager" },
  }, 201);
}

const SECONDARY_LEVELS = ["O-LEVEL", "A-LEVEL", "SECONDARY", "COMBINED"];
function isSecondaryLevel(level) {
  if (!level) return false;
  const l = String(level).trim().toUpperCase();
  return SECONDARY_LEVELS.some((v) => l.includes(v));
}

const SEC_HTML_MAP = {
  "/Academics.html": "/SecAcademics.html",
  "/ClassTeacher.html": "/SecClassTeacher.html",
  "/Index.html": "/SecIndex.html",
  "/Login.html": "/SecLogin.html",
  "/Parent.html": "/SecParent.html",
};

function isSuperAdminAuthorized(request, env) {
  const expectedKey = String(env.SUPER_ADMIN_KEY || "admin12345").trim();
  const headerKey = String(request.headers.get("x-super-admin-key") || "").trim();
  const authHeader = String(request.headers.get("Authorization") || "").replace(/^Bearer\s+/i, "").trim();
  return headerKey === expectedKey || authHeader === expectedKey;
}

async function handleSuperAdminApi(request, env, url) {
  if (!isSuperAdminAuthorized(request, env)) {
    return json({ ok: false, error: "Super Admin Key si sahihi." }, 401);
  }
  const path = url.pathname;

  if (path === "/api/superadmin/schools" && request.method === "GET") {
    const limit = Math.min(Number(url.searchParams.get("limit")) || 500, 500);
    const { results } = await env.DB.prepare(
      "SELECT * FROM schools ORDER BY created_at DESC LIMIT ?"
    ).bind(limit).all();
    return json({ ok: true, schools: results || [] });
  }

  const statusMatch = path.match(/^\/api\/superadmin\/schools\/([^/]+)\/status$/);
  if (statusMatch && request.method === "PATCH") {
    const id = statusMatch[1];
    const body = await request.json().catch(() => ({}));
    const status = String(body.status || "").toLowerCase();
    if (!["active", "suspended", "expired"].includes(status)) {
      return json({ ok: false, error: "Status si sahihi." }, 400);
    }
    await env.DB.prepare(
      "UPDATE schools SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?"
    ).bind(status, id).run();
    return json({ ok: true });
  }

  const updateMatch = path.match(/^\/api\/superadmin\/schools\/([^/]+)$/);
  if (updateMatch && request.method === "PATCH") {
    const id = updateMatch[1];
    const body = await request.json().catch(() => ({}));
    const plan = String(body.plan || "").toLowerCase();
    const expiresAt = body.subscription_expires_at ? String(body.subscription_expires_at) : null;

    if (plan && !PLANS.includes(plan)) {
      return json({ ok: false, error: "Kifurushi si sahihi." }, 400);
    }
    if (plan && expiresAt) {
      await env.DB.prepare(
        "UPDATE schools SET plan = ?, subscription_expires_at = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?"
      ).bind(plan, expiresAt, id).run();
    } else if (plan) {
      await env.DB.prepare(
        "UPDATE schools SET plan = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?"
      ).bind(plan, id).run();
    } else if (expiresAt) {
      await env.DB.prepare(
        "UPDATE schools SET subscription_expires_at = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?"
      ).bind(expiresAt, id).run();
    }
    return json({ ok: true });
  }

  return json({ ok: false, error: "Njia ya Super Admin haipo." }, 404);
}

async function serveFrontend(request, env, url) {
  const path = url.pathname;
  const isAsset = /\.[a-z0-9]+$/i.test(path) && !path.toLowerCase().endsWith(".html");
  if (isAsset) return env.ASSETS.fetch(request);

  if (path.startsWith("/superadmin")) {
    const p = path.endsWith("/") ? path + "index.html" : path + (path.endsWith(".html") ? "" : "/index.html");
    const res = await env.ASSETS.fetch(new URL(p, url.origin));
    if (res.ok) return res;
  }

  if (path === "/register" || path === "/register.html") {
    const res = await env.ASSETS.fetch(new URL("/register.html", url.origin));
    if (res.ok) return res;
  }

  const tenantResult = await resolveTenant(request, env);

  if (!tenantResult.ok) {
    if (path === "/" || path === "") {
      const regRes = await env.ASSETS.fetch(new URL("/register.html", url.origin));
      if (regRes.ok) return regRes;
      return new Response(landingFallbackHtml(env), {
        status: 200,
        headers: { "Content-Type": "text/html; charset=utf-8" },
      });
    }
    return new Response(tenantErrorHtml(tenantResult.error), {
      status: tenantResult.status || 404,
      headers: { "Content-Type": "text/html; charset=utf-8" },
    });
  }

  let pagePath = path;
  if (pagePath === "/" || pagePath === "") {
    pagePath = "/Index.html";
  } else if (pagePath.endsWith("/")) {
    pagePath = pagePath + "Index.html";
  } else if (!pagePath.toLowerCase().endsWith(".html")) {
    const seg = pagePath.slice(1);
    pagePath = "/" + seg.charAt(0).toUpperCase() + seg.slice(1) + ".html";
  }

  if (isSecondaryLevel(tenantResult.tenant.educationLevel) && SEC_HTML_MAP[pagePath]) {
    pagePath = SEC_HTML_MAP[pagePath];
  }

  let res = null;
  if (!pagePath.startsWith("/app/")) {
    res = await env.ASSETS.fetch(new URL("/app" + pagePath, url.origin));
  }
  if (!res || !res.ok) {
    res = await env.ASSETS.fetch(new URL(pagePath, url.origin));
  }
  if (!res.ok) return new Response("Page not found.", { status: 404 });

  const html = await res.text();
  const injected = injectSchoolContext(html, tenantResult.tenant);
  return new Response(injected, {
    status: 200,
    headers: { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store" },
  });
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const path = url.pathname;

    try {
      if (path === "/api/health") {
        return json({ ok: true, service: "new_cloudflare_saas", time: new Date().toISOString() });
      }
      if (path === "/api/check-subdomain" && request.method === "GET") {
        const sub = String(url.searchParams.get("subdomain") || "").trim().toLowerCase();
        if (!/^[a-z0-9]([a-z0-9-]{1,28}[a-z0-9])?$/.test(sub) || RESERVED_SUBDOMAINS.has(sub)) {
          return json({ ok: true, available: false });
        }
        const exists = await env.DB.prepare(
          "SELECT id FROM schools WHERE lower(trim(subdomain)) = ? LIMIT 1"
        ).bind(sub).first();
        return json({ ok: true, available: !exists });
      }
      if (path === "/api/register-school") {
        if (request.method !== "POST") return json({ error: "Method not allowed." }, 405);
        return await handleRegisterSchool(request, env);
      }
      if (path.startsWith("/api/superadmin/")) {
        return await handleSuperAdminApi(request, env, url);
      }
      if (path === "/api/school-info" && request.method === "GET") {
        const t = await resolveTenant(request, env);
        if (!t.ok) return json({ ok: false, error: t.error }, t.status);
        return json({ ok: true, school: publicSchoolInfo(t.tenant) });
      }
      if (path === "/api/gas") {
        if (request.method !== "POST") return json({ error: "Method not allowed." }, 405);
        const t = await resolveTenant(request, env);
        if (!t.ok) return json({ ok: false, error: t.error }, t.status);
        await ensurePagesLoaded(env, url.origin);
        return await handleGasApiRequest({ request, env, tenant: t.tenant, readPage, combinedScriptCode: COMBINED_SCRIPT });
      }
      return await serveFrontend(request, env, url);
    } catch (err) {
      return json({ ok: false, error: err?.message || "Hitilafu ya ndani ya seva." }, 500);
    }
  },
};
