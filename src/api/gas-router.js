// API Router & Plan Feature Guard
// Inashughulikia maombi ya POST /api/gas, inakagua Vifurushi (Basic, Pro, Advanced), Roles, na Session Tokens.

import { createD1Store } from "../emulator/store.js";
import { runLegacy } from "../emulator/runtime.js";

async function hashSessionToken(token) {
  const msgUint8 = new TextEncoder().encode(token);
  const hashBuffer = await crypto.subtle.digest("SHA-256", msgUint8);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
}

const PLAN_TIERS = {
  basic: 1,
  pro: 2,
  advanced: 3,
};

const METHOD_MIN_PLANS = {
  checkLogin: "basic",
  smsRestoreSession: "basic",
  smsEndSession: "basic",
  getLoginPage: "basic",
  getIndexPage: "basic",
  logout: "basic",
  getTeacherSession: "basic",
  changeTeacherPassword: "basic",
  psCall: "basic",
  getAcademicsData: "basic",
  saveStudentScores: "basic",
  getClassTeacherData: "basic",
  getHeadOfSchoolData: "basic",
  getSchoolManagerData: "basic",

  getTimetableData: "pro",
  saveTimetableData: "pro",
  generateTimetable: "pro",
  getLessonPlans: "pro",
  saveLessonPlan: "pro",
  approveLessonPlan: "pro",
  getSyllabusData: "pro",
  getSecondMasterData: "pro",
  saveDisciplineRecord: "pro",
  getTodReports: "pro",
  saveTodReport: "pro",
  getAccountantData: "pro",
  saveFeePayment: "pro",
  getFinancialSummary: "pro",
  sendSmsNotification: "pro",
  getSmsLogs: "pro",

  getParentPortalHtml: "advanced",
  getParentStudentData: "advanced",
  getParentFeeStatus: "advanced",
  getParentReportCard: "advanced",
  getAdvancedAnalytics: "advanced",
  getPerformanceTrends: "advanced",
};

const METHOD_ROLES = {
  getAcademicsData: ["academic", "academicmaster", "headofschool", "headteacher", "schoolmanager"],
  saveStudentScores: ["teacher", "classteacher", "academic", "academicmaster", "headofschool", "headteacher"],
  getClassTeacherData: ["classteacher", "headofschool", "headteacher", "schoolmanager"],
  getHeadOfSchoolData: ["headofschool", "headteacher", "schoolmanager"],
  getSchoolManagerData: ["schoolmanager"],
  getTimetableData: ["secondmaster", "academic", "academicmaster", "headofschool", "headteacher", "teacher", "schoolmanager"],
  saveTimetableData: ["secondmaster", "academicmaster", "schoolmanager"],
  getLessonPlans: ["teacher", "academic", "academicmaster", "headofschool", "headteacher"],
  saveLessonPlan: ["teacher"],
  approveLessonPlan: ["academic", "academicmaster", "headofschool", "headteacher"],
  getSecondMasterData: ["secondmaster", "discipline", "headofschool", "headteacher", "schoolmanager"],
  saveDisciplineRecord: ["secondmaster", "discipline", "headofschool", "headteacher"],
  getTodReports: ["teacher", "secondmaster", "discipline", "headofschool", "headteacher"],
  saveTodReport: ["teacher", "secondmaster", "discipline"],
  getAccountantData: ["accountant", "bursar", "schoolmanager", "headofschool", "headteacher"],
  saveFeePayment: ["accountant", "bursar"],
  getFinancialSummary: ["accountant", "bursar", "schoolmanager", "headofschool", "headteacher"],
  sendSmsNotification: ["schoolmanager", "headofschool", "headteacher", "secondmaster", "academic", "academicmaster"],
  getParentPortalHtml: ["parent", "schoolmanager", "headofschool", "headteacher"],
  getParentStudentData: ["parent"],
  getParentFeeStatus: ["parent"],
  getParentReportCard: ["parent"],
  getAdvancedAnalytics: ["schoolmanager", "headofschool", "headteacher", "academic", "academicmaster"],
  changeTeacherPassword: [
    "academic", "academicmaster", "classteacher", "teacher", "staff",
    "accountant", "bursar", "headofschool", "headteacher", "secondmaster",
    "discipline", "librarian", "storekeeper", "patron", "matron", "schoolmanager"
  ],
  psCall: [
    "schoolmanager", "headofschool", "headteacher", "accountant", "bursar",
    "academic", "academicmaster", "classteacher", "teacher", "secondmaster", "discipline"
  ],
};

function getDashboardPageName(role, educationLevel) {
  const isSecondary = /o-?level|a-?level|secondary/i.test(String(educationLevel || ""));
  const roleLower = String(role || "").toLowerCase();

  const ROLE_PAGES = {
    schoolmanager: "SchoolManager",
    headofschool: "HeadOfSchool",
    headteacher: "HeadOfSchool",
    secondmaster: "DeputeHeadmaster",
    discipline: "DeputeHeadmaster",
    academic: isSecondary ? "SecAcademics" : "Academics",
    academicmaster: isSecondary ? "SecAcademics" : "Academics",
    classteacher: isSecondary ? "SecClassTeacher" : "ClassTeacher",
    teacher: isSecondary ? "SubjectTeacher" : "Teacher",
    accountant: "Accountant",
    bursar: "Accountant",
    admin: "Admin",
    parent: isSecondary ? "SecParent" : "Parent",
    student: isSecondary ? "SecParent" : "Parent",
  };

  return ROLE_PAGES[roleLower] || "SchoolManager";
}

function injectDashboardBootScript(html, tenant, rawToken) {
  const level = tenant.educationLevel || tenant.education_level || "O-Level";
  const ctx = {
    id: tenant.id,
    name: tenant.name,
    subdomain: tenant.subdomain,
    plan: tenant.plan,
    educationLevel: level,
    education_level: level,
    logoUrl: tenant.logoUrl,
  };
  const bootScript = `<script>
    window.__SCHOOL__=${JSON.stringify(ctx)};
    localStorage.setItem("sms_token", ${JSON.stringify(rawToken)});
  </script><script src="/gas-compat.js"></script>`;

  return /<head[^>]*>/i.test(html)
    ? html.replace(/<head[^>]*>/i, (m) => m + bootScript)
    : bootScript + html;
}

export async function handleGasApiRequest({ request, env, tenant, readPage, combinedScriptCode }) {
  let body = {};
  try {
    body = await request.json();
  } catch (e) {
    return new Response(JSON.stringify({ error: "Muundo mbaya wa JSON payload." }), {
      status: 400,
      headers: { "Content-Type": "application/json" },
    });
  }

  const method = typeof body.method === "string" ? body.method.trim() : "";
  const args = Array.isArray(body.args) ? body.args : [];
  const token = typeof body.token === "string" ? body.token.trim() : "";

  if (!method || !/^[A-Za-z_$][\w$]{0,99}$/.test(method) || method.endsWith("_")) {
    return new Response(JSON.stringify({ error: "Njia ya API (method) haitambuliki au si sahihi." }), {
      status: 400,
      headers: { "Content-Type": "application/json" },
    });
  }

  const requiredPlan = METHOD_MIN_PLANS[method] || "basic";
  const tenantPlanLevel = PLAN_TIERS[tenant.plan] || 1;
  const requiredPlanLevel = PLAN_TIERS[requiredPlan] || 1;

  if (tenantPlanLevel < requiredPlanLevel) {
    return new Response(
      JSON.stringify({
        error: `Kipengele hiki kinahitaji Kifurushi cha '${requiredPlan.toUpperCase()}'. Kifurushi cha shule yako kwa sasa ni '${tenant.plan.toUpperCase()}'. Tafadhali fanya upgrade ya kifurushi.`,
        requiredPlan,
        currentPlan: tenant.plan,
      }),
      { status: 403, headers: { "Content-Type": "application/json" } }
    );
  }

  const store = createD1Store(env.DB);
  const appUrl = request.url;
  const schoolContext = {
    ...tenant,
    education_level: tenant.educationLevel || tenant.education_level || "O-Level",
  };

  if (method === "checkLogin") {
    return await handleLogin(env, store, schoolContext, args, readPage, combinedScriptCode, appUrl);
  }

  if (method === "smsRestoreSession") {
    return await handleRestoreSession(env, store, schoolContext, token, readPage, combinedScriptCode, appUrl);
  }

  if (method === "smsEndSession" || method === "logout") {
    if (token) {
      const tokenHash = await hashSessionToken(token);
      await env.DB.prepare(
        "DELETE FROM app_sessions WHERE school_id = ? AND token_hash = ?"
      ).bind(tenant.id, tokenHash).run();
    }
    return new Response(JSON.stringify({ result: true }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  }

  if (method === "getLoginPage" || method === "getIndexPage" || method === "getParentPortalHtml") {
    const isSecondary = /o-?level|a-?level|secondary/i.test(String(schoolContext.education_level || ""));
    let fallbackPage = "Login";
    if (method === "getIndexPage") fallbackPage = isSecondary ? "SecIndex" : "Index";
    else if (method === "getParentPortalHtml") fallbackPage = isSecondary ? "SecParent" : "Parent";
    else fallbackPage = isSecondary ? "SecLogin" : "Login";

    try {
      const htmlResult = await runLegacy({
        store,
        schoolId: tenant.id,
        school: schoolContext,
        fn: method,
        args: [],
        session: null,
        readPage,
        combinedScriptCode,
        appUrl,
        timezone: env.DEFAULT_TIMEZONE,
      });
      return new Response(JSON.stringify({ result: htmlResult }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    } catch (e) {
      return new Response(JSON.stringify({ result: readPage(fallbackPage) }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    }
  }

  if (!token) {
    return new Response(JSON.stringify({ error: "Muda wako wa kutumia mfumo umeisha. Tafadhali ingia tena (Missing Token)." }), {
      status: 401,
      headers: { "Content-Type": "application/json" },
    });
  }

  const tokenHash = await hashSessionToken(token);
  const accountStmt = env.DB.prepare(
    `SELECT a.username, a.full_name, a.role, a.position
     FROM app_sessions s 
     JOIN accounts a ON a.school_id = s.school_id AND lower(trim(a.username)) = lower(trim(s.username))
     WHERE s.school_id = ? AND s.token_hash = ? AND s.expires_at > ? AND lower(a.status) = 'active' AND s.role = a.role
     LIMIT 1`
  );
  const account = await accountStmt.bind(tenant.id, tokenHash, new Date().toISOString()).first();

  if (!account) {
    return new Response(JSON.stringify({ error: "Muda wako wa session umeisha au akaunti yako imefungwa. Tafadhali ingia tena." }), {
      status: 401,
      headers: { "Content-Type": "application/json" },
    });
  }

  const allowedRoles = METHOD_ROLES[method];
  if (allowedRoles && !allowedRoles.includes(account.role) && account.role !== "superadmin") {
    return new Response(
      JSON.stringify({ error: `Huruhusiwi kutekeleza kitendo hiki. Role yako (${account.role}) haina ruhusa.` }),
      { status: 403, headers: { "Content-Type": "application/json" } }
    );
  }

  const session = { username: account.username, role: account.role };

  try {
    const result = await runLegacy({
      store,
      schoolId: tenant.id,
      school: schoolContext,
      fn: method,
      args,
      session,
      readPage,
      combinedScriptCode,
      appUrl,
      timezone: env.DEFAULT_TIMEZONE,
    });

    return new Response(JSON.stringify({ result: result === undefined ? null : result }), {
      status: 200,
      headers: { "Content-Type": "application/json; charset=utf-8" },
    });
  } catch (error) {
    return new Response(
      JSON.stringify({ error: error.message || "Hitilafu imetokea wakati wa kutekeleza maombi yako." }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }
}

async function handleLogin(env, store, tenant, args, readPage, combinedScriptCode, appUrl) {
  try {
    const loginId = String(args.at(0) || "").trim();
    const password = String(args.at(1) || "");

    if (!loginId || !password) {
      return new Response(JSON.stringify({ result: { status: "fail", message: "Jina la mtumiaji/Email na neno la siri vinatakiwa." } }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    }

    const accStmt = env.DB.prepare(
      `SELECT a.username, a.full_name, a.password_hash, a.password_salt, a.role, a.position, a.status 
       FROM accounts a
       JOIN schools s ON s.id = a.school_id
       WHERE a.school_id = ? 
         AND (
           lower(trim(a.username)) = lower(trim(?)) 
           OR (lower(trim(s.contact_email)) = lower(trim(?)) AND lower(a.role) = 'schoolmanager')
         )
       LIMIT 1`
    );
    const account = await accStmt.bind(tenant.id, loginId, loginId).first();

    if (!account) {
      return new Response(JSON.stringify({ result: { status: "fail", message: "Jina la mtumiaji/Email au neno la siri si sahihi." } }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    }

    if (String(account.status).toLowerCase() !== "active") {
      return new Response(JSON.stringify({ result: { status: "fail", message: "Akaunti hii imesimamishwa." } }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    }

    const salt = account.password_salt;
    const msgUint8 = new TextEncoder().encode(password + salt);
    const hashBuffer = await crypto.subtle.digest("SHA-256", msgUint8);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    const calculatedHash = hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");

    if (calculatedHash !== account.password_hash) {
      return new Response(JSON.stringify({ result: { status: "fail", message: "Jina la mtumiaji/Email au neno la siri si sahihi." } }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    }

    const rawToken = crypto.randomUUID() + crypto.randomUUID().replace(/-/g, "");
    const tokenHash = await hashSessionToken(rawToken);
    const expiresAt = new Date(Date.now() + 12 * 3600 * 1000).toISOString();

    await env.DB.prepare(
      "INSERT INTO app_sessions (token_hash, school_id, username, role, expires_at) VALUES (?, ?, ?, ?, ?)"
    ).bind(tokenHash, tenant.id, account.username, account.role, expiresAt).run();

    const pageName = getDashboardPageName(account.role, tenant.education_level);
    let dashboardHtml = "";
    try {
      dashboardHtml = injectDashboardBootScript(readPage(pageName), tenant, rawToken);
    } catch (pageErr) {
      console.error("Dashboard page load error:", pageErr);
    }

    const legacyPayload = {
      status: "success",
      ok: true,
      username: account.username,
      fullName: account.full_name || account.username,
      role: account.role,
      position: account.position || "School Manager",
      sessionToken: rawToken,
      html: dashboardHtml,
    };

    return new Response(JSON.stringify({ result: legacyPayload }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    console.error("Login Error:", err);
    return new Response(JSON.stringify({ error: "Login Error: " + (err?.message || String(err)) }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }
}

async function handleRestoreSession(env, store, tenant, token, readPage, combinedScriptCode, appUrl) {
  if (!token) {
    return new Response(JSON.stringify({ result: { ok: false } }), { status: 200, headers: { "Content-Type": "application/json" } });
  }

  const tokenHash = await hashSessionToken(token);
  const account = await env.DB.prepare(
    `SELECT a.username, a.full_name, a.role, a.position 
     FROM app_sessions s 
     JOIN accounts a ON a.school_id = s.school_id AND lower(trim(a.username)) = lower(trim(s.username))
     WHERE s.school_id = ? AND s.token_hash = ? AND s.expires_at > ? AND lower(a.status) = 'active' LIMIT 1`
  ).bind(tenant.id, tokenHash, new Date().toISOString()).first();

  if (!account) {
    return new Response(JSON.stringify({ result: { ok: false } }), { status: 200, headers: { "Content-Type": "application/json" } });
  }

  const pageName = getDashboardPageName(account.role, tenant.education_level);
  let dashboardHtml = "";
  try {
    dashboardHtml = injectDashboardBootScript(readPage(pageName), tenant, token);
  } catch (e) {}

  return new Response(
    JSON.stringify({
      result: {
        ok: true,
        status: "success",
        username: account.username,
        fullName: account.full_name || account.username,
        role: account.role,
        position: account.position || "School Manager",
        sessionToken: token,
        html: dashboardHtml,
      },
    }),
    { status: 200, headers: { "Content-Type": "application/json" } }
  );
}
