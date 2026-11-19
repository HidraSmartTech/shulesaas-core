// Middleware ya Kutambua Shule (Tenant Isolation Middleware)
// Inachambua Subdomain au Query Parameter (?school=slug) na kutafuta shule kwenye Cloudflare D1

const PLATFORM_RESERVED = new Set([
  "www", "api", "app", "admin", "superadmin", "mail", "support",
  "help", "blog", "status", "cdn", "dashboard", "login", "register"
]);

export async function resolveTenant(request, env) {
  const url = new URL(request.url);
  const hostname = url.hostname.toLowerCase();

  let subdomain = null;

  // 1. Angalia Query Parameter (?school=feza au ?tenant=feza)
  const schoolParam = url.searchParams.get("school") || url.searchParams.get("tenant");
  if (schoolParam) {
    subdomain = schoolParam.trim().toLowerCase();
  } else {
    // 2. Angalia Subdomain kutoka kwenye URL
    const mainDomain = (env.MAIN_DOMAIN || "").toLowerCase().trim();

    if (mainDomain && hostname.endsWith("." + mainDomain)) {
      const prefix = hostname.slice(0, -(mainDomain.length + 1));
      if (prefix && !PLATFORM_RESERVED.has(prefix) && !prefix.includes(".")) {
        subdomain = prefix;
      }
    } else if (!hostname.endsWith(".workers.dev") && !hostname.endsWith(".pages.dev") && hostname !== "localhost") {
      // Custom domain fallback (e.g., shule.domainyako.co.tz)
      const parts = hostname.split(".");
      if (parts.length >= 3 && !PLATFORM_RESERVED.has(parts)) {
        subdomain = parts;
      }
    }
  }

  if (!subdomain) {
    return {
      ok: false,
      status: 400,
      error: "Tafadhali tumia Subdomain ya shule yako (mfano: https://shulesaas.co.tz) au weka parameter ya '?school=slug' kwenye URL.",
    };
  }

  // 3. Tafuta Shule kwenye Database ya Cloudflare D1
  try {
    const stmt = env.DB.prepare(
      `SELECT id, name, subdomain, education_level, plan, status, max_students,
              logo_url, contact_email, contact_phone, subscription_expires_at
       FROM schools WHERE lower(trim(subdomain)) = ? LIMIT 1`
    );
    const school = await stmt.bind(subdomain.toLowerCase()).first();

    if (!school) {
      return {
        ok: false,
        status: 404,
        error: `Shule yenye anwani ya '${subdomain}' haijasajiliwa au haipatikani kwenye mfumo.`,
      };
    }

    // 4. Angalia Hali ya Akaunti ya Shule (Status Check)
    const statusLower = String(school.status || "active").toLowerCase();
    if (statusLower === "suspended") {
      return {
        ok: false,
        status: 403,
        error: `Akaunti ya shule ya '${school.name}' imesimamishwa (Suspended). Tafadhali wasiliana na Msimamizi wa Mfumo.`,
      };
    }

    // 5. Angalia Tarehe ya Kuisha kwa Malipo (Subscription Expiry Check)
    if (school.subscription_expires_at) {
      const expiryDate = new Date(school.subscription_expires_at).getTime();
      const currentDate = Date.now();

      if (!Number.isNaN(expiryDate) && expiryDate < currentDate) {
        return {
          ok: false,
          status: 402,
          error: `Muda wa matumizi wa kifurushi cha shule ya '${school.name}' umeisha tarehe ${school.subscription_expires_at}. Tafadhali fanya malipo ili kuendelea kutumia mfumo.`,
        };
      }
    }

    const eduLevel = school.education_level || "O-Level";
    const lvlUpper = eduLevel.toUpperCase();
    const isPrimary = lvlUpper.includes("PRIMARY") || lvlUpper.includes("PRE") || lvlUpper === "NURSERY";

    return {
      ok: true,
      tenant: {
        id: school.id,
        name: school.name,
        subdomain: school.subdomain,
        educationLevel: eduLevel,
        education_level: eduLevel,
        isPrimary,
        isSecondary: !isPrimary,
        plan: (school.plan || "basic").toLowerCase(),
        status: statusLower,
        maxStudents: school.max_students || 500,
        logoUrl: school.logo_url || "",
        contactEmail: school.contact_email || "",
        contactPhone: school.contact_phone || "",
        subscriptionExpiresAt: school.subscription_expires_at,
      },
    };
  } catch (error) {
    return {
      ok: false,
      status: 500,
      error: `Hitilafu ya Server wakati wa kutafuta shule: ${error.message}`,
    };
  }
}
