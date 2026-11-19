// Middleware ya Kutambua Shule (Tenant Isolation Middleware)
// Inachambua Subdomain au Query Parameter (?school=slug) na kutafuta shule kwenye Cloudflare D1

export async function resolveTenant(request, env) {
  const url = new URL(request.url);
  const hostname = url.hostname.toLowerCase();
  
  let subdomain = null;

  // 1. Angalia kama kuna Query Parameter (?school=feza au ?tenant=feza) kwa ajili ya majaribio
  const schoolParam = url.searchParams.get("school") || url.searchParams.get("tenant");
  if (schoolParam) {
    subdomain = schoolParam.trim().toLowerCase();
  } else {
    // 2. Angalia Subdomain kutoka kwenye URL
    const mainDomain = (env.MAIN_DOMAIN || "").toLowerCase().trim();
    
    if (mainDomain && hostname.endsWith("." + mainDomain)) {
      const prefix = hostname.replace("." + mainDomain, "");
      if (prefix && prefix !== "www" && !prefix.includes(".")) {
        subdomain = prefix;
      }
    } else {
      // Direct IP / workers.dev / localhost handling
      const parts = hostname.split(".");
      if (parts.length >= 3 && parts[0] !== "www") {
        subdomain = parts[0];
      }
    }
  }

  // Ikiwa hakuna subdomain au parameter iliyopatikana
  if (!subdomain) {
    return {
      ok: false,
      status: 400,
      error: "Tafadhali tumia Subdomain ya shule yako (mfano: https://shule.domainyako.com) au weka parameter ya '?school=slug' kwenye URL.",
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
        error: `Shule yenye anwani ya '${subdomain}' haijajajajiliwa au haipatikani kwenye mfumo.`,
      };
    }

    // 4. Angalia Hali ya Akaunti ya Shule (Status Check)
    if (school.status === "suspended") {
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
      
      if (!isNaN(expiryDate) && expiryDate < currentDate) {
        return {
          ok: false,
          status: 402,
          error: `Muda wa matumizi wa kifurushi cha shule ya '${school.name}' umeisha tarehe ${school.subscription_expires_at}. Tafadhali fanya malipo ili kuendelea kutumia mfumo.`,
        };
      }
    }

    // Tenant imethibitishwa kikamilifu
    return {
      ok: true,
      tenant: {
        id: school.id,
        name: school.name,
        subdomain: school.subdomain,
        educationLevel: school.education_level || "O-Level",
        plan: (school.plan || "basic").toLowerCase(),
        status: school.status || "active",
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