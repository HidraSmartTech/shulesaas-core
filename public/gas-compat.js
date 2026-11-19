/* ============================================================================
 * public/gas-compat.js — Google Apps Script client compatibility layer
 * kwa ajili ya Multi-Tenant SaaS kwenye Cloudflare (Pages + Workers + D1)
 *
 * Inaigiza google.script.run ya GAS, na kila ombi hubeba:
 *   - Subdomain (moja kwa moja kupitia Host header), AU
 *   - ?school=<subdomain> kwa ajili ya majaribio kwenye workers.dev
 * pamoja na session token (localStorage: sms_token).
 * ========================================================================== */
(function () {
  "use strict";

  // --------------------------------------------------------------------------
  // 1. Tambua Muktadha wa Shule (School Context)
  // --------------------------------------------------------------------------
  var queryString = new URLSearchParams(window.location.search);
  var schoolParam = (queryString.get("school") || queryString.get("tenant") || "").trim().toLowerCase();
  var injectedSchool = window.__SCHOOL__ || null;

  function detectSubdomain() {
    if (schoolParam) return schoolParam;
    if (injectedSchool && injectedSchool.subdomain) return String(injectedSchool.subdomain).toLowerCase();
    var host = window.location.hostname.toLowerCase();
    var parts = host.split(".");
    if (parts.length >= 3 && parts[0] !== "www") return parts[0];
    return "";
  }

  var SCHOOL_SUBDOMAIN = detectSubdomain();

  function apiUrl() {
    // Subdomain hutumika moja kwa moja kupitia hostname. Tunaweka ?school=
    // tu kama ilivyopo kwenye URL ya ukurasa (majaribio ya workers.dev).
    return "/api/gas" + (schoolParam ? "?school=" + encodeURIComponent(schoolParam) : "");
  }

  // --------------------------------------------------------------------------
  // 2. Usimamizi wa Session Token
  // --------------------------------------------------------------------------
  var TOKEN_KEY = "sms_token";

  function getToken() {
    try { return localStorage.getItem(TOKEN_KEY) || ""; } catch (e) { return ""; }
  }
  function setToken(token) {
    try { localStorage.setItem(TOKEN_KEY, String(token)); } catch (e) { /* ignore */ }
  }
  function clearToken() {
    try {
      localStorage.removeItem(TOKEN_KEY);
      // Futa pia pane-state za zamani (kama ilivyokuwa kwenye old_primary)
      for (var i = localStorage.length - 1; i >= 0; i--) {
        var k = localStorage.key(i);
        if (k && k.indexOf("sms_pane_") === 0) localStorage.removeItem(k);
      }
    } catch (e) { /* ignore */ }
  }

  // --------------------------------------------------------------------------
  // 3. Msimbo wa Kuwasiliana na Server (fetch wrapper yenye Promise)
  // --------------------------------------------------------------------------
  var REQUEST_TIMEOUT_MS = 60000;

  function serverCall(method, args) {
    return new Promise(function (resolve, reject) {
      var controller = typeof AbortController !== "undefined" ? new AbortController() : null;
      var timer = controller
        ? setTimeout(function () { controller.abort(); }, REQUEST_TIMEOUT_MS)
        : null;

      fetch(apiUrl(), {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        credentials: "same-origin",
        body: JSON.stringify({
          method: String(method),
          args: Array.isArray(args) ? args : [],
          token: getToken(),
        }),
        signal: controller ? controller.signal : undefined,
      })
        .then(function (res) {
          if (timer) clearTimeout(timer);
          return res
            .json()
            .catch(function () { return {}; })
            .then(function (data) { return { status: res.status, data: data }; });
        })
        .then(function (r) {
          var data = r.data || {};

          if (r.status === 401) {
            clearToken();
            try {
              window.dispatchEvent(new CustomEvent("sms:session-expired"));
            } catch (e) { /* ignore */ }
            reject(new Error(data.error || "Muda wa session umeisha. Tafadhali ingia tena."));
            return;
          }

          if (r.status >= 400) {
            reject(new Error(data.error || "Ombi halijakamilika (" + r.status + ")."));
            return;
          }

          var result = data.result === undefined ? null : data.result;

          // Ikiwa server imetoa session token mpya (mfano baada ya login), ihifadhi
          if (result && typeof result === "object" && result.sessionToken) {
            setToken(result.sessionToken);
          }

          resolve(result);
        })
        .catch(function (err) {
          if (timer) clearTimeout(timer);
          if (err && err.name === "AbortError") {
            reject(new Error("Ombi limechukua muda mrefu. Angalia mtandao wako kisha jaribu tena."));
          } else {
            reject(err instanceof Error ? err : new Error(String(err)));
          }
        });
    });
  }

  // --------------------------------------------------------------------------
  // 4. google.script.run Emulation (Proxy inayokubali method yoyote)
  // --------------------------------------------------------------------------
  var runState = {
    _success: null,
    _failure: null,
    withSuccessHandler: function (fn) {
      this._success = typeof fn === "function" ? fn : null;
      return runProxy;
    },
    withFailureHandler: function (fn) {
      this._failure = typeof fn === "function" ? fn : null;
      return runProxy;
    },
  };

  var runProxy = new Proxy(runState, {
    get: function (target, prop) {
      if (prop === "withSuccessHandler" || prop === "withFailureHandler") {
        return target[prop].bind(target);
      }
      if (typeof prop !== "string" || prop === "_success" || prop === "_failure") {
        return undefined;
      }
      // Jina lolote jengine linachukuliwa kama method ya server
      return function () {
        var args = Array.prototype.slice.call(arguments);
        var onSuccess = target._success;
        var onFailure = target._failure;
        target._success = null;
        target._failure = null;

        serverCall(prop, args).then(
          function (result) {
            if (onSuccess) {
              try { onSuccess(result); } catch (e) { console.error("Success handler error:", e); }
            }
          },
          function (err) {
            if (onFailure) {
              try { onFailure(err); } catch (e) { console.error("Failure handler error:", e); }
            } else {
              console.error("google.script.run." + prop + " imeshindwa:", err && err.message ? err.message : err);
            }
          }
        );
      };
    },
  });

  // --------------------------------------------------------------------------
  // 5. Usakinishaji wa window.google.script (run, url, host)
  // --------------------------------------------------------------------------
  window.google = window.google || {};
  window.google.script = window.google.script || {};

  window.google.script.run = runProxy;

  window.google.script.url = {
    getLocation: function (callback) {
      var location = {
        hash: window.location.hash ? window.location.hash.substring(1) : "",
        parameter: {},
        parameters: {},
      };
      queryString.forEach(function (value, key) {
        location.parameter[key] = value;
        if (!location.parameters[key]) location.parameters[key] = [];
        location.parameters[key].push(value);
      });
      if (typeof callback === "function") callback(location);
      return location;
    },
  };

  window.google.script.host = {
    close: function () { window.close(); },
    setWidth: function () { /* no-op kwenye Cloudflare SaaS */ },
    setHeight: function () { /* no-op kwenye Cloudflare SaaS */ },
    editor: { focus: function () { /* no-op */ } },
  };

  // --------------------------------------------------------------------------
  // 6. API ya Kujifungia kwa Matumizi ya Kurasa (Promise-based helpers)
  // --------------------------------------------------------------------------
  window.SchoolSaaS = {
    call: serverCall,
    getToken: getToken,
    setToken: setToken,
    clearToken: clearToken,
    getSchool: function () {
      if (injectedSchool) return injectedSchool;
      return { subdomain: SCHOOL_SUBDOMAIN };
    },
    getSubdomain: function () { return SCHOOL_SUBDOMAIN; },
    getPlan: function () { return injectedSchool && injectedSchool.plan ? injectedSchool.plan : ""; },
    logout: function () {
      return serverCall("smsEndSession", [])
        .catch(function () { /* ignore */ })
        .then(function () {
          clearToken();
          var base = window.location.pathname.split("/")[1] || "";
          window.location.href = "/" + (schoolParam ? "?school=" + encodeURIComponent(schoolParam) : "");
        });
    },
    fetchSchoolInfo: function () {
      return fetch("/api/school-info" + (schoolParam ? "?school=" + encodeURIComponent(schoolParam) : ""))
        .then(function (res) { return res.json(); });
    },
  };
})();