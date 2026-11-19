/* ============================================================================
 * public/gas-compat.js — Google Apps Script client compatibility layer
 * kwa ajili ya Multi-Tenant SaaS kwenye Cloudflare (Pages + Workers + D1)
 * ========================================================================== */
(function () {
  "use strict";

  var queryString = new URLSearchParams(window.location.search);
  var schoolParam = (queryString.get("school") || queryString.get("tenant") || "").trim().toLowerCase();
  var injectedSchool = window.__SCHOOL__ || null;

  function detectSubdomain() {
    if (schoolParam) return schoolParam;
    if (injectedSchool && injectedSchool.subdomain) return String(injectedSchool.subdomain).toLowerCase();
    var host = window.location.hostname.toLowerCase();
    if (host.endsWith(".workers.dev") || host.endsWith(".pages.dev") || host === "localhost") return "";
    var parts = host.split(".");
    if (parts.length >= 3 && parts[0] !== "www") return parts[0];
    return "";
  }

  var SCHOOL_SUBDOMAIN = detectSubdomain();

  function withSchoolQuery(basePath) {
    return basePath + (schoolParam ? (basePath.indexOf("?") >= 0 ? "&" : "?") + "school=" + encodeURIComponent(schoolParam) : "");
  }

  function apiUrl() {
    return withSchoolQuery("/api/gas");
  }

  var TOKEN_KEY = "sms_token";

  function getToken() {
    try { return localStorage.getItem(TOKEN_KEY) || ""; } catch (e) { return ""; }
  }
  function setToken(token) {
    try { localStorage.setItem(TOKEN_KEY, String(token)); } catch (e) {}
  }
  function clearToken() {
    try {
      localStorage.removeItem(TOKEN_KEY);
      for (var i = localStorage.length - 1; i >= 0; i--) {
        var k = localStorage.key(i);
        if (k && k.indexOf("sms_pane_") === 0) localStorage.removeItem(k);
      }
    } catch (e) {}
  }

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
            } catch (e) {}
            reject(new Error(data.error || "Muda wa session umeisha. Tafadhali ingia tena."));
            return;
          }

          if (r.status >= 400) {
            reject(new Error(data.error || "Ombi halijakamilika (" + r.status + ")."));
            return;
          }

          var result = data.result === undefined ? null : data.result;

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

  function restApiCall(path, options) {
    options = options || {};
    var headers = Object.assign({ "Content-Type": "application/json" }, options.headers || {});
    var tok = getToken();
    if (tok) headers["Authorization"] = "Bearer " + tok;

    return fetch(withSchoolQuery(path), Object.assign({}, options, { headers: headers }))
      .then(function (res) {
        return res.json().catch(function () { return {}; }).then(function (data) {
          if (!res.ok) throw new Error(data.error || "Ombi limeshindikana (" + res.status + ").");
          return data;
        });
      });
  }

  function createRunner(successFn, failureFn) {
    return new Proxy({}, {
      get: function (_, prop) {
        if (prop === "withSuccessHandler") {
          return function (fn) {
            return createRunner(typeof fn === "function" ? fn : null, failureFn);
          };
        }
        if (prop === "withFailureHandler") {
          return function (fn) {
            return createRunner(successFn, typeof fn === "function" ? fn : null);
          };
        }
        if (typeof prop !== "string") return undefined;

        return function () {
          var args = Array.prototype.slice.call(arguments);
          serverCall(prop, args).then(
            function (result) {
              if (successFn) {
                try { successFn(result); } catch (e) { console.error("Success handler error:", e); }
              }
            },
            function (err) {
              if (failureFn) {
                try { failureFn(err); } catch (e) { console.error("Failure handler error:", e); }
              } else {
                console.error("google.script.run." + prop + " imeshindwa:", err && err.message ? err.message : err);
              }
            }
          );
        };
      }
    });
  }

  window.google = window.google || {};
  window.google.script = window.google.script || {};
  window.google.script.run = createRunner(null, null);

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
    setWidth: function () {},
    setHeight: function () {},
    editor: { focus: function () {} },
  };

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
        .catch(function () {})
        .then(function () {
          clearToken();
          window.location.href = withSchoolQuery("/");
        });
    },
    fetchSchoolInfo: function () {
      return fetch(withSchoolQuery("/api/school-info")).then(function (res) { return res.json(); });
    },
    getPayments: function (studentId) {
      var q = studentId ? "/api/payments?student_id=" + encodeURIComponent(studentId) : "/api/payments";
      return restApiCall(q, { method: "GET" });
    },
    recordPayment: function (payload) {
      return restApiCall("/api/payments", { method: "POST", body: JSON.stringify(payload || {}) });
    },
    getFeeStructures: function (academicYear) {
      var q = academicYear ? "/api/fee-structures?academic_year=" + encodeURIComponent(academicYear) : "/api/fee-structures";
      return restApiCall(q, { method: "GET" });
    },
    saveFeeStructure: function (payload) {
      return restApiCall("/api/fee-structures", { method: "POST", body: JSON.stringify(payload || {}) });
    },
  };
})();
