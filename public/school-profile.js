// The school's identity (name, address, motto, logo, report heading lines) comes from
// Settings and nowhere else. Signed-in pages load it with getSchoolProfile and can
//   • bind elements:   <span data-school="name"></span>   <img data-school-logo>
//   • read it in code: SchoolProfile.now().name   or   SchoolProfile.get(function (p) { ... })
(function () {
  "use strict";
  var EMPTY = { name: "", nameSet: false, code: "", council: "", address: "", phone: "", email: "",
                motto: "", logo: "", header1: "", header2: "" };
  var profile = null, waiting = [], loading = false;

  function bind(root, p) {
    root = root || document;
    var list = root.querySelectorAll("[data-school]");
    for (var i = 0; i < list.length; i++) {
      var el = list[i], key = el.getAttribute("data-school"), upper = /-upper$/.test(key);
      var value = p[upper ? key.replace(/-upper$/, "") : key];
      value = value == null ? "" : String(value);
      if (upper) value = value.toUpperCase();
      el.textContent = value || el.getAttribute("data-school-fallback") || "";
      if (!value && el.hasAttribute("data-school-hide-empty")) el.style.display = "none";
      else if (el.hasAttribute("data-school-hide-empty")) el.style.display = "";
    }
    var logos = root.querySelectorAll("img[data-school-logo]");
    for (var j = 0; j < logos.length; j++) {
      if (p.logo) { logos[j].src = p.logo; logos[j].style.display = ""; }
      else { logos[j].removeAttribute("src"); logos[j].style.display = "none"; }
    }
    var suffix = document.documentElement.getAttribute("data-title-suffix");
    if (suffix !== null) document.title = [p.name, suffix].filter(Boolean).join(" \u2014 ");
  }

  function finish(p) {
    profile = Object.assign({}, EMPTY, p || {});
    bind(document, profile);
    var cbs = waiting; waiting = [];
    cbs.forEach(function (cb) { try { cb(profile); } catch (e) { console.error(e); } });
    document.dispatchEvent(new CustomEvent("schoolprofile", { detail: profile }));
  }

  function load() {
    if (loading) return;
    loading = true;
    google.script.run
      .withSuccessHandler(function (p) { loading = false; finish(p); })
      .withFailureHandler(function () { loading = false; finish(EMPTY); })
      .getSchoolProfile();
  }

  window.SchoolProfile = {
    get: function (cb) { if (profile) cb(profile); else { waiting.push(cb); load(); } },
    now: function () { return profile || EMPTY; },
    bind: function (root) { bind(root, profile || EMPTY); },
    reload: function (cb) { profile = null; if (cb) waiting.push(cb); loading = false; load(); }
  };

  // Only fetch on page load when somebody is signed in; get() fetches on demand otherwise.
  function autoload() {
    var token = null;
    try { token = localStorage.getItem("sms_token"); } catch (e) { /* ignore */ }
    if (token) load();
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", autoload, { once: true });
  else autoload();
})();
