// Same idea as school-profile.js, for the public landing and login pages (no sign-in yet).
// Uses getPublicBranding: name, motto, logo, contact details and a few headline numbers.
//   <span data-school="name">  <img data-school-logo>  <b data-school-count="students">
(function () {
  "use strict";
  var EMPTY = { name: "", nameSet: false, motto: "", logo: "", address: "", phone: "", email: "", counts: {}, year: "" };
  var branding = null, waiting = [], loading = false;

  function bind(p) {
    var list = document.querySelectorAll("[data-school]");
    for (var i = 0; i < list.length; i++) {
      var el = list[i], key = el.getAttribute("data-school"), upper = /-upper$/.test(key);
      var value = p[upper ? key.replace(/-upper$/, "") : key];
      value = value == null ? "" : String(value);
      if (upper) value = value.toUpperCase();
      el.textContent = value || el.getAttribute("data-school-fallback") || "";
      if (el.hasAttribute("data-school-hide-empty")) el.style.display = value ? "" : "none";
    }
    var logos = document.querySelectorAll("img[data-school-logo]");
    for (var j = 0; j < logos.length; j++) {
      var img = logos[j], holder = img.parentElement;
      if (p.logo) { img.src = p.logo; img.style.display = ""; if (holder) holder.removeAttribute("data-no-logo"); }
      else { img.removeAttribute("src"); img.style.display = "none"; if (holder) holder.setAttribute("data-no-logo", "1"); }
    }
    var counts = document.querySelectorAll("[data-school-count]");
    for (var k = 0; k < counts.length; k++) {
      var n = (p.counts || {})[counts[k].getAttribute("data-school-count")];
      counts[k].textContent = n == null ? "\u2014" : String(n);
    }
    var yearEls = document.querySelectorAll("[data-school-year]");
    for (var y = 0; y < yearEls.length; y++) yearEls[y].textContent = p.year || String(new Date().getFullYear());
    var cur = document.querySelectorAll("[data-cur-year]");
    for (var c = 0; c < cur.length; c++) cur[c].textContent = String(new Date().getFullYear());
    var suffix = document.documentElement.getAttribute("data-title-suffix");
    if (suffix !== null) document.title = [p.name, suffix].filter(Boolean).join(" \u2014 ");
  }

  function finish(p) {
    branding = Object.assign({}, EMPTY, p || {});
    bind(branding);
    var cbs = waiting; waiting = [];
    cbs.forEach(function (cb) { try { cb(branding); } catch (e) { console.error(e); } });
  }

  function load() {
    if (loading) return;
    loading = true;
    google.script.run
      .withSuccessHandler(function (p) { loading = false; finish(p); })
      .withFailureHandler(function () { loading = false; finish(EMPTY); })
      .getPublicBranding();
  }

  window.SchoolPublic = { get: function (cb) { if (branding) cb(branding); else { waiting.push(cb); load(); } } };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", load, { once: true });
  else load();
})();
