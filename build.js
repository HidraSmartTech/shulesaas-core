import fs from "node:fs";

const files = [
  "core/DashboardAccess.gs", "core/RBAC.gs", "core/PrimarySetup.gs",
  "core/PrimaryApi.gs", "academics/Academictools.gs", "academics/LessonPlan.gs",
  "academics/LessonPlanSyllabus.gs", "academics/Features.gs", "operations/SecondMaster.gs",
  "operations/Timetableapi.gs", "operations/Todreportapi.gs", "portal/Parentportal.gs",
  "portal/AccountantBridge.gs", "portal/ParentAccess.gs", "operations/overrides.gs"
];

const read = (f) => fs.readFileSync("./src/legacy_gs/" + f, "utf8");
const getFns = (code) => [...new Set([...code.matchAll(/function\s+([a-zA-Z0-9_$]+)\s*\(/g)].map(m => m[1]))];

const priCode = read("core/PrimaryCode.gs");
const secCode = read("core/SecondaryCode.gs");
const priFns = getFns(priCode);
const secFns = getFns(secCode);
const coreFns = [...new Set([...priFns, ...secFns])];
const extraCode = files.map(read).join("\n;\n");
const allFns = [...new Set([...coreFns, ...getFns(extraCode)])];

const out = `export function runGas(sandbox, fnName, args) {
  const {
    Date, console, __HOST__, __SCHOOL__, Logger, Session,
    SpreadsheetApp, HtmlService, PropertiesService, CacheService,
    LockService, ScriptApp, Utilities
  } = sandbox;

  globalThis.__APP_ENV__ = "cloudflare";
  var calc, res, data, rows, sheet, ss, i, j, k, row, item, obj, list, map, out, tmp, val, key, id, name, code, err, msg, ok, status, total, count, sum, avg, min, max;

  const __PrimaryEngine = (function() {
    ${priCode}
    return { ${priFns.map(f => `${f}: typeof ${f} !== "undefined" ? ${f} : undefined`).join(", ")} };
  })();

  const __SecondaryEngine = (function() {
    ${secCode}
    return { ${secFns.map(f => `${f}: typeof ${f} !== "undefined" ? ${f} : undefined`).join(", ")} };
  })();

  function __resolveCoreEngine__() {
    const lvl = String(
      (__SCHOOL__ && __SCHOOL__.education_level) ||
      (__HOST__ && __HOST__.education_level) ||
      ""
    ).toUpperCase();
    const isPrimary = lvl.includes("PRIMARY") || lvl.includes("PRE") || lvl === "NURSERY";
    return isPrimary ? __PrimaryEngine : __SecondaryEngine;
  }

  ${coreFns.map(f => `function ${f}() {
    var p = __resolveCoreEngine__();
    var fb = (p === __PrimaryEngine) ? __SecondaryEngine : __PrimaryEngine;
    var fn = (p && p.${f}) || (fb && fb.${f});
    if (!fn) throw new Error("Missing function: ${f}");
    return fn.apply(p, arguments);
  }`).join("\n  ")}

  ${extraCode}

  const __fnMap = { ${allFns.join(", ")} };
  if (typeof __fnMap[fnName] !== "function") {
    throw new Error("Function not found: " + fnName);
  }
  const __fnRes = __fnMap[fnName].apply(null, args || []);
  return __fnRes === undefined ? null : __fnRes;
}
`;

fs.writeFileSync("./src/legacy_gs/compiled.js", out, "utf8");
console.log("TAYARI!");
