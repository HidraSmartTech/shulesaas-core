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
  const { Date, console, __HOST__, Logger, Session, SpreadsheetApp, HtmlService, PropertiesService, CacheService, LockService, ScriptApp, Utilities } = sandbox;
  const { Date, console, __HOST__, Logger, Session, SpreadsheetApp, HtmlService, PropertiesService, CacheService, LockService, ScriptApp, Utilities } = sandbox;
  const __APP_ENV__ = "cloudflare";
  const __PrimaryEngine = (function() { ${priCode}\n return { ${priFns.map(f => `${f}: typeof ${f}!=="undefined"?${f}:undefined`).join(",")} }; })();
  const __SecondaryEngine = (function() { ${secCode}\n return { ${secFns.map(f => `${f}: typeof ${f}!=="undefined"?${f}:undefined`).join(",")} }; })();
  function __resolveCoreEngine__() {
    var l = String((typeof school!=="undefined"&&school?.education_level)||(typeof __SCHOOL__!=="undefined"&&__SCHOOL__?.education_level)||"").toUpperCase();
    return (l.includes("PRIMARY")||l.includes("PRE")||l==="NURSERY") ? __PrimaryEngine : __SecondaryEngine;
  }
  ${coreFns.map(f => `function ${f}(){ var p=__resolveCoreEngine__(), fb=p===__PrimaryEngine?__SecondaryEngine:__PrimaryEngine, fn=p?.${f}||fb?.${f}; if(!fn) throw new Error("Missing ${f}"); return fn.apply(p, arguments); }`).join("\n")}
  ${extraCode}
  const map = { ${allFns.join(",")} };
  if (typeof map[fnName] !== "function") throw new Error("Function not found: " + fnName);
  const res = map[fnName].apply(null, args || []);
  return res === undefined ? null : res;
}`;

fs.writeFileSync("./src/legacy_gs/compiled.js", out, "utf8");
console.log("TAYARI!");
