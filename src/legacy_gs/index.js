
// src/legacy_gs/index.js — Bundle ya mafaili yote ya GAS (Multi-Level: Pre, Primary, O-Level & A-Level)
import PrimaryCode from "./core/PrimaryCode.gs";
import SecondaryCode from "./core/SecondaryCode.gs";
import DashboardAccess from "./core/DashboardAccess.gs";
import RBAC from "./core/RBAC.gs";
import PrimarySetup from "./core/PrimarySetup.gs";
import PrimaryApi from "./core/PrimaryApi.gs";
import Academictools from "./academics/Academictools.gs";
import LessonPlan from "./academics/LessonPlan.gs";
import LessonPlanSyllabus from "./academics/LessonPlanSyllabus.gs";
import Features from "./academics/Features.gs";
import SecondMaster from "./operations/SecondMaster.gs";
import Timetableapi from "./operations/Timetableapi.gs";
import Todreportapi from "./operations/Todreportapi.gs";
import Parentportal from "./portal/Parentportal.gs";
import AccountantBridge from "./portal/AccountantBridge.gs";
import ParentAccess from "./portal/ParentAccess.gs";
import Overrides from "./operations/overrides.gs";

// 1. Kusanya majina yote ya functions kutoka PrimaryCode.gs na SecondaryCode.gs
function extractFunctionNames(gsSource) {
  const names = new Set();
  const regex = /function\s+([a-zA-Z0-9_$]+)\s*\(/g;
  let match;
  while ((match = regex.exec(gsSource)) !== null) {
    names.add(match[1]);
  }
  return Array.from(names);
}

const primaryFns = extractFunctionNames(PrimaryCode);
const secondaryFns = extractFunctionNames(SecondaryCode);
const allCoreFns = Array.from(new Set([...primaryFns, ...secondaryFns]));

// 2. Funga kila Code.gs ndani ya boksi lake (Closure) ili zisigongane
const sandboxedCoreScript = `
var __PrimaryEngine = (function() {
  ${PrimaryCode}
  return {
    ${primaryFns.map((fn) => `${fn}: typeof ${fn} !== "undefined" ? ${fn} : undefined`).join(",\n    ")}
  };
})();

var __SecondaryEngine = (function() {
  ${SecondaryCode}
  return {
    ${secondaryFns.map((fn) => `${fn}: typeof ${fn} !== "undefined" ? ${fn} : undefined`).join(",\n    ")}
  };
})();

function __resolveCoreEngine__() {
  var lvl = "";
  if (typeof school !== "undefined" && school && school.education_level) {
    lvl = String(school.education_level);
  } else if (typeof __SCHOOL__ !== "undefined" && __SCHOOL__ && __SCHOOL__.education_level) {
    lvl = String(__SCHOOL__.education_level);
  } else if (typeof globalThis !== "undefined" && globalThis.school && globalThis.school.education_level) {
    lvl = String(globalThis.school.education_level);
  }
  lvl = lvl.trim().toUpperCase();
  if (lvl.indexOf("PRIMARY") !== -1 || lvl.indexOf("PRE") !== -1 || lvl === "NURSERY") {
    return __PrimaryEngine;
  }
  return __SecondaryEngine;
}

${allCoreFns
  .map(
    (fn) => `function ${fn}() {
  var primaryEng = __resolveCoreEngine__();
  var fallbackEng = primaryEng === __PrimaryEngine ? __SecondaryEngine : PrimaryEngine;
  var targetFn = (primaryEng && typeof primaryEng.${fn} === "function")
    ? primaryEng.${fn}
    : (fallbackEng && typeof fallbackEng.${fn} === "function" ? fallbackEng.${fn} : null);
  if (!targetFn) {
    throw new Error("Function '${fn}' haipatikani kwenye engine ya shule hii.");
  }
  return targetFn.apply(primaryEng, arguments);
}`
  )
  .join("\n")}
`;

// 3. Unganisha na mafaili mengine yote ya GAS kuwa COMBINED_SCRIPT moja
export const COMBINED_SCRIPT = [
  sandboxedCoreScript,
  DashboardAccess,
  RBAC,
  PrimarySetup,
  PrimaryApi,
  Academictools,
  LessonPlan,
  LessonPlanSyllabus,
  Features,
  SecondMaster,
  Timetableapi,
  Todreportapi,
  Parentportal,
  AccountantBridge,
  ParentAccess,
  Overrides, // Daima iwe ya mwisho: inaandika juu ya functions zilizotangulia
].join("\n;\n");
