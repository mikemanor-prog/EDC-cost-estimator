// Runs the web tool's model on each case and compares it with the LibreOffice recalc of the workbook.
// Usage: node tests/edc_compare_web_to_excel_v1_1.js [excel_results.json]
const fs = require("fs");
const html = fs.readFileSync(__dirname + "/../web/edc_estimator_v1.1.html", "utf8");
const js = html.match(/<script>([\s\S]*?)<\/script>/)[1];
const consts = js.slice(0, js.indexOf("const GEO_FIELDS"));
const model = js.slice(js.indexOf("// ---------- model ----------"), js.indexOf("// ---------- formatting ----------"));
const api = new Function(consts + model + "\nreturn { EXAMPLE, compute };")();
const cases = JSON.parse(fs.readFileSync(__dirname + "/edc_cases_v1_1.json"));
const xl = JSON.parse(fs.readFileSync(process.argv[2] || __dirname + "/edc_excel_results_v1_1.json"));

const PARTS = ["micro2", "micro3", "mcase", "a1s", "relMicro", "relMini", "relHigh"];
const parts = ["housing", "upper", "lower"];
const set = {
  "Estimator!B16": (S, v) => (S.layout = v.toLowerCase()), "Estimator!B17": (S, v) => (S.cavities = v),
  "Estimator!B18": (S, v) => (S.partResin.housing = v), "Estimator!B19": (S, v) => (S.partResin.upper = v),
  "Estimator!B20": (S, v) => (S.partResin.lower = v), "Estimator!B21": (S, v) => (S.labor = v), "Estimator!B22": (S, v) => (S.ops = v),
  "Assumptions!B34": (S, v) => (S.mold.runVol = v * 100), "Assumptions!B35": (S, v) => (S.mold.runArea = v * 100),
  "Assumptions!B38": (S, v) => (S.mold.clamp = v), "Assumptions!B43": (S, v) => (S.mold.b2t = v), "Assumptions!C45": (S, v) => (S.mold.b4r = v),
};
[7, 8, 9, 10].forEach((r, i) => (set["Estimator!B" + r] = (S, v) => (S.counts[PARTS[i]] = v)));
[12, 13, 14].forEach((r, i) => (set["Estimator!B" + r] = (S, v) => (S.counts[PARTS[4 + i]] = v)));
"BCD".split("").forEach((c, j) => {
  set["Estimator!" + c + "26"] = (S, v) => (S.press[parts[j]] = v);
  set["Estimator!" + c + "27"] = (S, v) => (S.mrate[parts[j]] = v);
});
const partRows = { Mass: "mass", "Material incl. runner": "material", "Shot volume incl. runner": "shotVol", "Shot weight incl. runner": "shotG",
  "Projected area incl. runner": "proj", "Rule-of-thumb clamp": "tonsRT", "Rule-of-thumb clamp (metric t)": "tonsRTm",
  "Press used": "tons", "Cycle time": "cycle", "Machine rate from band": "band", "Machine rate used": "rate", "Labor": "laborHr",
  "Molding": "molding", "Part cost": "cost" };
const top = { "Housing L": (r) => r.housing.L, "Housing W": (r) => r.housing.W, "Housing H": (r) => r.housing.H, "Housing volume": (r) => r.housing.vol,
  "Cover L": (r) => r.upper.L, "Cover W": (r) => r.upper.W, "Upper cover H": (r) => r.upper.H, "Lower cover H": (r) => r.lower.H,
  "Upper cover volume": (r) => r.upper.vol, "Lower cover volume": (r) => r.lower.vol,
  "Plastic cost per EDC set": (r) => r.total, "Total plastic": (r) => r.mass, "Stack height": (r) => r.stack };
const bumps = [null, "micro2", "micro3", "relMicro"];
let n = 0, bad = 0;
cases.forEach((c, i) => {
  const X = xl[i];
  const R = bumps.map((b) => {
    const S = JSON.parse(JSON.stringify(api.EXAMPLE));
    for (const [k, v] of Object.entries(c)) { if (!set[k]) throw new Error("no setter for " + k); set[k](S, v); }
    if (b) S.counts[b] += 1;
    return api.compute(S);
  });
  let part = null;
  for (const row of Object.values(X)) {
    const label = row.label.replace(/ \((mm|mm³|g|\$|\$\/hr|s|cm³|in²|US t)\)$/, "");
    let get = null;
    if (/^Mass \(g\)$/.test(row.label)) part = parts[parts.indexOf(part) + 1];   // a Mass row starts each part block
    if (top[label]) get = top[label];
    else if (part && partRows[label]) get = (r) => r[part][partRows[label]];
    if (!get) continue;
    row.vals.forEach((xv, j) => {
      n++;
      const jv = get(R[j]);
      if (!(Math.abs(xv - jv) <= 1e-9 * Math.max(1, Math.abs(xv)))) { bad++; console.log("case", i, "col", j, row.label, "xl", xv, "js", jv); }
    });
  }
  console.log("case", i, "total", R[0].total.toFixed(4), "tons", parts.map((p) => R[0][p].tons.toFixed(0)).join("/"));
});
console.log(`${n - bad}/${n} values match`);
process.exit(bad ? 1 : 0);
