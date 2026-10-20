// Runs the web tool's model on each case and compares it with the LibreOffice recalc of the workbook.
// Usage: node tests/edc_compare_web_to_excel_v1_2.js [excel_results.json]
const fs = require("fs");
const html = fs.readFileSync(__dirname + "/../web/edc_estimator_v1.2.html", "utf8");
const js = html.match(/<script>([\s\S]*?)<\/script>/)[1];
const consts = js.slice(0, js.indexOf("const GEO_FIELDS"));
const model = js.slice(js.indexOf("// ---------- model ----------"), js.indexOf("// ---------- formatting ----------"));
const api = new Function(consts + model + "\nreturn { EXAMPLE, compute };")();
const cases = JSON.parse(fs.readFileSync(__dirname + "/edc_cases_v1_2.json"));
const xl = JSON.parse(fs.readFileSync(process.argv[2] || __dirname + "/edc_excel_results_v1_2.json"));

const PARTS = ["micro2", "micro3", "mcase", "a1s", "relMicro", "relMini", "relHigh"];
const parts = ["housing", "upper", "lower"];
const set = {
  "Estimator!B16": (S, v) => (S.layout = v.toLowerCase()), "Estimator!B17": (S, v) => (S.cavities = v),
  "Estimator!B18": (S, v) => (S.partResin.housing = v), "Estimator!B19": (S, v) => (S.partResin.upper = v),
  "Estimator!B20": (S, v) => (S.partResin.lower = v), "Estimator!B21": (S, v) => (S.labor = v), "Estimator!B22": (S, v) => (S.ops = v),
  "Assumptions!B38": (S, v) => (S.gateDefault = v),
  "Assumptions!B43": (S, v) => (S.mold.clamp = v), "Assumptions!B48": (S, v) => (S.mold.b2t = v), "Assumptions!C50": (S, v) => (S.mold.b4r = v),
};
["Cold runner", "Hot-to-cold", "Hot runner"].forEach((g, i) => {
  set["Assumptions!B" + (35 + i)] = (S, v) => (S.gates[g][0] = v * 100);
  set["Assumptions!C" + (35 + i)] = (S, v) => (S.gates[g][1] = v * 100);
});
[7, 8, 9, 10].forEach((r, i) => (set["Estimator!B" + r] = (S, v) => (S.counts[PARTS[i]] = v)));
[12, 13, 14].forEach((r, i) => (set["Estimator!B" + r] = (S, v) => (S.counts[PARTS[4 + i]] = v)));
"BCD".split("").forEach((c, j) => {
  set["Estimator!" + c + "26"] = (S, v) => (S.press[parts[j]] = v);
  set["Estimator!" + c + "27"] = (S, v) => (S.mrate[parts[j]] = v);
  set["Estimator!" + c + "28"] = (S, v) => (S.partGate[parts[j]] = v);
});
const partRows = { Mass: "mass", "Gate / runner type used": "gate", "Runner volume used": (p) => p.runVol / 100,
  "Runner area used": (p) => p.runArea / 100, "Material incl. runner": "material", "Shot volume incl. runner": "shotVol", "Shot weight incl. runner": "shotG",
  "Projected area incl. runner": "proj", "Rule-of-thumb clamp": "tonsRT", "Rule-of-thumb clamp (metric t)": "tonsRTm",
  "Press size": "tons", "Cycle time": "cycle", "Machine rate from band": "band", "Machine rate used": "rate", "Labor": "laborHr",
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
    const label = row.label.replace(/ \((mm|mm³|g|\$|\$\/hr|s|cm³|in²|US t|%)\)$/, "");
    let get = null;
    if (/^Mass \(g\)$/.test(row.label)) part = parts[parts.indexOf(part) + 1];   // a Mass row starts each part block
    if (top[label]) get = top[label];
    else if (part && partRows[label]) { const k = partRows[label]; get = (r) => (typeof k === "function" ? k(r[part]) : r[part][k]); }
    if (!get) continue;
    row.vals.forEach((xv, j) => {
      n++;
      const jv = get(R[j]);
      const ok = typeof xv === "string" ? xv === jv : Math.abs(xv - jv) <= 1e-9 * Math.max(1, Math.abs(xv));
      if (!ok) { bad++; console.log("case", i, "col", j, row.label, "xl", xv, "js", jv); }
    });
  }
  console.log("case", i, "total", R[0].total.toFixed(4), "tons", parts.map((p) => R[0][p].tons.toFixed(0)).join("/"));
});
console.log(`${n - bad}/${n} values match`);
process.exit(bad ? 1 : 0);
