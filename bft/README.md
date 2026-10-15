# Battery Fuse Block (BFT) Cost Estimator

An Excel workbook and a web page with identical logic.

| File | What it is |
|---|---|
| `Battery_Fuse_Block_Cost_Estimator_v1.3.xlsx` | Baseline workbook (busbar geometry fixes, 2.0 mm cable lug placeholder) |
| `Battery_Fuse_Block_Cost_Estimator_v1.4.xlsx` | v1.4, built from v1.3 by `tools/build_v1_4.py` |
| `web/bft_estimator_v1.4.html` | Web tool, single file, open it in a browser |
| `tests/` | Checks that the web tool and the workbook agree |

## v1.4 changes

**Keepout zones** (new `Keepout Zones` sheet, web tab "Keepout zones"). There are two rectangular zones for a bolted relay or another box. An included zone sits at the far end of the fuse row, opposite the B+ stud, and adds to the base:
- Its width plus one divider is added to the base width.
- The base length is at least the zone length plus two walls.
- The wall height is at least the zone height.

Each zone has 0, 1 or 2 studs. Use 0 for a laser welded joint, which adds one laser weld joint cost. When the zone draws current from the bus, a copper branch at the bus thickness runs from the end of the input bus to the zone. The branch width is the larger of the current need and the bolt edge distance. The branch length is computed automatically, or you can override it. Studs count toward retained bolts, bosses, retention and joining adders, and bolt hardware.

**Busbar Sizer** (new sheet, web tab "Busbar sizer"):
- For a known bar (width, thickness, length) it gives the mass, cost, rated current and estimated temperature rise.
- It sweeps width : thickness ratios at a fixed current. Each row is the smallest bar for that ratio, and the cheapest bar that fits is flagged.
- The ampacity model is `I = k x derate x A^0.5 x p^0.39 x (rise/30)^0.61`. Its constants are on Assumptions rows 57 to 63. k = 5.95 is an approximate fit to DIN 43671 bare copper ratings, so calibrate it to test data.

## Rebuild and check

```
python bft/tools/build_v1_4.py                                              # writes the v1.4 workbook
python bft/tests/excel_recalc.py bft/tests/cases.json > bft/tests/excel_results.json   # LibreOffice recalc
node bft/tests/compare_web_to_excel.js                                      # 640/640 values must match
```
