# Battery Fuse Block (BFT) Cost Estimator

An Excel workbook and a web page with identical logic.

| File | What it is |
|---|---|
| `Battery_Fuse_Block_Cost_Estimator_v1.3.xlsx` | Baseline workbook (busbar geometry fixes, 2.0 mm cable lug placeholder) |
| `Battery_Fuse_Block_Cost_Estimator_v1.4.xlsx` | v1.4, built from v1.3 by `tools/build_v1_4.py` |
| `web/bft_estimator_v1.4.html` | Web tool v1.4, single file, open it in a browser |
| `Battery_Fuse_Block_Cost_Estimator_v1.5.xlsx` | v1.5, built by `tools/build_v1_5.py` (runs the v1.4 build, then applies v1.5) |
| `web/bft_estimator_v1.5.html` | Web tool v1.5 |
| `Battery_Fuse_Block_Cost_Estimator_v1.6.xlsx` | v1.6, built by `tools/build_v1_6.py` (runs the v1.5 build, then applies v1.6) |
| `web/bft_estimator_v1.6.html` | Web tool v1.6 |
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

## v1.5 changes

The Keepout Zones sheet is replaced by a **Layout** sheet (web tab "Layout"):
- **Sealed keepout zones.** Base walls stay at fuse height. A zone taller than the walls gets a cover bump-out, which adds the side walls of the bump to cover volume and raises the overall height. Placement is per zone. *End of row* puts the zone after the long-side fuses. *Back side* puts it behind the bus, which makes the block deeper, not longer. Bus notches are ignored for cost.
- **Short-side outputs.** On the short side you set how many of each fuse type go there, and which end they exit (opposite B+, or the B+ end). Those fuses turn 90 degrees, the bus becomes an L, and the cover gets an output extension on that end.
- **Pitch.** Each fuse type has a pitch override (blank = auto). You can also switch the walls between fuses off. Without walls, the auto pitch is body + creepage instead of body + divider, and sockets may overlap. A pitch below auto is flagged.
- **Battery fit.** Max length and depth (cover outside) are flagged, not enforced.

## v1.6 changes

- **OEM socket clearance diameter** is now a direct input per bolt size (Assumptions F31:F33). It is the socket plus the plastic protector plus the clearance to the protector. Socket OD and B17 are kept for reference only.
- **B+ corner, Left or Right.** This mirrors the drawing only, so size and cost are unchanged. When short-side outputs exit at the B+ end, the short-side column tucks under the B+ corner and B+ stays in the corner.
- **Short-side bus** has three options:
  - *One piece L*: copper is costed on its rectangular blank, so the scrap counts.
  - *Second bar laser welded*: one laser weld joint, plus a lap of one bus width.
  - *Second bar bolted*: one bolt of the B+ stud size plus a bolted joint, plus the same lap.
- **Copper scrap credit** (Assumptions B66, placeholder 0) is the share of the copper price recovered for blank scrap.

## Rebuild and check

```
python bft/tools/build_v1_4.py                                              # writes the v1.4 workbook
python bft/tests/excel_recalc.py bft/tests/cases.json > bft/tests/excel_results.json   # LibreOffice recalc
node bft/tests/compare_web_to_excel.js                                      # v1.4: 640/640 values must match
python bft/tools/build_v1_5.py
python bft/tests/excel_recalc_v1_5.py bft/tests/cases_v1_5.json > bft/tests/excel_results_v1_5.json
node bft/tests/compare_web_to_excel_v1_5.js                                 # v1.5: 1744/1744 values must match
python bft/tools/build_v1_6.py
python bft/tests/excel_recalc_v1_6.py bft/tests/cases_v1_6.json > bft/tests/excel_results_v1_6.json
node bft/tests/compare_web_to_excel_v1_6.js                                 # v1.6: 3164/3164 values must match
```
