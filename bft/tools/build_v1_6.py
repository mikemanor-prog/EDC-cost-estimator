"""Builds Battery_Fuse_Block_Cost_Estimator_v1.6.xlsx on top of the v1.5 build.

v1.6 adds:
  * OEM socket clearance diameter (socket + plastic protector + clearance) as a direct input
    per bolt size (Assumptions F31:F33). It used to be socket OD + 2 x extra clearance.
  * B+ corner Left or Right (a mirror image, so size and cost do not change).
  * When short-side outputs exit at the B+ end, the short-side column tucks under the B+ corner
    and B+ stays in the corner.
  * Short-side bus: one piece L (costed on its rectangular blank, scrap included), or a second
    bar laser welded or bolted to the main bus (lap = bus width, plus one joint cost).
  * Copper scrap credit (Assumptions B66), placeholder 0.

The web tool (bft/web/bft_estimator_v1.6.html) must match these formulas.
Run: python bft/tools/build_v1_6.py
"""
import copy
import os
import sys
import tempfile

from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_v1_5 as v5  # noqa: E402
from build_v1_4 import MONEY, N1, N2, calc, f, head, inp, unit  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "Battery_Fuse_Block_Cost_Estimator_v1.6.xlsx")
L, E, A = "Layout!", "Estimator!", "Assumptions!"
WALL, DIV = f"{A}$B$5", f"{A}$B$6"
RHO_T = f"{E}$B$7*{A}$B$13/1000"


def style(dst, src):
    dst.font = copy.copy(src.font)
    dst.fill = copy.copy(src.fill)
    dst.border = copy.copy(src.border)
    dst.number_format = src.number_format
    dst.alignment = copy.copy(src.alignment)


def build(out=OUT):
    tmp = os.path.join(tempfile.mkdtemp(), "v15.xlsx")
    v5.build(tmp)
    wb = load_workbook(tmp)
    est, asm, ly = wb["Estimator"], wb["Assumptions"], wb["Layout"]

    # ---------------------------------------------------------------- Assumptions
    asm["F30"] = "OEM socket clearance dia (mm)"
    for r, v in ((31, 23), (32, 19), (33, 17)):
        inp(asm.cell(r, 6), v, N1)
    asm["A34"] = ("Socket clearance dia (F) is the OEM value: socket + plastic protector + clearance to the "
                  "protector. Socket OD (E) and B17 are for reference only. Hardware cost is a placeholder.")
    asm["D17"] = "Not used from v1.6. Enter the OEM socket clearance diameter in F31:F33"
    head(asm, 65, "Copper blank and scrap")
    asm["A66"] = "Copper scrap credit"
    asm["A66"].font = f()
    inp(asm["B66"], 0, "0%")
    asm["C66"] = "fraction"
    asm["C66"].font = f()
    asm["D66"] = ("Placeholder. Share of the copper price recovered for blank scrap. "
                  "0 = scrap costs the full copper price")
    asm["D66"].font = f("595959", size=9)

    # ---------------------------------------------------------------- Layout inputs
    ly.column_dimensions["E"].width = 22
    for r, label, v, opts, note in (
            (5, "B+ corner", "Left", '"Left,Right"', "Mirror image only"),
            (6, "Short-side bus", "One piece L", '"One piece L,Second bar laser welded,Second bar bolted"',
             "One piece L is costed on its rectangular blank")):
        ly.cell(r, 5, label).font = f()
        inp(ly.cell(r, 6), v)
        unit(ly.cell(r, 7), note)
        d = DataValidation(type="list", formula1=opts, showErrorMessage=True)
        d.add(f"F{r}")
        ly.add_data_validation(d)

    head(ly, 89, "B+ CORNER, SHORT-SIDE BAR AND COPPER BLANK", 4)
    bend = f'AND(B51>0,B6="B+ end")'
    rows = [
        (90, "Long row start (from the B+ end inner wall)", f"=IF({bend},MAX({E}B35,B64),{E}B35)", N1,
         "mm. The short-side column tucks under the B+ corner when it exits at the B+ end"),
        (91, "Bus along the row, net mass", f"=B66*{E}B28*{RHO_T}", N1, "g"),
        (92, "Short-side bar length incl. lap", f'=IF(B51>0,B67+IF(F6="One piece L",0,{E}B28),0)', N1,
         "mm. A second bar laps the main bus by one bus width"),
        (93, "Short-side bar net mass", f"=B92*{E}B28*{RHO_T}", N1, "g"),
        (94, "Bus net mass", "=B91+B93", N1, "g, the part weight"),
        (95, "Bus blank mass", f'=IF(AND(B51>0,F6="One piece L"),B66*({E}B28+B67)*{RHO_T},B94)', N1,
         "g. One piece L = rectangle around the L"),
        (96, "Bus blank scrap", "=B95-B94", N1, "g"),
        (97, "Short-side bus joint cost",
         f'=IF(B51>0,IF(F6="Second bar laser welded",{A}$B$45,IF(F6="Second bar bolted",'
         f"INDEX({A}$I$31:$I$33,MATCH({E}$B$10,{A}$A$31:$A$33,0))+{A}$B$43,0)),0)", MONEY,
         "$. Bolted = one bolt of the B+ stud size plus a bolted joint"),
    ]
    for r, label, fm, fmt, note in rows:
        ly.cell(r, 1, label).font = f()
        calc(ly.cell(r, 2), fm, fmt)
        unit(ly.cell(r, 3), note)

    ly["B65"] = (f"=IF(B52>0,B90+B54-B56/2+B57/2,IF({bend},MAX({E}B35,B64-B63-{E}B27/2+{E}B28),{E}B35))")
    ly["B66"] = f'=IF(B51=0,B65,IF(B6="Opposite B+",{E}B35+B54+B34+C34+B63+{E}B27/2,B65))'

    # ---------------------------------------------------------------- Estimator
    est["A3"] = ("Revision v1.6: OEM socket clearance input, B+ corner left/right, short-side bus options "
                 "(one piece L on its blank, or a second bar welded or bolted), copper scrap. Builds on v1.5.")
    est["B36"] = (f'=MAX(B34+{L}B90+{L}B76+IF({L}B6="Opposite B+",{L}B64,0)+2*{WALL},'
                  f"{L}B77-{DIV}+2*{WALL})")
    est["A44"] = "Bus length (along the row plus the short-side bar incl. lap)"
    est["B44"] = f"={L}B66+{L}B92"
    est["A49"] = "Copper cost (part plus blank scrap)"
    est["B49"] = f"=(B48*(1+Assumptions!$B$15)+{L}B96*(1-Assumptions!$B$66))/1000*B9"
    for c in range(1, 3):
        style(est.cell(76, c), est.cell(75, c))
    est["A76"] = "Short-side bus joint"
    est["B76"] = f"={L}B97"
    est["A85"] = "Zone hardware, welds and bus joint"
    est["B85"] = "=B74+B75+B76"

    wb.save(out)
    return out


if __name__ == "__main__":
    print(build(*sys.argv[1:]))
