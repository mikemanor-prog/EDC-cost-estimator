"""Builds Battery_Fuse_Block_Cost_Estimator_Rev1.11.xlsx from Mike's Rev1.10 workbook.

Rev1.11 adds:
  * BOM Paste tab: one row per part, laid out like the BOM calculator (Part Number,
    Part Description, Total Unit Cost, Estimate Source / Rev, Notes) plus two columns,
    Material and Machine + Labor. The stamped busbar is the main row; the fuse block
    plastic, copper, purchased hardware and process adders follow.
  * Busbar Stamping: rule-of-thumb press size (cut length x thickness x shear strength,
    plus forming / stripping allowance and a capacity margin), rounded up to a standard
    press list. Press size and press rate are overrides (blank = rule of thumb / press table).
  * Resins tab copied (not linked) from the BV1XX EDC Cost Estimator Rev1.30, so the same
    material export pastes into both files. Estimator B8 picks the resin; price and density
    come from the Resins tab.
Cached values are written in, as in earlier builds.
Run: python bft/tools/build_rev1_11.py
"""
import copy
import os
import sys

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from build_v1_7 import add_cached_values  # noqa: E402

BFT = os.path.dirname(HERE)
SRC = os.path.join(BFT, "Battery_Fuse_Block_Cost_Estimator_Rev1.10.xlsx")
EDC = os.path.join(BFT, "source", "BV1XX_EDC_Cost_Estimator_Rev1.30.xlsm")
OUT = os.path.join(BFT, "Battery_Fuse_Block_Cost_Estimator_Rev1.11.xlsx")
REV = "Rev1.11"

HEAD_FILL = PatternFill("solid", fgColor="1F3A4D")
SUB_FILL = PatternFill("solid", fgColor="DCE6EC")
INPUT_FILL = PatternFill("solid", fgColor="FFF9C4")
TOT_FILL = PatternFill("solid", fgColor="E8F3E8")
THIN = Side(style="thin", color="B7BFCB")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
F_HEAD = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_LABEL = Font(name="Calibri", size=10)
F_BOLD = Font(name="Calibri", size=10, bold=True)
F_INPUT = Font(name="Calibri", size=10, color="0000FF")
F_LINK = Font(name="Calibri", size=10, color="008000")
F_NOTE = Font(name="Calibri", size=9, color="595959")
MONEY4 = '\\$0.0000;\\-\\$0.0000;"-"'

# Standard press list (metric t) and rate with operator ($/hr). Placeholders.
# Press list: size (metric t), $/hr with operator, bed length left to right (mm). Placeholders.
PRESSES = [(25, 55, 600), (45, 65, 750), (60, 75, 900), (80, 85, 1000), (110, 95, 1200), (150, 110, 1500),
           (200, 125, 1800), (300, 150, 2400)]
RESIN_ROWS = (4, 103)
RES = lambda col: f"Resins!${col}${RESIN_ROWS[0]}:${col}${RESIN_ROWS[1]}"


def put(ws, ref, value, font=F_LABEL, fill=None, fmt=None, border=None, wrap=False):
    c = ws[ref]
    c.value = value
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if border:
        c.border = border
    if wrap:
        c.alignment = Alignment(wrap_text=True, vertical="top")
    return c


def inp(ws, ref, value, fmt=None):
    return put(ws, ref, value, F_INPUT, INPUT_FILL, fmt, BOX)


# ---------------------------------------------------------------- Resins (copied)
def copy_resins(wb):
    src_wb = load_workbook(EDC, keep_vba=False)
    src = src_wb["Resins"]
    dst = wb.create_sheet("Resins", index=wb.sheetnames.index("Fuse Library") + 1)
    for row in src.iter_rows():
        for c in row:
            d = dst.cell(row=c.row, column=c.column, value=c.value)
            if c.has_style:
                d.font = copy.copy(c.font)
                d.fill = copy.copy(c.fill)
                d.border = copy.copy(c.border)
                d.alignment = copy.copy(c.alignment)
                d.number_format = c.number_format
                d.protection = copy.copy(c.protection)
    for k, dim in src.column_dimensions.items():
        dst.column_dimensions[k].width = dim.width
        dst.column_dimensions[k].hidden = dim.hidden
    for k, dim in src.row_dimensions.items():
        if dim.height:
            dst.row_dimensions[k].height = dim.height
    for m in src.merged_cells.ranges:
        dst.merge_cells(str(m))
    for dv in src.data_validations.dataValidation:
        dst.add_data_validation(copy.copy(dv))
    for cf in src.conditional_formatting:
        for rule in cf.rules:
            dst.conditional_formatting.add(str(cf.sqref), rule)
    dst.freeze_panes = src.freeze_panes
    dst.sheet_properties.tabColor = "4472C4"  # blue = set up once
    dst.sheet_view.showGridLines = src.sheet_view.showGridLines
    dst.sheet_view.zoomScale = src.sheet_view.zoomScale
    dst["A1"].value = "Resins (copied from the BV1XX EDC Cost Estimator Rev1.30, not linked)"
    wb.defined_names["ResinList"] = DefinedName(
        "ResinList", attr_text=src_wb.defined_names["ResinList"].attr_text)
    return dst


# ---------------------------------------------------------------- Estimator resin pick
def resin_pick(wb):
    est, asm = wb["Estimator"], wb["Assumptions"]
    est["A8"].value = "Plastic resin (base and cover)"
    est["B8"].value = "PA66-GF30"
    est["B8"].number_format = "@"
    est["C8"].value = (f'=IFERROR(TEXT(INDEX({RES("C")},MATCH($B$8,{RES("A")},0)),"$0.00")'
                       f'&" $/kg (Resins tab)","Pick a resin from the Resins tab")')
    est["C8"].font = F_LINK
    dv = DataValidation(type="list", formula1="=ResinList", allow_blank=False, showErrorMessage=True,
                        error="Pick a resin from the Resins tab (add new resins there).")
    est.add_data_validation(dv)
    dv.add("B8")
    price = f"INDEX({RES('C')},MATCH($B$8,{RES('A')},0))"
    est["B59"].value = f"=B58/1000*{price}"
    est["B67"].value = f"=B66/1000*{price}"
    asm["A10"].value = "Plastic density (from the Resins tab, resin on Estimator B8)"
    asm["B10"].value = f"=INDEX({RES('B')},MATCH(Estimator!$B$8,{RES('A')},0))"
    asm["B10"].font = F_LINK
    asm["D10"].value = "Linked. Set the density on the Resins tab (override column H)"
    est["A3"].value = (f"Revision {REV}: BOM Paste tab, busbar press size rule of thumb, and a Resins tab "
                       "copied from the EDC estimator (resin pick in B8). Fuse block geometry is unchanged from v1.9.")


# ---------------------------------------------------------------- Busbar press size
def press_size(wb):
    ws = wb["Busbar Stamping"]
    for col, w in (("G", 3), ("H", 44), ("I", 13), ("J", 13), ("K", 13), ("L", 48)):
        ws.column_dimensions[col].width = w
    r = 53
    for col in "HIJKL":
        ws[f"{col}{r}"].fill = HEAD_FILL
    put(ws, f"H{r}", "PRESS SIZE (RULE OF THUMB, METRIC t)", F_HEAD, HEAD_FILL)
    rows = [
        ("shear", "Strip shear strength (MPa)", 200, "0", "Placeholder. C110 half hard: about 0.7 x 290 MPa tensile"),
        ("perim", "Outline cut length (x 2 x (length + width) per part)", 1.3, "0.00",
         "Placeholder. Tabs, notches and relief cuts add cut length beyond the bounding box"),
        ("hole", "Perimeter per pierced feature (mm)", 16, "0.0", "Placeholder. A 5 mm hole is 16 mm"),
        ("cutoff", "Cutoff and carrier trim (x strip width, per hit)", 1, "0.00", "Placeholder. One cut across the strip"),
        ("allow", "Forming, stripping and pad allowance (% of cut force)", 0.3, "0%",
         "Placeholder. Stripping 10-20%, bends and spring pads the rest"),
        ("margin", "Press capacity margin (x)", 1.25, "0.00",
         "Run at 80% or less of rated tonnage; presses are rated near bottom of stroke"),
        ("dieEnd", "Die end allowance each side (mm)", 75, "0",
         "Placeholder. Die shoe beyond the first and last station"),
        ("ovrT", "Press size override (t), blank = rule of thumb", None, "0", "Enter the press the plant will run"),
        ("ovrR", "Press rate override ($/hr), blank = press table", None, "\\$0.00", "Press with operator"),
    ]
    P = {}
    for key, label, val, fmt, note in rows:
        r += 1
        put(ws, f"H{r}", label)
        inp(ws, f"I{r}", val, fmt)
        put(ws, f"L{r}", note, F_NOTE)
        P[key] = f"$I${r}"
    r += 2
    for col in "HIJKL":
        ws[f"{col}{r}"].fill = HEAD_FILL
    put(ws, f"H{r}", "Press size per layout", F_HEAD, HEAD_FILL)
    for col, t in (("I", "1 per hit"), ("J", "2 per hit"), ("K", "Selected")):
        put(ws, f"{col}{r}", t, F_HEAD, HEAD_FILL)
    hdr = r
    table0 = hdr + 12  # first press-list row, set below
    tableN = table0 + len(PRESSES) - 1
    sizes = f"$H${table0}:$H${tableN}"
    rates = f"$I${table0}:$I${tableN}"
    beds = f"$J${table0}:$J${tableN}"
    calc = [
        ("cut", "Cut length per hit (mm)", "0",
         lambda c, n: (f"={n}*({P['perim']}*2*($B$40+$B$41)+$B$67*{P['hole']})"
                       f"+{P['cutoff']}*{c}85")),
        ("force", "Cutting force (t)", "0.0", lambda c, n: f"={c}_cut*$B$16*{P['shear']}/9806.65"),
        ("total", "With forming and stripping (t)", "0.0", lambda c, n: f"={c}_force*(1+{P['allow']})"),
        ("rt", "Rule-of-thumb press size (t)", "0", lambda c, n: f"={c}_total*{P['margin']}"),
        ("die", "Progressive die length (mm)", "0", lambda c, n: f"={c}99*{c}84+2*{P['dieEnd']}"),
        ("std", "Next standard press (t)", "0",
         lambda c, n: (f'=IFERROR(INDEX({sizes},MATCH(1,INDEX(({sizes}>={c}_rt)*({beds}>={c}_die),0),0)),'
                       f'"Over press list")')),
        ("used", "Press size used (t)", "0",
         lambda c, n: f"=IF(ISNUMBER({P['ovrT']}),{P['ovrT']},IF(ISNUMBER({c}_std),{c}_std,{c}_rt))"),
        ("tbl", "Rate from press table ($/hr)", "\\$0.00",
         lambda c, n: (f"=IFERROR(INDEX({rates},MATCH(TRUE,INDEX({sizes}>={c}_used,0),0)),"
                       f"INDEX({rates},ROWS({rates})))")),
        ("rate", "Press rate used ($/hr)", "\\$0.00",
         lambda c, n: f"=IF(ISNUMBER({P['ovrR']}),{P['ovrR']},{c}_tbl)"),
    ]
    R = {}
    for i, (key, *_rest) in enumerate(calc):
        R[key] = hdr + 1 + i
    for key, label, fmt, fn in calc:
        rr = R[key]
        put(ws, f"H{rr}", label, F_BOLD if key in ("used", "rate") else F_LABEL)
        for col, srccol, n in (("I", "B", "$B$80"), ("J", "C", "$C$80")):
            f = fn(srccol, n)
            for k2, r2 in R.items():  # resolve the column's own earlier rows
                f = f.replace(f"{srccol}_{k2}", f"{col}{r2}")
            put(ws, f"{col}{rr}", f, F_BOLD if key in ("used", "rate") else F_LABEL, fmt=fmt)
        put(ws, f"K{rr}", f"=IF($B$44=2,J{rr},I{rr})", F_BOLD if key in ("used", "rate") else F_LABEL, fmt=fmt)
    put(ws, f"L{R['cut']}", "Parts per hit x (outline + pierced features) + cutoff. All die stations cut on every hit", F_NOTE)
    put(ws, f"L{R['force']}", "Cut length x thickness x shear strength", F_NOTE)
    put(ws, f"L{R['die']}", "Die stations x pitch + end allowance. Often sets the press, not tonnage", F_NOTE)
    put(ws, f"L{R['std']}", "Smallest press in the list with enough tonnage AND bed length for the die", F_NOTE)
    r = table0 - 1
    for col in "HIJKL":
        ws[f"{col}{r}"].fill = SUB_FILL
    put(ws, f"H{r}", "Press list: size (t)", F_BOLD, SUB_FILL)
    put(ws, f"I{r}", "$/hr with operator", F_BOLD, SUB_FILL)
    put(ws, f"J{r}", "Bed length (mm)", F_BOLD, SUB_FILL)
    put(ws, f"L{r}", "Placeholders. Replace with the stamping plant's presses, rates and bed sizes", F_NOTE, SUB_FILL)
    for i, (t, rate, bed) in enumerate(PRESSES):
        inp(ws, f"H{table0 + i}", t, "0")
        inp(ws, f"I{table0 + i}", rate, "\\$0.00")
        inp(ws, f"J{table0 + i}", bed, "0")
    assert table0 == R["rate"] + 3, "press list position"

    # Results now use the press rate for each layout
    ws["A56"].value = "Press rate used, selected layout ($/hr)"
    ws["B56"].value = f"=K{R['rate']}"
    ws["B56"].font = F_LINK
    ws["B56"].fill = PatternFill(fill_type=None)
    ws["F56"].value = f"From the PRESS SIZE block (H53). Override the press size or rate there"
    ws["B97"].value = f"=I{R['rate']}/B96"
    ws["C97"].value = f"=J{R['rate']}/C96"
    ws["B98"].value = f"=$B$57*I{R['rate']}/$B$58"
    ws["C98"].value = f"=$B$57*J{R['rate']}/$B$58"
    return {"used": f"$K${R['used']}", "rt": f"$K${R['rt']}"}


# ---------------------------------------------------------------- BOM Paste
def bom_paste(wb, press):
    ws = wb.create_sheet("BOM Paste", index=1)
    ws.sheet_properties.tabColor = "00B050"
    ws.sheet_view.showGridLines = False
    widths = {"A": 22, "B": 40, "C": 15, "D": 34, "E": 60, "F": 14, "G": 16}
    for k, w in widths.items():
        ws.column_dimensions[k].width = w
    put(ws, "A1", "BOM PASTE: one row per part, laid out for the BOM calculator", Font(name="Calibri", size=14, bold=True))
    put(ws, "A2", ("Enter part numbers and descriptions (yellow). Copy A:E, then Paste Special > Values into the BOM "
                   "calculator. F and G split the cost: add the same two columns after Notes in the BOM calculator "
                   "to carry them, or leave them off. Material + Machine/Labor does not equal Total on the stamped "
                   "busbar row: tooling, packaging and SG&A are only in Total."), F_NOTE, wrap=True)
    ws.merge_cells("A2:G2")
    ws.row_dimensions[2].height = 40
    heads = ["Part Number", "Part Description", "Total Unit Cost (USD)", "Estimate Source / Rev", "Notes",
             "Material (USD)", "Machine + Labor (USD)"]
    for i, h in enumerate(heads):
        c = ws.cell(4, 1 + i, h)
        c.font = F_HEAD
        c.fill = HEAD_FILL
        c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[4].height = 30
    E = "Estimator!"
    S = "'Busbar Stamping'!"
    src_est = f'"BFT {REV} Estimator"'
    rows = [
        ("Stamped PCB busbar",
         f"={S}$D$106", f'="BFT {REV} Busbar Stamping"',
         (f'="Progressive die, "&{S}$D$80&" per hit, "&TEXT({S}{press["used"]},"0")&" t press. Total also has tooling "'
          f'&TEXT({S}$D$101+{S}$D$102,"$0.0000")&", pack/inspect "&TEXT({S}$D$103,"$0.0000")&", SG&A "&TEXT({S}$D$105,"$0.0000")'),
         f"={S}$D$95", f"={S}$D$97+{S}$D$98"),
        ("Fuse block plastic base", f"={E}$B$78", f"={src_est}",
         f'="Resin only ("&{E}$B$8&", "&TEXT({E}$B$58,"0")&" g). Molding not estimated: use the EDC estimator for machine and labor"',
         f"={E}$B$78", "=0"),
        ("Fuse block cover", f"={E}$B$79", f"={src_est}",
         f'="Resin only ("&{E}$B$8&", "&TEXT({E}$B$66,"0")&" g). Molding not estimated"',
         f"={E}$B$79", "=0"),
        ("Fuse block busbar (copper)", f"={E}$B$80", f"={src_est}",
         f'="Copper and blank scrap only ("&TEXT({E}$B$48,"0")&" g). For a piece price use the Busbar Stamping row"',
         f"={E}$B$80", "=0"),
        ("Fuse bolts and nuts (purchased)", f"={E}$B$81", f"={src_est}", '="Purchased hardware"', f"={E}$B$81", "=0"),
        ("B+ connection part (purchased)", f"={E}$B$82", f"={src_est}", f'="B+ connection: "&{E}$B$12',
         f"={E}$B$82", "=0"),
        ("Keepout zone stud hardware (purchased)", f"={E}$B$74", f"={src_est}", '="Zero when no zone studs"',
         f"={E}$B$74", "=0"),
        ("Assembly process adders", f"={E}$B$83+{E}$B$84+{E}$B$75+{E}$B$76", f"={src_est}",
         f'="Bolt retention ("&{E}$B$13&"), joining ("&{E}$B$14&"), zone laser welds, short-side bus joint"',
         "=0", f"={E}$B$83+{E}$B$84+{E}$B$75+{E}$B$76"),
    ]
    r = 4
    for desc, total, srcf, note, mat, ml in rows:
        r += 1
        inp(ws, f"A{r}", None, "@")
        inp(ws, f"B{r}", desc, "@")
        put(ws, f"C{r}", total, F_BOLD, fmt=MONEY4, border=BOX)
        put(ws, f"D{r}", srcf, border=BOX)
        put(ws, f"E{r}", note, border=BOX)
        put(ws, f"F{r}", mat, fmt=MONEY4, border=BOX)
        put(ws, f"G{r}", ml, fmt=MONEY4, border=BOX)
    first_fb, last = 6, r
    r += 2
    put(ws, f"A{r}", "Check", F_BOLD)
    put(ws, f"B{r}", f'=IF(ABS(SUM(C{first_fb}:C{last})-{E}$B$86)<0.00005,"Fuse block rows add up to Estimator total "'
                     f'&TEXT({E}$B$86,"$0.0000"),"Fuse block rows do not add up to Estimator B86, check the tab")')
    ws.merge_cells(f"B{r}:E{r}")
    r += 1
    put(ws, f"A{r}", "Rows 6 to 12 are the fuse block (Estimator tab). Row 5 is a separate stamped busbar (Busbar Stamping tab). "
                     "Leave out rows you do not need when you paste.", F_NOTE)
    ws.freeze_panes = "A5"


# ---------------------------------------------------------------- Start Here
def start_here(wb):
    ws = wb["Start Here"]
    ws.insert_rows(12, 3)  # after the busbar steps, before the Busbar Sizer section
    put(ws, "A12", "BOM roll-up (green tab: BOM Paste)", F_HEAD, HEAD_FILL)
    ws["B12"].fill = HEAD_FILL
    put(ws, "A13", 1)
    put(ws, "B13", "BOM Paste: enter part numbers in column A, copy A:E and Paste Special > Values into the BOM "
                   "calculator. Material and Machine + Labor are in F and G.", wrap=False)
    r = 29  # Strip Stock row moved from 24 to 27; add Resins after Fuse Library (29)
    assert ws["A29"].value == "Fuse Library", ws["A29"].value
    ws.insert_rows(30, 1)
    put(ws, "A30", "Resins")
    put(ws, "B30", "Copied from the EDC estimator Rev1.30 (not linked). Paste the same material system export at M3; "
                   "new resins appear in the Estimator B8 resin list. Price and density come from here.")
    last = ws.max_row
    put(ws, f"A{last + 1}", REV, F_BOLD, SUB_FILL)
    put(ws, f"B{last + 1}", "BOM Paste tab (one row per part, material and machine/labor split). Busbar Stamping: "
                            "rule-of-thumb press size with press size and rate overrides and a press list. Resins tab "
                            "copied from the EDC estimator; Estimator B8 is now a resin pick.")
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for c in row:
            if isinstance(c.value, str) and "for example Rev1.11" in c.value:
                c.value = c.value.replace("for example Rev1.11", "for example Rev1.12")


def build(out=OUT, cache=True):
    wb = load_workbook(SRC)
    # openpyxl drops Rev1.10's cross-sheet (x14) dropdown on Busbar Stamping B6; put it back
    dv = DataValidation(type="list", formula1="='Strip Stock'!$A$5:$A$30", showErrorMessage=True)
    wb["Busbar Stamping"].add_data_validation(dv)
    dv.add("B6")
    copy_resins(wb)
    resin_pick(wb)
    press = press_size(wb)
    bom_paste(wb, press)
    start_here(wb)
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    if cache:
        add_cached_values(out)
    return out


if __name__ == "__main__":
    print(build(*sys.argv[1:]))
