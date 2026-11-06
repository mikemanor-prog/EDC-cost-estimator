"""Builds Battery_Fuse_Block_Cost_Estimator_Rev1.12.xlsm: Rev1.11 plus Clear buttons.

Rev1.12 adds red Clear buttons (macros, so the file is .xlsm) like the EDC estimator Rev1.30:
  * Busbar Stamping: clears the part inputs (part size, measured area, terminal counts,
    tabs / bends / pierces, pitch and strip overrides, press size and rate overrides).
    Material source, annual volume, rates, tooling costs and the press list are kept.
  * BOM Paste: clears the part numbers and resets the descriptions.
The VBA source is bft/tools/vba/ClearButtons_Rev1_12.bas. LibreOffice turns it into a
source-only vbaProject.bin (Excel compiles it on first open).
Run: python bft/tools/build_rev1_12.py
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile

from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_rev1_11 as r11  # noqa: E402
from build_v1_7 import add_cached_values  # noqa: E402

BFT = os.path.dirname(HERE)
OUT = os.path.join(BFT, "Battery_Fuse_Block_Cost_Estimator_Rev1.12.xlsm")
BAS = os.path.join(HERE, "vba", "ClearButtons_Rev1_12.bas")
REV = "Rev1.12"

# sheet title -> (macro, anchor col, anchor row, both zero-based, button text)
BUTTONS = {
    "Busbar Stamping": ("ClearBusbarStamping", 2, 0, "Clear part inputs"),
    "BOM Paste": ("ClearBOMPaste", 3, 0, "Clear part numbers"),
}


def check_macro_ranges(wb):
    """The macro clears fixed cells; fail the build if the layout moved."""
    bs = wb["Busbar Stamping"]
    expect = {"A7": "Manual: copper alloy", "A23": "Body length (mm)", "A26": "Measured total blank area",
              "A46": "Pitch override", "A47": "Strip width override", "A63": "Upturned PCB tabs",
              "A65": "Pierced features (holes", "H61": "Press size override", "H62": "Press rate override"}
    for ref, start in expect.items():
        assert str(bs[ref].value).startswith(start), (ref, bs[ref].value)
    assert str(bs["A28"].value).startswith("='Terminal Features'!$A$5") and str(bs["A35"].value).endswith("$A$12")
    bp = wb["BOM Paste"]
    assert bp["A4"].value == "Part Number" and bp["B12"].value == "Assembly process adders"
    src = open(BAS).read()
    for i in range(5, 13):
        assert f'"{bp[f"B{i}"].value}"' in src, bp[f"B{i}"].value


def build_xlsx(path):
    wb = load_workbook(r11.SRC)
    dv = DataValidation(type="list", formula1="='Strip Stock'!$A$5:$A$30", showErrorMessage=True)
    wb["Busbar Stamping"].add_data_validation(dv)
    dv.add("B6")
    r11.copy_resins(wb)
    r11.REV = REV
    r11.resin_pick(wb)
    press = r11.press_size(wb)
    r11.bom_paste(wb, press)
    r11.REV = "Rev1.11"
    r11.start_here(wb)  # adds the Rev1.11 history row
    r11.REV = REV
    sh = wb["Start Here"]
    last = sh.max_row
    r11.put(sh, f"A{last + 1}", REV, r11.F_BOLD, r11.SUB_FILL)
    r11.put(sh, f"B{last + 1}", "Red Clear buttons on Busbar Stamping (part inputs) and BOM Paste (part numbers). "
                                "The file is .xlsm because of the macros: click Enable Content to use them.")
    for row in sh.iter_rows():
        for c in row:
            if isinstance(c.value, str):
                c.value = c.value.replace("for example Rev1.12", "for example Rev1.13")
                if c.value.startswith("Save a new revision as Rev1.xx"):
                    c.value += " The file is .xlsm because of the Clear buttons."
    assert sh["A12"].value.startswith("BOM roll-up") and sh["A14"].value is None
    r11.insert_rows(sh, 15, 1)
    r11.put(sh, "A14", "Clear")
    r11.put(sh, "B14", "The red Clear button at the top of Busbar Stamping and BOM Paste empties that tab's part "
                       "inputs (it asks first). Click Enable Content when Excel shows the macro bar.")
    wb["Estimator"]["A3"].value = (f"Revision {REV}: Clear buttons (Busbar Stamping, BOM Paste). Rev1.11 added the BOM "
                                   "Paste tab, busbar press size rule of thumb and the Resins tab (resin pick in B8). "
                                   "Fuse block geometry is unchanged from v1.9.")
    wb.code_name = "ThisWorkbook"
    check_macro_ranges(wb)
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def make_vba_bin(workdir):
    """Let LibreOffice write a source-only vbaProject.bin holding the .bas module."""
    import uno
    from com.sun.star.beans import PropertyValue

    def prop(n, v):
        p = PropertyValue()
        p.Name, p.Value = n, v
        return p

    seed = os.path.join(workdir, "seed.xlsm")
    shutil.copy(r11.EDC, seed)  # has a VBA project; its Module1 is replaced below
    port = 2017
    proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore",
                             f"--accept=socket,host=localhost,port={port};urp;"])
    try:
        ctx = None
        for _ in range(90):
            try:
                local = uno.getComponentContext()
                res = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
                ctx = res.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(1)
        desk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(seed), "_blank", 0, (prop("Hidden", True),))
        libs = doc.BasicLibraries
        lib = libs.getByName("VBAProject")
        lib.replaceByName("Module1", "Option VBASupport 1\n" + open(BAS).read())
        libs.VBACompatibilityMode = True
        out = os.path.join(workdir, "vba.xlsm")
        doc.storeToURL(uno.systemPathToFileUrl(out), (prop("FilterName", "Calc MS Excel 2007 VBA XML"),))
        doc.close(True)
        try:
            desk.terminate()
        except Exception:
            pass
    finally:
        try:
            proc.wait(timeout=30)
        except subprocess.TimeoutExpired:
            proc.kill()
    with zipfile.ZipFile(out) as z:
        return z.read("xl/vbaProject.bin")


def drawing_xml(macro, col, row, text, shape_id=2):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<xdr:oneCellAnchor><xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>95250</xdr:colOff>'
            f'<xdr:row>{row}</xdr:row><xdr:rowOff>38100</xdr:rowOff></xdr:from>'
            '<xdr:ext cx="1737360" cy="320040"/>'
            f'<xdr:sp macro="[0]!{macro}" textlink=""><xdr:nvSpPr>'
            f'<xdr:cNvPr id="{shape_id}" name="Clear button" descr="Clears the part inputs on this tab"/>'
            '<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1737360" cy="320040"/></a:xfrm>'
            '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="C00000"/></a:solidFill>'
            '<a:ln><a:noFill/></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" '
            'wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r>'
            '<a:rPr lang="en-US" sz="1100" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>'
            f'<a:t>{text}</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>')


def to_xlsm(xlsx, vba_bin, out):
    """Add the VBA project and the button drawings, and mark the workbook macro-enabled."""
    zin = zipfile.ZipFile(xlsx)
    files = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    ct = files["[Content_Types].xml"].decode()
    ct = ct.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
    if 'Extension="bin"' not in ct:
        ct = ct.replace("<Default ", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ', 1)
    rels = files["xl/_rels/workbook.xml.rels"].decode()
    rels = rels.replace("</Relationships>",
                        '<Relationship Id="rIdVBA1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                        'Target="vbaProject.bin"/></Relationships>')
    files["xl/_rels/workbook.xml.rels"] = rels.encode()
    files["xl/vbaProject.bin"] = vba_bin

    # map sheet titles to sheetN.xml through workbook.xml and its rels
    wbxml = files["xl/workbook.xml"].decode()
    rid_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="/?(?:xl/)?([^"]+)"', rels))
    rid_target.update({k: v for v, k in re.findall(r'Target="/?(?:xl/)?([^"]+)"[^>]*Id="(rId\d+)"', rels)})
    sheets = {name: rid_target[rid] for name, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wbxml)}
    n_draw = max([int(m) for m in re.findall(r"xl/drawings/drawing(\d+)\.xml$", "\n".join(files), re.M)] + [0])
    for title, (macro, col, row, text) in BUTTONS.items():
        sheet_path = "xl/" + sheets[title]
        sx = files[sheet_path].decode()
        assert "<drawing " not in sx and "<legacyDrawing" not in sx, f"{title} already has a drawing"
        n_draw += 1
        dname = f"xl/drawings/drawing{n_draw}.xml"
        files[dname] = drawing_xml(macro, col, row, text).encode()
        ct = ct.replace("</Types>", f'<Override PartName="/{dname}" '
                        'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
        rel_path = sheet_path.replace("worksheets/", "worksheets/_rels/") + ".rels"
        rel = ('<Relationship Id="rIdBtn1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
               f'Target="../drawings/drawing{n_draw}.xml"/>')
        if rel_path in files:
            files[rel_path] = files[rel_path].decode().replace("</Relationships>", rel + "</Relationships>").encode()
        else:
            files[rel_path] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships '
                               'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                               f"{rel}</Relationships>").encode()
        if "xmlns:r=" not in sx[:600]:
            sx = sx.replace("<worksheet ", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
        # <drawing> goes after pageSetup / headerFooter / rowBreaks etc. and before legacyDrawing / tableParts / extLst
        m = re.search(r"<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b", sx)
        tag = '<drawing r:id="rIdBtn1"/>'
        sx = sx[:m.start()] + tag + sx[m.start():] if m else sx.replace("</worksheet>", tag + "</worksheet>")
        files[sheet_path] = sx.encode()
    files["[Content_Types].xml"] = ct.encode()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", files.pop("[Content_Types].xml"))
        for n, data in files.items():
            z.writestr(n, data)


def build(out=OUT):
    d = tempfile.mkdtemp()
    xlsx = os.path.join(d, "rev.xlsx")
    build_xlsx(xlsx)
    add_cached_values(xlsx)
    vba = make_vba_bin(d)
    to_xlsm(xlsx, vba, out)
    shutil.rmtree(d, ignore_errors=True)
    return out


if __name__ == "__main__":
    print(build(*sys.argv[1:]))
