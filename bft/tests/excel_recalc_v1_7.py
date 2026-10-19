import json, os, subprocess, sys, tempfile, openpyxl, warnings
warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "Battery_Fuse_Block_Cost_Estimator_v1.7.xlsx")
OUTS = {"Estimator": ["B26","B27","B28","B34","B35","B36","B37","B38","B39","B40","B41","B44","B45","B47","B48","B49","B51","B52","B53","B54","B55","B57","B58","B59","B66","B67","B70","B71","B72","B73","B74","B75","B76","B85","B86","B87","B88","E62","E63","E64","E65","E66","E13","E14","B15","B31","E15","B63","E67"],
        "Assumptions": ["B73","B74"],
        "Layout": [f"{c}{r}" for r in range(12,15) for c in "EFGHI"] + [f"{c}{r}" for c in "BC" for r in range(31,49)] + [f"B{r}" for r in list(range(51,73))+list(range(75,88))+list(range(90,98))],
        "Busbar Sizer": ["B6","B13","B14","B15","B16","B17","B18","B20","B21","B39","B40"] + [f"{c}{r}" for r in range(26,38) for c in "BCDEFGH"]+[f"I{r}" for r in range(26,38)]}
def run(cases):
    d = tempfile.mkdtemp()
    files = []
    for i, case in enumerate(cases):
        wb = openpyxl.load_workbook(SRC)
        for k, v in case.items():
            sh, c = k.split("!"); wb[sh][c] = v
        fn = f"{d}/c{i}.xlsx"; wb.save(fn); files.append(fn)
    os.makedirs(f"{d}/o")
    subprocess.run(["soffice","--headless","--calc","--convert-to","xlsx","--outdir",f"{d}/o"]+files, capture_output=True, timeout=300)
    res = []
    for i in range(len(cases)):
        wb = openpyxl.load_workbook(f"{d}/o/c{i}.xlsx", data_only=True)
        res.append({f"{sh}!{c}": wb[sh][c].value for sh, cs in OUTS.items() for c in cs})
    return res
if __name__ == "__main__":
    print(json.dumps(run(json.load(open(sys.argv[1]))), indent=0))
