# (c) 2026 Alain Yon
"""Extra checks from spec section 10.4, run through the app's engine on the demo sample data.

python test_extras.py
"""
import datetime
import glob
import os
import shutil
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, ROOT)
os.environ["PYTHONHASHSEED"] = "0"

from compare_xlsx import compare  # noqa: E402
import run_regression as rr  # noqa: E402
from cases import sample_dir  # noqa: E402

SRC = os.path.join(ROOT, "source_tools")
FAILS = []
MONTH = datetime.date.today().strftime("%b_%Y")          # OE names its approval file after the run month
MI20 = "MI20_Kitchen_Oct_2026.xlsx"


def check(name, cond, extra=""):
    print(("ok    " if cond else "FAIL  ") + name + (f"\n        {extra}" if extra and not cond else ""))
    if not cond:
        FAILS.append(name)


def new_root(*folders, name="ccs_x_"):
    root = tempfile.mkdtemp(prefix=name)
    for f in folders:
        dst = os.path.join(root, "tools", f)
        shutil.copytree(os.path.join(SRC, f), dst, ignore=shutil.ignore_patterns("samples", "__pycache__"))
        tf = os.path.join(sample_dir(f), "tool_files")
        if os.path.isdir(tf):
            shutil.copytree(tf, dst, dirs_exist_ok=True)
    os.makedirs(os.path.join(root, "inbox"))
    return root


def sample(folder, rel):
    return os.path.join(sample_dir(folder), "inputs", rel.replace("/", os.sep))


def golden(folder, case, name):
    return os.path.join(sample_dir(folder), "expected", case, name)


def engine_for(root):
    from app.engine import Engine
    return Engine(data_root=root, app_root=os.path.join(root, "_app"))


def run(eng, tool, inputs=None, options=None, answers=()):
    answers = list(answers)
    res = eng.start_run(tool, inputs=inputs or {}, options=options or {})
    if res.get("error"):
        return {"status": "not_started", "headline": res["error"], "outputs": [], "notices": []}, []
    cursor, lines = 0, []
    while True:
        p = eng.poll(res["run_id"], cursor)
        cursor = p["cursor"]
        for ev in p["events"]:
            if ev["type"] == "line":
                lines.append(ev["text"])
            elif ev["type"] == "prompt":
                eng.answer(res["run_id"], answers.pop(0) if answers else "")
        if p["done"]:
            return p["result"], lines
        time.sleep(0.05)


def same(a, b, tool_ignores):
    r = compare(a, b, tool_ignores)
    return not r["diffs"], "; ".join(r["diffs"][:4])


def main():
    ign = rr.load_manifest_ignores()
    from app import branding
    H, C = branding.HOTEL_PLANT, branding.CC_PLANT

    # ---- shared SAP data: sample's own MM60 / MI20 in sap_data, output identical to the .bat run (golden) ----
    root = new_root("oe_monthly")
    eng = engine_for(root)
    check("SAP data: a file with the wrong headers is refused",
          not eng.sap_replace("mm60", sample("oe_monthly", "MB52_OE_Oct_2026.xlsx"))["ok"] and eng.sap.path("mm60") is None)
    check("SAP data: MM60 loads", eng.sap_replace("mm60", sample("oe_monthly", "MM60_D100_Oct_2026.xlsx"))["ok"])
    bad = eng.sap_replace("mm60", sample("breakfast_tracker", MI20))
    check("SAP data: replacing MM60 with an MI20 is refused and the current file stays",
          not bad["ok"] and eng.sap.info("mm60")["original_name"] == "MM60_D100_Oct_2026.xlsx", str(bad))
    for f in ("MB52_OE_Oct_2026.xlsx", "MB51_Oct_2026.xlsx"):
        shutil.copy2(sample("oe_monthly", f), os.path.join(root, "inbox", f))
    oe_in = {"files": [os.path.join(root, "inbox", "MB52_OE_Oct_2026.xlsx"), os.path.join(root, "inbox", "MB51_Oct_2026.xlsx")]}
    res, lines = run(eng, "oe_monthly", oe_in)
    check("OE monthly with MM60 from SAP data finishes", res["status"] == "finished", res["headline"])
    tdir = os.path.join(root, "tools", "oe_monthly")
    appr_name = f"OE_Approval_{MONTH}.xlsx"
    for gname, pat in (("OE_Tracking_Workbook.xlsx", "OE_Tracking_Workbook.xlsx"), ("OE_Approval.xlsx", appr_name),
                       ("OE_Transfer_Form.xlsx", "OE_Transfer_Form.xlsx")):
        ok, why = same(golden("oe_monthly", "oe_monthly", gname), os.path.join(tdir, pat), ign["oe_monthly"])
        check(f"OE monthly (shared MM60): {gname} identical to the .bat run", ok, why)

    # ---- OE twice: an approval file with decisions is not overwritten, and the message is shown ----
    from openpyxl import load_workbook
    appr = os.path.join(tdir, appr_name)
    wb = load_workbook(appr)
    ws = wb["APPROVAL"]
    hdr = next(r for r in range(1, 15) if any(str(c.value or "").strip().lower() == "approve (y/n)" for c in ws[r]))
    col = next(c.column for c in ws[hdr] if str(c.value or "").strip().lower() == "approve (y/n)")
    ws.cell(hdr + 1, col).value = "Y"
    wb.save(appr)
    stamp = (os.path.getmtime(appr), os.path.getsize(appr))
    res2, lines2 = run(eng, "oe_monthly", oe_in)
    check("OE second run finishes", res2["status"] == "finished", res2["headline"])
    check("OE second run: approval file with decisions untouched", stamp == (os.path.getmtime(appr), os.path.getsize(appr)))
    check("OE second run: a _rebuilt_ approval file was written instead",
          bool(glob.glob(os.path.join(tdir, f"OE_Approval_{MONTH}_rebuilt_*.xlsx"))))
    check("OE second run: the script's message is surfaced on the result", any("left alone" in n for n in res2["notices"]), str(res2["notices"]))

    # ---- kitchen MI20 from SAP data: breakfast, canteen, coffee ----
    for tool, case, out, opts, target in (
            ("breakfast_tracker", "breakfast_tracker", "Breakfast_MonthEnd_Tool.xlsx", {"month": "October 2026"}, "Breakfast_MonthEnd_Tool.xlsx"),
            ("canteen_loss", "canteen_loss", "Canteen_MonthEnd_Tool.xlsx", {"month": "October 2026"}, "Canteen_MonthEnd_Tool.xlsx"),
            ("coffee_loss", "coffee_loss", "Coffee_Loss_Tracker.xlsx",
             dict(month="October 2026", qty_0="446", qty_1="12", qty_2="30", qty_3="500", qty_4="200", qty_5="15", qty_6="10", qty_7="5", qty_8="40", qty_9="25"),
             "Coffee_Loss_Tracker_October_2026.xlsx")):
        root = new_root(tool)
        eng = engine_for(root)
        eng.sap_replace("mi20_kitchen", sample(tool, MI20))
        if tool == "coffee_loss":   # a stale export left in the folder must be moved out, not read
            os.makedirs(os.path.join(root, "tools", tool, "Mi20_files"), exist_ok=True)
            shutil.copy2(sample("recount_selection", "L123_Variance_Oct_2026.xlsx"), os.path.join(root, "tools", tool, "Mi20_files", "old_export.xlsx"))
        res, lines = run(eng, tool, {}, opts)
        check(f"{tool} with MI20 from SAP data finishes", res["status"] == "finished", res["headline"])
        ok, why = same(golden(tool, case, out), os.path.join(root, "tools", tool, target), ign[tool])
        check(f"{tool} (shared MI20): output identical to the .bat run", ok, why)
        if tool == "coffee_loss":
            check("coffee: the stale MI20 was moved to _backup", bool(glob.glob(os.path.join(root, "tools", tool, "_backup", "*", "old_export.xlsx"))))

    # ---- comparator: previous MM60 from history, current from sap_data ----
    root = new_root("cost_comparator")
    eng = engine_for(root)
    eng.sap_replace("mm60", sample("cost_comparator", "MM60_D100_Sep_2026.xlsx"))
    eng.sap_replace("mm60", sample("cost_comparator", "MM60_D100_Oct_2026.xlsx"))
    res, lines = run(eng, "cost_comparator", {}, {"label_old": "September 2026", "label_new": "October 2026"})
    check("comparator with MM60 history + current finishes", res["status"] == "finished", res["headline"])
    ok, why = same(golden("cost_comparator", "cost_comparator", "MM60_Comparative.xlsx"),
                   os.path.join(root, "tools", "cost_comparator", "MM60_Comparative_September_2026_vs_October_2026.xlsx"), ign["cost_comparator"])
    check("comparator (shared MM60): output identical to the .bat run", ok, why)

    # ---- Theo CoS, embedded sheet: hard gate ----
    root = new_root("theo_cos")
    eng = engine_for(root)
    user_file = os.path.join(root, "inbox", "CC_Bar_Sales_Oct_2026.xlsx")
    shutil.copy2(sample("theo_cos", "CC_Bar_Sales_Oct_2026.xlsx"), user_file)
    before = open(user_file, "rb").read()
    wb = load_workbook(user_file, read_only=True, data_only=True)
    from openpyxl import Workbook
    alone = Workbook()
    for row in wb["MM60"].iter_rows(values_only=True):
        alone.active.append(list(row))
    wb.close()
    mm60_alone = os.path.join(root, "inbox", "MM60_CC.xlsx")
    alone.save(mm60_alone)
    loaded = eng.sap_replace("mm60", mm60_alone)
    check(f"MM60 for plant {C} is recognised and filed as the conference centre MM60, whichever card it is dropped on",
          loaded.get("ok") and loaded.get("kind") == "mm60_cc" and eng.sap.path("mm60") is None, str(loaded))
    eng.sap_replace("mm60_cc", sample("oe_monthly", "MM60_D100_Oct_2026.xlsx"))      # a hotel file dropped on the CC card
    check(f"MM60 for plant {H} is filed as the hotel MM60", eng.sap.info("mm60").get("plants") == [H]
          and eng.sap.info("mm60_cc").get("plants") == [C])
    # default: the Sales Mix is a conference centre file, so its plant's MM60 is the one written in
    res, lines = run(eng, "theo_cos", {"file": [user_file]})
    check(f"Theo picks the MM60 by the Sales Mix's plant ({C})", any(f"({C})" in ln for ln in lines),
          " | ".join(lines[:3]))
    check("Theo with the shared MM60 written into a temp copy finishes", res["status"] == "finished", res["headline"])
    out = os.path.join(root, "inbox", "CC_Bar_Sales_Oct_2026_CoS_Summary.xlsx")
    ok, why = (False, "no report next to the user's file") if not os.path.exists(out) else same(golden("theo_cos", "theo_cos", "CoS_Summary.xlsx"), out, ign["theo_cos"])
    check("Theo (embedded_sheet): report identical to the .bat run  [hard gate]", ok, why)
    check("Theo (embedded_sheet): the user's original file is unchanged", open(user_file, "rb").read() == before)
    check("Theo (embedded_sheet): the embed was actually used", any("temporary copy" in ln for ln in lines))

    # ---- a file locked in Excel fails pre-flight with the plain message ----
    root = new_root("breakfast_tracker")
    eng = engine_for(root)
    eng.sap_replace("mi20_kitchen", sample("breakfast_tracker", MI20))
    lock = os.path.join(root, "tools", "breakfast_tracker", "~$Breakfast_MonthEnd_Tool.xlsx")
    open(lock, "w").close()
    checks = eng.preflight("breakfast_tracker", {}, {"month": "October 2026"})
    msg = next((c["message"] for c in checks if c["status"] == "fail"), "")
    check("locked file fails pre-flight with the plain message",
          msg == "Close Breakfast_MonthEnd_Tool.xlsx in Excel first (someone else may have it open)", msg)
    check("locked file: the run is refused", "error" in eng.start_run("breakfast_tracker", inputs={}, options={"month": "October 2026"}))
    os.remove(lock)

    # ---- paths with spaces, & and ! ----
    root = new_root("recount_selection", "stock_age", name="ccs x & co! ")
    eng = engine_for(root)
    odd = os.path.join(root, "in box & more!")
    os.makedirs(odd)
    files = []
    for f in ("L123_Variance_Oct_2026.xlsx", "Pool_Bar_Sales_Oct_2026.xlsx", "Cocktail_Tracker.xlsx"):
        shutil.copy2(sample("recount_selection", f), os.path.join(odd, f))
        files.append(os.path.join(odd, f))
    res, lines = run(eng, "recount_selection", {"files": files}, {"mode": "build"}, answers=["Bar Manager", "25"])
    check("odd path (spaces, & and !): recount reaches the approval gate and finishes", res["status"] == "finished" and any("APPROVAL REQUIRED" in ln for ln in lines), res["headline"])
    outs = glob.glob(os.path.join(odd, "Recount_Selection_*.xlsx"))
    ok, why = (False, "no sheet written") if not outs else same(golden("recount_selection", "recount_gate", "Recount_Selection.xlsx"), outs[0], ign["recount_selection"])
    check("odd path: recount sheet identical to the .bat run, approver name included", ok, why)
    if outs:
        wb = load_workbook(outs[0])
        hits = [f"{ws.title}!{c.coordinate}" for ws in wb.worksheets for row in ws.iter_rows() for c in row
                if isinstance(c.value, str) and "Bar Manager" in c.value]
        print("        approver name found at: " + ", ".join(hits[:4]))
        check("recount: approver name is written on the sheet", bool(hits))
    res, lines = run(eng, "recount_selection", {"files": files}, {"mode": "build"}, answers=[""])
    check("recount: empty approver name cancels and writes nothing", res["status"] == "cancelled" and len(glob.glob(os.path.join(odd, "Recount_Selection_*.xlsx"))) == len(outs), res["headline"])
    src = os.path.join(odd, "Stock_Age_Source.xlsx")
    shutil.copy2(sample("stock_age", "Stock_Age_Source.xlsx"), src)
    res, lines = run(eng, "stock_age", {"file": [src]})
    ok, why = same(golden("stock_age", "stock_age", "Age_Analysis_Dashboard.xlsx"), os.path.join(root, "tools", "stock_age", "Age_Analysis_Dashboard.xlsx"), ign["stock_age"]) if res["status"] == "finished" else (False, res["headline"])
    check("odd path: stock age dashboard identical", ok, why)

    # ---- not-installed tools and one-run-at-a-time ----
    nroot = new_root("stock_age")
    os.remove(os.path.join(nroot, "tools", "stock_age", "build_age_analysis.py"))
    neng = engine_for(nroot)
    check("a tool whose script is not in its folder is reported as not installed", not neng.tool_view("stock_age")["installed"])
    check("a not-installed tool cannot be started", "error" in neng.start_run("stock_age", inputs={"file": [src]}))

    # ---- demo branding: the property name comes from config, never the Windows account ----
    st = engine_for(new_root()).state()["app"]
    check("demo: the property name comes from config/app_settings.json", st["property_name"] == branding.property_name())
    check("demo: the greeting uses the configured display name, not the Windows login",
          st["first_name"] == branding.user_display_name().split()[0] and st["user"] == branding.user_display_name())

    print("\nCellar --no-prompt and prompts-answered runs are regression cases: run_regression.py --only cellar_consumption")
    print("\n" + ("ALL PASSED" if not FAILS else f"{len(FAILS)} FAILED: " + "; ".join(FAILS)))
    return 0 if not FAILS else 1


if __name__ == "__main__":
    sys.exit(main())
