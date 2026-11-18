# (c) 2026 Alain Yon
"""Installer step 4 (copy files) and the SAP data migration, run with the private Python.

  python install_steps.py copy <package folder>     app code + tool folders, with the never-overwrite rules
  python install_steps.py where                     prints the data folder

Rules (spec 9.1 step 4, golden rule 3):
  * app code            -> %LOCALAPPDATA%\\CostControlSuiteDemo\\app      replaced
  * config              -> %LOCALAPPDATA%\\CostControlSuiteDemo\\config   copied only if missing (yours to edit)
  * tool folders        -> Documents\\Cost Control Suite Demo\\tools\\<id>
        anything missing is copied, including the demo's working files from
        sample_data\\<section>\\<id>\\tool_files (item maps, masters, configs)
        .py files are replaced, after a dated backup to _backup\\<yyyymmdd_hhmm>\\
        an existing .xlsx / .csv / .json / .txt (or anything else) is NEVER replaced
  * sample inputs       -> Documents\\Cost Control Suite Demo\\Sample inputs\\<id>
        the generated files to drop on each tool (copied only if missing)
  * MM60 / kitchen MI20 copies lying in tool folders move to tools\\<id>\\_backup\\ ; the newest of each
    becomes the first shared file on the SAP data page (only if none is loaded yet)
"""
import datetime
import filecmp
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
KITCHEN_TOOLS = ("coffee_loss", "breakfast_tracker", "canteen_loss")


def log(msg):
    print("    " + msg, flush=True)


def copy_app(pkg, app_root):
    src, dst = os.path.join(pkg, "app"), os.path.join(app_root, "app")
    if os.path.isdir(dst):
        shutil.rmtree(dst)
    shutil.copytree(src, dst, ignore=shutil.ignore_patterns("__pycache__"))
    log(f"app code -> {dst}")
    cfg_src, cfg_dst = os.path.join(pkg, "config"), os.path.join(app_root, "config")
    if os.path.isdir(cfg_src):
        os.makedirs(cfg_dst, exist_ok=True)
        for name in os.listdir(cfg_src):
            if not os.path.exists(os.path.join(cfg_dst, name)):
                shutil.copy2(os.path.join(cfg_src, name), os.path.join(cfg_dst, name))
        log(f"config -> {cfg_dst} (an existing app_settings.json is kept)")
    for extra in ("tests",):
        if os.path.isdir(os.path.join(pkg, extra)):
            d = os.path.join(app_root, extra)
            if os.path.isdir(d):
                shutil.rmtree(d)
            shutil.copytree(os.path.join(pkg, extra), d, ignore=shutil.ignore_patterns("__pycache__"))


def copy_tools(pkg, data_root):
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    src_root = os.path.join(pkg, "source_tools")
    counts = {"new": 0, "scripts": 0, "kept": 0}
    if not os.path.isdir(src_root):
        log("no source_tools folder in the package: tool folders left as they are")
        return
    for tool in sorted(os.listdir(src_root)):
        src = os.path.join(src_root, tool)
        if not os.path.isdir(src):
            continue
        dst = os.path.join(data_root, "tools", tool)
        for folder, subdirs, files in os.walk(src):
            subdirs[:] = [s for s in subdirs if s not in ("samples", "__pycache__", "_backup")]
            rel = os.path.relpath(folder, src)
            for name in files:
                s = os.path.join(folder, name)
                d = os.path.normpath(os.path.join(dst, rel, name))
                os.makedirs(os.path.dirname(d), exist_ok=True)
                if not os.path.exists(d):
                    shutil.copy2(s, d)
                    counts["new"] += 1
                elif name.lower().endswith(".py"):
                    if not filecmp.cmp(s, d, shallow=False):
                        b = os.path.join(dst, "_backup", stamp, rel)
                        os.makedirs(b, exist_ok=True)
                        shutil.copy2(d, os.path.join(b, name))
                        shutil.copy2(s, d)
                        counts["scripts"] += 1
                        log(f"{tool}\\{name} updated (old one kept in _backup\\{stamp})")
                else:
                    counts["kept"] += 1          # data file already there: never replaced
    # the demo's generated working files and sample inputs (never replace what is already there)
    samples = os.path.join(pkg, "sample_data")
    n_inputs = 0
    if os.path.isdir(samples):
        for section in sorted(os.listdir(samples)):
            sec = os.path.join(samples, section)
            if not os.path.isdir(sec):
                continue
            for tool in sorted(os.listdir(sec)):
                for sub, dest in (("tool_files", os.path.join(data_root, "tools", tool)),
                                  ("inputs", os.path.join(data_root, "Sample inputs", tool))):
                    src = os.path.join(sec, tool, sub)
                    if not os.path.isdir(src):
                        continue
                    for folder, _, files in os.walk(src):
                        rel = os.path.relpath(folder, src)
                        for name in files:
                            d = os.path.normpath(os.path.join(dest, rel, name))
                            if os.path.exists(d):
                                counts["kept"] += 1
                                continue
                            os.makedirs(os.path.dirname(d), exist_ok=True)
                            shutil.copy2(os.path.join(folder, name), d)
                            if sub == "inputs":
                                n_inputs += 1
                            else:
                                counts["new"] += 1
    log(f"tool folders -> {os.path.join(data_root, 'tools')}")
    log(f"{counts['new']} new files copied, {counts['scripts']} scripts updated, {counts['kept']} existing data files left untouched")
    if n_inputs:
        log(f"{n_inputs} sample input files -> {os.path.join(data_root, 'Sample inputs')}")


def migrate_sap(data_root, app_root):
    """Move stray MM60 / MI20 copies out of the tool folders; seed the SAP data page if it is empty."""
    sys.path.insert(0, app_root)
    from app.manifest import Manifest
    from app.sapdata import SapData
    from app.settings import Settings
    tools = os.path.join(data_root, "tools")
    found = {"mm60": [], "mi20_kitchen": []}
    if not os.path.isdir(tools):
        return
    sap = SapData(data_root, Settings(data_root))
    # A tool's own workbooks (item maps, month-end tools, costing files) are never touched, whatever
    # sheets they contain. Only a raw SAP export counts: one sheet, with the SAP header row on it.
    protected = set()
    for t in Manifest(os.path.join(app_root, "app", "tools.json")).tools:
        names = list(t.get("data_files", [])) + [f["file"] for f in t.get("open_files", [])]
        for v in (t.get("variants") or {}).values():
            names += list(v.get("data_files", [])) + [f["file"] for f in v.get("open_files", [])]
        protected.update(n.lower() for n in names)

    def raw_export(path, kind):
        if os.path.basename(path).lower() in protected:
            return False
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True)
            n = len(wb.sheetnames)
            wb.close()
        except Exception:
            return False
        return n == 1 and sap.inspect(kind, path).get("ok", False)
    for tool in sorted(os.listdir(tools)):
        tdir = os.path.join(tools, tool)
        if not os.path.isdir(tdir):
            continue
        for sub in ("", "Mi20_files"):
            folder = os.path.join(tdir, sub)
            if not os.path.isdir(folder):
                continue
            for name in os.listdir(folder):
                p = os.path.join(folder, name)
                if not name.lower().endswith(".xlsx") or name.startswith("~$") or not os.path.isfile(p):
                    continue
                if tool in KITCHEN_TOOLS and raw_export(p, "mi20_kitchen"):
                    found["mi20_kitchen"].append((tool, p))
                elif tool not in KITCHEN_TOOLS and tool != "cost_comparator" and raw_export(p, "mm60"):
                    found["mm60"].append((tool, p))
    user = os.environ.get("USERNAME", "installer")
    for kind, items in found.items():
        if not items:
            continue
        items.sort(key=lambda t: os.path.getmtime(t[1]), reverse=True)
        if sap.path(kind) is None:
            pick = items[0][1]
            ties = [p for _, p in items if abs(os.path.getmtime(p) - os.path.getmtime(pick)) < 60
                    and not filecmp.cmp(p, pick, shallow=False)]
            if ties:
                print(f"\n  Two {kind.upper()} files have the same date. Which one is current?")
                options = [pick] + ties
                for i, p in enumerate(options, 1):
                    print(f"    {i}) {p}")
                try:
                    ans = input("  Type the number (Enter for 1): ").strip().strip('"')
                    pick = options[int(ans) - 1] if ans else pick
                except (ValueError, IndexError, EOFError):
                    pass
            res = sap.replace(kind, pick, user)
            log(f"SAP data: {os.path.basename(pick)} is now the shared {kind.upper()}" if res.get("ok")
                else f"SAP data: {os.path.basename(pick)} was not loaded: {res.get('reason')}")
        for tool, p in items:
            b = os.path.join(tools, tool, "_backup")
            os.makedirs(b, exist_ok=True)
            target = os.path.join(b, os.path.basename(p))
            if os.path.exists(target):
                root, ext = os.path.splitext(os.path.basename(p))
                target = os.path.join(b, f"{root}_{datetime.datetime.now():%Y%m%d_%H%M%S}{ext}")
            shutil.move(p, target)
            log(f"moved {tool}\\{os.path.basename(p)} to {tool}\\_backup (tools now read it from SAP data)")


def main(argv):
    if not argv:
        print(__doc__)
        return 2
    pkg = os.path.abspath(argv[1]) if len(argv) > 1 else os.path.dirname(HERE)
    sys.path.insert(0, pkg)
    from app import paths
    data_root, app_root = paths.default_data_root(), paths.default_app_root()
    if argv[0] == "where":
        print(data_root)
        return 0
    if argv[0] == "copy":
        os.makedirs(data_root, exist_ok=True)
        os.makedirs(os.path.join(app_root, "logs"), exist_ok=True)
        try:
            probe = os.path.join(data_root, ".write_test")
            open(probe, "w").close()
            os.remove(probe)
        except OSError as e:
            print(f"    Cannot write to {data_root}: {e}")
            return 1
        copy_app(pkg, app_root)
        copy_tools(pkg, data_root)
        migrate_sap(data_root, app_root)
        return 0
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
