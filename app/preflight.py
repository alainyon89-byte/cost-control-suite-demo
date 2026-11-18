# (c) 2026 Alain Yon
"""Pre-flight checks shown in step 3 'Check and run' (spec section 7, 6, 10.4).

Each result: {"text", "status": ok | fail | warn | manual, "message"}.
Only "fail" blocks the Run button. An old SAP file is a warning, never a block.
"""
import glob
import os

from . import detect


def is_locked(path):
    """True when Excel (here or on a share) has the file open."""
    if not path or not os.path.exists(path):
        return False
    folder, name = os.path.split(path)
    if os.path.exists(os.path.join(folder, "~$" + name)):
        return True
    if len(name) > 8 and os.path.exists(os.path.join(folder, "~$" + name[2:])):
        return True
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        return False


def lock_message(path):
    return f"Close {os.path.basename(path)} in Excel first (someone else may have it open)"


def run_checks(tool, tool_dir, inputs, options, sap, script_ok=True, kind_of=None):
    out = []

    def add(text, status, message=""):
        out.append({"text": text, "status": status, "message": message})

    if not script_ok:
        add(f"{tool.get('script')} is in the tool folder", "fail",
            "Not installed yet. Add its folder and run the installer again.")
    # inputs
    for spec in tool.get("inputs", []):
        key = spec["key"]
        given = [p for p in inputs.get(key, []) if not str(p).startswith("@")]
        if spec.get("shared"):
            kind = kind_of(spec) if kind_of else ("mm60" if spec["shared"].startswith("mm60") else spec["shared"])
            if given:
                add(f"{spec['label']}: using {os.path.basename(given[0])} for this run only", "ok")
                continue
            if "@in_file" in inputs.get(key, []):
                add(f"{spec['label']}: using the MM60 sheet inside your file", "ok")
                continue
            if spec["shared"] == "mm60_previous":
                prev = sap.previous(kind)
                if prev:
                    add(f"{spec['label']}: {sap.original_name_of(prev)} from the SAP data history", "ok")
                else:
                    add(f"{spec['label']} is chosen", "fail",
                        f"There is no older {sap.info(kind)['label']} in the SAP data history yet. Pick the older file with 'Use a different file'.")
                continue
            info = sap.info(kind)
            if not info["loaded"]:
                if spec.get("deliver", {}).get("mode") == "embedded_sheet":
                    add(f"{info['label']} is not loaded on the SAP data page, so the MM60 sheet inside your file is used", "warn")
                elif spec.get("required", True):
                    add(f"{spec['label']} is loaded", "fail", f"Load the {info['label']} on the SAP data page first.")
                continue
            if info.get("old"):
                add(f"{info['label']} is {info['age_days']} days old", "warn",
                    f"Exported {info['exported'][:10]}. Replace it on the SAP data page if you have a newer export.")
            else:
                add(f"{info['label']} from SAP data, exported {info['exported'][:10]}", "ok")
            continue
        if spec.get("required") and not given:
            add(f"{spec['label']} added", "fail", "Add the file(s) in step 1.")
        for p in given:
            if not os.path.isfile(p):
                add(f"{os.path.basename(p)} can be found", "fail", "The file is no longer where it was. Add it again.")
    # options
    for opt in tool.get("options", []):
        if opt.get("required") and not str(options.get(opt["key"], "")).strip():
            add(f"{opt['label']} filled in", "fail", "Fill it in under step 2.")
    # manifest checks
    for chk in tool.get("preflight", []):
        text, rule = chk["text"], chk.get("check", "manual")
        name, _, arg = rule.partition(":")
        if name == "manual":
            add(text, "manual")
        elif name == "not_locked":
            p = os.path.join(tool_dir, arg)
            add(text, "fail", lock_message(p)) if is_locked(p) else add(text, "ok")
        elif name == "exists":
            p = os.path.join(tool_dir, arg)
            if not os.path.exists(p):
                add(text, "fail", f"{arg} is not in the tool folder ({tool_dir}).")
            elif is_locked(p):
                add(text, "fail", lock_message(p))
            else:
                add(text, "ok")
        elif name == "warn_missing":
            p = os.path.join(tool_dir, arg)
            add(text, "ok") if os.path.exists(p) else add(text, "warn", f"{arg} was not found in the tool folder.")
        elif name == "not_locked_input":
            locked = [p for p in inputs.get(arg, []) if is_locked(p)]
            add(text, "fail", lock_message(locked[0])) if locked else add(text, "ok")
        elif name == "no_sap_exports_in_folder":
            stale = []
            for p in glob.glob(os.path.join(tool_dir, "*.xlsx")):
                base = os.path.basename(p)
                if base.startswith("~$") or base.upper().startswith("OE_"):
                    continue
                if detect.sniff(p)["kind"] in ("mb52", "mm60", "mb51"):
                    stale.append(base)
            if stale:
                add(text, "warn", "Found " + ", ".join(stale) + ". The script uses the newest file of each type and "
                    "notes it on the CHECKS tab. Move old exports out of the tool folder to avoid this.")
            else:
                add(text, "ok")
        else:
            add(text, "manual")
    # a dropped file that is open in Excel: the scripts only read these, so warn rather than block
    for key, paths in inputs.items():
        for p in paths:
            if not str(p).startswith("@") and is_locked(p) and not any(
                    os.path.basename(p) in r["message"] for r in out):
                add(f"{os.path.basename(p)} is closed in Excel", "warn",
                    f"{os.path.basename(p)} looks open in Excel. Save and close it so the script reads the saved version.")
    return out


def blocking(results):
    return [r for r in results if r["status"] == "fail"]
