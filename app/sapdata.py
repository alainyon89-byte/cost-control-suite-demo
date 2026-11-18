# (c) 2026 Alain Yon
"""Shared SAP data (spec 6A): one MM60 and one kitchen MI20, loaded once.

  <data>\\sap_data\\MM60_current.xlsx
  <data>\\sap_data\\MI20_Kitchen_current.xlsx
  <data>\\sap_data\\sap_data.json     facts about each file
  <data>\\sap_data\\history\\           every replaced file; nothing is deleted

The files are taken exactly as exported: no merging, sorting or reformatting.
"""
import datetime
import json
import os
import shutil

from . import branding, detect

H, C = branding.HOTEL_PLANT, branding.CC_PLANT
KINDS = {
    "mm60": {"file": f"MM60_{H}_current.xlsx", "label": f"MM60 {branding.plant_name(H)} ({H})", "header_kind": "mm60",
             "prefix": f"MM60_{H}", "plant": H,
             "needs": "Plant, Material, Material Description, Material Group, Base Unit of Measure, Price"},
    "mm60_cc": {"file": f"MM60_{C}_current.xlsx", "label": f"MM60 {branding.plant_name(C)} ({C})", "header_kind": "mm60",
                "prefix": f"MM60_{C}", "plant": C,
                "needs": "Plant, Material, Material Description, Material Group, Base Unit of Measure, Price"},
    "mi20_kitchen": {"file": "MI20_Kitchen_current.xlsx", "label": "Kitchen MI20 (merged monthly file)",
                     "header_kind": "mi20", "prefix": "MI20_Kitchen",
                     "needs": "PhysInvDoc, Item, Material, Material Description, Plnt, Location, Book Qty, "
                              "Qty Counted, Difference qty, Difference"},
}


PLANTS = {H: "mm60", C: "mm60_cc"}          # an MM60 is identified by its Plant column
PLANT_NAMES = {H: branding.plant_name(H), C: branding.plant_name(C)}


def mm60_kind(plant):
    return PLANTS.get(str(plant or "").strip().upper(), "mm60")


def _now():
    return datetime.datetime.now().isoformat(timespec="seconds")


class SapData:
    def __init__(self, data_root, settings):
        self.dir = os.path.join(data_root, "sap_data")
        self.history_dir = os.path.join(self.dir, "history")
        self.meta_path = os.path.join(self.dir, "sap_data.json")
        self.settings = settings

    # ---- facts -------------------------------------------------------------
    def _meta(self):
        if os.path.exists(self.meta_path):
            try:
                with open(self.meta_path, encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def _save_meta(self, meta):
        os.makedirs(self.dir, exist_ok=True)
        tmp = self.meta_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2)
        os.replace(tmp, self.meta_path)

    def path(self, kind):
        p = os.path.join(self.dir, KINDS[kind]["file"])
        return p if os.path.exists(p) else None

    def info(self, kind):
        p = self.path(kind)
        base = {"kind": kind, "label": KINDS[kind]["label"], "loaded": bool(p), "path": p}
        if not p:
            return base
        m = self._meta().get(kind, {})
        exported = m.get("exported") or datetime.datetime.fromtimestamp(os.path.getmtime(p)).isoformat(timespec="seconds")
        try:
            age = (datetime.datetime.now() - datetime.datetime.fromisoformat(exported)).days
        except ValueError:
            age = None
        limit = int(self.settings.get("sap_age_warning_days") or 7)
        base.update(original_name=m.get("original_name") or os.path.basename(p), exported=exported,
                    exported_source=m.get("exported_source", "file's modified date"),
                    loaded_on=m.get("loaded_on"), loaded_by=m.get("loaded_by"), rows=m.get("rows"),
                    plants=m.get("plants", []), locations=m.get("locations", []),
                    age_days=age, warn_days=limit, old=(age is not None and age > limit))
        return base

    def status(self):
        return {k: self.info(k) for k in KINDS}

    def history(self, kind=None):
        out = []
        if os.path.isdir(self.history_dir):
            for name in sorted(os.listdir(self.history_dir), reverse=True):
                for k, spec in KINDS.items():
                    if name.startswith(spec["prefix"] + "_20") and (kind in (None, k)):
                        p = os.path.join(self.history_dir, name)
                        out.append({"kind": k, "name": name, "path": p,
                                    "modified": datetime.datetime.fromtimestamp(os.path.getmtime(p)).isoformat(timespec="seconds")})
        return out

    def previous(self, kind):
        """Newest replaced file of this kind (used by the MM60 comparator as the 'older' file)."""
        h = self.history(kind)
        if not h:
            return None
        return max(h, key=lambda r: r["modified"])["path"]

    def original_name_of(self, path):
        """Original export name for a file in the store or its history."""
        if path and os.path.dirname(os.path.abspath(path)) == os.path.abspath(self.history_dir):
            parts = os.path.basename(path).split("_", 3)
            for spec in KINDS.values():
                pre = spec["prefix"] + "_"
                if os.path.basename(path).startswith(pre):
                    rest = os.path.basename(path)[len(pre):]          # yyyymmdd_hhmm_<original>
                    bits = rest.split("_", 2)
                    if len(bits) == 3:
                        return bits[2]
            return parts[-1]
        for k in KINDS:
            if path and self.path(k) and os.path.abspath(path) == os.path.abspath(self.path(k)):
                return self._meta().get(k, {}).get("original_name") or os.path.basename(path)
        return os.path.basename(path)

    # ---- checking and replacing ---------------------------------------------
    def inspect(self, kind, path):
        """Check the column headers before anything is replaced. Returns dict(ok, reason, rows, plants, locations)."""
        spec = KINDS[kind]
        if not os.path.isfile(path):
            return {"ok": False, "reason": "That file could not be found."}
        if os.path.splitext(path)[1].lower() != ".xlsx":
            return {"ok": False, "reason": f"{os.path.basename(path)} is not an .xlsx file. Export it from SAP as a spreadsheet."}
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
        except Exception as e:
            return {"ok": False, "reason": f"{os.path.basename(path)} could not be opened ({type(e).__name__}). "
                                           "Is it still open in Excel, or not a real .xlsx file?"}
        try:
            ws = wb.worksheets[0]
            header, hrow, rows = None, 0, 0
            plants, locs = set(), set()
            idx = {}
            for i, r in enumerate(ws.iter_rows(values_only=True)):
                if header is None:
                    if i >= 10:
                        break
                    if detect.classify_headers(r or ()) == spec["header_kind"]:
                        header, hrow = [detect.norm(c) for c in r], i
                        for n, h in enumerate(header):
                            idx.setdefault(h, n)
                    continue
                if not r or all(c in (None, "") for c in r):
                    continue
                rows += 1
                for key, bag in (("plant", plants), ("plnt", plants), ("location", locs)):
                    if key in idx and idx[key] < len(r) and r[idx[key]] not in (None, ""):
                        bag.add(str(r[idx[key]]).strip())
        finally:
            wb.close()
        if header is None:
            found = detect.sniff(path)
            what = found["label"] if found["kind"] != "unknown" else "a file the app does not recognise"
            return {"ok": False, "reason": f"{os.path.basename(path)} does not have the {spec['label']} columns "
                                           f"({spec['needs']}). It looks like {what}. Nothing was replaced."}
        if rows == 0:
            return {"ok": False, "reason": f"{os.path.basename(path)} has the right columns but no rows. Nothing was replaced."}
        if spec.get("plant"):
            # MM60: the Plant column says which one it is, whichever card it was dropped on
            found = sorted(p.upper() for p in plants)
            if len(found) != 1 or found[0] not in PLANTS:
                return {"ok": False, "reason": f"{os.path.basename(path)} is an MM60 for plant "
                                               f"{', '.join(found) or '(no plant found)'}. The app keeps one MM60 for {H} "
                                               f"({PLANT_NAMES[H]}) and one for {C} ({PLANT_NAMES[C]}), each exported on "
                                               "its own. Nothing was replaced."}
            kind = PLANTS[found[0]]
        return {"ok": True, "kind": kind, "label": KINDS[kind]["label"], "rows": rows, "plants": sorted(plants),
                "locations": sorted(locs)}

    def replace(self, kind, path, user):
        """Refuses a file with the wrong headers and leaves the current file in place."""
        check = self.inspect(kind, path)
        if not check["ok"]:
            return check
        kind = check["kind"]
        spec = KINDS[kind]
        os.makedirs(self.history_dir, exist_ok=True)
        target = os.path.join(self.dir, spec["file"])
        meta = self._meta()
        moved = None
        if os.path.exists(target):
            old_name = meta.get(kind, {}).get("original_name") or spec["file"]
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
            moved = os.path.join(self.history_dir, f"{spec['prefix']}_{stamp}_{old_name}")
            n = 2
            while os.path.exists(moved):
                moved = os.path.join(self.history_dir, f"{spec['prefix']}_{stamp}_{n}_{old_name}")
                n += 1
            try:
                shutil.move(target, moved)
            except PermissionError:
                return {"ok": False, "reason": f"Close {spec['file']} in Excel first (someone else may have it open). Nothing was replaced."}
        shutil.copy2(path, target)          # copy2 keeps the export's own modified date
        exported = datetime.datetime.fromtimestamp(os.path.getmtime(path)).isoformat(timespec="seconds")
        meta[kind] = {"original_name": os.path.basename(path), "exported": exported,
                      "exported_source": "file's modified date (SAP exports carry no export date inside)",
                      "loaded_on": _now(), "loaded_by": user, "rows": check["rows"],
                      "plants": check["plants"], "locations": check["locations"]}
        self._save_meta(meta)
        result = {"ok": True, "kind": kind, "label": spec["label"], "moved_to_history": moved, "rows": check["rows"], "plants": check["plants"],
                  "locations": check["locations"], "pr_tool": None}
        if kind == "mm60" and self.settings.get("pr_tool_copy_mm60"):
            result["pr_tool"] = self._copy_to_pr_tool(target)
        return result

    def _copy_to_pr_tool(self, mm60_path):
        folder = (self.settings.get("pr_tool_data_folder") or "").strip().strip('"')
        if not folder or not os.path.isdir(folder):
            return "PR Tool copy skipped: set the PR Tool data folder in Settings first."
        try:
            old_dir = os.path.join(folder, "_old_mm60")
            for name in os.listdir(folder):
                if name.lower().startswith("mm60") and os.path.isfile(os.path.join(folder, name)):
                    os.makedirs(old_dir, exist_ok=True)
                    dest = os.path.join(old_dir, name)
                    if os.path.exists(dest):
                        root, ext = os.path.splitext(name)
                        dest = os.path.join(old_dir, f"{root}_{datetime.datetime.now():%H%M%S}{ext}")
                    shutil.move(os.path.join(folder, name), dest)
            new_name = f"mm60_{datetime.datetime.now():%Y%m%d}.xlsx"
            shutil.copy2(mm60_path, os.path.join(folder, new_name))
            return f"Also copied to the PR Tool as {new_name}."
        except Exception as e:
            return f"PR Tool copy failed: {e}"
