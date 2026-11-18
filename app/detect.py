# (c) 2026 Alain Yon
"""File sniffing: works out what a dropped file looks like (headers, sheet names).

This only LABELS files on screen and suggests tools on Home. The tool scripts
still get every file the same way the .bat gave it, and decide for themselves.
The MM60 / MB52 / MB51 header rules are the same ones oe_tracker.py uses.
"""
import os
import re
import warnings

warnings.filterwarnings("ignore")

KIND_LABELS = {
    "mm60": "MM60 price list",
    "mb52": "MB52 stock on hand",
    "mb51": "MB51 material documents",
    "mi20": "MI20 / stock count variance report",
    "micros_sales_mix": "Micros Sales Mix",
    "micros_report": "Micros report",
    "oe_approval": "OE approval file",
    "oe_workbook": "OE tracking workbook",
    "stock_age_source": "Stock age source workbook",
    "cellar_workbook": "Cellar Consumption workbook",
    "cocktail_tracker": "Cocktail Tracker / costing workbook",
    "sales_item_review": "Completed Sales Item Review",
    "recount_output": "Recount Selection workbook (an output, not a source file)",
    "pdf": "PDF (Micros report?)",
    "unknown": "Not recognised",
}
# which manifest "detect" kinds a sniffed kind satisfies
ALSO = {"mi20": {"mi20", "variance_report", "mi20_kitchen"}, "micros_sales_mix": {"micros_sales_mix", "micros_report"},
        "pdf": {"pdf", "micros_report"}}
SUGGEST = {
    "mm60": ["sap_data"], "mi20": ["sap_data", "recount_selection"],
    "mb52": ["oe_monthly"], "mb51": ["oe_monthly"],
    "micros_sales_mix": ["theo_cos", "food_sales_mix", "recount_selection", "cellar_consumption"],
    "micros_report": ["cellar_consumption"], "pdf": ["cellar_consumption"],
    "oe_approval": ["oe_approvals"], "stock_age_source": ["stock_age"],
    "cellar_workbook": ["cellar_consumption"], "cocktail_tracker": ["recount_selection"],
    "sales_item_review": ["recount_selection"],
}
DATE_SHEET = re.compile(r"^\s*\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\s*$")


def norm(x):
    return re.sub(r"\s+", " ", str(x if x is not None else "")).strip().lower()


def classify_headers(headers):
    h = {norm(x) for x in headers if x not in (None, "")}
    if {"material", "storage location", "movement type"} <= h and ("qty in unit of entry" in h or "quantity" in h):
        return "mb51"
    if {"material", "storage location", "unrestricted"} <= h:
        return "mb52"
    if {"physinvdoc", "material"} <= h and ("difference qty" in h or "difference" in h):
        return "mi20"
    if {"material", "price"} <= h and "storage location" not in h:
        return "mm60"
    return None


def read_head(path, max_rows=12, max_sheets=8):
    """{sheet name: [row tuples]} for the first rows of the first sheets, plus all sheet names."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        names = list(wb.sheetnames)
        heads = {}
        for ws in wb.worksheets[:max_sheets]:
            rows = []
            try:
                for i, r in enumerate(ws.iter_rows(values_only=True)):
                    if i >= max_rows:
                        break
                    rows.append(tuple(r or ()))
            except Exception:
                pass
            heads[ws.title] = rows
        return names, heads
    finally:
        wb.close()


def sniff(path):
    """Returns dict(kind, label, detail). Never raises."""
    name = os.path.basename(path)
    out = {"path": path, "name": name, "kind": "unknown", "label": KIND_LABELS["unknown"], "detail": ""}
    ext = os.path.splitext(name)[1].lower()
    if ext == ".pdf":
        out.update(kind="pdf", label=KIND_LABELS["pdf"])
        return out
    if ext not in (".xlsx", ".xlsm"):
        out["detail"] = "not an Excel workbook"
        return out
    try:
        names, heads = read_head(path)
    except Exception as e:
        out["detail"] = f"could not be read ({type(e).__name__})"
        return out
    low = [n.lower() for n in names]
    kind, detail = None, ""
    flat = [c for rows in heads.values() for r in rows for c in r if c not in (None, "")]
    flat_norm = {norm(c) for c in flat}
    if "approval" in low and "_info" in low:
        kind = "oe_approval"
    elif "confirm material no" in flat_norm:
        kind = "sales_item_review"
    elif "stock on hand" in low:
        kind = "stock_age_source"
    elif any(DATE_SHEET.match(n) for n in names) and "charged on doc" in " ".join(flat_norm):
        kind = "cellar_workbook"
    elif "report" in low and any(norm(c).startswith("sales mix items summar") for c in flat):
        kind = "micros_sales_mix"
        detail = "has an MM60 sheet" if "mm60" in low else "no MM60 sheet"
        for rows in heads.values():
            for r in rows:
                if r and norm(r[0]) == "business dates" and len(r) > 1 and r[1]:
                    detail = f"{r[1]} · {detail}"
    elif "quantity sold" in flat_norm or "qty sold" in flat_norm or (
            "business dates" in flat_norm and "major group" in flat_norm):
        kind = "micros_report"          # the upgraded Micros export: Major / Family Groups as rows
        for rows in heads.values():
            for r in rows:
                if r and norm(r[0]) == "business dates" and len(r) > 1 and r[1]:
                    detail = str(r[1])
    elif sum(1 for n in low if n.startswith("menu ")) >= 3:
        kind = "cocktail_tracker"
    elif any(n.startswith("recount selection") for n in low):
        kind = "recount_output"
    elif "summary" in low and "checks" in low and "locations" in low:
        kind = "oe_workbook"
    else:
        # a raw SAP export: the header row is on the first sheet (a workbook that merely contains
        # an MM60 or MI20 sheet further in is not an export)
        for r in (heads.get(names[0]) or [])[:10]:
            k = classify_headers(r)
            if k:
                kind = k
                break
    if kind:
        out.update(kind=kind, label=KIND_LABELS.get(kind, kind), detail=detail)
    if kind == "micros_sales_mix":
        out["plant"] = _sales_mix_plant(names, heads)
    return out


def _sales_mix_plant(names, heads):
    """D200 for the Conference Centre, D100 for the hotel: from the file's own MM60 sheet, else its Locations line."""
    from .branding import CC_PLANT, HOTEL_PLANT, plant_name
    for n in names:
        if n.lower() == "mm60":
            rows = heads.get(n) or []
            for i, r in enumerate(rows):
                cols = [norm(c) for c in r]
                if "plant" in cols and i + 1 < len(rows):
                    v = rows[i + 1][cols.index("plant")] if cols.index("plant") < len(rows[i + 1]) else None
                    if v:
                        return str(v).strip().upper()
    for rows in heads.values():
        for r in rows:
            if r and norm(r[0]) == "locations" and len(r) > 1 and r[1]:
                cc = plant_name(CC_PLANT).lower()
                return CC_PLANT if ("conference" in str(r[1]).lower() or cc in str(r[1]).lower()) else HOTEL_PLANT
    return HOTEL_PLANT


def matches(sniffed_kind, wanted_kind):
    return wanted_kind in ALSO.get(sniffed_kind, {sniffed_kind})


def suggest_tools(paths):
    """For files dropped on Home: which tools could take them. The app never runs on its own."""
    files, score = [], {}
    for p in paths:
        s = sniff(p)
        files.append(s)
        for tool in SUGGEST.get(s["kind"], []):
            score[tool] = score.get(tool, 0) + 1
    kinds = {f["kind"] for f in files}
    if {"mi20", "micros_sales_mix"} <= kinds:
        score["recount_selection"] = score.get("recount_selection", 0) + 3
    if "cellar_workbook" in kinds:
        score["cellar_consumption"] = score.get("cellar_consumption", 0) + 3
    if "mb52" in kinds:
        score["oe_monthly"] = score.get("oe_monthly", 0) + 3
    ranked = sorted(score, key=lambda t: -score[t])
    return {"files": files, "tools": ranked}
