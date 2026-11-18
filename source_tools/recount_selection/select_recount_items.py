#!/usr/bin/env python3
# ============================================================================
#  (c) 2026 Alain Yon. All rights reserved.
#  Recount Selection Builder  v2
# ============================================================================
"""
Recount Selection Builder  v2 - theoretical consumption model
=============================================================

The rule
--------
    Sales Mix usage  +  Cocktail ingredient usage  =  the variance

Whatever those two do NOT explain is what gets recounted.

For every line on the SAP variance report the script works out theoretical
consumption for the period:

  * DIRECT SALES  - the Micros Sales Mix, converted from selling units into
    SAP stock units. Spirits sold as tots divide by tots-per-bottle; wine sold
    by the glass (GLS lines) divides by glasses-per-bottle; a sales line
    carrying a bottle size ("VSOP Cognac 750ml") is a whole-bottle sale;
    draught converts by keg volume. All of it is configurable - see
    recount_config.csv, written on first run.

  * COCKTAIL INGREDIENTS - each cocktail sold on the Sales Mix is matched to
    its recipe in the Cocktail Tracker, and the recipe quantities are exploded
    onto the SAP material numbers they consume. Matching here is on material
    number, so it is exact.

Unexplained qty  = (book qty - counted qty) - theoretical consumption
Unexplained value = unexplained qty x SAP unit price

A line is EXPLAINED when the unaccounted gap falls inside the tolerance
(default: 10% of theoretical usage, or R250, whichever is larger). Explained
lines drop out of the recount. What is left is ranked by unexplained value and
the top 20 are selected.

The 20-line limit
-----------------
If more than MAX_ITEMS lines remain unexplained, the script writes nothing.
It prints the proposed selection and exits with code 2 so the .bat can ask for
sign-off. Re-run with --approve (and --approved-by "Name") to produce the
sheet.

Usage
-----
    python select_recount_items.py L123.xlsx SalesMix.xlsx Cocktail_Tracker.xlsx
    python select_recount_items.py L123.xlsx SalesMix.xlsx Cocktail_Tracker.xlsx ^
           --approve --approved-by "A Yon"

File order does not matter. Each file is identified by its contents.

Requires: Python 3 + openpyxl   (pip install openpyxl)
"""

import argparse
import csv
import datetime as _dt
import difflib
import os
import re
import sys
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is not installed. Run:  pip install openpyxl")


# ============================================================================
# SETTINGS - the policy dials
# ============================================================================
MAX_ITEMS = 20                  # recount lines allowed without sign-off

TOL_PERCENT = 0.10              # a gap within 10% of theoretical usage ...
TOL_VALUE = 250.00              # ... or under R250, counts as explained

TOT_ML = 25.0                   # spirit tot            -> 30 per 750ml bottle
WINE_GLASS_ML = 187.5           # wine glass            -> 4 per 750ml bottle
SPARKLING_GLASS_ML = 125.0      # sparkling flute       -> 6 per 750ml bottle
FORTIFIED_ML = 50.0             # port / sherry measure -> 15 per 750ml bottle
DRAUGHT_SERVING_LT = 0.5        # draught pint          -> 60 per 30lt keg

COCKTAIL_MATCH_CUTOFF = 0.82    # recipe name similarity needed to link a sale
PORTION_PRICE_RATIO = 0.35      # selling well under the unit cost = a portion, not a unit
GROUP_PORTION_RATIO = 0.45      # ... or well under the dearest line on the same stock item
CONFIG_FILE = "recount_config.csv"      # per family-group conversion rules
OVERRIDE_FILE = "item_overrides.csv"    # per item conversion overrides
ALIAS_FILE = "item_aliases.csv"         # force a sales line onto a material number
SUGGEST_MIN = 0.55                      # floor for "possible match worth a look"
REVIEW_CONFIDENT = 0.90                 # at or above this, a match is not worth reviewing
# ============================================================================


# ----------------------------------------------------------------------------
# Default conversion rules by Micros family group.
#   default_unit  : what a bare sales line means - "portion" or "whole"
#   servings      : servings per full stock unit when the line is a portion
#   by_volume     : Y = work the servings out from the keg size in the SAP
#                   description instead of the servings column
# ----------------------------------------------------------------------------
def default_family_rules():
    tots = round(750.0 / TOT_ML, 4)
    glasses = round(750.0 / WINE_GLASS_ML, 4)
    flutes = round(750.0 / SPARKLING_GLASS_ML, 4)
    fort = round(750.0 / FORTIFIED_ML, 4)
    return [
        # family fragment,     default_unit, servings, by_volume, note
        ("whiskey",            "portion", tots,    "N", "bare line = tot"),
        ("whisky",             "portion", tots,    "N", "bare line = tot"),
        ("cognac",             "portion", tots,    "N", "bare line = tot"),
        ("brandy",             "portion", tots,    "N", "bare line = tot"),
        ("white spirits",      "portion", tots,    "N", "bare line = tot"),
        ("liqueur",            "portion", tots,    "N", "bare line = tot"),
        ("rum",                "portion", tots,    "N", "bare line = tot"),
        ("shooters",           "portion", tots,    "N", "bare line = tot"),
        ("aperitif",           "portion", tots,    "N", "bare line = tot"),
        ("port",               "portion", fort,    "N", "50ml measure"),
        ("sherry",             "portion", fort,    "N", "50ml measure"),
        ("red wine",           "whole",   glasses, "N", "bare line = bottle, GLS = glass"),
        ("white wine",         "whole",   glasses, "N", "bare line = bottle, GLS = glass"),
        ("rose wine",          "whole",   glasses, "N", "bare line = bottle, GLS = glass"),
        ("sparkling wine",     "whole",   flutes,  "N", "bare line = bottle, GLS = flute"),
        ("beer",               "whole",   0,       "Y", "kegs convert by volume"),
        ("cooler",             "whole",   1,       "N", "each"),
        ("minerals cordial",   "portion", tots,    "N", "cordial shot from the bottle"),
        ("minerals",           "whole",   1,       "N", "each"),
        ("alcohol coffee",     "portion", tots,    "N", "spirit measure, no recipe held"),
        ("cocktails",          "recipe",  0,       "N", "exploded via Cocktail Tracker"),
    ]


def load_family_rules(folder):
    """Read recount_config.csv beside the script, creating it on first run."""
    path = os.path.join(folder, CONFIG_FILE)
    if not os.path.exists(path):
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["family_contains", "default_unit", "servings_per_stock_unit",
                        "by_keg_volume", "note"])
            for row in default_family_rules():
                w.writerow(row)
        print(f"  Created {CONFIG_FILE} with default conversions - edit it to change them.")

    rules = []
    with open(path, newline="", encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            frag = (row.get("family_contains") or "").strip().lower()
            if not frag:
                continue
            try:
                servings = float(row.get("servings_per_stock_unit") or 0)
            except ValueError:
                servings = 0.0
            rules.append({
                "frag": frag,
                "default": (row.get("default_unit") or "whole").strip().lower(),
                "servings": servings,
                "by_volume": (row.get("by_keg_volume") or "N").strip().upper().startswith("Y"),
            })
    return rules, path


def load_overrides(folder):
    """Optional per-item overrides: sales_item_contains,servings_per_stock_unit"""
    path = os.path.join(folder, OVERRIDE_FILE)
    if not os.path.exists(path):
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["sales_item_contains", "servings_per_stock_unit", "note"])
            w.writerow(["# example: Irish Whiskey,30,tot from a 750ml bottle", "", ""])
        return []
    out = []
    with open(path, newline="", encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            frag = (row.get("sales_item_contains") or "").strip()
            if not frag or frag.startswith("#"):
                continue
            try:
                out.append((frag.lower(), float(row.get("servings_per_stock_unit") or 1)))
            except ValueError:
                continue
    return out


def load_aliases(folder):
    """Force a Sales Mix line onto a material number the matcher cannot reach
    on names alone. One line per alias:  sales_item_contains,material_no"""
    path = os.path.join(folder, ALIAS_FILE)
    if not os.path.exists(path):
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["sales_item_contains", "material_no", "note"])
            w.writerow(["# Estate Red Blend", "90000123", "example - delete or edit"])
            w.writerow(["# Shooter Special", "SKIP", "example - SKIP means no stock line, stop asking"])
        return [], []
    out, skips = [], []
    with open(path, newline="", encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            frag = (row.get("sales_item_contains") or "").strip()
            mat = (row.get("material_no") or "").strip()
            if not frag or not mat or frag.startswith("#"):
                continue
            if mat.upper() == "SKIP":
                skips.append(frag.lower())     # decided: no stock line, stop asking
                continue
            out.append((frag.lower(), mat))
    return out, skips


# ============================================================================
# Header detection
# ============================================================================

VARIANCE_COLUMNS = {
    "doc":      ["physinvdoc", "phys inv doc", "physical inventory doc", "doc"],
    "line":     ["item", "line", "line item", "pos"],
    "material": ["material", "material no", "material number", "matnr"],
    "desc":     ["material description", "description", "item description"],
    "plant":    ["plnt", "plant", "werks"],
    "location": ["location", "storage location", "sloc", "lgort"],
    "book":     ["book qty", "book quantity", "book"],
    "counted":  ["qty counted", "quantity counted", "counted qty", "count qty"],
    "diffqty":  ["difference qty", "diff qty", "difference quantity"],
    "diffval":  ["difference", "difference value", "diff value"],
    "posval":   ["pos. diff. value", "pos diff value", "positive difference value"],
    "negval":   ["neg. diff. value", "neg diff value", "negative difference value"],
}

SALES_COLUMNS = {
    "item":     ["item", "menu item", "item name"],
    "qty":      ["qty sold", "quantity sold", "qty"],
    "family":   ["family group"],
    "avgprice": ["average price", "avg price"],
}

MASTER_COLUMNS = {
    "material": ["material", "material no", "material number", "matnr"],
    "desc":     ["material description", "description"],
    "uom":      ["base unit of measure", "base uom", "uom", "unit of measure"],
    "price":    ["price", "standard price", "moving price", "value"],
    "plant":    ["plant", "plnt", "werks"],
}


def _nh(v):
    return re.sub(r"\s+", " ", str(v or "").strip().lower())


def find_header_row(ws, required, spec, max_scan=25):
    last_row = ws.max_row or max_scan
    last_col = ws.max_column or 40
    for r in range(1, min(last_row, max_scan) + 1):
        headers = {}
        for c in range(1, last_col + 1):
            h = _nh(ws.cell(r, c).value)
            if not h:
                continue
            for key, aliases in spec.items():
                if key not in headers and h in aliases:
                    headers[key] = c
                    break
        if all(k in headers for k in required):
            return r, headers
    return None, None


# Sheet names that only ever appear on a workbook this script itself wrote -
# never on a source file. A completed review sheet is its own kind: it is not
# an error to hand one back, it is the normal way to hand one back, so it is
# never lumped in with a stale "old output, ignore it" file.
REVIEW_SHEETS = {"sales item review"}
PRIOR_OUTPUT_SHEETS = {"recount selection", "full analysis", "sold no variance",
                       "data quality", "sap stock lines"}


def classify_file(path):
    """Return 'variance', 'sales', 'cocktail', 'review', 'output' or None.

    Detection is entirely by sheet contents, never by filename - a review
    sheet is recognised however it has been renamed or saved.
    """
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return None
    try:
        names = {n.strip().lower() for n in wb.sheetnames}
        if names & REVIEW_SHEETS:
            return "review"
        if names & PRIOR_OUTPUT_SHEETS:
            return "output"
        if any(n.startswith("menu ") for n in names):
            return "cocktail"
        for ws in wb.worksheets:
            if find_header_row(ws, ["material", "desc", "diffqty"], VARIANCE_COLUMNS)[0]:
                return "variance"
        for ws in wb.worksheets:
            if find_header_row(ws, ["item", "qty", "family"], SALES_COLUMNS)[0]:
                return "sales"
        for ws in wb.worksheets:
            if find_header_row(ws, ["material", "price", "uom"], MASTER_COLUMNS)[0]:
                return "cocktail"
        return None
    finally:
        wb.close()


# ============================================================================
# Name handling
# ============================================================================

SIZE_TOKEN = re.compile(r"\b\d+(?:[.,]\d+)?\s*(?:ml|mls|lt|ltr|l|cl|kg|kgs|g|gr|gm)\b", re.I)
VOLUME_TOKEN = re.compile(r"(\d+(?:[.,]\d+)?)\s*(ml|lt|ltr|l)\b", re.I)
PACK_NOISE = re.compile(r"\b(nrb|can|bt|ea|bottle|btl|keg|draught|draft|pack)\b", re.I)
# "GLS" only. The word "glass" is excluded on purpose: it appears in product
# names such as "Crystal Glass Sparkling" and "Estate Glass Collection".
GLASS_MARK = re.compile(r"(?:\bgls\b|\bgl\b|\sg$|\sgl$)", re.I)


CODE_TOKEN = re.compile(r"^[a-z]{2,4}\d{2,4}$")     # SAP suffixes: ABC106, XYZ100, AB7


def norm_name(s):
    s = str(s or "").lower().replace("&", " and ")
    s = re.sub(r"[^a-z0-9. ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def strip_size(s):
    s = SIZE_TOKEN.sub(" ", s)
    s = re.sub(r"\b(19|20)\d{2}\b", " ", s)       # vintages
    return re.sub(r"\s+", " ", s).strip()


def squash(s):
    """Collapse to letters and digits only, so 'Red Blend' == 'RedBlend'."""
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def tokens(s):
    """Meaningful words, product codes dropped."""
    return [t for t in strip_size(norm_name(s)).split()
            if len(t) > 1 and not CODE_TOKEN.match(t)]


def token_score(a_toks, b_toks):
    """Share of a_toks found in b_toks. A short token matches a longer one it
    starts ('sauv' -> 'sauvignon', 'bl' -> 'blanc'), which is how Micros
    abbreviates wine names."""
    if not a_toks or not b_toks:
        return 0.0
    pool, hit = list(b_toks), 0
    for t in a_toks:
        for i, u in enumerate(pool):
            same = (t == u
                    or (len(t) >= 2 and u.startswith(t))
                    or (len(u) >= 2 and t.startswith(u))
                    # tolerate a typo: Chardonay / Chardonnay, Cero / Zero
                    or (min(len(t), len(u)) >= 5
                        and difflib.SequenceMatcher(None, t, u).ratio() >= 0.85))
            if same:
                hit += 1
                pool.pop(i)
                break
    return hit / len(a_toks)


def has_bottle_size(name):
    """A sales line naming a bottle size is a whole-unit sale."""
    return bool(SIZE_TOKEN.search(str(name or "")))


def is_glass_line(name):
    """Micros truncates at ~32 chars, so GLS can arrive as 'GL' or a bare 'G'."""
    return bool(GLASS_MARK.search(str(name or "").strip()))


def keg_litres(desc):
    """Total litres in a keg from its SAP description, e.g. 'Keg 30lt' -> 30."""
    best = 0.0
    for qty, unit in VOLUME_TOKEN.findall(str(desc or "")):
        v = float(qty.replace(",", "."))
        lt = v / 1000.0 if unit.lower() == "ml" else v
        best = max(best, lt)
    return best if best >= 5 else 0.0     # only keg-sized volumes


# ============================================================================
# Readers
# ============================================================================

def _num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").replace(" ", ""))
    except ValueError:
        return 0.0


def _txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _line_no(v):
    t = _txt(v)
    return t.zfill(3) if t.isdigit() and len(t) < 3 else t


def read_variance(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        hrow, cols = find_header_row(ws, ["material", "desc", "diffqty"], VARIANCE_COLUMNS)
        if not hrow:
            continue

        def v(r, k):
            c = cols.get(k)
            return ws.cell(r, c).value if c else None

        rows = []
        for r in range(hrow + 1, ws.max_row + 1):
            mat, desc = v(r, "material"), v(r, "desc")
            if mat in (None, "") and desc in (None, ""):
                continue
            if str(desc or "").strip().lower().startswith("total"):
                continue
            dq = _num(v(r, "diffqty"))
            dv = _num(v(r, "diffval"))
            val = abs(dv) if dv else abs(_num(v(r, "posval"))) + abs(_num(v(r, "negval")))
            rows.append({
                "doc": _txt(v(r, "doc")), "line": _line_no(v(r, "line")),
                "material": _txt(mat), "desc": str(desc or "").strip(),
                "plant": _txt(v(r, "plant")), "location": _txt(v(r, "location")),
                "book": _num(v(r, "book")), "counted": _num(v(r, "counted")),
                "diffqty": dq, "varvalue": val,
                "unitprice": (val / abs(dq)) if dq else 0.0,
            })
        if rows:
            return rows
    raise SystemExit(
        f"Could not find a variance table in {os.path.basename(path)}. "
        "Expected Material, Material Description and Difference qty columns."
    )


def read_sales(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        hrow, cols = find_header_row(ws, ["item", "qty"], SALES_COLUMNS)
        if not hrow:
            continue
        out = []
        for r in range(hrow + 1, ws.max_row + 1):
            name = ws.cell(r, cols["item"]).value
            if name in (None, "") or str(name).strip().lower().startswith("total"):
                continue
            out.append({
                "name": str(name).strip(),
                "family": _txt(ws.cell(r, cols["family"]).value) if "family" in cols else "",
                "qty": _num(ws.cell(r, cols["qty"]).value),
                "avgprice": _num(ws.cell(r, cols["avgprice"]).value) if "avgprice" in cols else 0.0,
            })
        if out:
            return out
    return []


def read_cocktail_tracker(path):
    """Return (recipes, master) where
         recipes = [{'name':..., 'sheet':..., 'ingredients':[(material, desc, qty, uom)]}]
         master  = {material: {'desc','uom','price'}}
    """
    wb = openpyxl.load_workbook(path, data_only=True)

    # --- SAP master list (material / description / UOM / price) --------------
    master = {}
    for ws in wb.worksheets:
        hrow, cols = find_header_row(ws, ["material", "price", "uom"], MASTER_COLUMNS)
        if not hrow:
            continue
        for r in range(hrow + 1, ws.max_row + 1):
            mat = _txt(ws.cell(r, cols["material"]).value)
            if not mat:
                continue
            master[mat] = {
                "desc": _txt(ws.cell(r, cols["desc"]).value) if "desc" in cols else "",
                "uom": _txt(ws.cell(r, cols["uom"]).value),
                "price": _num(ws.cell(r, cols["price"]).value),
            }
        if master:
            break

    # --- menu sheet -> recipe name -----------------------------------------
    menu_names = {}
    for ws in wb.worksheets:
        head = " ".join(_nh(ws.cell(1, c).value) for c in range(1, 6))
        if "summary" not in head and "summary" not in _nh(ws.title):
            continue
        for r in range(1, ws.max_row + 1):
            for c in range(1, (ws.max_column or 12)):
                a = _txt(ws.cell(r, c).value)
                b = _txt(ws.cell(r, c + 1).value)
                if re.fullmatch(r"(?i)menu\s*\d+", a or "") and b:
                    menu_names[a.strip().lower()] = b.strip()
        if menu_names:
            break

    # --- recipes ------------------------------------------------------------
    recipes = []
    for name in wb.sheetnames:
        if not name.strip().lower().startswith("menu "):
            continue
        ws = wb[name]
        hrow, cols = None, {}
        for r in range(1, min(ws.max_row, 12) + 1):
            labels = {_nh(ws.cell(r, c).value): c for c in range(1, (ws.max_column or 10) + 1)}
            mat_c = next((c for h, c in labels.items() if h.startswith("sap")), None)
            qty_c = next((c for h, c in labels.items() if h.startswith("qty")), None)
            if mat_c and qty_c:
                hrow = r
                cols = {"mat": mat_c, "qty": qty_c,
                        "desc": next((c for h, c in labels.items() if "description" in h), mat_c + 1),
                        "uom": next((c for h, c in labels.items() if h in ("uom", "unit")), qty_c + 1)}
                break
        if not hrow:
            continue
        ings = []
        for r in range(hrow + 1, ws.max_row + 1):
            mat = _txt(ws.cell(r, cols["mat"]).value)
            qty = _num(ws.cell(r, cols["qty"]).value)
            if not mat or qty == 0:
                continue
            ings.append((mat, _txt(ws.cell(r, cols["desc"]).value), qty,
                         _txt(ws.cell(r, cols["uom"]).value)))
        if ings:
            # recipe name from the summary tab, else from the sheet itself
            rn = menu_names.get(name.strip().lower()) or _txt(ws.cell(2, 3).value) or name
            recipes.append({"name": rn, "sheet": name, "ingredients": ings})
    return recipes, master


# ============================================================================
# Matching
# ============================================================================

BLOCKING = ("virgin", "mocktail", "non alcoholic", "alcohol free", "frozen")


def cocktail_key(s):
    s = norm_name(s)
    s = re.sub(r"\b(cocktail|cocktails|blh|mgh)\b", " ", s)
    return squash(s)


def blocking_tags(s):
    n = norm_name(s)
    return {t for t in BLOCKING if t in n}


def match_cocktails(sales, recipes):
    """Link cocktail sales lines to recipes. Returns (links, unmatched_sales,
    unused_recipes)."""
    rkeys = {}
    for rec in recipes:
        rkeys.setdefault(cocktail_key(rec["name"]), rec)
    links, unmatched, used = [], [], set()

    for s in sales:
        if "cocktail" not in s["family"].lower():
            continue
        if s["qty"] <= 0:
            continue
        key = cocktail_key(s["name"])
        tags = blocking_tags(s["name"])
        rec = None

        how = ""
        cand = rkeys.get(key)
        if cand and blocking_tags(cand["name"]) == tags:
            rec, how = cand, "exact"
        if rec is None:
            pool = [k for k, r in rkeys.items() if blocking_tags(r["name"]) == tags]
            starts = [k for k in pool if k.startswith(key) or key.startswith(k)]
            if len(starts) == 1:
                rec, how = rkeys[starts[0]], "prefix"
            else:
                near = difflib.get_close_matches(key, pool, n=1, cutoff=COCKTAIL_MATCH_CUTOFF)
                if near:
                    ratio = difflib.SequenceMatcher(None, key, near[0]).ratio()
                    rec, how = rkeys[near[0]], f"fuzzy {ratio:.0%}"
        if rec is None:
            unmatched.append(s)
            continue
        used.add(rec["sheet"])
        links.append({"sale": s, "recipe": rec, "how": how})

    unused = [r for r in recipes if r["sheet"] not in used]
    return links, unmatched, unused


TOKEN_MIN = 0.65      # share of the sales line's words that must be accounted for
SEQ_MIN = 0.55        # character similarity floor, size tokens included
COMBINED_MIN = 0.72   # overall score needed to accept a match
DECISIVE = 0.06       # winning margin that makes a match decisive
CLAIM_BONUS = 0.05    # nudge towards a stock line nothing else has claimed


def build_variance_index(vrows):
    for row in vrows:
        row["_sq"] = squash(strip_size(norm_name(row["desc"])))
        row["_sqsize"] = squash(norm_name(row["desc"]))
        row["_tok"] = tokens(row["desc"])
    return vrows


def _candidates(name, vrows):
    """Score every stock line against one sales line, best first."""
    clean = GLASS_MARK.sub(" ", norm_name(name)).strip()
    sq, sqs = squash(strip_size(clean)), squash(clean)
    if len(sq) < 4:
        return [], sq

    exact = [r for r in vrows if r["_sqsize"] == sqs]
    if len(exact) == 1:
        return [(1.0, exact[0], "exact")], sq

    stoks = tokens(clean)
    out = []
    for r in vrows:
        t1 = token_score(stoks, r["_tok"])
        if t1 < TOKEN_MIN:
            continue
        seq = difflib.SequenceMatcher(None, sqs, r["_sqsize"]).ratio()
        if seq < SEQ_MIN:
            continue
        t2 = token_score(r["_tok"], stoks)
        combined = 0.55 * t1 + 0.20 * t2 + 0.25 * seq
        if combined >= COMBINED_MIN:
            out.append((combined, r, f"matched {combined:.0%}"))
    out.sort(key=lambda x: -x[0])
    return out, sq


def suggest(name, vrows):
    """Best sub-threshold candidate, for the review list."""
    clean = GLASS_MARK.sub(" ", norm_name(name)).strip()
    sqs, stoks = squash(clean), tokens(clean)
    best = None
    for r in vrows:
        t1 = token_score(stoks, r["_tok"])
        seq = difflib.SequenceMatcher(None, sqs, r["_sqsize"]).ratio()
        score = 0.6 * t1 + 0.4 * seq
        if score >= SUGGEST_MIN and (best is None or score > best[0]):
            best = (score, r)
    return best


def assign_sales(sales_lines, vrows, aliases=()):
    """Resolve every sales line to a stock line.

    Round 1 takes the decisive matches and marks those stock lines claimed.
    Round 2 settles what is left, leaning towards a stock line nothing has
    claimed yet - which is what separates "Brut Cap Classique Nectar" from
    "Brut Cap Classique Rose" when the names alone cannot.
    """
    # Micros truncates at ~32 characters, so a bottle line and a glass line of
    # the same product can arrive under one identical name. They must resolve
    # to the same stock line, so score each distinct name once.
    by_material = {r["material"]: r for r in vrows}
    groups = {}
    for s in sales_lines:
        groups.setdefault(squash(GLASS_MARK.sub(" ", norm_name(s["name"]))), []).append(s)

    assigned, claimed, leftovers, reps = [], set(), [], []
    for g in groups.values():
        forced = next((by_material[m] for frag, m in aliases
                       if frag in g[0]["name"].lower() and m in by_material), None)
        if forced is not None:
            for member in g:
                assigned.append((member, forced, "alias", ""))
            claimed.add(id(forced))
            continue
        reps.append((_candidates(g[0]["name"], vrows)[0], g))

    for cands, s in reps:
        if not cands:
            leftovers.append((cands, s, "no match"))
            continue
        if len(cands) == 1 or cands[0][0] - cands[1][0] >= DECISIVE:
            for member in s:
                assigned.append((member, cands[0][1], cands[0][2], ""))
            claimed.add(id(cands[0][1]))
        else:
            leftovers.append((cands, s, "close call"))

    for cands, s, why in leftovers:
        if not cands:
            for member in s:
                assigned.append((member, None, "no match", ""))
            continue
        ranked = sorted(cands, key=lambda c: -(c[0] + (0 if id(c[1]) in claimed else CLAIM_BONUS)))
        best, runner = ranked[0], ranked[1] if len(ranked) > 1 else None
        note = f"close call vs {runner[1]['desc']}" if runner else why
        for member in s:
            assigned.append((member, best[1], best[2], note))
        claimed.add(id(best[1]))
    return assigned


# ============================================================================
# Conversion: Micros selling units -> SAP stock units
# ============================================================================

def servings_per_unit(sale, vrow, rules, overrides, group_top=0.0):
    """How many sales servings make up one SAP stock unit, plus a label.

    group_top is the dearest selling price among the sales lines pointing at
    this same stock item, which is how a bottle line is told apart from a glass
    line when Micros has truncated the "GLS" tag off the name.
    """
    lname = sale["name"].lower()
    for frag, servings in overrides:
        if frag in lname:
            return max(servings, 1e-9), f"override /{servings:g}"

    fam = sale["family"].lower()
    rule = next((r for r in rules if r["frag"] and r["frag"] in fam), None)
    wine_glasses = round(750.0 / WINE_GLASS_ML, 4)

    if rule and rule["by_volume"]:
        lt = keg_litres(vrow["desc"] if vrow else "")
        if lt:
            n = lt / DRAUGHT_SERVING_LT
            return n, f"keg {lt:g}lt /{n:g}"
        return 1.0, "each"

    n = rule["servings"] if rule and rule["servings"] > 1 else wine_glasses

    if is_glass_line(sale["name"]):
        return n, f"glass /{n:g}"

    # Price is checked before the name. Micros item names are unreliable about
    # units: "Don Julio Bianco 750ml" sells at R92 against a R657 bottle, so it
    # is a tot with the bottle size left in the name.
    price = sale["avgprice"]
    if price > 0 and n > 1:
        if group_top > 0 and price / group_top < GROUP_PORTION_RATIO:
            return n, f"portion by price /{n:g}"
        cost = (vrow or {}).get("unitprice") or 0.0
        if cost > 0 and price / cost < PORTION_PRICE_RATIO:
            return n, f"portion by price /{n:g}"

    if has_bottle_size(sale["name"]):
        return 1.0, "whole unit"
    if rule and rule["default"] == "portion" and rule["servings"] > 1:
        return rule["servings"], f"portion /{rule['servings']:g}"
    if rule is None:
        return 1.0, "each (no rule)"
    return 1.0, "whole unit"


# ============================================================================
# The model
# ============================================================================

def build_model(vrows, sales, recipes, rules, overrides, aliases=()):
    by_material = {r["material"]: r for r in vrows}
    build_variance_index(vrows)

    for r in vrows:
        r["sales_usage"] = 0.0
        r["cocktail_usage"] = 0.0
        r["sales_lines"] = []
        r["cocktail_lines"] = []

    # --- cocktails first: exact, by material number -------------------------
    links, ck_unmatched, ck_unused = match_cocktails(sales, recipes)
    off_report = {}
    for link in links:
        qty_sold = link["sale"]["qty"]
        for mat, desc, per, uom in link["recipe"]["ingredients"]:
            used = qty_sold * per
            row = by_material.get(mat)
            if row is None:
                d = off_report.setdefault(mat, {"desc": desc, "qty": 0.0, "uom": uom})
                d["qty"] += used
                continue
            row["cocktail_usage"] += used
            row["cocktail_lines"].append((link["sale"]["name"], qty_sold, per, used))

    # --- direct sales -------------------------------------------------------
    direct = [s for s in sales
              if "cocktail" not in s["family"].lower() and s["qty"] > 0]
    unmatched_sales, fuzzy_matches, review = [], [], []
    resolved = assign_sales(direct, vrows, aliases)
    group_top = {}
    for s, row, how, note in resolved:
        if row is not None:
            group_top[id(row)] = max(group_top.get(id(row), 0.0), s["avgprice"])
    for s, row, how, note in resolved:
        if row is None:
            s["why"] = how
            hint = suggest(s["name"], vrows)
            s["hint"] = (f"{hint[1]['desc']}  [material {hint[1]['material']}, "
                         f"{hint[0]:.0%} alike]") if hint else ""
            unmatched_sales.append(s)
            review.append({"sale": s, "row": None, "how": how, "note": "",
                           "conv": "", "hint": s["hint"],
                           "hint_mat": hint[1]["material"] if hint else "",
                           "hint_desc": hint[1]["desc"] if hint else ""})
            continue
        n, label = servings_per_unit(s, row, rules, overrides, group_top.get(id(row), 0.0))
        used = s["qty"] / n if n else s["qty"]
        row["sales_usage"] += used
        row["sales_lines"].append((s["name"], s["qty"], label, used))
        if how not in ("exact", "alias"):
            fuzzy_matches.append((s["name"], row["desc"], how, s["qty"], label, note))
            m = re.search(r"(\d+)%", how)
            score = int(m.group(1)) / 100.0 if m else 0.0
            if score < REVIEW_CONFIDENT or note:
                review.append({"sale": s, "row": row, "how": how, "note": note,
                               "conv": label, "hint": "", "score": score})

    # --- reconcile ----------------------------------------------------------
    for r in vrows:
        r["theoretical"] = r["sales_usage"] + r["cocktail_usage"]
        r["shortage"] = -r["diffqty"]                     # positive = short
        r["unexplained_qty"] = r["shortage"] - r["theoretical"]
        price = r["unitprice"] or 0.0
        r["unexplained_value"] = abs(r["unexplained_qty"]) * price
        r["theoretical_value"] = r["theoretical"] * price
        r["tolerance"] = max(TOL_VALUE, TOL_PERCENT * abs(r["theoretical_value"]))
        r["explained"] = r["unexplained_value"] <= r["tolerance"]

        if r["diffqty"] > 0:
            r["direction"] = "Counted more than book"
        elif r["unexplained_qty"] > 0:
            r["direction"] = "Short beyond sales"
        elif r["unexplained_qty"] < 0:
            r["direction"] = "Sales exceed the shortage"
        else:
            r["direction"] = ""

        if r["diffqty"] == 0 and r["theoretical"] <= 0:
            r["status"] = "No movement"
        elif r["diffqty"] == 0:
            r["status"] = "Sold, no variance"
        elif r["explained"]:
            r["status"] = "Explained by sales"
        else:
            r["status"] = "Unexplained"

    return {
        "cocktail_links": links,
        "cocktail_unmatched": ck_unmatched,
        "cocktail_unused": ck_unused,
        "unmatched_sales": unmatched_sales,
        "fuzzy_matches": fuzzy_matches,
        "review": review,
        "off_report_ingredients": off_report,
    }


def apply_master_prices(vrows, master):
    """SAP master price beats a price derived from the variance line, and is the
    only price available where the difference is zero."""
    for r in vrows:
        m = master.get(r["material"])
        if m and m["price"]:
            r["unitprice"] = m["price"]
            r["price_src"] = "SAP master"
        elif r["unitprice"]:
            r["price_src"] = "derived"
        else:
            r["price_src"] = "none"


def select(vrows, limit):
    pool = [r for r in vrows
            if r["status"] == "Unexplained" and abs(r["diffqty"]) > 0]
    pool.sort(key=lambda r: (-r["unexplained_value"], r["doc"], r["line"]))
    for i, r in enumerate(pool, 1):
        r["rank"] = i
    return pool, pool[:limit]


# ============================================================================
# Workbook
# ============================================================================

HEAD_FILL = PatternFill("solid", fgColor="1F3864")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
TITLE_FONT = Font(bold=True, size=14, color="1F3864")
HEAD_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MAIN_COLS = [
    ("Stock Sheet No", 15, "doc", "text"),
    ("Line No", 9, "line", "text"),
    ("Material No", 13, "material", "text"),
    ("Material Description", 42, "desc", "str"),
    ("Book Qty", 10, "book", "qty"),
    ("Counted Qty", 11, "counted", "qty"),
    ("Variance Qty", 12, "diffqty", "qty"),
    ("Sales Usage", 11, "sales_usage", "qty"),
    ("Cocktail Usage", 13, "cocktail_usage", "qty"),
    ("Theoretical Usage", 15, "theoretical", "qty"),
    ("Unexplained Qty", 14, "unexplained_qty", "qty"),
    ("Unexplained Value", 16, "unexplained_value", "rand"),
    ("Direction", 22, "direction", "str"),
    ("Rank", 7, "rank", "int"),
    ("Recount Qty", 12, "__input", ""),
    ("Comments", 28, "__input", ""),
]


def _put(ws, r, c, value, kind):
    cell = ws.cell(r, c)
    cell.border = BORDER
    if kind == "":
        cell.fill = INPUT_FILL
        return cell
    cell.value = value
    if kind == "qty":
        cell.number_format = "#,##0.0##"
    elif kind == "rand":
        cell.number_format = "#,##0.00"
    elif kind in ("text", "int"):
        cell.alignment = Alignment(horizontal="center")
    return cell


def write_sheet(selected, pool, vrows, diag, out_path, meta):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Recount Selection"

    ws["A1"] = "STOCK RECOUNT SELECTION"
    ws["A1"].font = TITLE_FONT
    info = [
        ("Outlet / Location", meta["location"]),
        ("Plant", meta["plant"]),
        ("Stock sheet numbers", meta["docs"]),
        ("Source variance report", meta["source"]),
        ("Sales Mix", meta["sales_src"]),
        ("Cocktail Tracker", meta["cocktail_src"]),
        ("Generated", meta["generated"]),
        ("Selection rule", "Sales Mix usage + cocktail ingredient usage should equal the "
                           "variance. Lines they do not explain, ranked by unexplained value."),
        ("Tolerance", f"{TOL_PERCENT:.0%} of theoretical usage or R{TOL_VALUE:,.0f}, "
                      "whichever is larger"),
        ("Conversions", f"tot {TOT_ML:g}ml, wine glass {WINE_GLASS_ML:g}ml, "
                        f"draught {DRAUGHT_SERVING_LT:g}lt"),
        ("Lines with a variance", f"{meta['with_variance']} of {meta['total_lines']}"),
        ("Explained by sales", f"{meta['explained']}"),
        ("Left unexplained", f"{meta['unexplained']}"),
        ("Value selected / unexplained",
         f"R {sum(r['unexplained_value'] for r in selected):,.2f} of "
         f"R {sum(r['unexplained_value'] for r in pool):,.2f}"),
        ("Approved by", meta["approved_by"] or "________________________"),
    ]
    r = 2
    for label, value in info:
        ws.cell(r, 1, label).font = Font(bold=True)
        ws.cell(r, 2, value)
        r += 1
    ws.column_dimensions["A"].width = 28

    hrow = r + 1
    for i, (label, width, _, _) in enumerate(MAIN_COLS, 1):
        c = ws.cell(hrow, i, label)
        c.fill, c.font = HEAD_FILL, HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.cell(hrow, 1).alignment = Alignment(horizontal="center", wrap_text=True)

    for i, item in enumerate(sorted(selected, key=lambda x: (x["doc"], x["line"]))):
        rr = hrow + 1 + i
        for ci, (_, _, key, kind) in enumerate(MAIN_COLS, 1):
            _put(ws, rr, ci, item.get(key), kind)

    ws.freeze_panes = ws.cell(hrow + 1, 1)
    last = get_column_letter(len(MAIN_COLS))
    ws.auto_filter.ref = f"A{hrow}:{last}{hrow + len(selected)}"
    ws.print_title_rows = f"{hrow}:{hrow}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    _write_analysis(wb, vrows)
    _write_no_variance(wb, vrows)
    _write_quality(wb, diag)
    wb.save(out_path)
    return out_path


def _sheet_head(ws, title, heads, widths, row=3):
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    for i, (h, w) in enumerate(zip(heads, widths), 1):
        c = ws.cell(row, i, h)
        c.fill, c.font = HEAD_FILL, HEAD_FONT
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row + 1, 1)


def _write_analysis(wb, vrows):
    ws = wb.create_sheet("Full Analysis")
    heads = ["Status", "Direction", "Rank", "Stock Sheet No", "Line No", "Material No",
             "Material Description", "Book Qty", "Counted Qty", "Variance Qty",
             "Sales Usage", "Cocktail Usage", "Theoretical Usage",
             "Unexplained Qty", "Unit Price", "Unexplained Value",
             "Explains What % of the Shortage", "Sales Lines Matched",
             "Cocktails Consuming This"]
    widths = [18, 24, 7, 15, 9, 13, 42, 10, 11, 12, 11, 13, 15, 14, 11, 16, 15, 46, 46]
    _sheet_head(ws, "Every variance line, with what sales and cocktails explain", heads, widths)
    rows = sorted([r for r in vrows if r["diffqty"] != 0 or r["theoretical"] > 0],
                  key=lambda r: (r["status"] != "Unexplained", -r["unexplained_value"]))
    for i, r in enumerate(rows):
        rr = 4 + i
        cover = ""
        if r["shortage"] > 0 and r["theoretical"] >= 0:
            cover = min(r["theoretical"] / r["shortage"], 9.99) if r["shortage"] else ""
        vals = [r["status"], r["direction"], r.get("rank", ""), r["doc"], r["line"], r["material"], r["desc"],
                r["book"], r["counted"], r["diffqty"], r["sales_usage"], r["cocktail_usage"],
                r["theoretical"], r["unexplained_qty"], r["unitprice"], r["unexplained_value"],
                cover,
                "; ".join(f"{n} x{q:g} ({lab})" for n, q, lab, _ in r["sales_lines"][:6]),
                "; ".join(f"{n} x{q:g}" for n, q, _, _ in r["cocktail_lines"][:6])]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(rr, ci, v)
            if ci in (8, 9, 10, 11, 12, 13, 14):
                c.number_format = "#,##0.0##"
            elif ci in (15, 16):
                c.number_format = "#,##0.00"
            elif ci == 17:
                c.number_format = "0%"
        if r["status"] == "Unexplained":
            ws.cell(rr, 1).fill = SUB_FILL


def _write_no_variance(wb, vrows):
    rows = [r for r in vrows if r["status"] == "Sold, no variance"]
    ws = wb.create_sheet("Sold No Variance")
    heads = ["Stock Sheet No", "Line No", "Material No", "Material Description",
             "Book Qty", "Counted Qty", "Variance Qty", "Theoretical Usage",
             "Theoretical Value"]
    widths = [15, 9, 13, 42, 10, 11, 12, 16, 16]
    _sheet_head(
        ws,
        "Sold during the period but showing no variance - review, not recount",
        heads, widths, row=4)
    ws["A2"] = ("Consumption was recorded on Micros yet the count matches book stock. "
                "Usually a posting or capture issue rather than stock at risk.")
    rows.sort(key=lambda r: -r["theoretical"] * (r["unitprice"] or 0))
    for i, r in enumerate(rows):
        rr = 5 + i
        vals = [r["doc"], r["line"], r["material"], r["desc"], r["book"], r["counted"],
                r["diffqty"], r["theoretical"], r["theoretical"] * (r["unitprice"] or 0)]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(rr, ci, v)
            if ci in (5, 6, 7, 8):
                c.number_format = "#,##0.0##"
            elif ci == 9:
                c.number_format = "#,##0.00"


def _write_quality(wb, diag):
    ws = wb.create_sheet("Data Quality")
    ws["A1"] = "Data quality - fix these and the model gets sharper"
    ws["A1"].font = TITLE_FONT
    r = 3

    def block(title, note, heads, widths, rows):
        nonlocal r
        ws.cell(r, 1, title).font = Font(bold=True, size=12, color="1F3864")
        r += 1
        if note:
            ws.cell(r, 1, note)
            r += 1
        for i, (h, w) in enumerate(zip(heads, widths), 1):
            c = ws.cell(r, i, h)
            c.fill, c.font = HEAD_FILL, HEAD_FONT
            c.alignment = Alignment(horizontal="center", wrap_text=True)
            if ws.column_dimensions[get_column_letter(i)].width or 0 < w:
                ws.column_dimensions[get_column_letter(i)].width = w
        r += 1
        if not rows:
            ws.cell(r, 1, "None")
            r += 2
            return
        for row in rows:
            for ci, v in enumerate(row, 1):
                c = ws.cell(r, ci, v)
                if isinstance(v, float):
                    c.number_format = "#,##0.0##"
            r += 1
        r += 1

    block("Cocktails sold with no recipe in the tracker",
          "Their ingredient usage is missing, so those stock items look shorter than they are.",
          ["Cocktail (Micros)", "Qty Sold", "Family Group"], [44, 12, 24],
          [(s["name"], s["qty"], s["family"]) for s in
           sorted(diag["cocktail_unmatched"], key=lambda x: -x["qty"])])

    block("Recipes in the tracker with no matching sales line",
          "Either the drink did not sell, or the name differs between Micros and the tracker.",
          ["Recipe (Tracker)", "Sheet"], [44, 14],
          [(r_["name"], r_["sheet"]) for r_ in diag["cocktail_unused"]])

    block("Cocktail ingredients not on the count sheet",
          "Consumed by recipes but absent from the bar variance report.",
          ["Material No", "Description", "Qty Consumed", "UOM"], [14, 44, 14, 8],
          [(m, d["desc"], d["qty"], d["uom"]) for m, d in
           sorted(diag["off_report_ingredients"].items(), key=lambda x: -x[1]["qty"])])

    block("Sales lines not matched to any stock item",
          "Sold on Micros but nothing on the variance report matched, so the usage is "
          "unallocated. Where a likely stock item is suggested, pin it by adding the "
          "sales name and material number to item_aliases.csv, then run again.",
          ["Item (Micros)", "Family Group", "Qty Sold", "Possible match worth a look"],
          [40, 24, 10, 56],
          [(s["name"], s["family"], s["qty"], s.get("hint", "")) for s in
           sorted(diag["unmatched_sales"], key=lambda x: -x["qty"]) if s["qty"] > 0])

    block("Matches made on similar rather than identical names",
          "Check these read correctly. A wrong match explains away a real shortage.",
          ["Item (Micros)", "Matched to (SAP)", "Score", "Qty Sold", "Conversion applied",
           "Note"], [40, 40, 12, 10, 22, 40],
          diag["fuzzy_matches"])



# ============================================================================
# Bulk review: build item_aliases.csv once, top it up when something new appears
# ============================================================================

REVIEW_COLS = [
    ("Item (Micros)", 40),
    ("Family Group", 24),
    ("Qty Sold", 10),
    ("Avg Price", 11),
    ("Status", 16),
    ("Currently matched to (SAP)", 40),
    ("Material No", 13),
    ("Score", 13),
    ("Conversion applied", 20),
    ("Alternative worth a look", 44),
    ("CONFIRM Material No", 20),
    ("Your note", 26),
]


def write_review(diag, vrows, out_path, meta):
    """One fillable sheet listing every sales line the matcher was not certain
    about. Fill the CONFIRM column, then run --import-review on this file."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales Item Review"

    ws["A1"] = "SALES ITEM REVIEW - confirm what each Micros line is in SAP"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Fill the CONFIRM Material No column for any line that is wrong or blank, "
                "then run the script again with --import-review on this file. "
                "Leave a row blank to accept what is shown, or type SKIP to stop being "
                "asked about it.")
    ws["A3"] = (f"Sales Mix: {meta['sales_src']}    Variance report: {meta['source']}    "
                f"Generated: {meta['generated']}")

    hrow = 5
    for i, (label, width) in enumerate(REVIEW_COLS, 1):
        c = ws.cell(hrow, i, label)
        c.fill, c.font = HEAD_FILL, HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = width

    rows = sorted(diag["review"], key=lambda x: (x["row"] is not None, -x["sale"]["qty"]))
    for i, item in enumerate(rows):
        rr = hrow + 1 + i
        s, row = item["sale"], item["row"]
        vals = [
            s["name"],
            s["family"],
            s["qty"],
            s["avgprice"],
            "NO MATCH" if row is None else "check this",
            row["desc"] if row else "",
            row["material"] if row else "",
            item["how"],
            item["conv"],
            item.get("hint") or item.get("note") or "",
            None,
            None,
        ]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(rr, ci, v)
            c.border = BORDER
            if ci == 3:
                c.number_format = "#,##0.#"
            elif ci == 4:
                c.number_format = "#,##0.00"
            elif ci in (7, 11):
                c.alignment = Alignment(horizontal="center")
            if ci in (11, 12):
                c.fill = INPUT_FILL
            if row is None and ci == 5:
                c.font = Font(bold=True, color="C00000")

    ws.freeze_panes = ws.cell(hrow + 1, 1)
    ws.auto_filter.ref = f"A{hrow}:{get_column_letter(len(REVIEW_COLS))}{hrow + len(rows)}"

    # A lookup of the stock lines, so the material number can be found in Excel
    ws2 = wb.create_sheet("SAP Stock Lines")
    heads = ["Material No", "Material Description", "Stock Sheet No", "Line No",
             "Book Qty", "Counted Qty", "Variance Qty", "Unit Price"]
    widths = [14, 46, 15, 9, 10, 11, 12, 12]
    _sheet_head(ws2, "Every line on the variance report - look up material numbers here",
                heads, widths)
    for i, r in enumerate(sorted(vrows, key=lambda x: x["desc"])):
        rr = 4 + i
        for ci, v in enumerate([r["material"], r["desc"], r["doc"], r["line"], r["book"],
                                r["counted"], r["diffqty"], r["unitprice"]], 1):
            c = ws2.cell(rr, ci, v)
            if ci in (5, 6, 7):
                c.number_format = "#,##0.0##"
            elif ci == 8:
                c.number_format = "#,##0.00"
    ws2.auto_filter.ref = f"A3:H{3 + len(vrows)}"

    wb.save(out_path)
    return out_path, len(rows)


def import_review(path, folder, vrows):
    """Fold the confirmed material numbers back into item_aliases.csv."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Sales Item Review"] if "Sales Item Review" in wb.sheetnames else wb.worksheets[0]

    hrow = None
    for r in range(1, min(ws.max_row, 20) + 1):
        labels = [_nh(ws.cell(r, c).value) for c in range(1, (ws.max_column or 12) + 1)]
        if "item (micros)" in labels:
            hrow = r
            cols = {lab: i + 1 for i, lab in enumerate(labels) if lab}
            break
    if hrow is None:
        sys.exit(f"{os.path.basename(path)} does not look like a review sheet "
                 "(no 'Item (Micros)' header row found).")

    item_c = cols.get("item (micros)")
    conf_c = cols.get("confirm material no")
    note_c = cols.get("your note")
    if not conf_c:
        sys.exit("No 'CONFIRM Material No' column found in the review sheet.")

    valid = {r["material"] for r in vrows}
    by_material = {r["material"]: r["desc"] for r in vrows}
    confirmed, skipped, bad = [], [], []
    for r in range(hrow + 1, ws.max_row + 1):
        name = _txt(ws.cell(r, item_c).value)
        raw = _txt(ws.cell(r, conf_c).value)
        if not name or not raw:
            continue
        note = _txt(ws.cell(r, note_c).value) if note_c else ""
        if raw.strip().upper() in ("SKIP", "IGNORE", "NONE", "N/A"):
            skipped.append((name, note))
            continue
        if raw not in valid:
            bad.append((name, raw))
            continue
        confirmed.append((name, raw, note or by_material.get(raw, "")))

    path_csv = os.path.join(folder, ALIAS_FILE)
    existing, order = {}, []
    if os.path.exists(path_csv):
        with open(path_csv, newline="", encoding="utf-8-sig") as fh:
            for row in csv.DictReader(fh):
                frag = (row.get("sales_item_contains") or "").strip()
                if not frag:
                    continue
                existing[frag.lower()] = [frag, (row.get("material_no") or "").strip(),
                                          (row.get("note") or "").strip()]
                order.append(frag.lower())

    added = changed = 0
    for name, mat, note in confirmed:
        key = name.lower()
        if key in existing:
            if existing[key][1] != mat:
                existing[key][1] = mat
                existing[key][2] = note
                changed += 1
        else:
            existing[key] = [name, mat, note]
            order.append(key)
            added += 1
    for name, note in skipped:
        key = name.lower()
        if key not in existing:
            existing[key] = [name, "SKIP", note or "no stock line for this item"]
            order.append(key)

    with open(path_csv, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["sales_item_contains", "material_no", "note"])
        for key in order:
            w.writerow(existing[key])

    rule("=")
    print("  ALIAS IMPORT")
    rule()
    print(f"  Review sheet      : {os.path.basename(path)}")
    print(f"  Confirmed         : {len(confirmed)}")
    print(f"  Added             : {added}")
    print(f"  Updated           : {changed}")
    if skipped:
        print(f"  Marked as skip    : {len(skipped)}")
    if bad:
        print(f"  REJECTED          : {len(bad)} - material number not on the variance report")
        for name, raw in bad[:10]:
            print(f"     {name[:44]:46} -> {raw}")
    print(f"  Aliases file      : {path_csv}")
    rule("=")
    return len(confirmed)


# ============================================================================
# Console
# ============================================================================

def rule(ch="-", n=100):
    print(ch * n)


def preview(selected, pool, vrows, diag, meta):
    rule("=")
    print("  RECOUNT SELECTION - sales and cocktails vs the variance")
    rule("=")
    print(f"  Variance report   : {meta['source']}")
    print(f"  Sales Mix         : {meta['sales_src']}")
    print(f"  Cocktail Tracker  : {meta['cocktail_src']}")
    print(f"  Outlet / Plant    : {meta['location']} / {meta['plant']}")
    print(f"  Stock sheets      : {meta['docs']}")
    rule()
    print(f"  Lines on report   : {meta['total_lines']}")
    print(f"  Lines w/ variance : {meta['with_variance']}")
    print(f"  Explained by sales: {meta['explained']}  "
          f"(within {TOL_PERCENT:.0%} or R{TOL_VALUE:,.0f})")
    print(f"  Unexplained       : {meta['unexplained']}   "
          f"R {sum(r['unexplained_value'] for r in pool):,.2f}")
    print(f"  Sold, no variance : {meta['sold_no_variance']} (review tab)")
    print(f"  Cocktails matched : {len(diag['cocktail_links'])} of "
          f"{len(diag['cocktail_links']) + len(diag['cocktail_unmatched'])} selling")
    rule()
    print(f"  PROPOSED RECOUNT - top {len(selected)} by unexplained value")
    rule()
    print(f"  {'#':>3} {'Sheet':<10} {'Ln':<4} {'Material':<8} {'Description':<38} "
          f"{'Var':>7} {'Theor':>7} {'Unexpl':>7} {'Value R':>11}")
    rule()
    for r in selected:
        print(f"  {r['rank']:>3} {r['doc']:<10} {r['line']:<4} {r['material']:<8} "
              f"{r['desc'][:38]:<38} {r['diffqty']:>7.1f} {r['theoretical']:>7.1f} "
              f"{r['unexplained_qty']:>7.1f} {r['unexplained_value']:>11,.2f}")
    rule()
    sel = sum(r["unexplained_value"] for r in selected)
    tot = sum(r["unexplained_value"] for r in pool)
    print(f"  Selected value    : R {sel:,.2f} of R {tot:,.2f} unexplained "
          f"({(sel / tot * 100) if tot else 0:.1f}%)")
    if len(pool) > len(selected):
        print(f"  Not selected      : {len(pool) - len(selected)} lines, R {tot - sel:,.2f}")
    rule("=")
    if diag["cocktail_unmatched"]:
        top = sorted(diag["cocktail_unmatched"], key=lambda x: -x["qty"])[:5]
        print("  WATCH: cocktails sold with no recipe, so their ingredient usage is missing:")
        for s in top:
            print(f"     {s['name'][:44]:46} {s['qty']:>6.0f} sold")
        if len(diag["cocktail_unmatched"]) > 5:
            print(f"     ... and {len(diag['cocktail_unmatched']) - 5} more, see the "
                  "Data Quality tab")
        rule("=")


# ============================================================================
# Main
# ============================================================================

def main():
    ap = argparse.ArgumentParser(
        description="Select recount lines the Sales Mix and Cocktail Tracker cannot explain."
    )
    ap.add_argument("files", nargs="+",
                    help="Variance report, Sales Mix and Cocktail Tracker .xlsx, any order")
    ap.add_argument("--approve", action="store_true", help="Sign off and write the sheet")
    ap.add_argument("--approved-by", default="", help="Name recorded on the sheet")
    ap.add_argument("--limit", type=int, default=MAX_ITEMS,
                    help=f"Lines to select (default {MAX_ITEMS})")
    ap.add_argument("--out", default="", help="Output .xlsx path")
    ap.add_argument("--export-review", nargs="?", const="auto", default="",
                    metavar="PATH",
                    help="Write a fillable sheet of every uncertain sales line, then stop")
    ap.add_argument("--import-review", default="", metavar="PATH",
                    help="Explicitly read a filled-in review sheet into item_aliases.csv "
                        "(normally not needed - just include the file, it is auto-detected)")
    args = ap.parse_args()

    here = os.path.dirname(os.path.abspath(__file__))
    rules, cfg_path = load_family_rules(here)
    overrides = load_overrides(here)
    aliases, skips = load_aliases(here)

    paths = {"variance": None, "sales": None, "cocktail": None, "review": None}
    unclassified = []
    for f in args.files:
        f = os.path.abspath(f)
        if not os.path.isfile(f):
            sys.exit(f"File not found: {f}")
        kind = classify_file(f)
        if kind == "output":
            print(f"  NOTE: {os.path.basename(f)} looks like a file this script wrote on a "
                  "previous run (a Recount Selection workbook), not a source file. Ignoring it.")
            continue
        if kind == "review":
            if paths["review"] is None:
                paths["review"] = f
            else:
                print(f"  NOTE: {os.path.basename(f)} is a second completed review sheet - "
                      "only the first one found is used. Import the rest separately.")
            continue
        if kind and paths[kind] is None:
            paths[kind] = f
        else:
            unclassified.append(f)

    if not paths["variance"] and unclassified:
        # Only guess when there is nowhere else for the file to go, and say so.
        paths["variance"] = unclassified.pop(0)
        print(f"  NOTE: could not identify {os.path.basename(paths['variance'])} by its "
              "contents - assuming it is the variance report.")

    if not paths["variance"]:
        sys.exit("No variance report recognised among the files supplied. Pass the bar "
                 "variance export (Material, Material Description, Difference qty columns), not a "
                 "file this script produced.")

    vrows = read_variance(paths["variance"])
    sales = read_sales(paths["sales"]) if paths["sales"] else []
    recipes, master = read_cocktail_tracker(paths["cocktail"]) if paths["cocktail"] else ([], {})

    if not sales:
        print("  NOTE: no Sales Mix supplied. Nothing can be explained, so every "
              "variance line stays unexplained.")
    if not recipes:
        print("  NOTE: no Cocktail Tracker supplied. Cocktail ingredient usage is "
              "not accounted for.")

    apply_master_prices(vrows, master)

    # A completed review sheet handed back among the inputs (however it was
    # renamed) is imported straight into item_aliases.csv, then the aliases
    # are reloaded and the run carries straight on into the normal build -
    # no separate --import-review pass needed.
    review_path = args.import_review or paths["review"]
    if review_path:
        if not os.path.isfile(review_path):
            sys.exit(f"Review sheet not found: {review_path}")
        import_review(review_path, here, vrows)
        aliases, skips = load_aliases(here)
        print("  Continuing with the refreshed aliases ...")
        print()

    diag = build_model(vrows, sales, recipes, rules, overrides, aliases)
    diag["review"] = [x for x in diag["review"]
                      if not any(k in x["sale"]["name"].lower() for k in skips)]
    pool, selected = select(vrows, max(1, args.limit))

    docs = sorted({r["doc"] for r in vrows if r["doc"]})
    meta = {
        "location": next((r["location"] for r in vrows if r["location"]), ""),
        "plant": next((r["plant"] for r in vrows if r["plant"]), ""),
        "docs": ", ".join(docs),
        "source": os.path.basename(paths["variance"]),
        "sales_src": os.path.basename(paths["sales"]) if paths["sales"] else "not supplied",
        "cocktail_src": os.path.basename(paths["cocktail"]) if paths["cocktail"] else "not supplied",
        "generated": _dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "total_lines": len(vrows),
        "with_variance": sum(1 for r in vrows if r["diffqty"] != 0),
        "explained": sum(1 for r in vrows if r["status"] == "Explained by sales"),
        "unexplained": len(pool),
        "sold_no_variance": sum(1 for r in vrows if r["status"] == "Sold, no variance"),
        "approved_by": args.approved_by.strip(),
    }

    if args.export_review:
        if not sales:
            sys.exit("A Sales Mix is needed to build a review sheet.")
        out = args.export_review
        if out == "auto":
            out = os.path.join(os.path.dirname(paths["variance"]),
                               f"Sales_Item_Review_{_dt.datetime.now():%Y%m%d}.xlsx")
            cand, n = out, 2
            while os.path.exists(cand):
                cand = out[:-5] + f"_{n}.xlsx"
                n += 1
            out = cand
        out, count = write_review(diag, vrows, out, meta)
        rule("=")
        print("  SALES ITEM REVIEW")
        rule()
        nomatch = sum(1 for x in diag["review"] if x["row"] is None)
        print(f"  {count} sales lines need a decision:")
        print(f"     {nomatch} matched nothing at all")
        print(f"     {count - nomatch} matched, but not confidently enough to trust")
        print(f"  {len(diag['fuzzy_matches']) - (count - nomatch)} further matches were "
              f"confident enough to leave alone.")
        print(f"  Aliases already in force: {len(aliases)}")
        print()
        print(f"  Written to:\n     {out}")
        print()
        print("  Fill the CONFIRM Material No column, then run:")
        print(f'     python select_recount_items.py <the three files> '
              f'--import-review "{os.path.basename(out)}"')
        rule("=")
        return 0

    preview(selected, pool, vrows, diag, meta)

    needs_approval = len(pool) > MAX_ITEMS or args.limit > MAX_ITEMS
    if needs_approval and not args.approve:
        print()
        print("  ** APPROVAL REQUIRED - NOTHING WRITTEN **")
        print(f"  {len(pool)} lines cannot be explained by sales, which is more than "
              f"the standard limit of {MAX_ITEMS}.")
        print()
        print("  To proceed, re-run with:")
        print('     --approve --approved-by "Your Name"')
        print(f'     add --limit N to recount more than {MAX_ITEMS} lines')
        print()
        rule("=")
        sys.exit(2)

    if not args.out:
        stamp = _dt.datetime.now().strftime("%Y%m%d")
        tag = re.sub(r"[^A-Za-z0-9_-]", "", meta["location"] or (docs[0] if docs else "SEL"))
        base = os.path.join(os.path.dirname(paths["variance"]),
                            f"Recount_Selection_{tag}_{stamp}")
        cand, n = base + ".xlsx", 2
        while os.path.exists(cand):          # never overwrite a signed-off sheet
            cand, n = f"{base}_{n}.xlsx", n + 1
        args.out = cand

    out = write_sheet(selected, pool, vrows, diag, args.out, meta)
    print()
    print(f"  Recount selection sheet written:\n     {out}")
    if meta["approved_by"]:
        print(f"  Approved by: {meta['approved_by']}")
    print(f"  Conversions read from: {cfg_path}")
    rule("=")
    return 0


if __name__ == "__main__":
    sys.exit(main())
