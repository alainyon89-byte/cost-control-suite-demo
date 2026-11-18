#!/usr/bin/env python3
"""
update_cellar_consumption.py  --  v2.2 (11 Nov 2026)

Pulls the quantity sold (and, where blank, the selling price) from a Micros
sales report and writes it into the "Charged on Doc" (and "Selling Price")
columns of a Cellar Consumption workbook.

MICROS REPORT FORMATS (v2.2) -- both are read automatically
  - Sales Mix Items Summary (old Micros): one row per item with "Item",
    "Qty Sold", "Gross Sales", "Family Group" and "Major Group" columns.
  - Upgraded Micros export ("Name" / "Quantity Sold" table): Major Group
    (Beverage Enterprise, Food) and Family Group (Ent - Red Wine, ...) are
    ROWS above their items, not columns. Group rows are told apart from
    items by the two "Percent of ... Group Sales" columns, the group is
    carried down onto each item, and the items are checked against the
    report's Total row. Sales are taken as "Gross Sales Before Discounts"
    (VAT included). Item names are no longer cut off at 32 characters.

HOW MATCHING WORKS (v2)
  Every Micros line is resolved to a SAP material code, then written to the
  sheet row carrying that SAP code. Names are only used to find the code:
    1. cellar_item_map.xlsx    -- your confirmed Micros name -> SAP code list
    2. exact name in MM60      -- beverage materials only, never food
    3. same name, size label differs (300ml vs 330ml, "750ml" missing)
    4. scored match (handles Micros' 32-character cut-off, SAP
       abbreviations, spelling slips like Chardonay/Chardonnay)
       -> asked on screen with numbered choices; your answer is saved to
          the item map (dated backup first) so it is never asked again.
       -> with --no-prompt, only strong unambiguous matches are used and
          they are listed under "Auto-matched, check these".

SERVES (spirits and cordials)
  Rows whose Cost is per serve (Cost formula ends in /30) are kept in
  serves end to end: Charged on Doc = Micros qty, and Usage is rewritten
  as (Open + Add - Close) so Variance compares tots with tots. Cordial
  rows are switched to the same basis (Cost /30 for a 750ml bottle at
  25ml). Wine sold by the glass is converted to bottle fractions. Pours
  and which categories count in serves live on the "House Rules" sheet of
  cellar_item_map.xlsx, keyed by SAP material group.

Usage:
    python update_cellar_consumption.py --files SalesMix.xlsx Cellar.xlsx
    python update_cellar_consumption.py --micros SalesMix.xlsx --cellar Cellar.xlsx --sheets "Bar 1,Bar 2"

Output: "<Cellar>_UPDATED.xlsx" next to the original (never overwritten),
plus "<Cellar>_UPDATE_LOG.txt" with everything printed on screen.
"""

import argparse
import datetime
import difflib
import os
import re
import shutil
import sys
from collections import Counter, defaultdict

import warnings

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.styles import Font, PatternFill

warnings.filterwarnings("ignore", message="Workbook contains no default style")
VERSION = "v2.2 (11 Nov 2026)"
DEFAULT_MAP_NAME = "cellar_item_map.xlsx"
COS_FLAG_LOW, COS_FLAG_HIGH = 0.05, 0.80   # house rule from the Theo CoS work
VAT_DIVISOR = 1.15

# Default house rules, used when the item map has no "House Rules" sheet.
# Material Group -> (category, pour ml, how the sheet counts it)
DEFAULT_HOUSE_RULES = {
    "BEV000100": ("Aperitifs & bitters", 25, "Serves"),
    "BEV000300": ("Brandy", 25, "Serves"),
    "BEV000500": ("Cognac", 25, "Serves"),
    "BEV000700": ("Liqueurs", 25, "Serves"),
    "BEV001100": ("Cordials & syrups", 25, "Serves"),
    "BEV001500": ("Rum", 25, "Serves"),
    "BEV001800": ("Whisky", 25, "Serves"),
    "BEV001900": ("Gin, vodka & white spirits", 25, "Serves"),
    "BEV001200": ("Port & dessert wine", 50, "Bottles"),
    "BEV001600": ("Sherry & fortified", 50, "Bottles"),
    "BEV001300": ("Red wine", 187.5, "Bottles"),
    "BEV001400": ("Rose wine", 187.5, "Bottles"),
    "BEV002000": ("White wine", 187.5, "Bottles"),
    "BEV002010": ("Auction wine", 187.5, "Bottles"),
    "BEV001700": ("Sparkling", 125, "Bottles"),
}

YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


# ----------------------------------------------------------------------
# Output: everything printed also goes to a log file
# ----------------------------------------------------------------------
class Log:
    def __init__(self):
        self.lines = []

    def __call__(self, text=""):
        print(text)
        self.lines.append(str(text))

    def record(self, text=""):
        """Log file only (e.g. an answer the console already shows)."""
        self.lines.append(str(text))

    def save(self, path):
        try:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write("\n".join(self.lines) + "\n")
            return True
        except OSError:
            return False


log = Log()


# ----------------------------------------------------------------------
# Text helpers
# ----------------------------------------------------------------------
def normalize(name):
    if name is None:
        return ""
    return re.sub(r"\s+", " ", str(name).strip().lower())


SIZE_TOKEN_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*(ml|cl|ltr|lt|litre|liter|l)\b", re.IGNORECASE)
TYPICAL_SIZES = ("200", "250", "275", "300", "330", "340", "355", "440", "500", "750", "1000")
STOPWORDS = {"wines", "wine", "estate", "the", "and", "of", "&"}
SERVE_MARKERS = {"gls"}          # "Botanical Gin GLS" = a tot / glass
SHORT_ABBREV = {"bl": "blanc", "sav": "sauvignon"}


def size_ml(name):
    """Bottle/can size in ml from a description ('750ml' -> 750, '1lt' -> 1000)."""
    if not name:
        return None
    m = SIZE_TOKEN_RE.search(str(name))
    if not m:
        return None
    val = float(m.group(1).replace(",", "."))
    unit = m.group(2).lower()
    if unit == "ml":
        return val
    if unit == "cl":
        return val * 10
    return val * 1000


def normalize_size_agnostic(name):
    key = normalize(name)
    key = SIZE_TOKEN_RE.sub(" ", key)
    key = re.sub(r"[^a-z0-9&]+$", "", key.strip())
    return re.sub(r"\s+", " ", key).strip()


def split_words(name):
    """Lower-case word list. Splits SAP's run-together names
    ('TrizanneSignatureWinesTSW' -> trizanne signature wines tsw)."""
    s = str(name or "")
    s = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", s)
    s = re.sub(r"(?<=[A-Z])(?=[A-Z][a-z])", " ", s)
    s = SIZE_TOKEN_RE.sub(" ", s)
    return re.findall(r"[a-z0-9]+", s.lower())


def name_tokens(name, is_micros):
    words = split_words(name)
    raw = str(name or "")
    if is_micros and words and words[-1].isdigit():
        last = words[-1]
        cut_off = 30 <= len(raw.rstrip()) <= 32   # old Micros cut names at 32; the upgraded one does not
        if len(last) >= 2 and (cut_off or any(sz.startswith(last) for sz in TYPICAL_SIZES)):
            words = words[:-1]              # truncated size, e.g. "330" / "33"
    out, serve_marker = [], False
    for w in words:
        if w in SERVE_MARKERS:
            serve_marker = True
            continue
        if w in STOPWORDS:
            continue
        if not is_micros and re.fullmatch(r"[a-z]+\d+", w):
            continue                        # SAP stock codes like abc100
        out.append(w)
    return out, serve_marker


def token_match(a, b):
    if a == b:
        return True
    if SHORT_ABBREV.get(a) == b or SHORT_ABBREV.get(b) == a:
        return True
    short, long_ = (a, b) if len(a) <= len(b) else (b, a)
    if len(short) >= 3 and long_.startswith(short) and not short.isdigit():
        return True                         # cab ~ cabernet, chardo ~ chard
    if len(short) >= 5 and difflib.SequenceMatcher(None, a, b).ratio() >= 0.8:
        return True                         # chardonay ~ chardonnay
    return False


def match_score(m_tokens, c_tokens, m_size=None, c_size=None):
    """Returns (score, share of Micros words found, brand_ok).
    An exact word counts 1, an abbreviation/spelling match 0.85."""
    if not m_tokens or not c_tokens:
        return 0.0, 0.0, False
    quality = []
    for m in m_tokens:
        if m in c_tokens:
            quality.append(1.0)
        elif any(token_match(m, c) for c in c_tokens):
            quality.append(0.85)
        else:
            quality.append(0.0)
    c_hit = sum(1 for c in c_tokens if any(token_match(m, c) for m in m_tokens))
    m_found = sum(1 for q in quality if q > 0) / len(m_tokens)
    brand_ok = token_match(m_tokens[0], c_tokens[0])
    score = sum(quality) / len(m_tokens) + 0.25 * c_hit / len(c_tokens) + (0.1 if brand_ok else 0.0)
    if m_size and c_size and abs(m_size - c_size) < 1:
        score += 0.05
    return score, m_found, brand_ok


def to_float(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").replace("R", "").strip())
    except ValueError:
        return None


def norm_code(v):
    """SAP code as int where numeric, else stripped text; None if blank."""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, int):
        return v
    s = str(v).strip()
    if not s:
        return None
    if re.fullmatch(r"\d+(\.0+)?", s):
        return int(float(s))
    return s


def fmt_num(x):
    if x is None:
        return ""
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    return f"{x:.4f}".rstrip("0").rstrip(".")


def money(x):
    return f"R{x:,.2f}"


def is_interactive():
    try:
        return sys.stdin is not None and sys.stdin.isatty()
    except Exception:
        return False


# ----------------------------------------------------------------------
# Micros report
# ----------------------------------------------------------------------
def _micros_line(item, qty, avg_price, sales, family, major):
    qty = to_float(qty) or 0.0
    avg = to_float(avg_price)
    sales = to_float(sales)
    if avg is None and sales is not None and qty:
        avg = sales / qty
    if sales is None and avg is not None:
        sales = avg * qty
    return {
        "raw_name": str(item),
        "key": normalize(item),
        "qty": qty,
        "avg_price": avg,
        "sales": sales or 0.0,
        "family_group": (str(family).strip() if family else ""),
        "major_group": (str(major).strip() if major else ""),
    }


class MicrosLines(list):
    """List of Micros lines plus what was learnt while reading them."""
    fmt = ""
    notes = ()


def _is_dash(v):
    return v is None or str(v).strip() in ("", "-", "--", "–", "—")


def _is_qty_header(v):
    """'Qty Sold' (old) or 'Quantity Sold' (upgraded), but not the
    '%' / 'Percent of Total Quantity Sold' columns."""
    if "%" in v or "percent" in v:
        return False
    return "qty sold" in v or "quantity sold" in v


def micros_columns(header):
    """Column positions (0-based) from a normalized header row, or {} if it is
    not a Micros item table. Understands the old Sales Mix Items Summary
    ("Item", "Qty Sold", "Gross Sales") and the upgraded Micros export
    ("Name", "Quantity Sold", "Gross Sales Before Discounts Total")."""
    cols = {}
    for i, v in enumerate(header):
        if v in ("item", "name"):
            if "item" not in cols:
                cols["item"], cols["style"] = i, v
        elif _is_qty_header(v):
            cols.setdefault("qty", i)
        elif "average price" in v:
            cols.setdefault("avg_price", i)
        elif v == "gross sales" or v.startswith("gross sales before discounts"):
            cols.setdefault("sales", i)
        elif v.startswith("gross sales after discounts"):
            cols.setdefault("sales_after", i)
        elif v.startswith("discounts"):
            cols.setdefault("discounts", i)
        elif v == "family group":
            cols.setdefault("family_group", i)
        elif v == "major group":
            cols.setdefault("major_group", i)
        elif v.startswith("percent of major group"):
            cols.setdefault("pct_major", i)
        elif v.startswith("percent of family group"):
            cols.setdefault("pct_family", i)
    if "item" not in cols or "qty" not in cols:
        return {}
    if "sales" not in cols and "sales_after" in cols:
        cols["sales"] = cols["sales_after"]
    return cols


def parse_micros_rows(rows, cols):
    """rows: lists of cell values (header row already removed). Handles both
    layouts and returns MicrosLines.

    Upgraded layout: Total, then Major Group, Family Group and item rows in
    order. A row with a number in 'Percent of Family Group Sales' is an item;
    a number only in 'Percent of Major Group Sales' is a Family Group; dashes
    in both is a Major Group. The current groups are carried onto each item."""
    def get(row, key):
        i = cols.get(key)
        return row[i] if i is not None and i < len(row) else None

    hierarchy = "pct_major" in cols and "pct_family" in cols
    if cols.get("style") == "name" and not hierarchy:
        raise ValueError(
            "This looks like the upgraded Micros report, but its 'Percent of Major Group Sales' / "
            "'Percent of Family Group Sales' columns are missing, so item rows cannot be told apart "
            "from group totals. Export the report with those columns included.")

    lines = MicrosLines()
    lines.fmt = "upgraded Micros" if hierarchy else "Sales Mix"
    major = family = ""
    total, discounts = None, 0.0
    for row in rows:
        item = get(row, "item")
        if item is None or not normalize(item):
            continue
        key = normalize(item)
        if key.startswith("total item"):
            continue
        if hierarchy:
            qty, sales = to_float(get(row, "qty")) or 0.0, to_float(get(row, "sales")) or 0.0
            if key == "total":
                total = (qty, sales)
                continue
            if _is_dash(get(row, "pct_family")):          # a group row, not an item
                if not _is_dash(get(row, "pct_major")):
                    family = str(item).strip()
                elif qty or sales:                         # an empty row of dashes changes nothing
                    major, family = str(item).strip(), ""
                continue
            fam, maj = family, major
        else:
            fam, maj = get(row, "family_group"), get(row, "major_group")
        lines.append(_micros_line(item, get(row, "qty"), get(row, "avg_price"), get(row, "sales"), fam, maj))
        discounts += to_float(get(row, "discounts")) or 0.0

    if hierarchy and not lines:
        raise ValueError("The upgraded Micros report was recognised but no item rows were found in it.")
    notes = []
    if hierarchy and total is not None:
        got_q, got_s = sum(ln["qty"] for ln in lines), sum(ln["sales"] for ln in lines)
        if abs(got_q - total[0]) > 0.5 or abs(got_s - total[1]) > 1.0:
            notes.append(f"WARNING: the item rows add up to {fmt_num(got_q)} sold / {money(got_s)} but the report's "
                         f"Total row says {fmt_num(total[0])} / {money(total[1])} -- some rows may have been misread.")
    if discounts:
        notes.append(f"Note: this report has discounts of {money(discounts)}. Sales used here are before discounts.")
    lines.notes = notes
    return lines


def load_micros_xlsx(path):
    """Returns a LIST of Micros lines (not a dict) so two items that old Micros
    truncated to the same 32-character name are both kept. The header row is
    searched for on every sheet (the upgraded export starts its table near
    row 41, below a summary block)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    for ws in [wb.active] + [w for w in wb.worksheets if w is not wb.active]:
        width = ws.max_column
        for r in range(1, min(ws.max_row, 300) + 1):
            cols = micros_columns([normalize(ws.cell(row=r, column=c).value) for c in range(1, width + 1)])
            if cols:
                rows = ([ws.cell(row=rr, column=c).value for c in range(1, width + 1)]
                        for rr in range(r + 1, ws.max_row + 1))
                return parse_micros_rows(rows, cols)
    raise ValueError(
        "Could not find the header row (looking for 'Item' or 'Name' together with 'Qty Sold' or "
        f"'Quantity Sold') in {path}. Is this a Micros sales report export?"
    )


def load_micros_pdf(path):
    """Same list from a Micros PDF export (pdfplumber). Tables on later pages
    that don't repeat the header reuse the last header's column layout, and
    the Major/Family Group rows of the upgraded layout carry across pages."""
    import pdfplumber

    cols, width, rows = None, None, []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                if not table:
                    continue
                found = micros_columns([normalize(c) for c in table[0]])
                if found:
                    cols, width, body = found, len(table[0]), table[1:]
                elif cols and len(table[0]) == width:
                    body = table                      # continuation page without a header
                else:
                    continue
                rows.extend(r for r in body if r)
    if cols is None or not rows:
        raise ValueError(
            f"Could not extract any item/qty rows from {path}. "
            "The PDF table layout may differ -- try exporting as .xlsx instead."
        )
    return parse_micros_rows(rows, cols)


def load_micros(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        return load_micros_xlsx(path)
    if ext == ".pdf":
        return load_micros_pdf(path)
    raise ValueError(f"Unsupported Micros file type: {ext} (expected .xlsx or .pdf)")


def get_micros_business_date(path):
    """'Business Dates' from the top of a Micros xlsx (US month/day/year)."""
    if os.path.splitext(path)[1].lower() != ".xlsx":
        return None
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        for r in range(1, 10):
            if normalize(ws.cell(row=r, column=1).value) == "business dates":
                raw = ws.cell(row=r, column=2).value
                if isinstance(raw, datetime.datetime):
                    return raw.date()
                m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})", str(raw or "").strip())
                if m:
                    month, day, year = (int(x) for x in m.groups())
                    year += 2000 if year < 100 else 0
                    try:
                        return datetime.date(year, month, day)
                    except ValueError:
                        return None
                return None
    except Exception:
        return None
    return None


def parse_sheet_name_as_date(name):
    """'04.09.2026' / '04,09,2026' / '04/09/2026' -> date (day.month.year)."""
    m = re.match(r"^(\d{1,2})\s*[.,/\-]\s*(\d{1,2})\s*[.,/\-]\s*(\d{2,4})$", str(name).strip())
    if not m:
        return None
    day, month, year = (int(x) for x in m.groups())
    year += 2000 if year < 100 else 0
    try:
        return datetime.date(year, month, day)
    except ValueError:
        return None


# ----------------------------------------------------------------------
# MM60 (SAP master)
# ----------------------------------------------------------------------
class MM60:
    def __init__(self, ws):
        self.by_code = {}
        self.bev = []          # candidate list for matching: beverage materials only
        if ws is None:
            return
        header_row, cols = None, {}
        for r in range(1, min(ws.max_row, 5) + 1):
            vals = {normalize(ws.cell(row=r, column=c).value): c for c in range(1, ws.max_column + 1)}
            if "material" in vals and "material description" in vals:
                header_row, cols = r, vals
                break
        if not header_row:
            return
        c_code, c_desc = cols["material"], cols["material description"]
        c_group, c_price = cols.get("material group"), cols.get("price")
        c_val = cols.get("valuation class")
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            code = norm_code(row[c_code - 1] if len(row) >= c_code else None)
            desc = row[c_desc - 1] if len(row) >= c_desc else None
            if code is None or desc is None:
                continue
            group = str(row[c_group - 1]).strip() if c_group and row[c_group - 1] else ""
            price = to_float(row[c_price - 1]) if c_price else None
            val_class = str(row[c_val - 1]).strip() if c_val and row[c_val - 1] is not None else ""
            rec = {"code": code, "desc": str(desc).strip(), "group": group, "price": price}
            if code not in self.by_code:
                self.by_code[code] = rec
            is_bev = group.upper().startswith("BEV") or val_class == "3031"
            if is_bev and "do not use" not in normalize(desc):
                toks, _ = name_tokens(desc, is_micros=False)
                rec["tokens"], rec["size"] = toks, size_ml(desc)
                self.bev.append(rec)
        self.exact = {}
        for rec in self.bev:
            self.exact.setdefault(normalize(rec["desc"]), rec["code"])
        sa, ambiguous = {}, set()
        for rec in self.bev:
            k = normalize_size_agnostic(rec["desc"])
            if k in sa and sa[k] != rec["code"]:
                ambiguous.add(k)
            else:
                sa[k] = rec["code"]
        self.size_agnostic = {k: v for k, v in sa.items() if k not in ambiguous}

    def desc(self, code):
        rec = self.by_code.get(code)
        return rec["desc"] if rec else None


# ----------------------------------------------------------------------
# Item map + house rules (cellar_item_map.xlsx)
# ----------------------------------------------------------------------
MAP_HEADERS = ["Micros Item", "SAP Code", "SAP Description", "Serves per Unit", "Source", "Notes"]


def load_item_map(path):
    """normalized Micros name -> {"sap": code | "IGNORE", "serves": float|None}"""
    entries = {}
    if not path or not os.path.exists(path):
        return entries
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Item Map" not in wb.sheetnames:
        return entries
    ws = wb["Item Map"]
    hdr = {normalize(ws.cell(row=1, column=c).value): c for c in range(1, ws.max_column + 1)}
    c_item, c_sap = hdr.get("micros item"), hdr.get("sap code")
    c_serves = hdr.get("serves per unit")
    if not c_item or not c_sap:
        return entries
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=c_item).value
        sap = ws.cell(row=r, column=c_sap).value
        if not normalize(name) or sap in (None, ""):
            continue
        code = "IGNORE" if normalize(sap) == "ignore" else norm_code(sap)
        serves = to_float(ws.cell(row=r, column=c_serves).value) if c_serves else None
        entries.setdefault(normalize(name), {"sap": code, "serves": serves})
    return entries


def load_house_rules(path):
    rules = dict(DEFAULT_HOUSE_RULES)
    if not path or not os.path.exists(path):
        return rules
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception:
        return rules
    if "House Rules" not in wb.sheetnames:
        return rules
    ws = wb["House Rules"]
    hdr = {normalize(ws.cell(row=1, column=c).value): c for c in range(1, ws.max_column + 1)}
    c_g, c_cat = hdr.get("material group"), hdr.get("category")
    c_pour, c_unit = hdr.get("pour ml"), hdr.get("sheet counts in")
    if not (c_g and c_pour):
        return rules
    loaded = {}
    for r in range(2, ws.max_row + 1):
        g = ws.cell(row=r, column=c_g).value
        pour = to_float(ws.cell(row=r, column=c_pour).value)
        if not g or not pour:
            continue
        unit = normalize(ws.cell(row=r, column=c_unit).value) if c_unit else "bottles"
        cat = ws.cell(row=r, column=c_cat).value if c_cat else ""
        loaded[str(g).strip().upper()] = (cat or "", pour, "Serves" if unit.startswith("serve") else "Bottles")
    return loaded or rules


def append_to_item_map(path, new_entries, mm60):
    """Appends confirmed answers. Takes a dated backup first and never
    touches existing rows. Returns True on success."""
    if not new_entries:
        return True
    try:
        if os.path.exists(path):
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
            shutil.copy2(path, os.path.splitext(path)[0] + f"_backup_{stamp}.xlsx")
            wb = openpyxl.load_workbook(path)
            ws = wb["Item Map"] if "Item Map" in wb.sheetnames else wb.create_sheet("Item Map", 0)
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Item Map"
        if ws.max_row == 1 and ws.cell(row=1, column=1).value is None:
            for c, h in enumerate(MAP_HEADERS, start=1):
                ws.cell(row=1, column=c).value = h
                ws.cell(row=1, column=c).font = Font(bold=True)
        today = datetime.date.today().strftime("%d %b %Y")
        for e in new_entries:
            r = ws.max_row + 1
            ws.cell(row=r, column=1).value = e["name"]
            ws.cell(row=r, column=2).value = e["sap"]
            ws.cell(row=r, column=3).value = mm60.desc(e["sap"]) if e["sap"] != "IGNORE" else "(not a stock item)"
            ws.cell(row=r, column=5).value = f"Confirmed on screen {today}"
        wb.save(path)
        return True
    except PermissionError:
        return False


# ----------------------------------------------------------------------
# Resolving Micros lines to SAP codes
# ----------------------------------------------------------------------
def candidates_for(line, mm60, sheet_codes, top=5):
    m_toks, _ = name_tokens(line["raw_name"], is_micros=True)
    m_size = size_ml(line["raw_name"])
    scored = []
    for rec in mm60.bev:
        s, cov, brand = match_score(m_toks, rec["tokens"], m_size, rec["size"])
        if cov <= 0:
            continue
        on_sheet = rec["code"] in sheet_codes
        scored.append({"code": rec["code"], "desc": rec["desc"], "price": rec["price"],
                       "score": s + (0.05 if on_sheet else 0.0), "cov": cov, "brand": brand,
                       "on_sheet": on_sheet})
    scored.sort(key=lambda d: d["score"], reverse=True)
    return scored[:top]


def strong_auto_match(cands):
    """Only used without prompting: every Micros word must be found in the
    SAP name, same brand word first, and a clear gap to the runner-up.
    ('Estate Reserve Blend Nad' is NOT auto-matched to 'Estate
    Reserve' -- 'Nad' is missing.)"""
    if not cands:
        return None
    best = cands[0]
    second = cands[1]["score"] if len(cands) > 1 else 0.0
    if best["brand"] and best["cov"] >= 0.999 and best["score"] - second >= 0.15:
        return best
    return None


def ask_user(line, cands, mm60):
    log()
    log(f"  New Micros item: '{line['raw_name'].strip()}'  ({fmt_num(line['qty'])} sold"
        + (f" @ {money(line['avg_price'])}" if line["avg_price"] else "")
        + (f", {line['family_group']}" if line["family_group"] else "") + ")")
    for i, c in enumerate(cands, start=1):
        price = money(c["price"]) if c["price"] is not None else "no price"
        tag = "  [on this sheet]" if c["on_sheet"] else ""
        log(f"     {i}) {c['code']:<7} {c['desc'][:42]:<42} {price:>10}{tag}")
    # Enter only accepts option 1 when it is a clear match; otherwise Enter skips,
    # so a quick Enter can never save a weak guess to the item map.
    default = cands[0]["code"] if strong_auto_match(cands) else "SKIP"
    log(f"     Enter = {'1' if default != 'SKIP' else 'skip'} | number = pick | type a SAP code | "
        "S = skip this time | I = ignore always (not stock)")
    for _ in range(3):
        try:
            ans = input("  Your choice: ").strip()
        except EOFError:
            return "SKIP"
        log.record(f"  Your choice: {ans}")
        if ans == "":
            return default
        if ans.lower() == "s":
            return "SKIP"
        if ans.lower() == "i":
            return "IGNORE"
        if ans.isdigit() and 1 <= int(ans) <= len(cands) and len(ans) <= 2:
            return cands[int(ans) - 1]["code"]
        code = norm_code(ans)
        if code in mm60.by_code:
            return code
        log("     Not a listed number or a SAP code found in MM60 -- try again.")
    return "SKIP"


def lookup_item_map(item_map, key):
    """Item map lookup. Old Micros cut names at 32 characters and the upgraded
    Micros does not, so a map entry saved from a cut-off name still applies to
    the full name (and the other way round) -- but only when the long name
    starts with the whole short one, the short one is 30+ characters, and
    every entry that fits points at the same SAP code."""
    hit = item_map.get(key)
    if hit:
        return hit
    fits = [v for mk, v in item_map.items()
            if (len(mk) >= 30 and key.startswith(mk)) or (len(key) >= 30 and mk.startswith(key))]
    if fits and len({str(v["sap"]) for v in fits}) == 1:
        return fits[0]
    return None


def resolve_lines(lines, mm60, sheet_codes, item_map, interactive):
    """Adds "sap" and "how" to every line. Returns new item-map entries."""
    new_entries = []
    asked = {}
    # Names of the items already on the sheet(s): an exact or size-only match
    # against these wins (e.g. Micros 'Dry Gin' -> the sheet's 'Dry Gin 750ml'
    # even though MM60 also has 'Dry Gin 1lt').
    sheet_exact, sheet_sa, sa_clash = {}, {}, set()
    for code in sheet_codes:
        d = mm60.desc(code)
        if not d:
            continue
        sheet_exact.setdefault(normalize(d), code)
        k = normalize_size_agnostic(d)
        if k in sheet_sa and sheet_sa[k] != code:
            sa_clash.add(k)
        sheet_sa.setdefault(k, code)
    for k in sa_clash:
        sheet_sa.pop(k, None)

    for ln in lines:
        ln["sap"], ln["how"], ln["suggest"] = None, None, None
        if normalize(ln["major_group"]) == "food":
            ln["how"] = "food"
            continue
        hit = lookup_item_map(item_map, ln["key"])
        if hit:
            ln["sap"], ln["how"] = hit["sap"], "map"
            ln["serves_override"] = hit.get("serves")
            if hit["sap"] == "IGNORE":
                ln["how"] = "ignored"
            continue
        code = sheet_exact.get(ln["key"])
        if code is None:
            code = mm60.exact.get(ln["key"])
        if code is not None:
            ln["sap"], ln["how"] = code, "exact"
            continue
        sa_key = normalize_size_agnostic(ln["raw_name"])
        code = sheet_sa.get(sa_key)
        if code is None:
            code = mm60.size_agnostic.get(sa_key)
        if code is not None:
            ln["sap"], ln["how"] = code, "size"
            continue
        if ln["qty"] <= 0:
            continue                      # nothing sold: no need to ask
        if ln["key"] in asked:            # same truncated name seen already this run
            ln["sap"], ln["how"] = asked[ln["key"]]
            continue
        cands = candidates_for(ln, mm60, sheet_codes)
        if interactive and cands:
            choice = ask_user(ln, cands, mm60)
            if choice == "SKIP":
                ln["suggest"] = cands[0] if cands else None
                asked[ln["key"]] = (None, None)
                continue
            ln["sap"], ln["how"] = choice, ("ignored" if choice == "IGNORE" else "asked")
            asked[ln["key"]] = (ln["sap"], ln["how"])
            new_entries.append({"name": ln["raw_name"].strip(), "sap": choice})
            continue
        best = strong_auto_match(cands)
        if best:
            ln["sap"], ln["how"], ln["auto_from"] = best["code"], "auto", best["desc"]
        else:
            ln["suggest"] = cands[0] if cands else None
    return new_entries


# ----------------------------------------------------------------------
# Cellar sheet layout
# ----------------------------------------------------------------------
def find_cellar_header(ws):
    for r in range(1, min(ws.max_row, 15) + 1):
        row_vals = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None:
                row_vals.setdefault(normalize(v), c)
        if "item description" in row_vals and "charged on doc" in row_vals:
            return r, row_vals
    raise ValueError("Could not find a header row with 'Item Description' and 'Charged on Doc' in this sheet.")


def find_sections(ws, cols):
    """Category sections from each subtotal row's =SUM(K a:K b) formula."""
    tr_col = cols.get("total revenue")
    if tr_col is None:
        return []
    from openpyxl.utils import get_column_letter
    letter = get_column_letter(tr_col)
    pattern = re.compile(rf"^=SUM\({letter}(\d+):{letter}(\d+)\)$", re.IGNORECASE)
    desc_col = cols.get("item description")
    sections = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=tr_col).value
        if isinstance(v, str):
            m = pattern.match(v.replace(" ", ""))
            if m:
                label_val = ws.cell(row=r, column=desc_col).value if desc_col else None
                sections.append({"header_row": r, "body_start": int(m.group(1)), "body_end": int(m.group(2)),
                                 "label": label_val.strip() if isinstance(label_val, str) else ""})
    return sections


def broad_class(label):
    t = normalize(label).replace("ent - ", "").replace("prop - ", "")
    for key, cls in (("red wine", "red wine"), ("white wine", "white wine"), ("rose", "rose wine"),
                     ("sparkling", "sparkling"), ("whisk", "whisky"), ("brandy", "brandy"),
                     ("rum", "rum"), ("white spirit", "white spirits"), ("liqueur", "liqueur"),
                     ("beer", "beer"), ("cooler", "beer"), ("mineral", "minerals")):
        if key in t:
            return cls
    return None


# ----------------------------------------------------------------------
# One sheet
# ----------------------------------------------------------------------
def process_sheet(wb, sheet_name, lines, mm60, rules):
    ws = wb[sheet_name]
    header_row, cols = find_cellar_header(ws)
    c_sap = cols.get("sap code") or cols.get("sap no")
    c_charged = cols["charged on doc"]
    c_cost, c_price = cols.get("cost"), cols.get("selling price")
    c_open, c_add, c_close, c_usage = cols.get("open"), cols.get("add"), cols.get("close"), cols.get("usage")
    if c_sap is None:
        log(f"\n=== Sheet '{sheet_name}' === no 'Sap Code' column found, skipped.")
        return
    from openpyxl.utils import get_column_letter as L

    sections = find_sections(ws, cols)
    section_rows = set()
    for s in sections:
        section_rows.update(range(s["body_start"], s["body_end"] + 1))

    def section_of(r):
        for i, s in enumerate(sections):
            if s["body_start"] <= r <= s["body_end"]:
                return i
        return None

    # --- rows on the sheet, by SAP code ---------------------------------
    rows_by_code = defaultdict(list)
    for r in range(header_row + 1, ws.max_row + 1):
        if sections and r not in section_rows:
            continue
        code = norm_code(ws.cell(row=r, column=c_sap).value)
        if code is not None:
            rows_by_code[code].append(r)

    # --- how each Micros line converts ----------------------------------
    def line_serves(ln, code):
        """Serves per bottle for this Micros line (1 = sold whole)."""
        if ln.get("serves_override"):
            return ln["serves_override"]
        rec = mm60.by_code.get(code) or {}
        rule = rules.get(str(rec.get("group", "")).upper())
        if not rule:
            return 1.0
        _, pour, _ = rule
        bottle = size_ml(rec.get("desc"))
        _, gls = name_tokens(ln["raw_name"], is_micros=True)
        unit_price = rec.get("price")
        by_serve = gls or (ln["avg_price"] and unit_price and ln["avg_price"] < unit_price)
        if not by_serve:
            return 1.0
        return bottle / pour if bottle else None

    sold = defaultdict(list)         # code -> [(line, serves)]
    ignored, food = [], []
    for ln in lines:
        if ln["how"] == "food":
            if ln["qty"] > 0:
                food.append(ln)
            continue
        if ln["how"] == "ignored":
            if ln["qty"] > 0:
                ignored.append(ln)
            continue
        if ln["sap"] is None:
            continue
        sold[ln["sap"]].append((ln, line_serves(ln, ln["sap"])))

    # --- insert sold items that have no row yet --------------------------
    added, not_placed, section_full = [], [], []
    claimed = set()
    fg_votes, mg_votes = defaultdict(Counter), defaultdict(Counter)
    for code, rows in rows_by_code.items():
        sec = section_of(rows[0])
        if sec is None:
            continue
        grp = (mm60.by_code.get(code) or {}).get("group")
        if grp:
            mg_votes[grp][sec] += 1
        for ln, _ in sold.get(code, []):
            if ln["family_group"]:
                fg_votes[normalize(ln["family_group"])][sec] += 1

    def pick_section(ln, code):
        fg = normalize(ln["family_group"])
        label_hit = next((i for i, s in enumerate(sections) if s["label"] and normalize(s["label"]) == fg), None)
        for votes in (fg_votes.get(fg), mg_votes.get((mm60.by_code.get(code) or {}).get("group"))):
            if votes:
                top = votes.most_common()
                best_n = top[0][1]
                tied = [sec for sec, n in top if n == best_n]
                if label_hit in tied:
                    return label_hit
                return tied[0]
        return label_hit

    def free_row(sec):
        s = sections[sec]
        for r in range(s["body_start"], s["body_end"] + 1):
            if r not in claimed and norm_code(ws.cell(row=r, column=c_sap).value) is None:
                return r
        return None

    def ensure_formulas(r, sec):
        """A reused row may have been cleared by an earlier run -- copy the
        formulas back from another row in the same section."""
        s = sections[sec]
        keyed_in = {c_sap, c_open, c_add, c_close, c_charged, c_price}   # counts / inputs: never copied
        for c in range(1, ws.max_column + 1):
            if c in keyed_in or ws.cell(row=r, column=c).value not in (None, ""):
                continue
            for d in range(s["body_start"], s["body_end"] + 1):
                v = ws.cell(row=d, column=c).value
                # only real formulas that point at cells (not a typed sum like =72+96+72)
                if d != r and isinstance(v, str) and v.startswith("=") and re.search(r"[A-Z]+\$?\d", v):
                    origin = f"{L(c)}{d}"
                    ws.cell(row=r, column=c).value = Translator(v, origin=origin).translate_formula(f"{L(c)}{r}")
                    break

    for code, items in sold.items():
        if code in rows_by_code:
            continue
        qty = sum(ln["qty"] for ln, _ in items)
        if qty <= 0:
            continue
        ln0 = items[0][0]
        if not sections:
            not_placed.append((ln0, code))
            continue
        sec = pick_section(ln0, code)
        if sec is None:
            not_placed.append((ln0, code))
            continue
        r = free_row(sec)
        if r is None:
            section_full.append((ln0, code, sections[sec]["label"]))
            continue
        claimed.add(r)
        ws.cell(row=r, column=c_sap).value = code
        ensure_formulas(r, sec)
        ws.row_dimensions[r].hidden = False
        rows_by_code[code].append(r)
        added.append((r, ln0, code, sections[sec]["label"] or "top section"))

    # --- serve basis per row, and formula patches -------------------------
    patched, unit_warnings = [], []
    row_unit = {}
    for code, rows in rows_by_code.items():
        rec = mm60.by_code.get(code) or {}
        rule = rules.get(str(rec.get("group", "")).upper())
        bottle = size_ml(rec.get("desc"))
        for r in rows:
            cost_f = ws.cell(row=r, column=c_cost).value if c_cost else None
            m = re.search(r"/\s*(\d+(?:\.\d+)?)\s*$", cost_f) if isinstance(cost_f, str) else None
            if m:
                n = float(m.group(1))
                if rule and bottle and rule[2] == "Serves" and abs(bottle / rule[1] - n) > 0.01:
                    unit_warnings.append(f"Row {r} {rec.get('desc')}: Cost divides by {fmt_num(n)}, but a "
                                         f"{fmt_num(bottle)}ml bottle at {fmt_num(rule[1])}ml holds "
                                         f"{fmt_num(bottle / rule[1])} serves -- check the Cost formula.")
                row_unit[r] = (n, False)
            elif rule and rule[2] == "Serves" and bottle and abs(bottle / rule[1] - round(bottle / rule[1])) < 1e-9 \
                    and bottle / rule[1] >= 10:
                row_unit[r] = (bottle / rule[1], True)
            else:
                row_unit[r] = (1.0, False)

            n, patch_cost = row_unit[r]
            if n == 1:
                continue
            notes = []
            if patch_cost:
                if isinstance(cost_f, str) and cost_f.startswith("="):
                    ws.cell(row=r, column=c_cost).value = f"{cost_f}/{fmt_num(n)}"
                    notes.append(f"Cost per {fmt_num(rule[1])}ml serve (/{fmt_num(n)})")
                else:
                    unit_warnings.append(f"Row {r} {rec.get('desc')}: Cost is not a formula, left per bottle.")
                    row_unit[r] = (1.0, False)
                    continue
            
            if notes:
                patched.append((r, rec.get("desc") or code, "; ".join(notes)))

    # --- write Charged on Doc and Selling Price ---------------------------
    written, duplicates, overwrote, kept_manual, conversions = [], [], [], [], []
    price_kept = []
    revenue_written = 0.0
    lines_written = set()
    for code, rows in rows_by_code.items():
        items = sold.get(code, [])
        first = rows[0]
        n_row = row_unit[first][0]
        if len(rows) > 1:
            duplicates.append((code, rows))
        if not items:
            cell = ws.cell(row=first, column=c_charged)
            if cell.value in (None, ""):
                cell.value = 0
            elif to_float(cell.value) not in (None, 0.0):
                kept_manual.append((first, mm60.desc(code) or code, cell.value))
            for r in rows[1:]:
                if ws.cell(row=r, column=c_charged).value in (None, ""):
                    ws.cell(row=r, column=c_charged).value = 0
            continue

        charged, sales, unconvertible = 0.0, 0.0, []
        for ln, s in items:
            if s is None:
                unconvertible.append(ln)
                continue
            contrib = ln["qty"] * n_row / s
            charged += contrib
            sales += ln["sales"] if ln["sales"] else (ln["avg_price"] or 0) * ln["qty"]
            lines_written.add(id(ln))
            if abs(n_row / s - 1) > 1e-9:
                conversions.append((first, ln, s, n_row, contrib))
        for ln in unconvertible:
            ln["reason"] = (f"SAP {code} {mm60.desc(code)} -- sold by the serve but MM60 has no bottle size "
                            "to convert with; set 'Serves per Unit' in the item map")
        charged = round(charged, 4)
        if abs(charged - round(charged)) < 1e-9:
            charged = int(round(charged))

        cell = ws.cell(row=first, column=c_charged)
        old = cell.value
        if old not in (None, "") and to_float(old) not in (None, 0.0) and abs((to_float(old) or 0) - charged) > 1e-6:
            overwrote.append((first, mm60.desc(code) or code, old, charged))
        cell.value = charged
        for r in rows[1:]:
            ws.cell(row=r, column=c_charged).value = 0

        price_used = None
        if c_price:
            pcell = ws.cell(row=first, column=c_price)
            existing = to_float(pcell.value)
            calc = round(sales / charged, 2) if charged else None
            if existing in (None, 0.0):
                if calc is not None:
                    pcell.value = calc
                price_used = calc
            else:
                price_used = existing
                if calc and abs(existing - calc) > 0.01 * calc:
                    price_kept.append((first, mm60.desc(code) or code, existing, calc))
        if price_used and charged:
            revenue_written += price_used * charged
        written.append((first, mm60.desc(code) or code, charged, items, n_row, price_used))

    # --- checks ------------------------------------------------------------
    not_in_mm60, zero_price, cos_flags, misplaced = [], [], [], []
    for code, rows in rows_by_code.items():
        rec = mm60.by_code.get(code)
        if mm60.by_code and rec is None:
            for r in rows:
                ws.cell(row=r, column=c_sap).fill = YELLOW
                not_in_mm60.append((r, code))
            continue
        if rec and (rec["price"] in (None, 0.0)):
            zero_price.append((rows[0], rec["desc"], code))
    for r, desc, charged, items, n_row, price in written:
        rec = mm60.by_code.get(norm_code(ws.cell(row=r, column=c_sap).value)) or {}
        if charged and price and rec.get("price"):
            cos = (rec["price"] / n_row) / (price / VAT_DIVISOR)
            if cos < COS_FLAG_LOW or cos > COS_FLAG_HIGH:
                cos_flags.append((r, desc, cos))
        sec = section_of(r)
        if sec is not None and sections[sec]["label"]:
            fam = items[0][0]["family_group"]
            a, b = broad_class(fam), broad_class(sections[sec]["label"])
            if a and b and a != b:
                misplaced.append((r, desc, fam, sections[sec]["label"]))

    # --- counts vs sales: used but not sold, sold but not counted -------------
    def simple_value(v):
        """Number from a typed count or a typed sum like =72+96+72; None otherwise."""
        if v is None or v == "":
            return 0.0
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, str) and re.fullmatch(r"=\+?[\d.\s+\-]+", v):
            try:
                return float(sum(float(x) for x in re.findall(r"[+-]?\s*\d+(?:\.\d+)?", v[1:].replace(" ", ""))))
            except ValueError:
                return None
        return None

    used_not_sold, sold_not_counted = [], []
    if c_open and c_add and c_close:
        for code, rows in rows_by_code.items():
            vals = [[simple_value(ws.cell(row=r, column=c).value) for c in (c_open, c_add, c_close)] for r in rows]
            if any(v is None for row in vals for v in row):
                continue
            n = row_unit[rows[0]][0]
            usage = sum(o + a - c for o, a, c in vals)
            counted = any(any(v for v in row) for row in vals)
            charged = to_float(ws.cell(row=rows[0], column=c_charged).value) or 0.0
            unit = " serves" if n != 1 else ""
            if usage > 0 and charged == 0:
                used_not_sold.append(f"Row {rows[0]} {mm60.desc(code)}: {fmt_num(usage)}{unit} used, 0 on Micros")
            if charged > 0 and not counted:
                sold_not_counted.append(f"Row {rows[0]} {mm60.desc(code)}: {fmt_num(charged)}{unit} sold, no Open/Add/Close")

    # --- hide and clear rows that are still empty ---------------------------
    hidden = 0
    for s in sections:
        for r in range(s["body_start"], s["body_end"] + 1):
            if norm_code(ws.cell(row=r, column=c_sap).value) is None:
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).value = None
                ws.row_dimensions[r].hidden = True
                hidden += 1
            else:
                ws.row_dimensions[r].hidden = False

    # --- report -------------------------------------------------------------
    micros_bev = [ln for ln in lines if ln["how"] not in ("food", "ignored") and ln["qty"] > 0]
    micros_sales = sum(ln["sales"] for ln in micros_bev)
    missing = [ln for ln in micros_bev if id(ln) not in lines_written]
    diff = micros_sales - revenue_written

    title = ws.cell(row=2, column=3).value
    log(f"\n=== Sheet '{sheet_name}'" + (f"  ({title})" if title else "") + " ===")
    status = "OK" if abs(diff) < 1 else "CHECK -- see 'Not written' below"
    log(f"Control total: Micros beverage sales {money(micros_sales)} | written to sheet "
        f"{money(revenue_written)} | difference {money(diff)}  [{status}]")

    log(f"\nWritten to Charged on Doc ({len(written)} rows):")
    for r, desc, charged, items, n_row, price in sorted(written):
        unit = " serves" if n_row != 1 else ""
        src = ", ".join(sorted({f"'{ln['raw_name'].strip()}'" for ln, _ in items}))
        same = len(items) == 1 and normalize(items[0][0]["raw_name"]) == normalize(desc)
        log(f"   Row {r:<4} {str(desc)[:40]:<40} {fmt_num(charged):>8}{unit}" + ("" if same else f"   <- {src}"))

    via = defaultdict(list)
    for ln in lines:
        if ln["qty"] > 0 and ln["how"] in ("map", "size", "auto", "asked"):
            via[ln["how"]].append(ln)
    if via["asked"]:
        log(f"\nConfirmed by you on screen and saved to the item map ({len(via['asked'])}):")
        for ln in via["asked"]:
            log(f"   '{ln['raw_name'].strip()}' -> SAP {ln['sap']} {mm60.desc(ln['sap'])}")
    if via["size"]:
        log(f"\nSame name, size label differs (auto, {len(via['size'])}):")
        for ln in via["size"]:
            log(f"   '{ln['raw_name'].strip()}' -> {mm60.desc(ln['sap'])}")
    if via["auto"]:
        log(f"\nAuto-matched, CHECK THESE ({len(via['auto'])}) -- add them to the item map once confirmed:")
        for ln in via["auto"]:
            log(f"   '{ln['raw_name'].strip()}' -> SAP {ln['sap']} {ln.get('auto_from')}")
    if via["map"]:
        log(f"\nMatched through the item map: {len(via['map'])} Micros lines.")

    if conversions:
        log(f"\nServe conversions ({len(conversions)}):")
        for r, ln, s, n_row, contrib in conversions:
            what = (f"{fmt_num(ln['qty'])} serves = {fmt_num(contrib)} bottles" if n_row == 1
                    else f"{fmt_num(ln['qty'])} x {fmt_num(n_row / s)} = {fmt_num(contrib)} serves")
            log(f"   Row {r} '{ln['raw_name'].strip()}': {what}")
    if patched:
        log(f"\nRows switched to serves so Variance compares like with like ({len(patched)}):")
        for r, desc, note in sorted(patched):
            log(f"   Row {r:<4} {str(desc)[:40]:<40} {note}")
    if unit_warnings:
        log(f"\nServe warnings ({len(unit_warnings)}):")
        for w in unit_warnings:
            log(f"   {w}")

    if duplicates:
        log(f"\nSame SAP code on more than one row ({len(duplicates)}) -- Micros qty written to the first row, "
            "0 on the others so sales are not counted twice. Delete the extra row if it is a stray copy:")
        for code, rows in duplicates:
            log(f"   SAP {code} {mm60.desc(code)}: rows {', '.join(map(str, rows))}")
    if added:
        log(f"\nAdded as new rows ({len(added)}) -- no Open/Add/Close counts on these yet, so until they are "
            "filled in the Variance shows the full qty sold:")
        for r, ln, code, label in added:
            note = "" if normalize(ln["raw_name"]) == normalize(mm60.desc(code)) else f"  <- '{ln['raw_name'].strip()}'"
            log(f"   Row {r} [{label}]: SAP {code} {mm60.desc(code)}{note}")

    if missing:
        log(f"\nNot written -- sold on Micros but not on the sheet ({len(missing)}, {money(sum(ln['sales'] for ln in missing))}):")
        placed_codes = {code for code, rows in rows_by_code.items()}
        for ln in missing:
            reason = ln.get("reason")
            if not reason:
                if ln["sap"] is None:
                    sug = ln.get("suggest")
                    reason = "no SAP match" + (f" (closest: {sug['code']} {sug['desc']})" if sug else "")
                elif any(x[0] is ln for x in not_placed):
                    reason = f"SAP {ln['sap']} {mm60.desc(ln['sap'])} -- could not tell which section it belongs in"
                elif any(x[0] is ln for x in section_full):
                    reason = f"SAP {ln['sap']} -- its section has no free row left"
                elif ln["sap"] not in placed_codes:
                    reason = f"SAP {ln['sap']} not placed"
                else:
                    reason = "not written"
            log(f"   '{ln['raw_name'].strip()}'  {fmt_num(ln['qty'])} sold, {money(ln['sales'])}  -- {reason}")
    if ignored:
        log(f"\nIgnored (marked 'not a stock item' in the item map, {len(ignored)}): "
            + ", ".join(f"'{ln['raw_name'].strip()}'" for ln in ignored))
    if food:
        log(f"\nExcluded (Major Group = Food, {len(food)}): " + ", ".join(f"'{ln['raw_name'].strip()}'" for ln in food))

    if overwrote:
        log(f"\nCharged on Doc already had a different value and was replaced ({len(overwrote)}):")
        for r, desc, old, new in overwrote:
            log(f"   Row {r} {desc}: {old} -> {fmt_num(new)}")
    if kept_manual:
        log(f"\nNothing on Micros, existing Charged on Doc kept ({len(kept_manual)}):")
        for r, desc, v in kept_manual:
            log(f"   Row {r} {desc}: {v}")
    if price_kept:
        log(f"\nSelling Price already filled and kept, differs from Micros ({len(price_kept)}):")
        for r, desc, old, calc in price_kept:
            log(f"   Row {r} {desc}: sheet {money(old)} vs Micros {money(calc)}")

    if not_in_mm60:
        log(f"\nSAP codes NOT found in MM60 ({len(not_in_mm60)}) -- highlighted yellow:")
        for r, code in not_in_mm60:
            log(f"   Row {r}: SAP {code}")
    if zero_price:
        log(f"\nMM60 price is R0 ({len(zero_price)}) -- cost on these rows will show R0:")
        for r, desc, code in zero_price:
            log(f"   Row {r}: SAP {code} {desc}")
    if cos_flags:
        log(f"\nCoS outside {COS_FLAG_LOW:.0%}-{COS_FLAG_HIGH:.0%} (cost vs selling price ex VAT) -- "
            f"often a wrong unit or price ({len(cos_flags)}):")
        for r, desc, cos in cos_flags:
            log(f"   Row {r} {desc}: {cos:.1%}")
    if misplaced:
        log(f"\nSitting in a different category section than Micros puts it ({len(misplaced)}) -- "
            "its revenue rolls into that section's subtotal:")
        for r, desc, fam, label in misplaced:
            log(f"   Row {r} {desc}: Micros '{fam}', sheet section '{label}'")
    if used_not_sold or sold_not_counted:
        log("\nCheck with the bar -- similar quantities on both lists usually mean an item was rung up "
            "under another item's POS button:")
        for t in used_not_sold:
            log(f"   {t}")
        for t in sold_not_counted:
            log(f"   {t}")
    if hidden:
        log(f"\nHid and cleared {hidden} empty template rows.")


# ----------------------------------------------------------------------
# File handling
# ----------------------------------------------------------------------
def looks_like_micros_xlsx(path):
    if os.path.splitext(path)[1].lower() != ".xlsx":
        return False
    try:
        load_micros_xlsx(path)
        return True
    except Exception:
        return False


def looks_like_cellar_xlsx(path):
    if os.path.splitext(path)[1].lower() != ".xlsx":
        return False
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        for ws in wb.worksheets:
            try:
                find_cellar_header(ws)
                return True
            except ValueError:
                continue
    except Exception:
        return False
    return False


def classify_files(paths):
    cellar_paths, micros_paths = [], []
    for p in paths:
        if not os.path.exists(p):
            sys.exit(f"File not found: '{p}'")
        if os.path.splitext(p)[1].lower() == ".pdf":
            micros_paths.append(p)
            continue
        is_cellar, is_micros = looks_like_cellar_xlsx(p), looks_like_micros_xlsx(p)
        if is_cellar and not is_micros:
            cellar_paths.append(p)
        elif is_micros and not is_cellar:
            micros_paths.append(p)
        elif is_cellar and is_micros:
            sys.exit(f"Could not tell if '{p}' is the Cellar workbook or a Micros report. Pass it with --cellar/--micros.")
        else:
            sys.exit(f"Could not recognise '{p}' as a Cellar workbook or a Micros sales report. "
                     "Is it open in Excel, or the wrong file?")
    if len(cellar_paths) != 1:
        sys.exit(f"Expected exactly one Cellar Consumption workbook among the files given, found {len(cellar_paths)}.")
    if not micros_paths:
        sys.exit("No Micros sales report found among the files given.")
    return cellar_paths[0], micros_paths


def resolve_sheet(wb, name):
    """Finds the real sheet for a typed name, ignoring case, spaces and the
    . , / - separators in dates (12.10.2026 = 12,10,2026 = '12-10-2026 ')."""
    if name in wb.sheetnames:
        return name
    flat = lambda t: re.sub(r"[^a-z0-9]", "", str(t).lower())
    hits = [n for n in wb.sheetnames if flat(n) == flat(name)]
    d = parse_sheet_name_as_date(name)
    if not hits and d:
        hits = [n for n in wb.sheetnames if parse_sheet_name_as_date(n) == d]
    return hits[0] if len(hits) == 1 else None


def selectable_sheets(wb):
    return [ws.title for ws in wb.worksheets if ws.sheet_state == "visible" and ws.title.lower() != "mm60"]


def dated_sheets(wb):
    out = {}
    for s in selectable_sheets(wb):
        d = parse_sheet_name_as_date(s)
        if d:
            out[d] = s
    return out


def auto_pair_micros_to_sheets(wb, micros_paths):
    by_date = dated_sheets(wb)
    pairs, unmatched = [], []
    for mp in micros_paths:
        bdate = get_micros_business_date(mp)
        sheet = by_date.get(bdate) if bdate else None
        (pairs.append((sheet, mp)) if sheet else unmatched.append((mp, bdate)))
    if unmatched:
        log("\nCould not auto-match these Micros reports to a dated sheet:")
        for mp, bdate in unmatched:
            reason = f"its Business Date ({bdate}) has no matching dated sheet" if bdate else "no Business Date could be read"
            log(f"   {mp}  -- {reason}")
        log("   Available dated sheets: " + ", ".join(f"{s} ({d})" for d, s in sorted(by_date.items())))
    return pairs


def save_workbook(wb, out_path):
    try:
        wb.calculation.fullCalcOnLoad = True       # Excel recalculates every formula on open
    except Exception:
        pass
    try:
        wb.save(out_path)
        return out_path
    except PermissionError:
        alt = os.path.splitext(out_path)[0] + datetime.datetime.now().strftime("_%H%M%S") + ".xlsx"
        wb.save(alt)
        log(f"\n'{os.path.basename(out_path)}' is open in Excel, so the result was saved as '{os.path.basename(alt)}' instead.")
        return alt


def main():
    ap = argparse.ArgumentParser(description="Populate Cellar Consumption 'Charged on Doc' from a Micros sales report.")
    ap.add_argument("--files", nargs="+", help="Every file for this run, any order: the Cellar workbook and one or more Micros reports.")
    ap.add_argument("--micros", nargs="+", help="Micros sales report(s), old or upgraded format (.xlsx or .pdf).")
    ap.add_argument("--cellar", help="Cellar Consumption workbook (.xlsx).")
    ap.add_argument("--sheet", help="One sheet to update.")
    ap.add_argument("--sheets", help="Comma-separated sheets, e.g. 'Bar 1,Bar 2'. With several Micros files, paired in order.")
    ap.add_argument("--auto", action="store_true", help="Pick a sheet without asking if none matches the Business Date.")
    ap.add_argument("--map", help=f"Item map workbook (default: {DEFAULT_MAP_NAME} next to this script).")
    ap.add_argument("--no-prompt", action="store_true", help="Never ask on screen; use only strong matches and list them for checking.")
    ap.add_argument("--out", help="Output path (default: <cellar>_UPDATED.xlsx).")
    args = ap.parse_args()
    try:
        sys.stdout.reconfigure(errors="replace")   # odd characters in SAP names never crash the run
    except Exception:
        pass

    log(f"Cellar Consumption updater {VERSION}")
    if args.files:
        if args.micros or args.cellar:
            sys.exit("Use either --files on its own, or --micros/--cellar -- not both.")
        cellar_path, micros_paths = classify_files(args.files)
    else:
        if not args.micros or not args.cellar:
            sys.exit("Need either --files, or both --micros and --cellar.")
        cellar_path, micros_paths = args.cellar, list(args.micros)
        if len(micros_paths) == 1 and os.path.splitext(micros_paths[0])[1].lower() != ".pdf":
            if (looks_like_cellar_xlsx(micros_paths[0]) and not looks_like_micros_xlsx(micros_paths[0])
                    and looks_like_micros_xlsx(cellar_path) and not looks_like_cellar_xlsx(cellar_path)):
                log("Note: the two files look swapped -- swapping them automatically.")
                micros_paths[0], cellar_path = cellar_path, micros_paths[0]

    interactive = is_interactive() and not args.no_prompt
    map_path = args.map or os.path.join(os.path.dirname(os.path.abspath(__file__)), DEFAULT_MAP_NAME)
    item_map = load_item_map(map_path)
    rules = load_house_rules(map_path)
    log(f"Item map: {map_path} ({len(item_map)} entries)" if os.path.exists(map_path)
        else f"Item map: not found at {map_path} -- one will be created when you confirm a match.")

    wb = openpyxl.load_workbook(cellar_path)
    mm60 = MM60(wb["MM60"] if "MM60" in wb.sheetnames else None)
    if not mm60.by_code:
        log("WARNING: no MM60 sheet found in the Cellar workbook -- SAP codes cannot be looked up.")

    # --- decide which Micros report goes to which sheet -----------------
    jobs = []
    if len(micros_paths) > 1 and not args.sheets:
        pairs = auto_pair_micros_to_sheets(wb, micros_paths)
        if not pairs:
            sys.exit("Could not auto-match any of the given Micros reports to a dated sheet in this workbook.")
        log("Auto-matched Micros reports to sheets by date:")
        for s, mp in pairs:
            log(f"   {os.path.basename(mp)}  ->  '{s}'")
        jobs = [(mp, [s]) for s, mp in pairs]
    elif len(micros_paths) > 1:
        names = [s.strip() for s in args.sheets.split(",") if s.strip()]
        if len(names) != len(micros_paths):
            sys.exit(f"--sheets has {len(names)} sheet(s) but there are {len(micros_paths)} Micros file(s) -- they pair one-to-one.")
        jobs = [(mp, [s]) for s, mp in zip(names, micros_paths)]
    else:
        mp = micros_paths[0]
        if args.sheets:
            names = [s.strip() for s in args.sheets.split(",") if s.strip()]
        elif args.sheet:
            names = [args.sheet]
        else:
            bdate = get_micros_business_date(mp)
            match = dated_sheets(wb).get(bdate) if bdate else None
            candidates = selectable_sheets(wb)
            if match:
                names = [match]
                log(f"Sheet '{match}' picked automatically from the report's Business Date ({bdate:%d %b %Y}).")
            elif len(candidates) == 1:
                names = candidates
                log(f"Only one sheet to update: '{candidates[0]}'.")
            elif args.auto or not interactive:
                date_like = [s for s in candidates if parse_sheet_name_as_date(s)]
                names = [date_like[0] if date_like else candidates[0]]
                log(f"Auto-selected sheet: {names[0]}")
            else:
                log("Which sheet should be updated?")
                for i, s in enumerate(candidates, start=1):
                    log(f"   {i}) {s}")
                typed = input("Sheet number(s) or name(s), comma-separated: ").strip()
                names = []
                for part in [p.strip() for p in typed.split(",") if p.strip()]:
                    names.append(candidates[int(part) - 1] if part.isdigit() and 1 <= int(part) <= len(candidates) else part)
        jobs = [(mp, names)]

    fixed_jobs, missing = [], []
    for mp, names in jobs:
        real = []
        for n in names:
            r = resolve_sheet(wb, n)
            if r is None:
                missing.append(n)
            elif r != n:
                log(f"Sheet '{n}' matched to '{r}'.")
            real.append(r or n)
        fixed_jobs.append((mp, real))
    jobs = fixed_jobs
    if missing:
        shown = [f"'{n}'" + ("" if wb[n].sheet_state == "visible" else " (hidden)") for n in wb.sheetnames]
        sys.exit(f"Sheet(s) not found: {', '.join(missing)}\nSheets in {os.path.basename(cellar_path)}: {', '.join(shown)}")

    # --- run ---------------------------------------------------------------
    all_new = []
    for mp, names in jobs:
        log(f"\nReading Micros report: {os.path.basename(mp)}")
        lines = load_micros(mp)
        log(f"  -> {len(lines)} lines, {sum(1 for ln in lines if ln['qty'] > 0)} with sales "
            f"({getattr(lines, 'fmt', '') or 'Micros'} format).")
        for note in getattr(lines, "notes", ()):
            log(f"  {note}")
        sheet_codes = set()
        for s in names:
            try:
                hr, cols = find_cellar_header(wb[s])
                c_sap = cols.get("sap code") or cols.get("sap no")
                for r in range(hr + 1, wb[s].max_row + 1):
                    code = norm_code(wb[s].cell(row=r, column=c_sap).value) if c_sap else None
                    if code is not None:
                        sheet_codes.add(code)
            except ValueError:
                pass
        new_entries = resolve_lines(lines, mm60, sheet_codes, item_map, interactive)
        for e in new_entries:
            item_map[normalize(e["name"])] = {"sap": e["sap"], "serves": None}
        all_new.extend(new_entries)
        for s in names:
            process_sheet(wb, s, lines, mm60, rules)

    if all_new:
        existed = os.path.exists(map_path)
        if append_to_item_map(map_path, all_new, mm60):
            log(f"\nSaved {len(all_new)} confirmed match(es) to the item map"
                + (" (dated backup of the old map taken first)." if existed else f" (new file: {map_path})."))
        else:
            log(f"\nCould not save to the item map -- is {os.path.basename(map_path)} open in Excel? "
                "Close it and add these rows by hand:")
            for e in all_new:
                log(f"   {e['name']}  ->  {e['sap']}")

    out_path = args.out or (os.path.splitext(cellar_path)[0] + "_UPDATED.xlsx")
    saved = save_workbook(wb, out_path)
    log(f"\nSaved: {saved}")
    log_path = os.path.splitext(saved)[0].replace("_UPDATED", "") + "_UPDATE_LOG.txt"
    if log.save(log_path):
        print(f"Log:   {log_path}")


if __name__ == "__main__":
    main()
