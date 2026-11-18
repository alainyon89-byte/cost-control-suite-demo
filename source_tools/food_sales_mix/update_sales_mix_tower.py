"""
update_sales_mix.py

Automates updating the Tower Wing food Sales Mix workbook from a monthly
Micros (Simphony) sales-mix export.

WHAT IT DOES
    1. Reads the "Business Dates" cell in the Micros export to work out
       which month/year to update.
    2. Finds that month's column in the sales mix workbook (sheet matches
       the year, e.g. "2026"). If the column doesn't exist, it stops and
       tells you instead of guessing where to put it.
    3. Matches each Micros "Item" to a row in the sales mix by name:
         - exact match (case/whitespace-insensitive)        -> written in
         - known alias (item_aliases_tower.json)                  -> written in
         - known new item (item_additions_tower.json, "pending")  -> row created,
           then written in
         - known exclusion (item_exclusions_tower.json)           -> skipped, no
           flag (you've already confirmed it's not the same dish)
         - fuzzy match >= AUTO_MATCH_THRESHOLD               -> written in,
           logged as an auto-match so you can sanity check it
         - fuzzy match >= REVIEW_THRESHOLD but below         -> NOT written,
           listed in the report for you to decide
         - no reasonable match                                -> skipped,
           listed in the report in case it's a new item worth adding
    4. Backs up the original workbook before writing anything.
    5. Writes every value as a plain number (matching the existing sheet's
       convention - no formulas), using the number format already present
       in the target cell.
    6. Produces a run report (printed + saved as a .txt file next to the
       workbook) showing what was updated, auto-matched, flagged, and
       skipped.

CONFIG FILES (same folder as this script) — Tower Wing variant
    item_aliases_tower.json     POS name -> Sales Mix row label to write into.
                           Add an entry here whenever a flagged item turns
                           out to be the same dish under a different name.
    item_exclusions_tower.json  POS names confirmed NOT the same as any similar
                           mix row. Stops them being flagged again.
    item_additions_tower.json   POS names confirmed as genuinely new dishes, with
                           the category they belong under. The script adds
                           a new row for each on first run, then treats it
                           like any other mix item afterwards.

    This is the Tower Wing copy of update_sales_mix.py — it reads/writes
    the _tower config files only, so it can share a folder with the Main Hotel
    version (which uses the plain item_aliases.json / item_exclusions.json /
    item_additions.json names) without either one touching the other's config.

USAGE
    python update_sales_mix.py <sales_mix.xlsx> <micros_export.xlsx>

    Or just run run_update.bat, which prompts for the Micros export file.
"""

import sys
import json
import shutil
import difflib
import datetime
from pathlib import Path

import openpyxl
import pandas as pd

AUTO_MATCH_THRESHOLD = 0.90   # auto-write and log
REVIEW_THRESHOLD = 0.60       # below this, don't even mention it (too weak to be useful)

MONTH_NAME_ALIASES = {
    "febuary": "february",
    "januray": "january",
    "decmber": "december",
}

SCRIPT_DIR = Path(__file__).resolve().parent
ALIASES_PATH = SCRIPT_DIR / "item_aliases_tower.json"
EXCLUSIONS_PATH = SCRIPT_DIR / "item_exclusions_tower.json"
ADDITIONS_PATH = SCRIPT_DIR / "item_additions_tower.json"

CATEGORY_HEADER_FILL = "FFE2EFDA"  # matches existing category row shading


def normalize_month(name):
    name = str(name).strip().lower()
    return MONTH_NAME_ALIASES.get(name, name)


def normalize_item(name):
    return " ".join(str(name).strip().lower().split())


def load_json_config(path, default):
    """Returns (data, found, error). found=False means the file did not
    exist at this path - the caller should surface that loudly rather than
    silently proceeding with an empty default."""
    if not path.exists():
        return default, False, None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f), True, None
    except json.JSONDecodeError as e:
        return default, True, f"'{path.name}' exists but is not valid JSON: {e}"


def save_json_config(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def find_pos_header_row(ws, max_scan=15):
    for r in range(1, max_scan + 1):
        if str(ws.cell(row=r, column=1).value).strip().lower() == "item":
            return r
    raise ValueError("Could not find the 'Item' header row in the Micros export "
                      "within the first 15 rows. The export format may have changed.")


def get_business_month_year(ws):
    for r in range(1, 10):
        label = ws.cell(row=r, column=1).value
        if label and str(label).strip().lower() == "business dates":
            raw = ws.cell(row=r, column=2).value
            break
    else:
        raise ValueError("Could not find a 'Business Dates' row in the Micros export.")

    start_str = str(raw).split("-")[0].strip()
    dt = datetime.datetime.strptime(start_str, "%m/%d/%Y")
    return dt.strftime("%B").lower(), dt.year


def load_pos_data(pos_path):
    wb = openpyxl.load_workbook(pos_path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    month, year = get_business_month_year(ws)
    header_row = find_pos_header_row(ws)

    df = pd.read_excel(pos_path, sheet_name=wb.sheetnames[0], skiprows=header_row - 1)
    df.columns = [str(c).strip() for c in df.columns]

    if "Item" not in df.columns or "Qty Sold" not in df.columns:
        raise ValueError(f"Expected 'Item' and 'Qty Sold' columns, found: {df.columns.tolist()}")

    df = df[df["Item"].notna()]
    df = df[df["Item"].astype(str).str.strip() != "Total Item Sales:"]

    df["Item"] = df["Item"].astype(str).str.strip()
    summed = df.groupby("Item", as_index=False)["Qty Sold"].sum()

    return month, year, summed


def find_year_sheet(wb, year):
    year_str = str(year)
    if year_str in wb.sheetnames:
        return wb[year_str]
    raise ValueError(
        f"No sheet named '{year_str}' exists in the sales mix workbook. "
        f"Available sheets: {wb.sheetnames}. Create the year sheet first, "
        f"or tell me the correct sheet to use."
    )


def find_month_column(ws, month_name, year):
    target_month = normalize_month(month_name)
    for c in range(1, ws.max_column + 1):
        y = ws.cell(row=1, column=c).value
        m = ws.cell(row=2, column=c).value
        if y is None or m is None:
            continue
        if str(y).strip() == str(year) and normalize_month(m) == target_month:
            return c
    return None


def build_mix_name_index(ws):
    """Returns norm_name -> (row_num, display_label) for every item row."""
    index = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v is None:
            continue
        label = str(v).strip()
        if label.lower() == "sales mix for month of":
            continue
        index[normalize_item(label)] = (r, label)
    return index


def find_category_row(ws, category_name):
    target = normalize_item(category_name)
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v and normalize_item(v) == target:
            return r
    return None


def find_next_category_row(ws, after_row):
    """Finds the row number of the next category header after a given row.
    Category rows are bold with the green fill and have no data in month
    columns."""
    for r in range(after_row + 1, ws.max_row + 1):
        cell = ws.cell(row=r, column=1)
        if cell.value and cell.font and cell.font.bold:
            fill = cell.fill.fgColor.rgb if cell.fill and cell.fill.fgColor else None
            if fill == CATEGORY_HEADER_FILL:
                return r
    return ws.max_row + 1


def insert_new_item_row(ws, category_name, item_label):
    """Inserts a new row for item_label at the end of category_name's
    block, copying formatting from the last existing item row in that
    block. Returns the new row number."""
    cat_row = find_category_row(ws, category_name)
    if cat_row is None:
        raise ValueError(f"Category '{category_name}' not found in the sheet. "
                          f"Check item_additions.json for a typo, or add the "
                          f"category header to the workbook first.")

    next_cat_row = find_next_category_row(ws, cat_row)
    insert_at = next_cat_row  # push new row in right before the next category
    template_row = insert_at - 1  # last existing item row in this category, for style copy

    ws.insert_rows(insert_at)

    for c in range(1, ws.max_column + 1):
        src = ws.cell(row=template_row, column=c)
        dst = ws.cell(row=insert_at, column=c)
        dst.font = src.font.copy()
        dst.fill = src.fill.copy()
        dst.border = src.border.copy()
        dst.alignment = src.alignment.copy()
        dst.number_format = src.number_format

    ws.cell(row=insert_at, column=1).value = item_label
    return insert_at


def match_items(pos_df, mix_index, aliases, exclusions):
    """Returns exact, auto, flagged, skipped.
    Each match entry: (pos_item, row_num, mix_label, qty, similarity)."""
    exact = []
    auto = []
    flagged = []
    skipped = []

    mix_norms = list(mix_index.keys())
    exclusions_norm = {normalize_item(x) for x in exclusions}
    aliases_norm = {normalize_item(k): v for k, v in aliases.items()}

    for _, row in pos_df.iterrows():
        pos_item = row["Item"]
        qty = row["Qty Sold"]
        norm = normalize_item(pos_item)

        if norm in exclusions_norm:
            continue  # confirmed not-the-same, don't even list it

        if norm in aliases_norm:
            target_norm = normalize_item(aliases_norm[norm])
            if target_norm in mix_index:
                row_num, label = mix_index[target_norm]
                exact.append((pos_item, row_num, label, qty, 1.0))
                continue
            flagged.append((pos_item, aliases_norm[norm] + " (ALIAS TARGET NOT FOUND)", qty, 1.0))
            continue

        if norm in mix_index:
            row_num, label = mix_index[norm]
            exact.append((pos_item, row_num, label, qty, 1.0))
            continue

        best_norm, best_ratio = None, 0.0
        for cand in mix_norms:
            ratio = difflib.SequenceMatcher(None, norm, cand).ratio()
            if ratio > best_ratio:
                best_norm, best_ratio = cand, ratio

        if best_ratio >= AUTO_MATCH_THRESHOLD:
            row_num, label = mix_index[best_norm]
            auto.append((pos_item, row_num, label, qty, best_ratio))
        elif best_ratio >= REVIEW_THRESHOLD:
            flagged.append((pos_item, mix_index[best_norm][1], qty, best_ratio))
        else:
            skipped.append((pos_item, qty))

    return exact, auto, flagged, skipped


def write_values(ws, month_col, matches):
    for pos_item, row_num, mix_label, qty, ratio in matches:
        cell = ws.cell(row=row_num, column=month_col)
        cell.value = int(qty) if float(qty).is_integer() else float(qty)


def apply_pending_additions(ws, additions_config, pos_df):
    """For each pending addition not already present in the sheet, insert
    a new row. Returns updated additions_config (pending cleared of what
    was just added), a list of (pos_name, mix_name, category, row_num)
    added this run, and a list of (pos_name, mix_name, row_num) that were
    already present and therefore skipped."""
    pending = additions_config.get("pending", [])
    added_this_run = []
    already_present = []

    # Resolved entries (added just now, or already present) are dropped from
    # pending; anything not touched this run (e.g. a bad category name that
    # raised before insertion) is deliberately left in place rather than lost.
    resolved_norms = set()

    for entry in pending:
        pos_name = entry["pos_name"]
        mix_name = entry.get("mix_name", pos_name)
        category = entry["category"]
        norm_mix = normalize_item(mix_name)

        mix_index = build_mix_name_index(ws)
        if norm_mix in mix_index:
            existing_row = mix_index[norm_mix][0]
            already_present.append((pos_name, mix_name, existing_row))
            resolved_norms.add(norm_mix)
            continue

        row_num = insert_new_item_row(ws, category, mix_name)
        added_this_run.append((pos_name, mix_name, category, row_num))
        resolved_norms.add(norm_mix)

    additions_config["pending"] = [
        e for e in pending
        if normalize_item(e.get("mix_name", e["pos_name"])) not in resolved_norms
    ]
    return additions_config, added_this_run, already_present


def main():
    if len(sys.argv) != 3:
        print("Usage: python update_sales_mix.py <sales_mix.xlsx> <micros_export.xlsx>")
        sys.exit(1)

    mix_path = Path(sys.argv[1])
    pos_path = Path(sys.argv[2])

    if not mix_path.exists():
        print(f"ERROR: Sales mix file not found: {mix_path}")
        sys.exit(1)
    if not pos_path.exists():
        print(f"ERROR: Micros export file not found: {pos_path}")
        sys.exit(1)

    backup_dir = mix_path.parent / "backups"
    backup_dir.mkdir(exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"{mix_path.stem}_backup_{timestamp}{mix_path.suffix}"
    shutil.copy2(mix_path, backup_path)

    aliases_raw, aliases_found, aliases_err = load_json_config(ALIASES_PATH, {})
    aliases = {k: v for k, v in aliases_raw.items() if not k.startswith("_")}
    exclusions_cfg, exclusions_found, exclusions_err = load_json_config(EXCLUSIONS_PATH, {"items": []})
    exclusions = exclusions_cfg.get("items", [])
    additions_cfg, additions_found, additions_err = load_json_config(ADDITIONS_PATH, {"pending": []})

    print("--- Config files loaded (Tower Wing) ---")
    print(f"  Script folder: {SCRIPT_DIR}")
    print(f"  item_aliases_tower.json:    {'FOUND, ' + str(len(aliases)) + ' entries' if aliases_found else 'NOT FOUND - using none'}"
          + (f"  [ERROR: {aliases_err}]" if aliases_err else ""))
    print(f"  item_exclusions_tower.json: {'FOUND, ' + str(len(exclusions)) + ' entries' if exclusions_found else 'NOT FOUND - using none'}"
          + (f"  [ERROR: {exclusions_err}]" if exclusions_err else ""))
    print(f"  item_additions_tower.json:  {'FOUND, ' + str(len(additions_cfg.get('pending', []))) + ' pending' if additions_found else 'NOT FOUND - using none'}"
          + (f"  [ERROR: {additions_err}]" if additions_err else ""))
    if not (aliases_found and exclusions_found and additions_found):
        print("  WARNING: one or more Tower Wing config files were not found next to this script.")
        print(f"  Make sure item_aliases_tower.json, item_exclusions_tower.json, and")
        print(f"  item_additions_tower.json are saved in: {SCRIPT_DIR}")
    print()

    month, year, pos_df = load_pos_data(pos_path)
    print(f"Detected period: {month.title()} {year}")

    wb = openpyxl.load_workbook(mix_path, data_only=False)
    ws = find_year_sheet(wb, year)

    month_col = find_month_column(ws, month, year)
    if month_col is None:
        print(f"ERROR: No column found for {month.title()} {year} in sheet '{year}'.")
        print("Add the month column to the workbook first (matching the existing "
              "year/month header format), then re-run.")
        sys.exit(1)

    col_letter = openpyxl.utils.get_column_letter(month_col)
    print(f"Target column: {col_letter} (sheet '{year}')")

    additions_cfg, added_this_run, already_present = apply_pending_additions(ws, additions_cfg, pos_df)
    if added_this_run:
        month_col = find_month_column(ws, month, year)

    existing_vals = [ws.cell(row=r, column=month_col).value for r in range(1, ws.max_row + 1)]
    existing_count = sum(1 for v in existing_vals if v is not None)
    if existing_count > 0:
        print(f"NOTE: {existing_count} cells already have data in column {col_letter}. "
              f"They will be overwritten only where a match is found for that row.")

    mix_index = build_mix_name_index(ws)
    exact, auto, flagged, skipped = match_items(pos_df, mix_index, aliases, exclusions)

    write_values(ws, month_col, exact)
    write_values(ws, month_col, auto)

    wb.save(mix_path)
    save_json_config(ADDITIONS_PATH, additions_cfg)

    report_lines = []
    report_lines.append(f"Sales Mix Update Report - {month.title()} {year}")
    report_lines.append(f"Run at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report_lines.append(f"Backup saved to: {backup_path}")
    report_lines.append("")

    if added_this_run:
        report_lines.append(f"--- NEW ITEM ROWS ADDED ({len(added_this_run)}) ---")
        for pos_name, mix_name, category, row_num in added_this_run:
            report_lines.append(f"  Added row {row_num} '{mix_name}' under '{category}' "
                                 f"(from POS item '{pos_name}')")
        report_lines.append("")

    if already_present:
        report_lines.append(f"--- PENDING ADDITIONS ALREADY IN THE SHEET ({len(already_present)}) ---")
        report_lines.append("  These are in item_additions_tower.json but already exist as rows - no action taken.")
        for pos_name, mix_name, existing_row in already_present:
            report_lines.append(f"  '{mix_name}' already at row {existing_row} (from POS item '{pos_name}')")
        report_lines.append("")

    report_lines.append(f"Exact matches written (incl. aliases & new items): {len(exact)}")
    report_lines.append(f"Auto-matched (>= {int(AUTO_MATCH_THRESHOLD*100)}% similarity) written: {len(auto)}")
    report_lines.append(f"Flagged for manual review (not written): {len(flagged)}")
    report_lines.append(f"Skipped, no reasonable match (out of scope): {len(skipped)}")
    report_lines.append("")

    if auto:
        report_lines.append("--- AUTO-MATCHED (please spot-check) ---")
        for pos_item, row_num, mix_label, qty, ratio in auto:
            report_lines.append(f"  POS '{pos_item}' -> Mix row {row_num} '{mix_label}' "
                                 f"(qty={qty}, similarity={ratio:.0%})")
        report_lines.append("")

    if flagged:
        report_lines.append("--- FLAGGED FOR YOUR REVIEW (not written to sheet) ---")
        report_lines.append("  Reply 'join' / 'not the same' / 'add' per line like last time, "
                             "and I'll update item_aliases_tower.json / item_exclusions_tower.json / "
                             "item_additions_tower.json accordingly.")
        for pos_item, mix_label, qty, ratio in flagged:
            report_lines.append(f"  POS '{pos_item}' (qty={qty}) looks like it might be "
                                 f"'{mix_label}' (similarity={ratio:.0%})")
        report_lines.append("")

    if skipped:
        report_lines.append(f"--- SKIPPED, no match found ({len(skipped)} items) ---")
        for pos_item, qty in skipped:
            report_lines.append(f"  '{pos_item}' (qty={qty})")
        report_lines.append("")

    report_text = "\n".join(report_lines)
    print("\n" + report_text)

    report_path = mix_path.parent / f"update_report_{month}_{year}_{timestamp}.txt"
    report_path.write_text(report_text, encoding="utf-8")
    print(f"\nFull report saved to: {report_path}")


if __name__ == "__main__":
    main()
