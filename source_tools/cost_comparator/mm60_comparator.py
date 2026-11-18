# ============================================================
#  MM60 COST COMPARATIVE TOOL
#  © 2026 Alain Yon. All rights reserved.
#  Version: 1.0 | Last updated: July 8, 2026
#
#  Compares unit prices between two SAP MM60 price file exports
#
#  HOW TO RUN (do NOT double-click the file):
#  1. Update the CONFIG section below
#  2. Run:  Run_Cost_Comparator.bat
#  3. Output file appears in the same folder
# ============================================================

# -- CONFIG -- PROMPTED AT STARTUP ------

def get_file_inputs():
    """Prompt user for file paths and labels."""
    print("\n" + "="*60)
    print("  MM60 Cost Comparator -- File Input")
    print("="*60)
    print("  NOTE: Filenames are CASE-SENSITIVE on Mac/Linux")
    print("        (Windows is case-insensitive)")
    
    while True:
        file1 = input("\nEnter FIRST MM60 file (e.g., 'MM60_January.xlsx'): ").strip()
        if file1 and file1.endswith(('.xlsx', '.xls')):
            break
        print("  [ERROR] Enter a valid Excel filename (.xlsx or .xls)")
    
    while True:
        file2 = input("Enter SECOND MM60 file (e.g., 'MM60_June.xlsx'): ").strip()
        if file2 and file2.endswith(('.xlsx', '.xls')):
            break
        print("  [ERROR] Enter a valid Excel filename (.xlsx or .xls)")
    
    label1 = input("Enter label for first file (e.g., 'January 2026'): ").strip()
    if not label1:
        label1 = file1.replace('.xlsx', '').replace('.xls', '')
    
    label2 = input("Enter label for second file (e.g., 'June 2026'): ").strip()
    if not label2:
        label2 = file2.replace('.xlsx', '').replace('.xls', '')
    
    return file1, file2, label1, label2

# ------

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, sys

PROPERTY_NAME = os.environ.get("CCS_PROPERTY_NAME", "Demo Hotel")   # set by the app from config/app_settings.json

# ── CATEGORY MAP ─────────────────────────────────────────────────
CATEGORY_MAP = {
    "BEV000100": "Aperitifs & Vermouths",
    "BEV000200": "Beer (Bottled & Can)",
    "BEV000300": "Brandy & Potstill",
    "BEV000400": "Cider & Alcoholic Mixers",
    "BEV000500": "Cognac",
    "BEV000600": "Draught Equipment & Gas",
    "BEV000700": "Liqueurs & Shooters",
    "BEV000800": "Soft Drinks & Energy",
    "BEV000900": "Still & Sparkling Water",
    "BEV001000": "Juices (Beverage)",
    "BEV001100": "Bar Cordials & Syrups",
    "BEV001200": "Port & Dessert Wine",
    "BEV001300": "Red Wine",
    "BEV001310": "Red Wine (Other)",
    "BEV001400": "Rose Wine",
    "BEV001500": "Rum",
    "BEV001600": "Sherry & Fortified Wine",
    "BEV001700": "Sparkling Wine & Champagne",
    "BEV001800": "Whisky & Whiskey",
    "BEV001900": "Gin & Vodka",
    "BEV002000": "White Wine",
    "BEV002010": "White Wine (Other)",
    "FOO000100": "Bread & Fresh Bakery",
    "FOO000200": "Butter & Spreads",
    "FOO000400": "Dairy, Cheese & Ice Cream",
    "FOO000500": "Dry Baking & Patisserie",
    "FOO000600": "Sauces, Condiments & Jams",
    "FOO000700": "Dry Goods, Grains & Cereals",
    "FOO000800": "Herbs, Spices & Seasonings",
    "FOO000900": "Pasta",
    "FOO001000": "Snacks & Confectionery",
    "FOO001100": "Canned Fish, Fruit & Veg",
    "FOO001200": "Frozen Goods",
    "FOO001300": "Fresh Seafood",
    "FOO001400": "Pastry & Frozen Snacks",
    "FOO001500": "Fruit Platters & Garnish",
    "FOO001600": "Halaal Sausages & Deli",
    "FOO001700": "Coffee, Tea & Hot Beverages",
    "FOO001800": "Juices (Food-side)",
    "FOO001900": "Beef, Veal & Game",
    "FOO002000": "Nuts & Seeds",
    "FOO002100": "Poultry & Duck",
    "FOO002200": "Prepared Salads & Platters",
    "FOO002300": "Processed Meats & Deli",
    "FOO002400": "Tea",
    "FOO002500": "Fresh Vegetables & Herbs",
    "CLE000100": "Cleaning & Chemical Supplies",
    "COV000100": "Covers & Disposable Service",
    "CRO000100": "Crockery",
    "CUT000100": "Cutlery",
    "FBL000100": "F&B Linen & Table Covers",
    "GLA000100": "Glassware",
    "GUE000100": "Guest Supplies",
    "HOL000100": "Hollowware & Buffet Equip",
    "MAI000800": "Maintenance Supplies",
    "MAI001000": "Maintenance Supplies",
    "MAI001500": "Maintenance Supplies",
    "PAP000100": "Paper, Packaging & Stationery",
    "RLI000100": "Room Linen",
    "ROO000100": "Room Supplies & Equipment",
    "STA000100": "Stationery",
    "TOB000100": "Tobacco & Cigars",
    "UTE000100": "Kitchen Utensils & Equipment",
}

# ── COLOURS ───────────────────────────────────────────────────────
DARK_BLUE  = "1F3864"
MID_BLUE   = "2E75B6"
LIGHT_BLUE = "D6E4F0"
ORANGE     = "F4B942"
WHITE      = "FFFFFF"
GREY       = "F2F2F2"
TOTAL_BG   = "FFF2CC"
RED_BG     = "FCE4D6"
GREEN_BG   = "E2EFDA"
LINK_GRN   = "008000"
INPUT_BLUE = "0000FF"
HYPERLINK  = "1F3864"

def tb():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def cs(ws, ref, val=None, bold=False, fc="000000", sz=10, bg=None,
       ah="left", av="center", wrap=False, fmt=None, italic=False, border=True):
    cell = ws[ref] if isinstance(ref, str) else ref
    if val is not None:
        cell.value = val
    cell.font = Font(name="Arial", bold=bold, color=fc, size=sz, italic=italic)
    if bg:
        cell.fill = fill(bg)
    cell.alignment = Alignment(horizontal=ah, vertical=av, wrap_text=wrap)
    if border:
        cell.border = tb()
    if fmt:
        cell.number_format = fmt
    return cell

# ── LOAD & CLEAN ──────────────────────────────────────────────────
def load_mm60(filepath, label):
    print(f"  Loading {label} from: {filepath}")
    if not os.path.exists(filepath):
        print(f"\n  ERROR: File not found — {filepath}")
        print(f"  Make sure the file is in the same folder as this script.")
        print(f"  Script is running from: {os.getcwd()}")
        input("\n  Press Enter to close...")
        sys.exit(1)
    df = pd.read_excel(filepath)
    df = df[~df['Material Description'].str.strip().str.upper().eq('DO NOT USE')].copy()
    df['Price'] = pd.to_numeric(df['Price'], errors='coerce').fillna(0)
    df['Category'] = df['Material Group'].map(CATEGORY_MAP).fillna(df['Material Group'])
    print(f"  {label}: {len(df):,} items loaded")
    return df

# ── DATA DUMP SHEET ───────────────────────────────────────────────
def build_data_sheet(wb, sheet_title, df, month_label, tab_color):
    ws = wb.create_sheet(sheet_title)
    ws.sheet_properties.tabColor = tab_color
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:F1")
    cs(ws, "A1", f"{PROPERTY_NAME.upper()}  |  SAP MM60 PRICE FILE  |  {month_label.upper()}",
       bold=True, sz=12, bg=DARK_BLUE, fc=WHITE, ah="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:F2")
    cs(ws, "A2",
       "Source: SAP MM60 export  |  Do Not Use items excluded  |  Do not edit — re-import from SAP each month",
       italic=True, sz=9, fc="595959", ah="left", border=False)
    ws.row_dimensions[2].height = 13

    cs(ws, "A3", "Period:", bold=True, bg=GREY, sz=9)
    ws.merge_cells("B3:C3")
    cs(ws, "B3", month_label, fc=INPUT_BLUE, bold=True, bg=LIGHT_BLUE, sz=9)
    ws.row_dimensions[3].height = 15

    hdrs   = ["Material #", "Material Group", "Category", "Description", "UOM", "Unit Price (R)"]
    widths = [12, 12, 30, 52, 7, 16]
    for c, (h, w) in enumerate(zip(hdrs, widths), 1):
        ltr = get_column_letter(c)
        ws.column_dimensions[ltr].width = w
        cs(ws, f"{ltr}5", h, bold=True, bg=MID_BLUE, fc=WHITE, ah="center", sz=9)
    ws.row_dimensions[5].height = 16

    row = 6
    current_cat = None
    for _, r in df.sort_values(['Category', 'Material Description']).iterrows():
        cat = r['Category']
        if cat != current_cat:
            ws.merge_cells(f"A{row}:F{row}")
            cs(ws, f"A{row}", f"  \u25b8  {cat}  [{r['Material Group']}]",
               bold=True, sz=9, bg=LIGHT_BLUE, fc=MID_BLUE, ah="left")
            ws.row_dimensions[row].height = 13
            row += 1
            current_cat = cat
        bg_r = GREY if row % 2 == 0 else WHITE
        cs(ws, f"A{row}", str(r['Material']), sz=9, bg=bg_r, ah="center")
        cs(ws, f"B{row}", r['Material Group'], sz=9, bg=bg_r, ah="center")
        cs(ws, f"C{row}", cat, sz=9, bg=bg_r)
        cs(ws, f"D{row}", r['Material Description'], sz=9, bg=bg_r)
        cs(ws, f"E{row}", r['Base Unit of Measure'], sz=9, bg=bg_r, ah="center")
        cs(ws, f"F{row}", r['Price'], sz=9, bg=bg_r, ah="right", fmt='R #,##0.00')
        ws.row_dimensions[row].height = 12
        row += 1

    ws.merge_cells(f"A{row}:E{row}")
    cs(ws, f"A{row}", "TOTAL CATALOGUE UNIT PRICE SUM", bold=True, bg=TOTAL_BG, sz=9, ah="right")
    cs(ws, f"F{row}", f"=SUM(F6:F{row-1})", bold=True, bg=TOTAL_BG, fmt='R #,##0.00', ah="right")
    ws.row_dimensions[row].height = 15
    ws.freeze_panes = "A6"
    return ws, row

# ── SUMMARY SHEET ─────────────────────────────────────────────────
def build_summary(wb, df1, df2, label1, label2, total_row1, total_row2):
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_properties.tabColor = "2E75B6"
    ws.sheet_view.showGridLines = False

    for i, w in enumerate([38,16,16,16,14,3,14,16,16,14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Title
    ws.merge_cells("A1:J1")
    cs(ws, "A1",
       f"{PROPERTY_NAME.upper()}  |  UNIT PRICE COMPARATIVE  |  {label1.upper()} vs {label2.upper()}",
       bold=True, sz=13, bg=DARK_BLUE, fc=WHITE, ah="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:J2")
    cs(ws, "A2",
       "Source: SAP MM60 unit price exports  |  Variance = Month 2 minus Month 1  |  "
       "No quantities applied  |  Click any category name to jump to item detail",
       italic=True, sz=9, fc="595959", ah="left", border=False)
    ws.row_dimensions[2].height = 13

    cs(ws, "A4", f"{label1}:", bold=True, bg=GREY, sz=9)
    ws.merge_cells("B4:D4")
    cs(ws, "B4", "='Month 1 Data'!B3", fc=LINK_GRN, bold=True, bg=LIGHT_BLUE, sz=9)
    cs(ws, "A5", f"{label2}:", bold=True, bg=GREY, sz=9)
    ws.merge_cells("B5:D5")
    cs(ws, "B5", "='Month 2 Data'!B3", fc=LINK_GRN, bold=True, bg=LIGHT_BLUE, sz=9)
    ws.row_dimensions[4].height = 14
    ws.row_dimensions[5].height = 14

    # ── Section 1: Overall basket ────────────────────────────────
    ws.row_dimensions[7].height = 18
    ws.merge_cells("A7:E7")
    cs(ws, "A7", "OVERALL BASKET COMPARISON", bold=True, sz=11,
       bg=MID_BLUE, fc=WHITE, ah="center")

    for c, h in enumerate([
        "", f"{label1} Total (R)", f"{label2} Total (R)", "Variance (R)", "Variance (%)"
    ], 1):
        cs(ws, f"{get_column_letter(c)}8", h, bold=True, bg=DARK_BLUE,
           fc=WHITE, ah="center", sz=9)
    ws.row_dimensions[8].height = 15

    cs(ws, "A9", "Total Catalogue Unit Price Sum", bold=True, sz=10)
    cs(ws, "B9", f"='Month 1 Data'!F{total_row1}", fc=LINK_GRN,
       fmt='R #,##0.00', ah="right")
    cs(ws, "C9", f"='Month 2 Data'!F{total_row2}", fc=LINK_GRN,
       fmt='R #,##0.00', ah="right")
    cs(ws, "D9", "=C9-B9", bold=True,
       fmt='R #,##0.00;(R #,##0.00);"-"', ah="right")
    cs(ws, "E9", '=IFERROR((C9-B9)/B9,"-")', bold=True, fmt="0.0%", ah="center")
    ws.row_dimensions[9].height = 18

    ws.merge_cells("A10:E10")
    cs(ws, "A10",
       '=IF(D9>0,"⚠  Unit prices INCREASED overall — review category drivers below",'
       'IF(D9<0,"✔  Unit prices DECREASED overall","→  No overall change"))',
       italic=True, sz=9, border=False, bg=WHITE)
    ws.row_dimensions[10].height = 13

    # ── Section 2: Category breakdown ───────────────────────────
    ws.row_dimensions[12].height = 18
    ws.merge_cells("A12:E12")
    cs(ws, "A12", "UNIT PRICE COMPARISON BY SAP MATERIAL GROUP",
       bold=True, sz=11, bg=MID_BLUE, fc=WHITE, ah="center")

    for c, h in enumerate([
        "Category  [click to jump to items ▼]",
        f"{label1} Total (R)", f"{label2} Total (R)",
        "Variance (R)", "Variance (%)"
    ], 1):
        cs(ws, f"{get_column_letter(c)}13", h, bold=True, bg=DARK_BLUE,
           fc=WHITE, ah="center", sz=9)
    ws.row_dimensions[13].height = 15

    # Write category rows — store row numbers for hyperlink pass later
    cat_summary_rows = {}   # code -> summary row number
    row = 14

    for i, (code, label) in enumerate(sorted(CATEGORY_MAP.items())):
        bg_r = LIGHT_BLUE if i % 2 == 0 else WHITE

        # Column A — placeholder text for now, hyperlink added after item section
        c1 = ws.cell(row=row, column=1)
        c1.value = f"{label}  [{code}]"
        c1.font = Font(name="Arial", sz=9, color="000000")
        c1.fill = fill(bg_r); c1.border = tb()
        c1.alignment = Alignment(horizontal="left", vertical="center")

        for col, sname, maxr in [
            (2, "Month 1 Data", total_row1-1),
            (3, "Month 2 Data", total_row2-1)
        ]:
            cx = ws.cell(row=row, column=col)
            cx.value = f'=SUMIF(\'{sname}\'!B6:B{maxr},"{code}",\'{sname}\'!F6:F{maxr})'
            cx.font = Font(name="Arial", color=LINK_GRN, sz=9)
            cx.fill = fill(bg_r); cx.border = tb()
            cx.number_format = 'R #,##0.00'
            cx.alignment = Alignment(horizontal="right", vertical="center")

        d = ws.cell(row=row, column=4)
        d.value = f"=C{row}-B{row}"
        d.font = Font(name="Arial", sz=9)
        d.fill = fill(bg_r); d.border = tb()
        d.number_format = 'R #,##0.00;(R #,##0.00);"-"'
        d.alignment = Alignment(horizontal="right", vertical="center")

        e = ws.cell(row=row, column=5)
        e.value = f'=IFERROR((C{row}-B{row})/B{row},"-")'
        e.font = Font(name="Arial", sz=9)
        e.fill = fill(bg_r); e.border = tb()
        e.number_format = "0.0%"
        e.alignment = Alignment(horizontal="center", vertical="center")

        cat_summary_rows[code] = row
        ws.row_dimensions[row].height = 13
        row += 1

    # Grand total
    cs(ws, f"A{row}", "GRAND TOTAL", bold=True, bg=TOTAL_BG, sz=10, ah="right")
    cs(ws, f"B{row}", f"=SUM(B14:B{row-1})", bold=True, bg=TOTAL_BG,
       fmt='R #,##0.00', ah="right")
    cs(ws, f"C{row}", f"=SUM(C14:C{row-1})", bold=True, bg=TOTAL_BG,
       fmt='R #,##0.00', ah="right")
    cs(ws, f"D{row}", f"=C{row}-B{row}", bold=True, bg=TOTAL_BG,
       fmt='R #,##0.00;(R #,##0.00);"-"', ah="right")
    cs(ws, f"E{row}", f'=IFERROR((C{row}-B{row})/B{row},"-")', bold=True,
       bg=TOTAL_BG, fmt="0.0%", ah="center")
    ws.row_dimensions[row].height = 18
    grand_row = row
    row += 2

    # ── Section 3: Top movers ────────────────────────────────────
    ws.merge_cells(f"A{row}:E{row}")
    cs(ws, f"A{row}", "TOP COST MOVERS BY CATEGORY",
       bold=True, sz=11, bg=ORANGE, fc=DARK_BLUE, ah="center")
    ws.row_dimensions[row].height = 18

    fc_r = 14
    lc_r = grand_row - 1

    movers = [
        ("Highest R increase",
         f'=IFERROR(INDEX(A{fc_r}:A{lc_r},MATCH(MAX(D{fc_r}:D{lc_r}),D{fc_r}:D{lc_r},0)),"N/A")',
         f'=IFERROR(MAX(D{fc_r}:D{lc_r}),"-")', 'R #,##0.00', RED_BG),
        ("Largest R decrease",
         f'=IFERROR(INDEX(A{fc_r}:A{lc_r},MATCH(MIN(D{fc_r}:D{lc_r}),D{fc_r}:D{lc_r},0)),"N/A")',
         f'=IFERROR(MIN(D{fc_r}:D{lc_r}),"-")', 'R #,##0.00', GREEN_BG),
        ("Highest % increase",
         f'=IFERROR(INDEX(A{fc_r}:A{lc_r},MATCH(MAX(E{fc_r}:E{lc_r}),E{fc_r}:E{lc_r},0)),"N/A")',
         f'=IFERROR(MAX(E{fc_r}:E{lc_r}),"-")', '0.0%', RED_BG),
    ]
    for j, (lbl, cat_f, val_f, val_fmt, bg_r) in enumerate(movers):
        r = row + 1 + j
        cs(ws, f"A{r}", lbl, bold=True, bg=GREY, sz=9)
        cx = ws.cell(row=r, column=2)
        cx.value = cat_f
        cx.font = Font(name="Arial", bold=True, sz=9)
        cx.fill = fill(LIGHT_BLUE); cx.border = tb()
        cx.alignment = Alignment(horizontal="left")
        ws.merge_cells(f"B{r}:C{r}")
        dx = ws.cell(row=r, column=4)
        dx.value = val_f
        dx.font = Font(name="Arial", bold=True, sz=9)
        dx.fill = fill(bg_r); dx.border = tb()
        dx.number_format = val_fmt
        dx.alignment = Alignment(horizontal="right")
        ws.row_dimensions[r].height = 14
    row += len(movers) + 3

    # ── Section 4: Item-level comparison ────────────────────────
    ws.merge_cells(f"A{row}:H{row}")
    cs(ws, f"A{row}",
       "ITEM-LEVEL PRICE COMPARISON — ALL ITEMS (matched on Material number)",
       bold=True, sz=11, bg=MID_BLUE, fc=WHITE, ah="center")
    ws.row_dimensions[row].height = 18

    item_hdrs  = ["Material #", "Category", "Description", "UOM",
                  f"{label1} (R)", f"{label2} (R)", "Variance (R)", "Variance (%)"]
    item_widths = [12, 30, 48, 7, 16, 16, 16, 14]
    for col, (h, w) in enumerate(zip(item_hdrs, item_widths), 1):
        ltr = get_column_letter(col)
        ws.column_dimensions[ltr].width = max(
            float(ws.column_dimensions[ltr].width or 0), w)
        cs(ws, f"{ltr}{row+1}", h, bold=True, bg=DARK_BLUE, fc=WHITE, ah="center", sz=9)
    ws.row_dimensions[row+1].height = 15

    # Build merged item dataframe
    # Combine descriptive columns from BOTH months so items that only exist
    # in one month (new or discontinued SKUs) still get a correct Category,
    # Description, Material Group and UOM instead of falling through to "Unknown".
    desc_cols = ['Material', 'Material Description', 'Material Group',
                 'Category', 'Base Unit of Measure']
    combined_desc = pd.concat([df1[desc_cols], df2[desc_cols]], ignore_index=True)
    combined_desc = combined_desc.drop_duplicates(subset='Material', keep='first')

    merged = pd.merge(
        combined_desc,
        df1[['Material', 'Price']].rename(columns={'Price': 'P1'}),
        on='Material', how='outer'
    )
    merged = pd.merge(
        merged,
        df2[['Material', 'Price']].rename(columns={'Price': 'P2'}),
        on='Material', how='outer'
    ).fillna({'P1': 0, 'P2': 0})
    merged['Material Description'] = merged['Material Description'].fillna('')
    merged['Category'] = merged['Category'].fillna(
        merged['Material Group'].map(CATEGORY_MAP).fillna(
            merged['Material Group'].fillna('Unknown')))
    merged['Base Unit of Measure'] = merged['Base Unit of Measure'].fillna('')
    merged['Var_R']   = merged['P2'] - merged['P1']
    merged['Var_Pct'] = np.where(
        merged['P1'] > 0,
        (merged['P2'] - merged['P1']) / merged['P1'],
        np.nan
    )
    merged = merged.sort_values(['Category', 'Material Description'])

    item_start  = row + 2
    irow        = item_start
    current_cat = None

    # cat_label -> first item row (for hyperlink targets)
    # We key by the display label matching what's in cat_summary_rows
    cat_item_rows = {}   # display label string -> row number of category header

    for _, r in merged.iterrows():
        cat = r['Category']

        if cat != current_cat:
            # Record the row this category lands on for hyperlinking
            # Find the code for this category to match cat_summary_rows
            cat_item_rows[cat] = irow

            ws.merge_cells(f"A{irow}:H{irow}")
            cell_hdr = ws.cell(row=irow, column=1)
            cell_hdr.value = f"  \u25b8  {cat}"
            cell_hdr.font  = Font(name="Arial", bold=True, sz=9,
                                  color=MID_BLUE)
            cell_hdr.fill  = fill(LIGHT_BLUE)
            cell_hdr.border = tb()
            cell_hdr.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[irow].height = 13
            irow += 1
            current_cat = cat

        bg_r = GREY if irow % 2 == 0 else WHITE
        cs(ws, f"A{irow}", str(int(r['Material'])), sz=9, bg=bg_r, ah="center")
        cs(ws, f"B{irow}", cat, sz=9, bg=bg_r)
        cs(ws, f"C{irow}", r['Material Description'], sz=9, bg=bg_r)
        cs(ws, f"D{irow}", r['Base Unit of Measure'], sz=9, bg=bg_r, ah="center")
        cs(ws, f"E{irow}", r['P1'], sz=9, bg=bg_r, ah="right", fmt='R #,##0.00')
        cs(ws, f"F{irow}", r['P2'], sz=9, bg=bg_r, ah="right", fmt='R #,##0.00')
        cs(ws, f"G{irow}", r['Var_R'], sz=9, bg=bg_r, ah="right",
           fmt='R #,##0.00;(R #,##0.00);"-"')

        pct_cell = ws.cell(row=irow, column=8)
        if pd.notna(r['Var_Pct']):
            pct_cell.value = r['Var_Pct']
            pct_cell.number_format = "0.0%"
        else:
            pct_cell.value = "-"
        pct_cell.font  = Font(name="Arial", sz=9)
        pct_cell.fill  = fill(bg_r)
        pct_cell.border = tb()
        pct_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[irow].height = 12
        irow += 1

    # Totals row
    ws.merge_cells(f"A{irow}:D{irow}")
    cs(ws, f"A{irow}", "TOTALS", bold=True, bg=TOTAL_BG, sz=9, ah="right")
    cs(ws, f"E{irow}", f"=SUM(E{item_start}:E{irow-1})", bold=True,
       bg=TOTAL_BG, fmt='R #,##0.00', ah="right")
    cs(ws, f"F{irow}", f"=SUM(F{item_start}:F{irow-1})", bold=True,
       bg=TOTAL_BG, fmt='R #,##0.00', ah="right")
    cs(ws, f"G{irow}", f"=F{irow}-E{irow}", bold=True, bg=TOTAL_BG,
       fmt='R #,##0.00;(R #,##0.00);"-"', ah="right")
    cs(ws, f"H{irow}", f'=IFERROR((F{irow}-E{irow})/E{irow},"-")', bold=True,
       bg=TOTAL_BG, fmt="0.0%", ah="center")
    ws.row_dimensions[irow].height = 15

    # ── ADD HYPERLINKS: category summary → item section ──────────
    # Now we know where every category landed in the item section,
    # go back and update each category summary row with a hyperlink
    print("  Adding category hyperlinks...")
    linked = 0
    for code, label in sorted(CATEGORY_MAP.items()):
        summary_row = cat_summary_rows.get(code)
        target_row  = cat_item_rows.get(label)   # label is the display name

        if summary_row and target_row:
            cell = ws.cell(row=summary_row, column=1)
            # Internal hyperlink: #SheetName!CellRef
            cell.hyperlink = f"#Summary!A{target_row}"
            cell.font = Font(
                name="Arial", sz=9,
                color=DARK_BLUE,        # dark blue — visible on both light/white rows
                bold=False,
                underline="single"
            )
            linked += 1

    print(f"  {linked} hyperlinks added")

    ws.freeze_panes = "A4"
    print(f"  Summary complete — {irow - item_start:,} item rows written")

# ── MAIN ──────────────────────────────────────────────────────────
def main():
    global FILE_MONTH1, FILE_MONTH2, LABEL_MONTH1, LABEL_MONTH2
    
    FILE_MONTH1, FILE_MONTH2, LABEL_MONTH1, LABEL_MONTH2 = get_file_inputs()
    
    print("\n" + "="*55)
    print(f"  {PROPERTY_NAME.upper()} — MM60 COST COMPARATIVE TOOL")
    print("="*55)

    df1 = load_mm60(FILE_MONTH1, LABEL_MONTH1)
    df2 = load_mm60(FILE_MONTH2, LABEL_MONTH2)

    wb = Workbook()

    print(f"\n  Building data sheets...")
    _, tr1 = build_data_sheet(wb, "Month 1 Data", df1, LABEL_MONTH1, "70AD47")
    _, tr2 = build_data_sheet(wb, "Month 2 Data", df2, LABEL_MONTH2, "ED7D31")

    print(f"  Building summary...")
    build_summary(wb, df1, df2, LABEL_MONTH1, LABEL_MONTH2, tr1, tr2)

    wb.move_sheet("Summary", offset=-wb.sheetnames.index("Summary"))

    # ── Add Attribution Sheet ─────────────────────────────────────────────────
    if 'Attribution' in wb.sheetnames:
        del wb['Attribution']
    
    att_sheet = wb.create_sheet('Attribution', 0)
    from datetime import datetime
    att_sheet['A1'] = '© 2026 Alain Yon. All rights reserved.'
    att_sheet['A2'] = ''
    att_sheet['A3'] = 'Script: mm60_comparator.py v1.0'
    att_sheet['A4'] = f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'
    att_sheet['A5'] = ''
    att_sheet['A6'] = 'This workbook and associated analysis are the intellectual property'
    att_sheet['A7'] = 'of Alain Yon.'
    
    att_sheet['A1'].font = Font(bold=True, size=12)
    att_sheet.column_dimensions['A'].width = 70
    
    # ── Add Print Footer to All Sheets ────────────────────────────────────────
    from openpyxl.worksheet.header_footer import HeaderFooter
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_margins.left = 0.5
        ws.page_margins.right = 0.5
        hf = HeaderFooter()
        hf.oddFooter.center.text = "© 2026 Alain Yon | mm60_comparator.py v1.0 | Page &P of &N"
        ws.header_footer = hf

    m1       = LABEL_MONTH1.replace(" ", "_")
    m2       = LABEL_MONTH2.replace(" ", "_")
    out_file = f"MM60_Comparative_{m1}_vs_{m2}.xlsx"

    wb.save(out_file)
    print(f"\n  ✔  Output saved: {out_file}")
    print(f"  Location: {os.path.join(os.getcwd(), out_file)}")
    print("="*55)
    input("\n  Press Enter to close...")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n  ERROR: {e}")
        import traceback
        traceback.print_exc()
        input("\n  Press Enter to close...")
