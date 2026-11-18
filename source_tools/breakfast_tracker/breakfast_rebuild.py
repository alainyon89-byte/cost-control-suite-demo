"""
=============================================================================
COST CONTROL SUITE  |  Breakfast Month-End Rebuild Script
© 2026 Alain Yon. All rights reserved.
Version: 2.0 | Last updated: July 8, 2026
=============================================================================
HOW TO USE EACH MONTH:
  1. Save your Mi20 SAP export as an .xlsx file in this folder
  2. Run:  Run_Breakfast_Report.bat
  3. Open Breakfast_MonthEnd_Tool.xlsx
  4. Manually enter Breakfast Revenue and Breakfast Covers in the KPI block

REQUIREMENTS (one-time install):
  pip install openpyxl pandas

FILES NEEDED IN THE SAME FOLDER:
  - breakfast_rebuild.py            (this script — stays permanently)
  - Mi20_kitchen.xlsx               (SAP export — swap each month)
  - Breakfast_Item_Master.xlsx      (item list — update if chef adds items)
  - Breakfast_MonthEnd_Tool.xlsx    (output — rebuilt each month)
=============================================================================
"""

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG  — update this line each month
# ─────────────────────────────────────────────────────────────────────────────
SAP_FILE    = "Mi20_kitchen.xlsx"    # SAP Mi20 export filename
# ─────────────────────────────────────────────────────────────────────────────

import os, sys
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

def get_month_label():
    """Prompt user for month label."""
    while True:
        label = input("\nEnter month label (e.g., 'June 2026'): ").strip()
        if label:
            return label
        print("  [ERROR] Month label cannot be empty. Try again.")

# ── Palette ───────────────────────────────────────────────────────────────────
DARK_TEAL   = "1F4E79"
MID_TEAL    = "2E75B6"
LIGHT_BLUE  = "D6E4F0"
WHITE       = "FFFFFF"
AMBER       = "FFD966"
LIGHT_AMBER = "FFF2CC"
RED_LIGHT   = "FCE4D6"
GREEN_LIGHT = "E2EFDA"
GREY_HDR    = "595959"

def _f(h): return PatternFill("solid", start_color=h, fgColor=h)
def _b(s="thin"):
    sd = Side(style=s); return Border(left=sd, right=sd, top=sd, bottom=sd)
def _ft(bold=False, colour="000000", size=10, italic=False):
    return Font(name="Arial", bold=bold, color=colour, size=size, italic=italic)
def _al(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def style(cell, bold=False, colour="000000", size=10, bg=None,
          h="center", v="center", wrap=False, num=None, italic=False, border=True):
    cell.font      = _ft(bold=bold, colour=colour, size=size, italic=italic)
    cell.alignment = _al(h=h, v=v, wrap=wrap)
    if bg:   cell.fill   = _f(bg)
    if num:  cell.number_format = num
    if border: cell.border = _b()

# ── Paths ─────────────────────────────────────────────────────────────────────
MONTH_LABEL = get_month_label()

script_dir  = os.path.dirname(os.path.abspath(__file__))
sap_path    = os.path.join(script_dir, SAP_FILE)
master_path = os.path.join(script_dir, "Breakfast_Item_Master.xlsx")
out_path    = os.path.join(script_dir, "Breakfast_MonthEnd_Tool.xlsx")

# ── Validate files ────────────────────────────────────────────────────────────
if not os.path.exists(sap_path):
    sys.exit(f"\n  ERROR: Cannot find SAP file:\n  {sap_path}\n"
             f"  Check SAP_FILE in the config and make sure the file is in the same folder.\n")

if not os.path.exists(master_path):
    sys.exit(f"\n  ERROR: Cannot find Breakfast_Item_Master.xlsx\n  {master_path}\n"
             f"  This file must be in the same folder as the script.\n")

# ── Load SAP file ─────────────────────────────────────────────────────────────
print(f"\n  Loading SAP file: {SAP_FILE}")
df_sap = pd.read_excel(sap_path, header=0)
num_cols = df_sap.shape[1]
print(f"  Columns detected: {num_cols}")

if num_cols >= 19:
    print("  Format: Full SAP PhysInv export (19 columns)")
    df_sap = df_sap.iloc[:, :12].copy()
elif num_cols >= 12:
    print(f"  Format: Mi20 kitchen export ({num_cols} columns — using first 12)")
    df_sap = df_sap.iloc[:, :12].copy()
else:
    sys.exit(f"\n  ERROR: Unrecognised format — {num_cols} columns.\n"
             f"  Expected at least 12 columns (Mi20 format).\n")

df_sap.columns = ['PhysInvDoc','Item','Material','Description','Plnt','Location',
                   'Book_qty','Count_qty','Diff_qty','Diff_value','Pos_diff','Neg_diff']
df_sap = df_sap[df_sap['Description'].notna() &
                pd.to_numeric(df_sap['Diff_value'], errors='coerce').notna()].copy()
df_sap['Diff_value']   = pd.to_numeric(df_sap['Diff_value'], errors='coerce').abs()
df_sap['Material_str'] = df_sap['Material'].astype(float).astype(int).astype(str)

# ── Load breakfast item master ────────────────────────────────────────────────
print(f"  Loading Breakfast_Item_Master.xlsx")
df_master = pd.read_excel(master_path, header=2)
df_master = df_master[pd.to_numeric(df_master['Material Code'], errors='coerce').notna()].copy()
df_master['Material_str'] = df_master['Material Code'].astype(float).astype(int).astype(str)
df_master['Breakfast_Pct'] = pd.to_numeric(df_master['Breakfast %'], errors='coerce').fillna(1.0)
bfast_codes = set(df_master['Material_str'])
split_map   = dict(zip(df_master['Material_str'], df_master['Breakfast_Pct']))
print(f"  Master items: {len(bfast_codes)}")

# ── Filter & consolidate ──────────────────────────────────────────────────────
data = df_sap[df_sap['Material_str'].isin(bfast_codes)].copy()
print(f"  Rows after breakfast filter: {len(data)}")

con = data.groupby(['Material_str','Description'], sort=False).agg(
    Total_cost = ('Diff_value', 'sum'),
).reset_index()
con['Breakfast_Pct']     = con['Material_str'].map(split_map).fillna(1.0)
con['Total_cost_adj']    = con['Total_cost'] * con['Breakfast_Pct']

# Show splits applied
splits = con[con['Breakfast_Pct'] < 1.0]
if len(splits) > 0:
    print(f"\n  Split % applied to {len(splits)} item(s):")
    for _, sr in splits.iterrows():
        print(f"    {sr['Description']}: {sr['Breakfast_Pct']*100:.0f}%"
              f" -> R{sr['Total_cost_adj']:,.2f} (was R{sr['Total_cost']:,.2f})")

# Remove 0% items
excluded = con[con['Breakfast_Pct'] == 0.0]
if len(excluded) > 0:
    print(f"\n  Excluded {len(excluded)} item(s) at 0%:")
    for _, er in excluded.iterrows():
        print(f"    {er['Description']}")
con = con[con['Breakfast_Pct'] > 0.0].copy()
con = con.sort_values('Total_cost_adj', ascending=False).reset_index(drop=True)

grand_total  = df_sap['Diff_value'].sum()
bfast_total  = con['Total_cost_adj'].sum()
bfast_pct    = bfast_total / grand_total if grand_total else 0

print(f"\n  {'─'*45}")
print(f"  Total food cost (Mi20):  R {grand_total:>12,.2f}")
print(f"  Breakfast cost:          R {bfast_total:>12,.2f}")
print(f"  Breakfast % of food:       {bfast_pct*100:>11.1f}%")
print(f"  {'─'*45}")
print(f"\n  REMINDER: Enter Breakfast Revenue + Covers manually in the KPI block after opening.\n")

# ── Preserve trend history ────────────────────────────────────────────────────
trend_history = {}
if os.path.exists(out_path):
    print(f"  Preserving Trend history from existing workbook...")
    try:
        ex_wb = openpyxl.load_workbook(out_path, data_only=True)
        if "Trend" in ex_wb.sheetnames:
            ex_t = ex_wb["Trend"]
            for col in range(2, 14):
                food_val  = ex_t.cell(5, col).value
                bfast_val = ex_t.cell(6, col).value
                rev_val   = ex_t.cell(8, col).value
                cov_val   = ex_t.cell(10, col).value
                if isinstance(food_val, (int, float)) and food_val:
                    trend_history[col] = {
                        'food':    food_val,
                        'bfast':   bfast_val,
                        'bfast_pct': ex_t.cell(7, col).value,
                        'revenue': rev_val,
                        'covers':  cov_val,
                    }
        ex_wb.close()
        print(f"  Months preserved: {len(trend_history)}")
    except Exception as e:
        print(f"  Warning: could not read trend history ({e}). Starting fresh.")

month_short = MONTH_LABEL.split()[0][:3]
try:    month_col = MONTHS.index(month_short) + 2
except: month_col = None

# ════════════════════════════════════════════════════════════════════════════
# BUILD WORKBOOK
# ════════════════════════════════════════════════════════════════════════════
wb = openpyxl.Workbook()

# ── SAP DUMP ──────────────────────────────────────────────────────────────────
ws_d = wb.active
ws_d.title = "SAP Dump"
ws_d.sheet_properties.tabColor = "2E75B6"

ws_d.merge_cells("A1:L1")
ws_d["A1"] = f"SAP KITCHEN DUMP (Mi20)  |  {MONTH_LABEL}  |  Paste full export here from row 3"
style(ws_d["A1"], bold=True, colour=WHITE, size=11, bg=DARK_TEAL, h="left")
ws_d.row_dimensions[1].height = 28
ws_d.merge_cells("A2:L2")
ws_d["A2"] = "Breakfast filter applied automatically on Summary tab using Breakfast_Item_Master.xlsx"
style(ws_d["A2"], colour=GREY_HDR, size=9, bg=LIGHT_AMBER, h="left", border=False)
ws_d.row_dimensions[2].height = 18

mi20_cols = ['PhysInvDoc','Item','Material','Material Description','Plnt','Location',
             'Book Qty','Qty Counted','Difference qty','Difference',
             'Pos. diff. value','neg. diff. value']
for ci, col in enumerate(mi20_cols, 1):
    style(ws_d.cell(3, ci, col), bold=True, colour=WHITE, size=9, bg=MID_TEAL)

for ri, row_vals in enumerate(df_sap.drop(columns='Material_str').values.tolist(), 4):
    for ci, val in enumerate(row_vals, 1):
        c = ws_d.cell(ri, ci, val)
        style(c, size=9, h="left" if ci==4 else "center",
              bg=LIGHT_BLUE if ri%2==0 else WHITE,
              num='R#,##0.00' if ci in [10,11,12] else None)

ws_d.freeze_panes = "A4"
for i, w in enumerate([12,6,10,36,6,8,9,9,12,12,14,14], 1):
    ws_d.column_dimensions[get_column_letter(i)].width = w

# ── SUMMARY ───────────────────────────────────────────────────────────────────
ws_s = wb.create_sheet("Summary")
ws_s.sheet_properties.tabColor = "375623"

ws_s.merge_cells("A1:E1")
ws_s["A1"] = "BREAKFAST COST SUMMARY  |  Month-End Report"
style(ws_s["A1"], bold=True, colour=WHITE, size=13, bg=DARK_TEAL, h="left")
ws_s.row_dimensions[1].height = 32

ws_s["A2"] = "Reporting Month:"
style(ws_s["A2"], bold=True, size=10, bg=LIGHT_AMBER, h="right", border=False)
ws_s["B2"] = MONTH_LABEL
style(ws_s["B2"], bold=True, colour="0000FF", size=11, bg=LIGHT_AMBER, h="left", border=False)
ws_s.row_dimensions[2].height = 22

ws_s.merge_cells("A3:E3")
ws_s["A3"] = (f"Filtered to {len(con)} breakfast items from full Mi20  |  "
              f"Ranked by cost (highest first)  |  "
              f"Adjust splits in Breakfast_Item_Master.xlsx if needed")
style(ws_s["A3"], colour=GREY_HDR, size=9, bg=LIGHT_AMBER, h="left", wrap=True, border=False)
ws_s.row_dimensions[3].height = 20

for ci, h in enumerate(["#","Material Code","Material Description",
                         "Cost (R)","% of Breakfast"], 1):
    style(ws_s.cell(4, ci, h), bold=True, colour=WHITE, size=10, bg=MID_TEAL)
ws_s.row_dimensions[4].height = 24

DR = 5
for ri, row in con.iterrows():
    r  = DR + ri
    bg = LIGHT_BLUE if ri%2==0 else WHITE
    ws_s.cell(r, 1, ri+1)
    style(ws_s.cell(r,1), colour=GREY_HDR, size=9, bg=bg)
    ws_s.cell(r, 2, int(row['Material_str']))
    style(ws_s.cell(r,2), size=10, bg=bg)
    ws_s.cell(r, 3, row['Description'])
    style(ws_s.cell(r,3), size=10, bg=bg, h="left")
    ws_s.cell(r, 4, round(row['Total_cost_adj'], 2))
    style(ws_s.cell(r,4), size=10, bg=bg, num='R#,##0.00')
    tr_ref = DR + len(con)
    ws_s.cell(r, 5, f"=D{r}/D${tr_ref}")
    style(ws_s.cell(r,5), size=10, bg=bg, num='0.0%')
    ws_s.row_dimensions[r].height = 18

# Total row
tr = DR + len(con)
ws_s.merge_cells(f"A{tr}:C{tr}")
ws_s.cell(tr, 1, f"TOTAL BREAKFAST COST  |  {MONTH_LABEL}")
style(ws_s.cell(tr,1), bold=True, colour=DARK_TEAL, size=11, bg=AMBER, h="right")
ws_s.cell(tr, 4, f"=SUM(D{DR}:D{tr-1})")
style(ws_s.cell(tr,4), bold=True, colour=DARK_TEAL, size=11, bg=AMBER, num='R#,##0.00')
ws_s.cell(tr, 5, "100.0%")
style(ws_s.cell(tr,5), bold=True, colour=DARK_TEAL, size=11, bg=AMBER, num='0.0%')
ws_s.row_dimensions[tr].height = 26

# KPI block
kb = tr + 2
ws_s.merge_cells(f"A{kb-1}:E{kb-1}")
ws_s.cell(kb-1, 1, f"KPI SUMMARY  |  {MONTH_LABEL}")
style(ws_s.cell(kb-1,1), bold=True, colour=WHITE, size=11, bg=DARK_TEAL, h="left")
ws_s.row_dimensions[kb-1].height = 22

kpi_rows = [
    ("Total Food Cost (Mi20)",        round(grand_total,2),              'R#,##0.00', "0000FF"),
    ("Breakfast Cost",                f"=D{tr}",                         'R#,##0.00', None),
    ("Breakfast % of Total Food",     f"=D{kb+1}/D{kb}",                 '0.0%',      None),
    (None, None, None, None),
    ("Breakfast Revenue (enter here)", 0,                                'R#,##0.00', "0000FF"),
    ("Breakfast Cost % of Revenue",   f"=IFERROR(D{kb+1}/D{kb+4},0)",   '0.0%',      None),
    (None, None, None, None),
    ("Breakfast Covers (enter here)", 0,                                 '#,##0',     "0000FF"),
    ("Cost per Cover (R)",            f"=IFERROR(D{kb+1}/D{kb+7},0)",   'R#,##0.00', None),
]
for i, (label, val, num, fc) in enumerate(kpi_rows):
    r = kb + i
    if not label:
        ws_s.row_dimensions[r].height = 6; continue
    is_input = "enter here" in label
    bg_row = LIGHT_AMBER if is_input else (LIGHT_BLUE if i%2==0 else WHITE)
    ws_s.merge_cells(f"A{r}:C{r}")
    style(ws_s.cell(r,1,label), bold=True, size=10, bg=bg_row, h="right", border=False)
    vc = ws_s.cell(r, 4, val)
    style(vc, bold=True, colour=fc or "000000", size=10, bg=bg_row, num=num,
          italic=is_input)
    ws_s.cell(r,5).fill = _f(bg_row)
    ws_s.row_dimensions[r].height = 20

# Reminder note
note_row = kb + len(kpi_rows) + 1
ws_s.merge_cells(f"A{note_row}:E{note_row}")
ws_s.cell(note_row, 1,
    "↑  Enter Breakfast Revenue and Covers manually each month — "
    "these come from your revenue report and POS covers, not SAP")
style(ws_s.cell(note_row,1), colour=GREY_HDR, size=9,
      bg=LIGHT_AMBER, h="left", italic=True, border=False)

for i, w in enumerate([4,13,42,16,14], 1):
    ws_s.column_dimensions[get_column_letter(i)].width = w
ws_s.freeze_panes = "A5"

# ── TREND ─────────────────────────────────────────────────────────────────────
ws_t = wb.create_sheet("Trend")
ws_t.sheet_properties.tabColor = "FFD966"

ws_t.merge_cells("A1:N1")
ws_t["A1"] = "BREAKFAST  |  Month-on-Month Trend"
style(ws_t["A1"], bold=True, colour=WHITE, size=13, bg=DARK_TEAL, h="left")
ws_t.row_dimensions[1].height = 32
ws_t.merge_cells("A2:N2")
ws_t["A2"] = ("Actuals populate automatically each month.  "
              "Revenue and Covers rows need manual entry after each run — see Summary tab.")
style(ws_t["A2"], colour=GREY_HDR, size=9, bg=LIGHT_AMBER, h="left", border=False)
ws_t.row_dimensions[2].height = 18

ws_t.cell(4, 1, "").border = _b()
for mi, m in enumerate(MONTHS, 2):
    style(ws_t.cell(4, mi, m), bold=True, colour=WHITE, size=10, bg=MID_TEAL)
style(ws_t.cell(4, 14, "YTD"), bold=True, colour=WHITE, size=10, bg=DARK_TEAL)
ws_t.row_dimensions[4].height = 24

trend_defs = [
    (5,  "Total Food Cost (R)",     'R#,##0',    'food'),
    (6,  "Breakfast Cost (R)",      'R#,##0',    'bfast'),
    (7,  "Breakfast % of Food",     '0.0%',      'bfast_pct'),
    (8,  "Breakfast Revenue (R)",   'R#,##0',    'revenue'),
    (9,  "Cost % of Revenue",       '0.0%',      None),
    (10, "Breakfast Covers",        '#,##0',     'covers'),
    (11, "Cost per Cover (R)",      'R#,##0.00', None),
    (12, "Breakfast Allocation",    'R#,##0.00', None),
    (13, "CoS per cover",           '0.0%',      None),
]

for row_num, label, num_fmt, hist_key in trend_defs:
    ws_t.cell(row_num, 1, label)
    is_manual = label in ["Breakfast Revenue (R)", "Breakfast Covers"]
    style(ws_t.cell(row_num,1), bold=True, size=10,
          bg=LIGHT_AMBER if is_manual else LIGHT_BLUE, h="left")
    ws_t.row_dimensions[row_num].height = 22

    for col in range(2, 14):
        c      = ws_t.cell(row_num, col)
        col_bg = LIGHT_BLUE if row_num%2==0 else WHITE
        if is_manual:
            col_bg = LIGHT_AMBER

        # Restore history
        if col in trend_history and col != month_col and hist_key:
            hist_val = trend_history[col].get(hist_key)
            if hist_val is not None:
                style(c, colour="0000FF", size=10, bg=col_bg, num=num_fmt)
                c.value = hist_val
                continue

        # Calculated rows — formula only, no history
        if row_num == 9:
            col_l = get_column_letter(col)
            c.value = f"=IFERROR({col_l}6/{col_l}8,0)"
            c.number_format = num_fmt
            c.border = _b(); c.alignment = _al()
            c.fill = _f(col_bg); c.font = _ft(size=10)
            continue
        if row_num == 11:
            col_l = get_column_letter(col)
            c.value = f"=IFERROR({col_l}6/{col_l}10,0)"
            c.number_format = num_fmt
            c.border = _b(); c.alignment = _al()
            c.fill = _f(col_bg); c.font = _ft(size=10)
            continue
        if row_num == 12:
            c.value = "=275/1.15"
            c.number_format = num_fmt
            c.border = _b(); c.alignment = _al()
            c.fill = _f(col_bg); c.font = _ft(size=10)
            continue
        if row_num == 13:
            col_l = get_column_letter(col)
            c.value = f"=IFERROR({col_l}11/{col_l}12,0)"
            c.number_format = num_fmt
            c.border = _b(); c.alignment = _al()
            c.fill = _f(col_bg); c.font = _ft(size=10)
            continue

        # Current month
        if col == month_col:
            if row_num == 5:   val = round(grand_total, 2)
            elif row_num == 6: val = round(bfast_total, 2)
            elif row_num == 7: val = round(bfast_pct, 4)
            else:              val = None
            if val is not None:
                style(c, colour="0000FF", size=10, bg=col_bg, num=num_fmt)
                c.value = val
            else:
                style(c, colour=GREY_HDR, size=9, bg=col_bg,
                      num=num_fmt, italic=True)
        else:
            style(c, size=10, bg=col_bg, num=num_fmt)

    # YTD column
    if row_num in (12, 13):
        ytd = ws_t.cell(row_num, 14)
        style(ytd, size=10, bg=DARK_TEAL)
        ws_t.row_dimensions[row_num].height = 22
        continue
    if row_num == 7:
        ytd_f = "=IFERROR(N6/N5,0)"
    elif row_num == 9:
        ytd_f = "=IFERROR(N6/N8,0)"
    elif row_num == 11:
        ytd_f = "=IFERROR(N6/N10,0)"
    else:
        ytd_f = f"=SUM(B{row_num}:M{row_num})"
    ytd = ws_t.cell(row_num, 14, ytd_f)
    style(ytd, bold=True, colour=WHITE, size=10, num=num_fmt, bg=DARK_TEAL)

ws_t.column_dimensions["A"].width = 24
for col in range(2, 15):
    ws_t.column_dimensions[get_column_letter(col)].width = 13
ws_t.freeze_panes = "B5"

# ── Add Attribution Sheet ─────────────────────────────────────────────────────
if 'Attribution' in wb.sheetnames:
    del wb['Attribution']

att_sheet = wb.create_sheet('Attribution', 0)
from datetime import datetime
att_sheet['A1'] = '© 2026 Alain Yon. All rights reserved.'
att_sheet['A2'] = ''
att_sheet['A3'] = 'Script: breakfast_rebuild.py v2.0'
att_sheet['A4'] = f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'
att_sheet['A5'] = ''
att_sheet['A6'] = 'This workbook and associated analysis are the intellectual property'
att_sheet['A7'] = 'of Alain Yon.'

att_sheet['A1'].font = Font(bold=True, size=12)
att_sheet.column_dimensions['A'].width = 70

# ── Add Print Footer to All Sheets ────────────────────────────────────────────
from openpyxl.worksheet.header_footer import HeaderFooter
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    hf = HeaderFooter()
    hf.oddFooter.center.text = "© 2026 Alain Yon | breakfast_rebuild.py v2.0 | Page &P of &N"
    ws.header_footer = hf

# ── Save ──────────────────────────────────────────────────────────────────────
wb.active = ws_s
wb.save(out_path)
print(f"\n  ✓ Saved: Breakfast_MonthEnd_Tool.xlsx")
print(f"  Location: {out_path}\n")
