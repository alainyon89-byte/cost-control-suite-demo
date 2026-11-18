"""
=============================================================================
COST CONTROL SUITE  |  Staff Canteen Month-End Rebuild Script
© 2026 Alain Yon. All rights reserved.
Version: 2.0 | Last updated: July 8, 2026
=============================================================================
HOW TO USE EACH MONTH:
  1. Save your Mi20 SAP export as an .xlsx file in this folder
  2. Update the three CONFIG lines below
  3. Run:  Run_Canteen_Report.bat
  4. Output saved as Canteen_MonthEnd_Tool.xlsx in the same folder

REQUIREMENTS (one-time install):
  pip install openpyxl pandas

FILES NEEDED IN THE SAME FOLDER:
  - canteen_rebuild.py          (this script)
  - MI20_kitchen.xlsx           (or whatever you name the SAP export)
  - Canteen_Item_Master.xlsx    (generated automatically on first run — keep it)
  - Canteen_MonthEnd_Tool.xlsx  (output — rebuilt each month)
=============================================================================
"""

# -----------------------------------------------------------------------------
# CONFIG  — update this line each month
# -----------------------------------------------------------------------------
SAP_FILE     = "MI20_kitchen.xlsx"    # SAP Mi20 export filename
BUDGET_ALLOC = 250000                 # staff canteen budget allocation for this month (R), demo figure
# -----------------------------------------------------------------------------

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

# Palette
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

# Paths
MONTH_LABEL  = get_month_label()
script_dir   = os.path.dirname(os.path.abspath(__file__))
sap_path     = os.path.join(script_dir, SAP_FILE)
master_path  = os.path.join(script_dir, "Canteen_Item_Master.xlsx")
out_path     = os.path.join(script_dir, "Canteen_MonthEnd_Tool.xlsx")

# Validate SAP file
if not os.path.exists(sap_path):
    sys.exit(f"\n  ERROR: Cannot find SAP file:\n  {sap_path}\n"
             f"  Check SAP_FILE in the config and make sure the file is in the same folder.\n")

print(f"\n  Loading SAP file: {SAP_FILE}")
df_sap   = pd.read_excel(sap_path, header=0)
num_cols = df_sap.shape[1]
print(f"  Columns detected: {num_cols}")

if num_cols >= 19:
    print("  Format: Full SAP PhysInv export (19 columns)")
    df_sap.columns = ['stock_count','PhysInvDoc','Item','Material','Description','Batch','Plnt',
                       'SLoc','UoM','Book_qty','Count_qty','Diff_qty','Diff_value',
                       'Pos_diff','Neg_diff','Notes','Unit_price','Monthly_cost','Total_consumption'
                       ][:num_cols]
elif num_cols >= 12:
    print(f"  Format: Mi20 kitchen export ({num_cols} columns)")
    df_sap = df_sap.iloc[:, :12].copy()
    df_sap.columns = ['PhysInvDoc','Item','Material','Description','Plnt','Location',
                       'Book_qty','Count_qty','Diff_qty','Diff_value','Pos_diff','Neg_diff']
else:
    sys.exit(f"\n  ERROR: Unrecognised format — {num_cols} columns found.\n"
             f"  Expected 12 (Mi20) or 19 (full PhysInv). Check the export.\n")

df_sap['Material_str'] = df_sap['Material'].dropna().astype(float).astype(int).astype(str)

# Load or build canteen item master
if os.path.exists(master_path):
    print(f"  Loading canteen item master: Canteen_Item_Master.xlsx")
    df_master = pd.read_excel(master_path, header=2)
    df_master = df_master[pd.to_numeric(df_master['Material Code'], errors='coerce').notna()].copy()
    df_master['Material_str'] = df_master['Material Code'].astype(float).astype(int).astype(str)
    canteen_codes = set(df_master['Material_str'])
    print(f"  Master items: {len(canteen_codes)}")
else:
    print(f"\n  NOTE: No Canteen_Item_Master.xlsx found — using all SAP items.\n")
    canteen_codes = None

# Filter & consolidate
data = df_sap[df_sap['Description'].notna() &
              pd.to_numeric(df_sap['Diff_value'], errors='coerce').notna()].copy()
data['Diff_value'] = pd.to_numeric(data['Diff_value'], errors='coerce').abs()
data['Diff_qty']   = pd.to_numeric(data['Diff_qty'],   errors='coerce').abs()

if canteen_codes:
    data = data[data['Material_str'].isin(canteen_codes)].copy()
    print(f"  Rows after canteen filter: {len(data)}")

consolidated = data.groupby(['Material_str','Description'], sort=False).agg(
    Total_qty  = ('Diff_qty',   'sum'),
    Total_cost = ('Diff_value', 'sum'),
).reset_index()
consolidated['Unit_price'] = consolidated.apply(
    lambda r: round(r['Total_cost']/r['Total_qty'],2) if r['Total_qty'] else 0, axis=1)

# Apply canteen split % from master
if os.path.exists(master_path):
    df_split = pd.read_excel(master_path, header=2)
    if df_split.shape[1] >= 4:
        df_split.columns = ['Material_Code','Description','Chef_Notes',
                             'Canteen_Pct','Split_Notes'][:df_split.shape[1]]
        df_split = df_split[pd.to_numeric(df_split['Material_Code'], errors='coerce').notna()].copy()
        df_split['Material_str'] = df_split['Material_Code'].astype(float).astype(int).astype(str)
        df_split['Canteen_Pct']  = pd.to_numeric(df_split['Canteen_Pct'], errors='coerce').fillna(1.0)
        split_map = dict(zip(df_split['Material_str'], df_split['Canteen_Pct']))
        consolidated['Canteen_Pct'] = consolidated['Material_str'].map(split_map).fillna(1.0)
        splits_applied = consolidated[consolidated['Canteen_Pct'] < 1.0]
        if len(splits_applied) > 0:
            print(f"\n  Split % applied to {len(splits_applied)} item(s):")
            for _, sr in splits_applied.iterrows():
                print(f"    {sr['Description']}: {sr['Canteen_Pct']*100:.0f}%"
                      f" -> R{sr['Total_cost']*sr['Canteen_Pct']:,.2f}"
                      f" (was R{sr['Total_cost']:,.2f})")
    else:
        consolidated['Canteen_Pct'] = 1.0
else:
    consolidated['Canteen_Pct'] = 1.0

consolidated['Total_cost_adj'] = consolidated['Total_cost'] * consolidated['Canteen_Pct']
consolidated['Total_qty_adj']  = consolidated['Total_qty']  * consolidated['Canteen_Pct']

excluded = consolidated[consolidated['Canteen_Pct'] == 0.0]
if len(excluded) > 0:
    print(f"\n  Excluded {len(excluded)} item(s) at 0% canteen:")
    for _, er in excluded.iterrows():
        print(f"    {er['Description']}")
consolidated = consolidated[consolidated['Canteen_Pct'] > 0.0].copy()
consolidated = consolidated.sort_values('Total_cost_adj', ascending=False).reset_index(drop=True)
grand_total  = consolidated['Total_cost_adj'].sum()
over_budget  = grand_total > BUDGET_ALLOC

print(f"\n  {'─'*45}")
print(f"  Unique canteen items:  {len(consolidated)}")
print(f"  Grand total:           R {grand_total:>12,.2f}")
print(f"  Budget:                R {BUDGET_ALLOC:>12,.2f}")
variance = BUDGET_ALLOC - grand_total
print(f"  Variance (Under)/Over: R {variance:>12,.2f}  {'OK' if variance >= 0 else 'OVER BUDGET'}")
print(f"  {'─'*45}\n")

# Determine month column
month_short = MONTH_LABEL.split()[0][:3]
try:
    month_col = MONTHS.index(month_short) + 2
except ValueError:
    month_col = None
print(f"  Month detected: {month_short} -> column {month_col}")

# Preserve trend history from existing workbook
trend_history = {}
if os.path.exists(out_path):
    print(f"  Preserving Trend history from existing workbook...")
    try:
        ex_wb = openpyxl.load_workbook(out_path, data_only=True)
        if "Trend" in ex_wb.sheetnames:
            ex_t = ex_wb["Trend"]
            for col in range(2, 14):
                cost_val = ex_t.cell(5, col).value
                bud_val  = ex_t.cell(6, col).value
                cov_val  = ex_t.cell(9, col).value
                cpc_val  = ex_t.cell(10, col).value
                if isinstance(cost_val, (int, float)) and cost_val:
                    trend_history[col] = {
                        'cost':    cost_val,
                        'budget':  bud_val,
                        'covers':  cov_val,
                        'cpc':     cpc_val,
                    }
        ex_wb.close()
        print(f"  Months with existing data: {len(trend_history)}")
    except Exception as e:
        print(f"  Warning: could not read trend history ({e}). Starting fresh.")

# ============================================================
# BUILD WORKBOOK
# ============================================================
wb = openpyxl.Workbook()

# ── SHEET 1: SAP DUMP ──────────────────────────────────────
ws_d = wb.active
ws_d.title = "SAP Dump"
ws_d.sheet_properties.tabColor = "2E75B6"

ws_d.merge_cells("A1:L1")
ws_d["A1"] = f"SAP KITCHEN DUMP  |  {MONTH_LABEL}  |  Full kitchen export — canteen filter applied on Consolidated tab"
style(ws_d["A1"], bold=True, colour=WHITE, size=11, bg=DARK_TEAL, h="left")
ws_d.row_dimensions[1].height = 28
ws_d.merge_cells("A2:L2")
ws_d["A2"] = f"Paste full Mi20 export here from row 3 each month."
style(ws_d["A2"], colour=GREY_HDR, size=9, bg=LIGHT_AMBER, h="left", border=False)
ws_d.row_dimensions[2].height = 20

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

# ── SHEET 2: CANTEEN ITEMS ─────────────────────────────────
ws_m = wb.create_sheet("Canteen Items")
ws_m.sheet_properties.tabColor = "C55A11"
ws_m.merge_cells("A1:E1")
ws_m["A1"] = "CANTEEN ITEM MASTER  |  Filter reference + Split %"
style(ws_m["A1"], bold=True, colour=WHITE, size=12, bg=DARK_TEAL, h="left")
ws_m.row_dimensions[1].height = 28
ws_m.merge_cells("A2:E2")
ws_m["A2"] = "Canteen % in column D drives cost allocation. 100% = full canteen. 0% = excluded. Add items: new row with code + description + %. Remove: delete row."
style(ws_m["A2"], colour=GREY_HDR, size=9, bg=LIGHT_AMBER, h="left", wrap=True, border=False)
ws_m.row_dimensions[2].height = 30
for ci, h in enumerate(["Material Code","Material Description","Chef Notes","Canteen %","Split Notes"], 1):
    style(ws_m.cell(3, ci, h), bold=True, colour=WHITE, size=10, bg=MID_TEAL)
ws_m.row_dimensions[3].height = 22

if os.path.exists(master_path):
    df_disp = pd.read_excel(master_path, header=2)
    df_disp = df_disp[pd.to_numeric(df_disp['Material Code'], errors='coerce').notna()].copy()
    df_disp['Canteen_Pct'] = pd.to_numeric(df_disp['Canteen %'], errors='coerce').fillna(1.0)
    active_codes = set(consolidated['Material_str'])
    df_disp = df_disp.sort_values('Material Description').reset_index(drop=True)
    for ri, row in df_disp.iterrows():
        r   = ri + 4
        bg  = LIGHT_BLUE if ri%2==0 else WHITE
        pct = row['Canteen_Pct']
        pct_bg = RED_LIGHT if pct==0 else (AMBER if pct<0.5 else (LIGHT_AMBER if pct<1.0 else GREEN_LIGHT))
        ws_m.cell(r, 1, int(row['Material Code']))
        style(ws_m.cell(r,1), size=10, bg=bg)
        ws_m.cell(r, 2, row['Material Description'])
        style(ws_m.cell(r,2), size=10, bg=bg, h="left")
        note_val = row.get('Chef Notes (from SAP Dec)', '')
        note_val = note_val if pd.notna(note_val) else ''
        ws_m.cell(r, 3, note_val)
        style(ws_m.cell(r,3), size=9, colour=GREY_HDR, bg=bg, h="left", italic=True)
        ws_m.cell(r, 4, pct)
        style(ws_m.cell(r,4), bold=True, colour="0000FF", size=10, bg=pct_bg, num='0%')
        split_note = row.get('Split Reason / Notes', '')
        split_note = split_note if pd.notna(split_note) else ''
        ws_m.cell(r, 5, split_note)
        style(ws_m.cell(r,5), size=9, colour=GREY_HDR, bg=bg, h="left")
        ws_m.row_dimensions[r].height = 18
for i, w in enumerate([14,40,28,12,42], 1):
    ws_m.column_dimensions[get_column_letter(i)].width = w
ws_m.freeze_panes = "A4"

# ── SHEET 3: CONSOLIDATED ──────────────────────────────────
ws_c = wb.create_sheet("Consolidated")
ws_c.sheet_properties.tabColor = "375623"

ws_c.merge_cells("A1:H1")
ws_c["A1"] = "STAFF CANTEEN  |  Consolidated Month-End Cost Summary"
style(ws_c["A1"], bold=True, colour=WHITE, size=13, bg=DARK_TEAL, h="left")
ws_c.row_dimensions[1].height = 32

ws_c["A2"] = "Reporting Month:"
style(ws_c["A2"], bold=True, size=10, bg=LIGHT_AMBER, h="right", border=False)
ws_c["B2"] = MONTH_LABEL
style(ws_c["B2"], bold=True, colour="0000FF", size=11, bg=LIGHT_AMBER, h="left", border=False)
ws_c.merge_cells("B2:D2")
ws_c["F2"] = "Budget Allocation (R):"
style(ws_c["F2"], bold=True, size=10, bg=LIGHT_AMBER, h="right", border=False)
ws_c["G2"] = BUDGET_ALLOC
style(ws_c["G2"], bold=True, colour="0000FF", size=11, bg=LIGHT_AMBER, num='R#,##0', border=False)
ws_c.row_dimensions[2].height = 22

ws_c.merge_cells("A3:H3")
ws_c["A3"] = (f"Filtered to {len(consolidated)} canteen items  |  Chef reviewed & signed off  |  Ranked by cost")
style(ws_c["A3"], colour=GREY_HDR, size=9, bg=LIGHT_AMBER, h="left", wrap=True, border=False)
ws_c.row_dimensions[3].height = 20

for ci, h in enumerate(["#","Material Code","Material Description","Canteen %",
                         "Adj Qty","Avg Unit Price (R)","Adjusted Cost (R)","% of Total"], 1):
    style(ws_c.cell(4, ci, h), bold=True, colour=WHITE, size=10, bg=MID_TEAL)
ws_c.row_dimensions[4].height = 24

DR = 5
for ri, row in consolidated.iterrows():
    r   = DR + ri
    bg  = LIGHT_BLUE if ri%2==0 else WHITE
    pct_bg = LIGHT_AMBER if row['Canteen_Pct'] < 1.0 else bg
    ws_c.cell(r, 1, ri+1)
    style(ws_c.cell(r,1), colour=GREY_HDR, size=9, bg=bg)
    for ci, val, num, lft, cbg in [
        (2, int(row['Material_str']),          None,         False, bg),
        (3, row['Description'],                None,         True,  bg),
        (4, row['Canteen_Pct'],                '0%',         False, pct_bg),
        (5, round(row['Total_qty_adj'],3),     '#,##0.000',  False, bg),
        (6, row['Unit_price'],                 'R#,##0.00',  False, bg),
        (7, round(row['Total_cost_adj'],2),    'R#,##0.00',  False, bg),
    ]:
        c = ws_c.cell(r, ci, val)
        style(c, size=10, h="left" if lft else "center", bg=cbg, num=num,
              bold=(ci==4 and row['Canteen_Pct']<1.0),
              colour="FF8C00" if (ci==4 and row['Canteen_Pct']<1.0) else "000000")
    tr_ref = DR + len(consolidated)
    c = ws_c.cell(r, 8, f"=G{r}/G${tr_ref}")
    style(c, size=10, bg=bg, num='0.0%')
    ws_c.row_dimensions[r].height = 18

tr = DR + len(consolidated)
ws_c.merge_cells(f"A{tr}:F{tr}")
ws_c.cell(tr, 1, f"TOTAL CANTEEN COST  |  {MONTH_LABEL}")
style(ws_c.cell(tr,1), bold=True, colour=DARK_TEAL, size=11, bg=AMBER, h="right")
ws_c.cell(tr, 7, f"=SUM(G{DR}:G{tr-1})")
style(ws_c.cell(tr,7), bold=True, colour=DARK_TEAL, size=11, bg=AMBER, num='R#,##0.00')
ws_c.cell(tr, 8, "100.0%")
style(ws_c.cell(tr,8), bold=True, colour=DARK_TEAL, size=11, bg=AMBER, num='0.0%')
ws_c.row_dimensions[tr].height = 26

kb = tr + 2
ws_c.merge_cells(f"A{kb-1}:H{kb-1}")
ws_c.cell(kb-1, 1, f"KPI SUMMARY  |  {MONTH_LABEL}")
style(ws_c.cell(kb-1,1), bold=True, colour=WHITE, size=11, bg=DARK_TEAL, h="left")
ws_c.row_dimensions[kb-1].height = 22

kpi_rows = [
    ("Budget Allocation (R)",          BUDGET_ALLOC,                               'R#,##0.00', "0000FF"),
    ("Total Canteen Cost (R)",         f"=G{tr}",                                  'R#,##0.00', None),
    ("Variance — (Over) / Under",      f"=G{kb}-G{kb+1}",                          'R#,##0.00', None),
    (None,None,None,None),
    ("Lunch / Dinner Covers",          0,                                           '#,##0',     "0000FF"),
    ("Breakfast Covers",               0,                                           '#,##0',     "0000FF"),
    ("Total Covers",                   f"=G{kb+4}+G{kb+5}",                        '#,##0',     None),
    (None,None,None,None),
    ("Lunch / Dinner Rate (R/cover)",  30.47,                                       'R#,##0.00', "0000FF"),
    ("Breakfast Rate (R/cover)",       18.59,                                       'R#,##0.00', "0000FF"),
    ("Revenue from Staff Meals (R)",   f"=G{kb+8}*G{kb+4}+G{kb+9}*G{kb+5}",       'R#,##0.00', None),
    (None,None,None,None),
    ("Cost per Cover (R)",             f"=IFERROR(G{kb+1}/G{kb+6},0)",             'R#,##0.00', None),
    ("Loss per Cover (R)",             f"=IFERROR((G{kb+1}-G{kb+10})/G{kb+6},0)",  'R#,##0.00', None),
]
for i, (label, val, num, fc) in enumerate(kpi_rows):
    r = kb + i
    if not label:
        ws_c.row_dimensions[r].height = 6; continue
    is_var = "Variance" in label
    bg_row = (RED_LIGHT if over_budget else GREEN_LIGHT) if is_var else (
              LIGHT_BLUE if i%2==0 else WHITE)
    ws_c.merge_cells(f"A{r}:F{r}")
    style(ws_c.cell(r,1,label), bold=True, size=10, bg=bg_row, h="right", border=False)
    vc = ws_c.cell(r, 7, val)
    var_col = ("FF0000" if over_budget else "375623") if is_var else (fc or "000000")
    style(vc, bold=True, colour=var_col, size=10, bg=bg_row, num=num)
    ws_c.cell(r,8).fill = _f(bg_row)
    ws_c.row_dimensions[r].height = 20

for i, w in enumerate([4,13,42,10,16,18,16,12], 1):
    ws_c.column_dimensions[get_column_letter(i)].width = w
ws_c.freeze_panes = "A5"

# ── SHEET 4: TREND ─────────────────────────────────────────
ws_t = wb.create_sheet("Trend")
ws_t.sheet_properties.tabColor = "FFD966"

ws_t.merge_cells("A1:P1")
ws_t["A1"] = "STAFF CANTEEN  |  Month-on-Month Cost Trend"
style(ws_t["A1"], bold=True, colour=WHITE, size=13, bg=DARK_TEAL, h="left")
ws_t.row_dimensions[1].height = 32
ws_t.merge_cells("A2:P2")
ws_t["A2"] = "Actuals populate automatically each month. Blue = editable input. Green = under budget. Red = over."
style(ws_t["A2"], colour=GREY_HDR, size=9, bg=LIGHT_AMBER, h="left", border=False)
ws_t.row_dimensions[2].height = 20

ws_t.cell(4,1,"").border = _b()
for mi, m in enumerate(MONTHS, 2):
    style(ws_t.cell(4,mi,m), bold=True, colour=WHITE, size=10, bg=MID_TEAL)
for ci, lbl in [(14,"YTD Total"),(15,"Budget YTD"),(16,"Var YTD")]:
    style(ws_t.cell(4,ci,lbl), bold=True, colour=WHITE, size=10, bg=DARK_TEAL)
ws_t.row_dimensions[4].height = 24

trend_defs = [
    (5,  "Canteen Cost (R)",   'R#,##0',    False),
    (6,  "Budget (R)",         'R#,##0',    False),
    (7,  "Variance (R)",       'R#,##0',    True),
    (8,  "",                   None,        False),
    (9,  "Total Covers",       '#,##0',     False),
    (10, "Cost per Cover (R)", 'R#,##0.00', False),
]

for row_num, label, num_fmt, is_var in trend_defs:
    ws_t.cell(row_num, 1, label)
    style(ws_t.cell(row_num,1), bold=True, size=10,
          bg=AMBER if is_var else LIGHT_BLUE, h="left")
    ws_t.row_dimensions[row_num].height = 22 if label else 8

    for col in range(2, 14):
        c      = ws_t.cell(row_num, col)
        col_bg = LIGHT_BLUE if row_num%2==0 else WHITE

        if is_var:
            col_l = get_column_letter(col)
            c.value = f"=IF(AND({col_l}6=0,{col_l}5=0),\"\",{col_l}6-{col_l}5)"
            c.number_format = num_fmt or 'R#,##0'
            c.border = _b()
            c.alignment = _al()
            if col == month_col:
                c.fill = _f(GREEN_LIGHT if not over_budget else RED_LIGHT)
                c.font = _ft(size=10, colour="375623" if not over_budget else "FF0000")
            else:
                c.fill = _f(LIGHT_AMBER)
                c.font = _ft(size=10, colour=GREY_HDR)
            continue

        if not label:
            continue

        # Restore history for months already processed
        if col in trend_history and col != month_col:
            hist = trend_history[col]
            if row_num == 5:    val = hist.get('cost')
            elif row_num == 6:  val = hist.get('budget')
            elif row_num == 9:  val = hist.get('covers')
            elif row_num == 10: val = hist.get('cpc')
            else:               val = None
            if val is not None:
                style(c, colour="0000FF", size=10, bg=col_bg, num=num_fmt)
                c.value = val
                continue

        # Current month
        if col == month_col:
            if row_num == 5:   val = round(grand_total, 2)
            elif row_num == 6: val = BUDGET_ALLOC
            else:              val = None
            if val is not None:
                style(c, colour="0000FF", size=10, bg=col_bg, num=num_fmt)
                c.value = val
            else:
                style(c, colour=GREY_HDR, size=9, bg=col_bg, num=num_fmt, italic=True)
        else:
            style(c, size=10, bg=col_bg, num=num_fmt)

    # YTD column
    if is_var:
        ytd = ws_t.cell(row_num, 14, "=SUM(B6:M6)-SUM(B5:M5)")
        style(ytd, bold=True, size=10, num=num_fmt, bg=AMBER,
              colour="375623" if not over_budget else "FF0000")
        ws_t.cell(row_num,15).fill = _f(AMBER)
        var_ytd = ws_t.cell(row_num, 16, f"=N{row_num}")
        style(var_ytd, bold=True, size=10, num=num_fmt,
              bg=GREEN_LIGHT if not over_budget else RED_LIGHT,
              colour="375623" if not over_budget else "FF0000")
    elif label:
        ytd = ws_t.cell(row_num, 14,
                        f"=IFERROR(N5/N9,0)" if row_num==10
                        else f"=SUM(B{row_num}:M{row_num})")
        style(ytd, bold=True, colour=WHITE, size=10, num=num_fmt, bg=DARK_TEAL)
        ws_t.cell(row_num,15).fill = _f(DARK_TEAL)
        ws_t.cell(row_num,16).fill = _f(DARK_TEAL)

ws_t.column_dimensions["A"].width = 22
for col in range(2, 17):
    ws_t.column_dimensions[get_column_letter(col)].width = 13
ws_t.freeze_panes = "B5"

# ── Add Attribution Sheet ─────────────────────────────────────────────────────
if 'Attribution' in wb.sheetnames:
    del wb['Attribution']

att_sheet = wb.create_sheet('Attribution', 0)
from datetime import datetime
att_sheet['A1'] = '© 2026 Alain Yon. All rights reserved.'
att_sheet['A2'] = ''
att_sheet['A3'] = 'Script: canteen_rebuild.py v2.0'
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
    hf.oddFooter.center.text = "© 2026 Alain Yon | canteen_rebuild.py v2.0 | Page &P of &N"
    ws.header_footer = hf

# Save workbook
wb.active = ws_c
wb.save(out_path)

# Save / update Canteen_Item_Master.xlsx if it doesn't exist
if not os.path.exists(master_path):
    print(f"\n  Creating Canteen_Item_Master.xlsx from current data...")
    wb_m = openpyxl.Workbook()
    ws_mi = wb_m.active
    ws_mi.title = "Canteen Items"
    ws_mi.cell(1,1,"Material Code")
    ws_mi.cell(1,2,"Material Description")
    all_items = data[['Material_str','Description']].drop_duplicates().sort_values('Description')
    for ri, row in enumerate(all_items.itertuples(), 2):
        ws_mi.cell(ri, 1, int(row.Material_str))
        ws_mi.cell(ri, 2, row.Description)
    wb_m.save(master_path)
    print(f"  Saved: Canteen_Item_Master.xlsx  ({len(all_items)} items)")

print(f"\n  Saved: Canteen_MonthEnd_Tool.xlsx")
print(f"  Location: {out_path}")
print(f"\n  REMINDER: Open the Consolidated tab and update the covers figures")
print(f"  (Lunch/Dinner covers, Breakfast covers) in the KPI block.\n")
