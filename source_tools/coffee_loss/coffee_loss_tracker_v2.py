"""
Coffee Loss Tracker v2 — hotel coffee shop
© 2026 Alain Yon. All rights reserved.
Version: 2.0 | Last updated: June 19, 2026
================================================
Compares theoretical ingredient consumption (from Micros POS sales × recipes)
against actual Mi20 ingredient consumption to quantify and value coffee losses.

Recipe costs and quantities are read live from Coffee_Costing_Updated.xlsx
so they stay in sync whenever the costing file is updated.

MONTHLY CONFIG — update these three lines only:
"""

COSTING_FILE   = "Coffee_Costing_Updated.xlsx"   # costing workbook (same folder)
MI20_FOLDER    = "Mi20_files"                     # folder containing Mi20 .xlsx exports

# Drink names (must match costing file exactly)
DRINK_NAMES = [
    "Americano",
    "Single Esspresso",
    "Double Esspresso",
    "Cappuccino",
    "Café Latte",
    "Decaf Americano",
    "Decaf Cappuccino",
    "Decaf Café Latte ",
    "Café Mocha",
    "Cortado",
]

# ─────────────────────────────────────────────────────────────────────────────
# DO NOT EDIT BELOW THIS LINE
# ─────────────────────────────────────────────────────────────────────────────

import os
import glob
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              numbers as xl_numbers)
from openpyxl.utils import get_column_letter
from datetime import datetime

MISC_PCT = 0.03  # 3% miscellaneous cost loading


# ── 1. READ RECIPES FROM COSTING FILE ────────────────────────────────────────

def load_recipes(costing_path):
    """
    Read each Menu sheet from the costing workbook.
    Returns:
        recipes  : {drink_name: {selling_price, total_cost, ingredients: [...]}}
        sap_info : {sap_nr: {desc, uom, price}}
    """
    if not os.path.exists(costing_path):
        raise FileNotFoundError(
            f"Costing file not found: {costing_path}\n"
            "Place Coffee_Costing_Updated.xlsx in the same folder as this script."
        )

    wb = load_workbook(costing_path, data_only=True)

    # Build SAP lookup from MM60 sheet
    sap_info = {}
    if 'MM60' in wb.sheetnames:
        ws_mm60 = wb['MM60']
        for row in ws_mm60.iter_rows(min_row=2, values_only=True):
            if row[1] and str(row[1]).isdigit():
                sap_info[int(row[1])] = {
                    'desc': row[2],
                    'uom':  row[4],
                    'price': float(row[5]) if row[5] else 0.0
                }

    recipes = {}
    menu_sheets = [s for s in wb.sheetnames if s.startswith('Menu ')]

    for sheet_name in menu_sheets:
        ws = wb[sheet_name]
        drink = ws['C2'].value
        if not drink:
            continue

        selling_price = ws['G33'].value or 0
        total_cost    = ws['G28'].value or 0
        misc_cost     = ws['G27'].value or 0
        ingredient_subtotal = total_cost - misc_cost  # before 3% loading

        ingredients = []
        for row in range(5, 27):
            sap_nr = ws[f'B{row}'].value
            qty    = ws[f'D{row}'].value
            uom    = ws[f'E{row}'].value
            price  = ws[f'F{row}'].value
            cost   = ws[f'G{row}'].value

            if sap_nr and str(sap_nr).strip().lstrip('0123456789') == '':
                try:
                    sap_int = int(float(sap_nr))
                    if qty and price:
                        ingredients.append({
                            'sap_nr':    sap_int,
                            'desc':      sap_info.get(sap_int, {}).get('desc', f'SAP {sap_int}'),
                            'qty_per_serve': float(qty),
                            'uom':       uom,
                            'unit_price': float(price),
                            'line_cost': float(cost) if cost else 0.0,
                        })
                except (ValueError, TypeError):
                    pass

        recipes[drink] = {
            'selling_price':         selling_price,
            'ingredient_subtotal':   ingredient_subtotal,
            'misc_cost':             misc_cost,
            'total_cost':            total_cost,
            'ingredients':           ingredients,
        }

    return recipes, sap_info


# ── 2. LOAD MI20 DATA ─────────────────────────────────────────────────────────

def load_mi20(folder):
    """
    Loads all Mi20 Physical Inventory variance exports from folder.
    Returns a DataFrame with columns: material_code, description, quantity
    Aggregates across all files found.
    """
    files = glob.glob(os.path.join(folder, '*.xlsx')) + \
            glob.glob(os.path.join(folder, '*.xls'))

    if not files:
        print(f"  [WARNING] No Mi20 files found in '{folder}'. "
              "Actual consumption will show as 0.")
        return pd.DataFrame(columns=['material_code', 'description', 'quantity'])

    frames = []
    for f in files:
        try:
            df = pd.read_excel(f)
            df.columns = [str(c).strip() for c in df.columns]
            frames.append(df)
            print(f"  Loaded Mi20: {os.path.basename(f)} ({len(df)} rows)")
        except Exception as e:
            print(f"  [WARNING] Could not read {f}: {e}")

    if not frames:
        return pd.DataFrame(columns=['material_code', 'description', 'quantity'])

    combined = pd.concat(frames, ignore_index=True)

    # Detect material and quantity columns (flexible naming)
    col_map = {}
    for col in combined.columns:
        cl = col.lower()
        if 'material' in cl and 'desc' not in cl:
            col_map['material'] = col
        elif 'description' in cl:
            col_map['description'] = col
        elif 'quantity' in cl or 'qty' in cl:
            col_map['quantity'] = col

    if 'material' not in col_map or 'quantity' not in col_map:
        print(f"  [WARNING] Could not identify material/quantity columns in Mi20.")
        print(f"  Found columns: {list(combined.columns)}")
        return pd.DataFrame(columns=['material_code', 'description', 'quantity'])

    out = combined[[col_map['material'], col_map.get('description', col_map['material']), col_map['quantity']]].copy()
    out.columns = ['material_code', 'description', 'quantity']
    out['material_code'] = pd.to_numeric(out['material_code'], errors='coerce')
    out['quantity']      = pd.to_numeric(out['quantity'],      errors='coerce').fillna(0)
    out = out.dropna(subset=['material_code'])
    out['material_code'] = out['material_code'].astype(int)

    # Aggregate (multiple Mi20 exports may cover different periods)
    agg = out.groupby('material_code').agg(
        description=('description', 'first'),
        quantity=('quantity', 'sum')
    ).reset_index()

    return agg


# ── 3. CALCULATE THEORETICAL VS ACTUAL ───────────────────────────────────────

def calculate_variances(recipes, sales_qty, mi20_df):
    """
    For each SAP ingredient used in recipes:
      theoretical_qty  = sum(sales_qty[drink] × recipe_qty_per_serve) across all drinks
      actual_qty       = Mi20 consumption
      variance_qty     = actual - theoretical  (positive = more used than expected = loss)
      variance_value   = variance_qty × unit_price × (1 + MISC_PCT)
    """
    # Build ingredient aggregate: {sap_nr: {desc, uom, unit_price, theoretical_qty}}
    ingredient_totals = {}

    for drink, data in recipes.items():
        sold = sales_qty.get(drink, 0)
        for ing in data['ingredients']:
            sap = ing['sap_nr']
            if sap not in ingredient_totals:
                ingredient_totals[sap] = {
                    'desc':      ing['desc'],
                    'uom':       ing['uom'],
                    'unit_price': ing['unit_price'],
                    'theoretical_qty': 0.0,
                    'theoretical_cost': 0.0,
                    'drinks_used_in': []
                }
            theo_qty = sold * ing['qty_per_serve']
            ingredient_totals[sap]['theoretical_qty']  += theo_qty
            ingredient_totals[sap]['theoretical_cost'] += theo_qty * ing['unit_price']
            if drink not in ingredient_totals[sap]['drinks_used_in']:
                ingredient_totals[sap]['drinks_used_in'].append(drink)

    # Build variance rows
    rows = []
    mi20_lookup = {row['material_code']: row['quantity']
                   for _, row in mi20_df.iterrows()} if not mi20_df.empty else {}

    for sap, data in ingredient_totals.items():
        actual_qty    = mi20_lookup.get(sap, 0.0)
        theo_qty      = data['theoretical_qty']
        variance_qty  = actual_qty - theo_qty
        unit_price    = data['unit_price']
        variance_cost = variance_qty * unit_price
        # Apply 3% misc loading to the variance cost
        variance_cost_with_misc = variance_cost * (1 + MISC_PCT)

        rows.append({
            'SAP Nr':               sap,
            'Item Description':     data['desc'],
            'UOM':                  data['uom'],
            'Unit Price (R)':       unit_price,
            'Theoretical Qty':      round(theo_qty, 4),
            'Actual Qty (Mi20)':    round(actual_qty, 4),
            'Variance Qty':         round(variance_qty, 4),
            'Variance Cost (R)':    round(variance_cost_with_misc, 2),
            'Loss / (Gain)':        round(variance_cost_with_misc, 2),
            'Used In':              ', '.join(data['drinks_used_in']),
        })

    return rows


# ── 4. BUILD DRINK SUMMARY TABLE ──────────────────────────────────────────────

def build_drink_summary(recipes, sales_qty):
    rows = []
    total_revenue = 0
    total_cost = 0

    for drink, data in recipes.items():
        sold = sales_qty.get(drink, 0)
        revenue = sold * (data['selling_price'] / 1.15)  # excl VAT
        cost    = sold * data['total_cost']
        rows.append({
            'Drink':              drink,
            'Selling Price (R)':  data['selling_price'],
            'Recipe Cost (R)':    round(data['total_cost'], 4),
            'COS %':              round(data['total_cost'] / (data['selling_price'] / 1.15), 4)
                                  if data['selling_price'] else 0,
            'Units Sold':         sold,
            'Revenue excl VAT':   round(revenue, 2),
            'Total Cost':         round(cost, 2),
        })
        total_revenue += revenue
        total_cost    += cost

    return rows, total_revenue, total_cost


# ── 5. WRITE EXCEL OUTPUT ─────────────────────────────────────────────────────

PAL = {
    'title_bg':  '042F2E',
    'hdr_bg':    '0F4C5C',
    'input_bg':  'FEF3C7',
    'loss_bg':   'FEE2E2',
    'gain_bg':   'DCFCE7',
    'total_bg':  '134E4A',
    'zebra':     'F0FDFA',
    'white':     'FFFFFF',
    'amber':     'F59E0B',
    'red':       'DC2626',
    'green':     '16A34A',
    'title_txt': 'FFFFFF',
    'hdr_txt':   'FFFFFF',
    'total_txt': 'FFFFFF',
}

def _fill(hex_color):
    return PatternFill('solid', fgColor=hex_color)

def _font(color='000000', bold=False, size=11):
    return Font(color=color, bold=bold, size=size, name='Calibri')

def _border():
    side = Side(style='thin', color='CCCCCC')
    return Border(left=side, right=side, top=side, bottom=side)

def _center():
    return Alignment(horizontal='center', vertical='center', wrap_text=True)

def _right():
    return Alignment(horizontal='right', vertical='center')

def _left():
    return Alignment(horizontal='left', vertical='center')


def write_title_row(ws, row, text, ncols, start_col=1):
    ws.merge_cells(start_row=row, start_column=start_col,
                   end_row=row, end_column=start_col + ncols - 1)
    cell = ws.cell(row=row, column=start_col, value=text)
    cell.fill  = _fill(PAL['title_bg'])
    cell.font  = _font(PAL['title_txt'], bold=True, size=13)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[row].height = 28


def write_header_row(ws, row, headers, start_col=1):
    for i, h in enumerate(headers, start=start_col):
        c = ws.cell(row=row, column=i, value=h)
        c.fill      = _fill(PAL['hdr_bg'])
        c.font      = _font(PAL['hdr_txt'], bold=True, size=10)
        c.alignment = _center()
        c.border    = _border()
    ws.row_dimensions[row].height = 32


def write_data_row(ws, row, values, formats=None, zebra=False, loss_col=None, start_col=1):
    bg = PAL['zebra'] if zebra else PAL['white']
    for i, val in enumerate(values, start=start_col):
        c = ws.cell(row=row, column=i, value=val)
        c.border    = _border()
        c.alignment = _right() if isinstance(val, (int, float)) else _left()

        # Loss/gain colouring on specified column
        if loss_col and i == loss_col:
            if isinstance(val, (int, float)):
                c.fill = _fill(PAL['loss_bg']) if val > 0 else \
                         (_fill(PAL['gain_bg']) if val < 0 else _fill(bg))
                c.font = _font(PAL['red'] if val > 0 else
                               (PAL['green'] if val < 0 else '000000'), bold=True)
                continue

        c.fill = _fill(bg)
        c.font = _font()

        if formats and i - start_col < len(formats) and formats[i - start_col]:
            c.number_format = formats[i - start_col]


def write_total_row(ws, row, values, start_col=1):
    for i, val in enumerate(values, start=start_col):
        c = ws.cell(row=row, column=i, value=val)
        c.fill      = _fill(PAL['total_bg'])
        c.font      = _font(PAL['total_txt'], bold=True)
        c.border    = _border()
        c.alignment = _right() if isinstance(val, (int, float)) else _left()


def create_output(recipes, sales_qty, mi20_df, month_label, output_path):
    wb = Workbook()

    # ── Sheet 1: Drink Summary ────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = 'Drink Summary'

    drink_rows, total_rev, total_cost_val = build_drink_summary(recipes, sales_qty)

    write_title_row(ws1, 1, f'Coffee Drink Summary — {month_label}', 7)
    headers = ['Drink', 'Sell Price (R)', 'Recipe Cost (R)', 'COS %',
               'Units Sold', 'Revenue excl VAT (R)', 'Total Cost (R)']
    write_header_row(ws1, 2, headers)

    fmts = [None, '#,##0.00', '#,##0.0000', '0.0%', '#,##0', '#,##0.00', '#,##0.00']
    for idx, row_data in enumerate(drink_rows):
        r = 3 + idx
        vals = [row_data['Drink'], row_data['Selling Price (R)'],
                row_data['Recipe Cost (R)'], row_data['COS %'],
                row_data['Units Sold'], row_data['Revenue excl VAT'],
                row_data['Total Cost']]
        write_data_row(ws1, r, vals, fmts, zebra=(idx % 2 == 0))

    total_row = 3 + len(drink_rows)
    total_cos = total_cost_val / total_rev if total_rev else 0
    write_total_row(ws1, total_row,
                    ['TOTAL', '', '', total_cos, '',
                     round(total_rev, 2), round(total_cost_val, 2)])
    ws1.cell(total_row, 4).number_format = '0.0%'
    ws1.cell(total_row, 6).number_format = '#,##0.00'
    ws1.cell(total_row, 7).number_format = '#,##0.00'

    col_widths = [26, 15, 15, 10, 12, 22, 15]
    for i, w in enumerate(col_widths, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    ws1.freeze_panes = 'A3'

    # ── Sheet 2: Recipe Cards ─────────────────────────────────────────────────
    ws2 = wb.create_sheet('Recipe Cards')
    r = 1
    write_title_row(ws2, r, f'Coffee Recipe Cards (incl. 3% Misc) — {month_label}', 7)
    r += 1

    for drink, data in recipes.items():
        write_header_row(ws2, r,
            [drink, 'SAP Nr', 'Ingredient', 'QTY per Serve', 'UOM',
             'Unit Price (R)', 'Line Cost (R)'])
        r += 1
        for idx, ing in enumerate(data['ingredients']):
            vals = ['', ing['sap_nr'], ing['desc'], ing['qty_per_serve'],
                    ing['uom'], ing['unit_price'], ing['line_cost']]
            fmts = [None, '#,##0', None, '#,##0.0000', None, '#,##0.00', '#,##0.00']
            write_data_row(ws2, r, vals, fmts, zebra=(idx % 2 == 0))
            r += 1

        # Misc cost row
        misc_val = data['misc_cost']
        vals = ['', '—', 'Misc Cost (3%)', '', '', '', misc_val]
        write_data_row(ws2, r, vals, [None]*6 + ['#,##0.00'], zebra=True)
        r += 1

        # Total row
        write_total_row(ws2, r,
            ['', '', 'Total Recipe Cost', '', '',
             '', round(data['total_cost'], 4)])
        ws2.cell(r, 7).number_format = '#,##0.0000'
        r += 2  # blank row between drinks

    col_widths2 = [22, 10, 38, 14, 8, 16, 14]
    for i, w in enumerate(col_widths2, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    ws2.freeze_panes = 'A2'

    # ── Sheet 3: Ingredient Loss Analysis ────────────────────────────────────
    ws3 = wb.create_sheet('Ingredient Loss Analysis')

    variance_rows = calculate_variances(recipes, sales_qty, mi20_df)

    write_title_row(ws3, 1, f'Ingredient Loss Analysis (vs Mi20) — {month_label}', 9)
    headers3 = ['SAP Nr', 'Ingredient', 'UOM', 'Unit Price (R)',
                'Theoretical Qty', 'Actual Qty (Mi20)',
                'Variance Qty', 'Variance Cost (R)\nincl. 3% Misc',
                'Used In Drinks']
    write_header_row(ws3, 2, headers3)

    fmts3 = ['#,##0', None, None, '#,##0.00', '#,##0.0000',
             '#,##0.0000', '#,##0.0000', '#,##0.00', None]

    total_loss = 0
    for idx, row_data in enumerate(variance_rows):
        r3 = 3 + idx
        vals = [row_data['SAP Nr'], row_data['Item Description'],
                row_data['UOM'], row_data['Unit Price (R)'],
                row_data['Theoretical Qty'], row_data['Actual Qty (Mi20)'],
                row_data['Variance Qty'], row_data['Loss / (Gain)'],
                row_data['Used In']]
        write_data_row(ws3, r3, vals, fmts3, zebra=(idx % 2 == 0),
                       loss_col=8)
        total_loss += row_data['Loss / (Gain)']

    total_row3 = 3 + len(variance_rows)
    write_total_row(ws3, total_row3,
        ['', 'TOTAL VARIANCE (incl. 3% misc)', '', '', '', '',
         '', round(total_loss, 2), ''])
    ws3.cell(total_row3, 8).number_format = '#,##0.00'
    ws3.cell(total_row3, 8).fill = _fill(PAL['loss_bg'] if total_loss > 0 else PAL['gain_bg'])
    ws3.cell(total_row3, 8).font = Font(
        color=PAL['red'] if total_loss > 0 else PAL['green'],
        bold=True, size=11, name='Calibri')

    col_widths3 = [10, 36, 8, 14, 16, 18, 14, 22, 36]
    for i, w in enumerate(col_widths3, 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    ws3.freeze_panes = 'A3'

    # ── Sheet 4: Mi20 Raw Data ────────────────────────────────────────────────
    if not mi20_df.empty:
        ws4 = wb.create_sheet('Mi20 Raw Data')
        write_title_row(ws4, 1, f'Mi20 Raw Consumption Data — {month_label}', 3)
        write_header_row(ws4, 2, ['Material Code', 'Description', 'Quantity'])
        for idx, (_, row) in enumerate(mi20_df.iterrows()):
            r4 = 3 + idx
            vals = [row['material_code'], row['description'], row['quantity']]
            write_data_row(ws4, r4, vals, ['#,##0', None, '#,##0.0000'],
                           zebra=(idx % 2 == 0))
        for i, w in enumerate([14, 48, 14], 1):
            ws4.column_dimensions[get_column_letter(i)].width = w
        ws4.freeze_panes = 'A3'

    # ── Sheet 5: Legend ───────────────────────────────────────────────────────
    ws5 = wb.create_sheet('Legend & Notes')
    write_title_row(ws5, 1, 'Coffee Loss Tracker v2 — Legend & Notes', 3)

    notes = [
        ('', ''),
        ('HOW TO USE THIS TOOL', ''),
        ('Step 1', f'Update SALES_QTY in the script with Micros POS counts for {month_label}'),
        ('Step 2', f'Drop Mi20 export files into the "{MI20_FOLDER}" folder'),
        ('Step 3', f'Update MONTH_LABEL in the script'),
        ('Step 4', 'Run: python coffee_loss_tracker_v2.py'),
        ('', ''),
        ('COST METHODOLOGY', ''),
        ('Recipe costs', f'Read live from {COSTING_FILE} — update the costing file to refresh prices'),
        ('Misc loading', '3% applied to all ingredient costs to account for unmeasured losses'),
        ('Variance cost', 'Variance Qty × Unit Price × 1.03'),
        ('', ''),
        ('COLOUR CODING', ''),
        ('Red cells (Loss)',  'Actual consumption exceeds theoretical — investigate'),
        ('Green cells (Gain)', 'Actual consumption below theoretical — possible under-recording'),
        ('Amber cells',       'User input required'),
        ('', ''),
        ('SHEETS', ''),
        ('Drink Summary',         'Revenue, cost and COS% by drink for the month'),
        ('Recipe Cards',          'Full ingredient breakdown per drink including 3% misc cost'),
        ('Ingredient Loss Analysis', 'Core output: theoretical vs actual vs variance per SAP ingredient'),
        ('Mi20 Raw Data',         'Aggregated Mi20 consumption data as loaded'),
        ('', ''),
        ('Generated', datetime.now().strftime('%Y-%m-%d %H:%M')),
    ]

    for idx, (label, note) in enumerate(notes):
        r5 = 2 + idx
        cell_label = ws5.cell(r5, 1, label)
        cell_note  = ws5.cell(r5, 2, note)
        if label in ('HOW TO USE THIS TOOL', 'COST METHODOLOGY',
                     'COLOUR CODING', 'SHEETS'):
            cell_label.font = _font(bold=True, size=11)
            cell_label.fill = _fill(PAL['hdr_bg'])
            cell_label.font = _font(PAL['hdr_txt'], bold=True)
        else:
            cell_label.font = _font(bold=True)
            cell_note.font  = _font()

    ws5.column_dimensions['A'].width = 30
    ws5.column_dimensions['B'].width = 65

    # ── Add Attribution Sheet ─────────────────────────────────────────────────
    if 'Attribution' in wb.sheetnames:
        del wb['Attribution']
    
    att_sheet = wb.create_sheet('Attribution', 0)
    att_sheet['A1'] = '© 2026 Alain Yon. All rights reserved.'
    att_sheet['A2'] = ''
    att_sheet['A3'] = 'Script: coffee_loss_tracker_v2.py v2.0'
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
        hf.oddFooter.center.text = "© 2026 Alain Yon | coffee_loss_tracker_v2.py v2.0 | Page &P of &N"
        ws.header_footer = hf

    wb.save(output_path)
    return total_loss


# ── INPUT COLLECTION ─────────────────────────────────────────────────────────

def get_month_label():
    """Prompt user for month label."""
    while True:
        label = input("\nEnter month label (e.g., 'June 2026'): ").strip()
        if label:
            return label
        print("  [ERROR] Month label cannot be empty. Try again.")

def get_sales_quantities():
    """Prompt user for sales quantities for each drink."""
    sales_qty = {}
    print("\nEnter sales quantities for each drink (enter 0 if none sold):")
    print("-" * 60)
    
    for drink in DRINK_NAMES:
        while True:
            try:
                qty_input = input(f"  {drink}: ").strip()
                qty = int(qty_input) if qty_input else 0
                if qty >= 0:
                    sales_qty[drink] = qty
                    break
                else:
                    print(f"    [ERROR] Must be 0 or positive. Try again.")
            except ValueError:
                print(f"    [ERROR] Invalid input. Enter a whole number.")
    
    return sales_qty

# ── MAIN ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    # Collect user input
    MONTH_LABEL = get_month_label()
    SALES_QTY = get_sales_quantities()
    
    print(f"\n{'='*60}")
    print(f"  Coffee Loss Tracker v2 — {MONTH_LABEL}")
    print(f"{'='*60}\n")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    costing_path = os.path.join(script_dir, COSTING_FILE)
    mi20_folder  = os.path.join(script_dir, MI20_FOLDER)

    print("→ Loading recipes from costing file...")
    recipes, sap_info = load_recipes(costing_path)
    print(f"  {len(recipes)} drinks loaded: {', '.join(recipes.keys())}")

    print("\n→ Loading Mi20 consumption data...")
    os.makedirs(mi20_folder, exist_ok=True)
    mi20_df = load_mi20(mi20_folder)
    if not mi20_df.empty:
        print(f"  {len(mi20_df)} ingredients loaded from Mi20")
    else:
        print("  No Mi20 data — run with sales data only to check recipe costs")

    print("\n→ Building output workbook...")
    output_filename = f"Coffee_Loss_Tracker_{MONTH_LABEL.replace(' ', '_')}.xlsx"
    output_path = os.path.join(script_dir, output_filename)

    total_loss = create_output(recipes, SALES_QTY, mi20_df, MONTH_LABEL, output_path)

    print(f"\n{'='*60}")
    print(f"  Output: {output_filename}")
    total_units = sum(SALES_QTY.values())
    print(f"  Total units sold entered: {total_units}")
    if total_units == 0:
        print("  [NOTE] All sales quantities are 0 — update SALES_QTY and rerun")
    print(f"  Net variance (loss): R{total_loss:,.2f}")
    print(f"{'='*60}\n")
