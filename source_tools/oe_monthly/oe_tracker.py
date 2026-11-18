#!/usr/bin/env python3
"""
=============================================================================
COST CONTROL SUITE  |  OE Tracker v3.3  (monthly build)
(c) 2026 Alain Yon. All rights reserved.
Version: 3.3 | Last updated: 6 November 2026
=============================================================================
Builds the monthly Operational Equipment (OE) tracking workbook and the
F&B Manager approval file from three SAP exports:

    MB52  stock on hand per storage location   (required)
    MM60  unit prices                           (required)
    MB51  scrapping, movement types 551 / 552   (optional)

The SAP files can have any name. They are recognised by their column
headers. Drop them in this folder, or drag them onto Run_OE_Monthly.bat.

Carried forward from the current OE_Tracking_Workbook.xlsx (or, if that is
missing, the newest copy in the Archive folder):
    * Par Level (EA), Notes and Par Group on every location tab
    * the LOCATIONS tab (names, "On Approval" Y/N, notes, approval file
      and approvers)
    * First Seen (used to flag items that are new this month)

Outputs
    OE_Tracking_Workbook.xlsx       live workbook (previous one archived first)
    OE_Approval_<Mon>_<YYYY>.xlsx   send to the F&B Manager
    OE_Approval_<File>_<Mon>_<YYYY>.xlsx
                                    one per "Approval File" on the LOCATIONS
                                    tab (e.g. Linen Main), two-step sign-off
    OE_Transfer_Form.xlsx           one-page transfer form for operations
                                    (item lists from MB52, prices from MM60,
                                    no quantities on hand anywhere in the file)

Requires Python 3.8+ and openpyxl  (pip install openpyxl)
"""

import os
import re
import sys
import glob
import math
import shutil
import warnings
from datetime import datetime, date, timedelta
from collections import defaultdict

warnings.filterwarnings('ignore', message='Workbook contains no default style')
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
except ImportError:
    print('ERROR: openpyxl is not installed. Run:  python -m pip install openpyxl')
    sys.exit(1)

VERSION = '3.3'
PLANT = 'D100'
PROPERTY_NAME = os.environ.get("CCS_PROPERTY_NAME", "Demo Hotel")   # set by the app from config/app_settings.json
WORKBOOK_NAME = 'OE_Tracking_Workbook.xlsx'
ARCHIVE_DIR = 'Archive'
HIGH_STOCK_RATIO = 0.5       # HIGH urgency when stock is below 50% of par
PRICE_DIFF_CHECK = 0.05      # flag MB52 unit value vs MM60 price differences > 5%
MAX_TRANSFER_SOURCES = 3
LOC_RE = re.compile(r'^L\d{3}$')

# Storage location names, plant D100 (demo hotel: Main Hotel and Tower Wing, TW).
# Only used to seed the LOCATIONS tab; after that the tab is the master.
DEFAULT_LOCATION_NAMES = {
    'L073': 'KITCHEN', 'L075': 'BAR - CELLAR', 'L076': 'ROOM SERVICE',
    'L077': 'MINI BAR', 'L080': 'KITCHEN CLEANING', 'L081': 'STATIONERY',
    'L082': 'CLEANING SUPPLIES', 'L083': 'GUEST SUPPLIES', 'L084': 'TOBACCO',
    'L095': 'CROCKERY', 'L096': 'CUTLERY', 'L097': 'GLASSWARE',
    'L098': 'HOLLOWARE', 'L099': 'UTENSILS', 'L100': 'F&B LINEN',
    'L101': 'ROOMS LINEN', 'L102': 'ROOMS OTHER', 'L107': 'CC TRANSFERS',
    'L108': 'CAFE', 'L112': 'BANQUETING', 'L118': 'HOTEL OTHER',
    'L123': 'POOL BAR', 'L222': 'BAR TW', 'L223': 'MINI BAR TW',
    'L224': 'PAPER SUPPLIES TW', 'L225': 'STATIONERY TW',
    'L226': 'CLEANING SUPPLIES TW', 'L227': 'GUEST SUPPLIES TW',
    'L228': 'CROCKERY TW', 'L229': 'CUTLERY TW', 'L230': 'GLASSWARE TW',
    'L231': 'HOLLOWARE TW', 'L232': 'UTENSILS TW', 'L233': 'F&B LINEN TW',
    'L234': 'ROOMS LINEN TW', 'L235': 'ROOMS OTHER TW', 'L236': 'FLOORS TW',
    'L237': 'HOTEL OTHER', 'L239': 'STILL WATER',
}

# Location tab columns (header text is what the next run reads back)
COLS = ['Material', 'Description', 'Unit Price (MM60)', 'Stock (MB52)',
        'Hold Value (MB52)', 'Scrap Qty (MB51)', 'Scrap Value (MB51)',
        'Par Level (EA)', 'Order Qty', 'Order Value', 'Urgency', 'Flag',
        'Notes', 'First Seen', 'Par Group (order code)']
HDR_ROW = 4
FIRST_ROW = 5

# Colours (same family as v2)
C_TITLE = 'FF6B35'
C_HEAD = '004E89'
C_SECTION = 'F77F00'
C_TOTAL = 'F0E5CF'
C_INPUT = 'FFF2CC'
C_BAND = 'F7F9FC'
C_HIGH_FILL, C_HIGH_FONT = 'F8CBAD', '9C0006'
C_MED_FILL, C_MED_FONT = 'FFE699', '7F6000'
C_LINK = '0563C1'
C_GREY = '808080'

F_MONEY = '#,##0.00'
F_RAND = '"R "#,##0.00'
F_QTY = '#,##0'
F_QTY_DEC = '#,##0.00'

THIN = Side(style='thin', color='D9D9D9')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def fill(colour):
    return PatternFill(start_color=colour, end_color=colour, fill_type='solid')


def parse_number(v):
    """Number from a cell value. Handles '1,234.56', '1234,56', SAP '12-'."""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        if isinstance(v, float) and math.isnan(v):
            return None
        return float(v)
    s = str(v).strip().replace(' ', '').replace(' ', '')
    if s == '':
        return None
    neg = False
    if s.endswith('-'):
        neg, s = True, s[:-1]
    if s.startswith('(') and s.endswith(')'):
        neg, s = True, s[1:-1]
    s = s.replace('R', '') if s.startswith('R') else s
    if ',' in s and '.' in s:
        s = s.replace(',', '')
    elif ',' in s:
        if re.fullmatch(r'-?\d{1,3}(,\d{3})+', s):
            s = s.replace(',', '')
        else:
            s = s.replace(',', '.')
    try:
        n = float(s)
    except ValueError:
        return None
    return -n if neg else n


def parse_material(v):
    n = parse_number(v)
    if n is None or n <= 0 or n != int(n):
        return None
    return int(n)


def parse_date(v):
    if v is None or v == '':
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return (datetime(1899, 12, 30) + timedelta(days=int(v))).date()
    s = str(v).strip()
    for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y', '%Y/%m/%d', '%d-%m-%Y',
                '%Y-%m-%d %H:%M:%S'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def text(v):
    return '' if v is None else str(v).strip()


def norm(h):
    return re.sub(r'\s+', ' ', text(h)).lower()


def qty_format(x):
    return F_QTY if x is None or float(x).is_integer() else F_QTY_DEC


def clean_num(x):
    """Store whole numbers as int so Excel shows 12 not 12.0."""
    if x is None:
        return None
    x = round(float(x), 6)
    return int(x) if x.is_integer() else x


def order_qty(par, stock):
    """Same rule as the sheet formula: ROUNDUP(ROUND(MAX(par-stock,0),4),0)."""
    if par is None or par <= 0:
        return 0
    return int(math.ceil(round(max(par - stock, 0.0), 4)))


def urgency(par, stock):
    if par is None or par <= 0 or stock >= par:
        return ''
    return 'HIGH' if stock < par * HIGH_STOCK_RATIO else 'MEDIUM'


def file_is_locked(path):
    """True when Excel has the file open (Windows locks it)."""
    if not os.path.exists(path):
        return False
    try:
        with open(path, 'r+b'):
            return False
    except PermissionError:
        return True
    except OSError:
        return False


def period_slug(period):
    return re.sub(r'[^A-Za-z0-9]+', '_', period).strip('_')


# ---------------------------------------------------------------------------
# SAP file detection and loading
# ---------------------------------------------------------------------------
def read_table(path):
    """Return (headers, rows) from the first sheet that has a 'Material' header."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            rows = ws.iter_rows(values_only=True)
            buffer = []
            header = None
            for i, r in enumerate(rows):
                if i >= 10:
                    break
                buffer.append(r)
                if r and any(norm(c) == 'material' for c in r):
                    header = [text(c) for c in r]
                    break
            if header is None:
                continue
            data = [list(r) for r in rows]
            return header, data
        return None, []
    finally:
        wb.close()


def classify(headers):
    h = {norm(x) for x in headers if x}
    if {'material', 'storage location', 'movement type'} <= h and \
            ('qty in unit of entry' in h or 'quantity' in h):
        return 'MB51'
    if {'material', 'storage location', 'unrestricted'} <= h:
        return 'MB52'
    if {'material', 'price'} <= h and 'storage location' not in h:
        return 'MM60'
    return None


def col_index(headers, *names, required=True):
    hn = [norm(x) for x in headers]
    for name in names:
        if norm(name) in hn:
            return hn.index(norm(name))
    if required:
        raise ValueError(f"column '{names[0]}' not found. Columns: {headers}")
    return None


def find_sap_files(folder, explicit):
    """Find MB52/MM60/MB51 by header. Explicit (dragged) files win."""
    found = defaultdict(list)
    candidates = []
    for p in explicit:
        candidates.append((p, True))
    for p in glob.glob(os.path.join(folder, '*.xlsx')):
        base = os.path.basename(p)
        if base.startswith('~$') or base.upper().startswith('OE_'):
            continue
        if any(os.path.abspath(p) == os.path.abspath(e) for e in explicit):
            continue
        candidates.append((p, False))
    for p, is_explicit in candidates:
        try:
            headers, _ = read_table_headers_only(p)
        except Exception as e:  # unreadable file: skip, but say so
            print(f'  ! Skipped {os.path.basename(p)} (could not read: {e})')
            continue
        kind = classify(headers) if headers else None
        if kind:
            found[kind].append((p, is_explicit, os.path.getmtime(p)))
        elif is_explicit:
            print(f'  ! {os.path.basename(p)} is not an MB52, MM60 or MB51 export - ignored')
    chosen, notes = {}, []
    for kind in ('MB52', 'MM60', 'MB51'):
        lst = found.get(kind, [])
        if not lst:
            continue
        lst.sort(key=lambda t: (t[1], t[2]), reverse=True)  # explicit first, then newest
        chosen[kind] = lst[0][0]
        if len(lst) > 1:
            others = ', '.join(os.path.basename(t[0]) for t in lst[1:])
            notes.append(f'{kind}: {len(lst)} files found. Using '
                         f'{os.path.basename(lst[0][0])}; ignored {others}. '
                         f'Move old SAP exports out of the folder to avoid this.')
    return chosen, notes


def read_table_headers_only(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            for i, r in enumerate(ws.iter_rows(values_only=True)):
                if i >= 10:
                    break
                if r and any(norm(c) == 'material' for c in r):
                    return [text(c) for c in r], None
        return None, None
    finally:
        wb.close()


def plant_ok(row, i_plant):
    if i_plant is None:
        return True
    p = text(row[i_plant]).upper()
    return p == '' or p == PLANT


def load_mb52(path, checks):
    headers, rows = read_table(path)
    i_mat = col_index(headers, 'Material')
    i_desc = col_index(headers, 'Material Description', 'Material Desc', required=False)
    i_loc = col_index(headers, 'Storage Location')
    i_qty = col_index(headers, 'Unrestricted')
    i_val = col_index(headers, 'Value Unrestricted', required=False)
    i_plant = col_index(headers, 'Plant', required=False)
    i_uom = col_index(headers, 'Base Unit of Measure', required=False)
    stock, other_plant, control = {}, 0, None
    for r in rows:
        r = r + [None] * (len(headers) - len(r))
        mat = parse_material(r[i_mat])
        loc = text(r[i_loc]).upper()
        if mat is None or not loc:
            if mat is None and not loc and i_val is not None:
                v = parse_number(r[i_val])
                if v is not None:
                    control = v  # SAP total line at the bottom
            continue
        if not plant_ok(r, i_plant):
            other_plant += 1
            continue
        qty = parse_number(r[i_qty]) or 0.0
        val = parse_number(r[i_val]) if i_val is not None else None
        key = (loc, mat)
        if key in stock:
            stock[key]['qty'] += qty
            if val is not None:
                stock[key]['value'] = (stock[key]['value'] or 0) + val
            checks.append(('Duplicate MB52 line (summed)', loc, mat,
                           stock[key]['desc'], 'Two MB52 lines for the same item and location'))
        else:
            stock[key] = {'desc': text(r[i_desc]) if i_desc is not None else '',
                          'qty': qty, 'value': val,
                          'uom': text(r[i_uom]) if i_uom is not None else ''}
    if other_plant:
        checks.append(('Ignored rows', '', None, '',
                       f'MB52: {other_plant} rows from a plant other than {PLANT}'))
    return stock, control


def load_mm60(path, checks):
    headers, rows = read_table(path)
    i_mat = col_index(headers, 'Material')
    i_desc = col_index(headers, 'Material Description', 'Material Desc', required=False)
    i_price = col_index(headers, 'Price')
    i_plant = col_index(headers, 'Plant', required=False)
    i_uom = col_index(headers, 'Base Unit of Measure', required=False)
    prices, other_plant = {}, 0
    for r in rows:
        r = r + [None] * (len(headers) - len(r))
        mat = parse_material(r[i_mat])
        if mat is None:
            continue
        if not plant_ok(r, i_plant):
            other_plant += 1
            continue
        desc = text(r[i_desc]) if i_desc is not None else ''
        prices[mat] = {'price': parse_number(r[i_price]), 'desc': desc,
                       'uom': text(r[i_uom]) if i_uom is not None else '',
                       'dnu': desc.lower().startswith('dnu') or 'do not use' in desc.lower()}
    if other_plant:
        checks.append(('Ignored rows', '', None, '',
                       f'MM60: {other_plant} rows from a plant other than {PLANT}'))
    return prices


def load_mb51(path, checks):
    headers, rows = read_table(path)
    i_mat = col_index(headers, 'Material')
    i_desc = col_index(headers, 'Material Description', required=False)
    i_loc = col_index(headers, 'Storage Location')
    i_mvt = col_index(headers, 'Movement Type')
    i_qty = col_index(headers, 'Qty in unit of entry', 'Quantity')
    i_amt = col_index(headers, 'Amt.in Loc.Cur.', 'Amount in LC', required=False)
    i_date = col_index(headers, 'Posting Date', required=False)
    i_plant = col_index(headers, 'Plant', required=False)
    i_unit = col_index(headers, 'Unit of Entry', required=False)
    scrap = {}
    ignored_mvt, other_plant, control = defaultdict(int), 0, None
    dates = []
    for r in rows:
        r = r + [None] * (len(headers) - len(r))
        mat = parse_material(r[i_mat])
        loc = text(r[i_loc]).upper()
        if mat is None or not loc:
            if mat is None and not loc and i_amt is not None:
                v = parse_number(r[i_amt])
                if v is not None:
                    control = v
            continue
        if not plant_ok(r, i_plant):
            other_plant += 1
            continue
        mvt = text(r[i_mvt])
        mvt = mvt[:-2] if mvt.endswith('.0') else mvt
        if mvt not in ('551', '552'):
            ignored_mvt[mvt] += 1
            continue
        qty = parse_number(r[i_qty]) or 0.0
        amt = (parse_number(r[i_amt]) or 0.0) if i_amt is not None else 0.0
        d = parse_date(r[i_date]) if i_date is not None else None
        if d:
            dates.append(d)
        key = (loc, mat)
        s = scrap.setdefault(key, {'qty': 0.0, 'value': 0.0, 'desc': '', 'units': set()})
        s['qty'] -= qty      # 551 posts negative; show scrap as a positive number
        s['value'] -= amt
        s['desc'] = s['desc'] or (text(r[i_desc]) if i_desc is not None else '')
        if i_unit is not None and text(r[i_unit]):
            s['units'].add(text(r[i_unit]).upper())
    for mvt, n in ignored_mvt.items():
        checks.append(('Ignored rows', '', None, '',
                       f'MB51: {n} rows with movement type {mvt or "(blank)"} (only 551/552 used)'))
    if other_plant:
        checks.append(('Ignored rows', '', None, '',
                       f'MB51: {other_plant} rows from a plant other than {PLANT}'))
    date_range = (min(dates), max(dates)) if dates else None
    return scrap, control, date_range


# ---------------------------------------------------------------------------
# Previous workbook (par levels, notes, locations, first seen)
# ---------------------------------------------------------------------------
def find_previous_workbook(folder):
    live = os.path.join(folder, WORKBOOK_NAME)
    if os.path.exists(live):
        return live
    arch = glob.glob(os.path.join(folder, ARCHIVE_DIR, 'OE_Tracking_Workbook*.xlsx'))
    arch = [a for a in arch if not os.path.basename(a).startswith('~$')]
    if arch:
        return max(arch, key=os.path.getmtime)
    return None


def read_previous_workbook(path, checks):
    prev = {'par': {}, 'notes': {}, 'first_seen': {}, 'desc': {}, 'price': {},
            'group': {}, 'items': set(), 'locations': {}, 'loc_order': []}
    wb_f = load_workbook(path, data_only=False)
    wb_v = load_workbook(path, data_only=True)
    try:
        for name in wb_f.sheetnames:
            ws_f, ws_v = wb_f[name], wb_v[name]
            if name.upper() == 'LOCATIONS':
                read_locations_tab(ws_v, prev)
                continue
            loc = name.strip().upper()
            if not LOC_RE.match(loc):
                continue
            hdr_row, cmap = None, {}
            for r in range(1, 11):
                vals = [norm(ws_f.cell(r, c).value) for c in range(1, ws_f.max_column + 1)]
                if 'material' in vals and any(v.startswith('par level') for v in vals):
                    hdr_row = r
                    for c, v in enumerate(vals, 1):
                        if v == 'material':
                            cmap['mat'] = c
                        elif v == 'description':
                            cmap['desc'] = c
                        elif v.startswith('par level'):
                            cmap['par'] = c
                        elif v == 'notes':
                            cmap['notes'] = c
                        elif v == 'first seen':
                            cmap['first'] = c
                        elif v.startswith('par group'):
                            cmap['group'] = c
                        elif v.startswith('unit price') or v.startswith('mm60 cost'):
                            cmap['price'] = c
                    break
            if hdr_row is None:
                checks.append(('Previous workbook', loc, None, '',
                               f'Tab {name}: header row not found, par levels not read'))
                continue
            for r in range(hdr_row + 1, ws_f.max_row + 1):
                raw_mat = ws_f.cell(r, cmap['mat']).value
                if isinstance(raw_mat, str) and raw_mat.strip().startswith('='):
                    continue  # totals row (v2 put =SUM() over material numbers here)
                mat = parse_material(raw_mat)
                if mat is None:
                    continue
                key = (loc, mat)
                prev['items'].add(key)
                if 'desc' in cmap:
                    prev['desc'][key] = text(ws_v.cell(r, cmap['desc']).value)
                if 'price' in cmap:
                    pr = parse_number(ws_v.cell(r, cmap['price']).value)
                    if pr and pr > 0:
                        prev['price'][mat] = pr
                pv = ws_v.cell(r, cmap['par']).value
                if pv is None:
                    pf = ws_f.cell(r, cmap['par']).value
                    pv = pf if not (isinstance(pf, str) and pf.startswith('=')) else None
                if pv is not None and text(pv) != '':
                    n = parse_number(pv)
                    if n is None or n < 0:
                        checks.append(('Par level ignored', loc, mat, prev['desc'].get(key, ''),
                                       f'Par Level "{pv}" is not a number of 0 or more'))
                    else:
                        prev['par'][key] = n
                if 'notes' in cmap:
                    nv = text(ws_v.cell(r, cmap['notes']).value)
                    if nv:
                        prev['notes'][key] = nv
                if 'first' in cmap:
                    fv = text(ws_v.cell(r, cmap['first']).value)
                    if fv:
                        prev['first_seen'][key] = fv
                if 'group' in cmap:
                    gv = ws_v.cell(r, cmap['group']).value
                    if gv not in (None, '') and text(gv) != '':
                        g = parse_material(gv)
                        if g is None:
                            checks.append(('Par group ignored', loc, mat, prev['desc'].get(key, ''),
                                           f'Par Group "{gv}" is not a material number'))
                        else:
                            prev['group'][key] = g
    finally:
        wb_f.close()
        wb_v.close()
    return prev


def read_locations_tab(ws, prev):
    hdr, cmap = None, {}
    for r in range(1, 11):
        vals = [norm(ws.cell(r, c).value) for c in range(1, 12)]
        if 'code' in vals:
            hdr = r
            for c, v in enumerate(vals, 1):
                if v == 'code':
                    cmap['code'] = c
                elif v == 'name':
                    cmap['name'] = c
                elif v.startswith('on approval'):
                    cmap['appr'] = c
                elif v == 'notes':
                    cmap['notes'] = c
                elif v == 'approval file':
                    cmap['file'] = c
                elif v == 'approver 1':
                    cmap['a1'] = c
                elif v == 'approver 2':
                    cmap['a2'] = c
            break
    if hdr is None:
        return
    for r in range(hdr + 1, ws.max_row + 1):
        code = text(ws.cell(r, cmap['code']).value).upper()
        if not LOC_RE.match(code):
            continue
        appr = text(ws.cell(r, cmap['appr']).value).upper() if 'appr' in cmap else 'Y'
        prev['locations'][code] = {
            'name': text(ws.cell(r, cmap['name']).value) if 'name' in cmap else '',
            'appr': 'N' if appr in ('N', 'NO') else 'Y',
            'notes': text(ws.cell(r, cmap['notes']).value) if 'notes' in cmap else '',
            'file': text(ws.cell(r, cmap['file']).value) if 'file' in cmap else '',
            'a1': text(ws.cell(r, cmap['a1']).value) if 'a1' in cmap else '',
            'a2': text(ws.cell(r, cmap['a2']).value) if 'a2' in cmap else '',
        }
        prev['loc_order'].append(code)


# ---------------------------------------------------------------------------
# Build the item list
# ---------------------------------------------------------------------------
def build_items(stock, prices, scrap, prev, period, checks):
    par = prev['par'] if prev else {}
    notes = prev['notes'] if prev else {}
    first_seen = prev['first_seen'] if prev else {}
    prev_items = prev['items'] if prev else set()
    groups = prev.get('group', {}) if prev else {}

    keys = set(stock)
    keys |= {k for k, v in par.items() if v > 0}
    keys |= set(notes)
    keys |= set(scrap)
    for k, v in par.items():
        if k not in keys:   # par 0, no note, gone from stock, nothing scrapped
            desc = prices.get(k[1], {}).get('desc') or prev['desc'].get(k, '')
            checks.append(('Removed (par 0, no stock left)', k[0], k[1], desc,
                           'Par was 0 and the item is no longer in MB52, so it was dropped'))

    items = []
    for key in keys:
        loc, mat = key
        st = stock.get(key)
        mm = prices.get(mat)
        sc = scrap.get(key)
        desc = (mm or {}).get('desc') or (st or {}).get('desc') or \
            (sc or {}).get('desc') or (prev['desc'].get(key, '') if prev else '')
        if desc.startswith('='):
            desc = ' ' + desc
        price = (mm or {}).get('price')
        qty = st['qty'] if st else 0.0
        hold = (st['value'] if st and st['value'] is not None else
                (qty * price if price else 0.0))
        p = par.get(key)
        flags = []
        dnu = bool(mm and mm['dnu'])
        if not st and p and p > 0:
            flags.append('Not in MB52 - stock 0')
            checks.append(('Par item not in MB52 (stock 0)', loc, mat, desc,
                           f'Par {clean_num(p)}; treated as zero stock and ordered in full'))
        if not st and sc:
            flags.append('Scrapped out - none left')
            checks.append(('Scrapped out (none left)', loc, mat, desc,
                           f'{clean_num(sc["qty"])} scrapped; item no longer in MB52 here'))
        if not price:
            last = prev['price'].get(mat) if prev else None
            what = 'not in MM60' if mm is None else 'R0 in MM60'
            if last:
                price = last
                flags.append(f'Price {what} - last known price used')
                checks.append(('MM60 price missing or R0', loc, mat, desc,
                               f'Price {what}; used R{last:,.2f} from the previous workbook'))
            else:
                flags.append(f'Price {what}')
                checks.append(('MM60 price missing or R0', loc, mat, desc,
                               f'Price {what} and no earlier price; order value shows as 0'))
        if dnu:
            flags.append('Do Not Use (MM60) - not ordered')
            checks.append(('Do Not Use in MM60', loc, mat, desc, 'Excluded from orders and approvals'))
        if st and price and qty:
            unit = (st['value'] or 0) / qty
            if price and abs(unit - price) / price > PRICE_DIFF_CHECK:
                checks.append(('MB52 value vs MM60 price', loc, mat, desc,
                               f'MB52 R{unit:,.2f} per unit vs MM60 R{price:,.2f}'))
        if sc and st and sc['units'] and st['uom'] and sc['units'] != {st['uom'].upper()}:
            checks.append(('MB51 unit differs from MB52', loc, mat, desc,
                           f'MB51 unit {"/".join(sorted(sc["units"]))} vs MB52 {st["uom"]}; '
                           f'scrap qty may be in a different unit'))
        fs = first_seen.get(key, '')
        if prev and key not in prev_items:
            fs = period
        if prev and st and fs == period and p is None:
            flags.append('New this month')
            checks.append(('New item (no par yet)', loc, mat, desc,
                           f'Stock {clean_num(qty)}; set a Par Level if this item should be kept stocked'))
        items.append({
            'loc': loc, 'mat': mat, 'desc': desc, 'price': price, 'stock': qty,
            'hold': hold, 'scrap_qty': sc['qty'] if sc else 0.0,
            'scrap_val': sc['value'] if sc else 0.0, 'par': p,
            'flags': flags, 'note': notes.get(key, ''),
            'first_seen': fs, 'dnu': dnu, 'group': groups.get(key),
        })

    # Par groups: several SAP codes held against one par. The par sits on the
    # order code line; members carry the order code in "Par Group". The order
    # line is compared with the combined stock of the group.
    by_key = {(i['loc'], i['mat']): i for i in items}
    members = defaultdict(list)
    for i in items:
        if i['group']:
            members[(i['loc'], i['group'])].append(i)
    for (loc, oc), mem in members.items():
        head = by_key.get((loc, oc))
        if head is None or head['group'] != oc:
            for i in mem:
                i['group'] = None
                i['flags'].append(f'Par group {oc} not found - stock counted on its own')
                checks.append(('Par group order code not found', loc, i['mat'], i['desc'],
                               f'Par Group {oc} has no line of its own on this tab with Par Group {oc}; '
                               f'the line is treated as ungrouped'))
            continue
        others = [i for i in mem if i is not head]
        head['basis'] = sum(i['stock'] for i in mem)
        head['group_text'] = ' + '.join(f'{i["mat"]} ({clean_num(i["stock"]):,})' for i in mem)
        if others:
            head['flags'].append(f'Group par: stock = {head["group_text"]}')
        for i in others:
            i['flags'].append(f'In par group {oc} - ordered on {oc}')
            if i['par'] and i['par'] > 0:
                checks.append(('Par on a group member', loc, i['mat'], i['desc'],
                               f'Par {clean_num(i["par"])} set on a member of group {oc}; '
                               f'it still orders on its own stock. Clear it or the group may over-order'))
    for i in items:
        basis = i.get('basis', i['stock'])
        i['basis'] = basis
        i['group_text'] = i.get('group_text', '')
        oq = 0 if i['dnu'] else order_qty(i['par'], basis)
        i['order'] = oq
        i['order_val'] = oq * (i['price'] or 0.0)
        i['urgency'] = '' if i['dnu'] or oq == 0 else urgency(i['par'], basis)
        i['flag'] = '; '.join(i.pop('flags'))
    items.sort(key=lambda d: (d['loc'], d['desc'].lower(), d['mat']))
    return items


def build_locations(items, prev, checks):
    """Location settings: previous LOCATIONS tab first, defaults for new codes."""
    prev_locs = prev['locations'] if prev else {}
    codes = sorted({i['loc'] for i in items} | set(prev_locs))
    locs = {}
    for c in codes:
        if c in prev_locs:
            locs[c] = dict(prev_locs[c])
            if not locs[c]['name']:
                locs[c]['name'] = DEFAULT_LOCATION_NAMES.get(c, '')
        else:
            locs[c] = {'name': DEFAULT_LOCATION_NAMES.get(c, ''), 'appr': 'Y', 'notes': '',
                       'file': '', 'a1': '', 'a2': ''}
        if not locs[c]['name'] and any(i['loc'] == c for i in items):
            checks.append(('Location has no name', c, None, '',
                           'Add a name on the LOCATIONS tab (used on approvals and transfers)'))
    return locs


def transfer_sources(items, locs):
    """For each material: locations holding stock above a par level (> 0)."""
    excess = defaultdict(list)
    for i in items:
        if i['par'] and i['par'] > 0 and i['basis'] > i['par']:
            excess[i['mat']].append((i['loc'], i['basis'] - i['par']))
    for m in excess:
        excess[m].sort(key=lambda t: -t[1])
    return excess


def transfer_text(item, excess, locs):
    parts = []
    for loc, ex in excess.get(item['mat'], []):
        if loc == item['loc']:
            continue
        label = f'{loc} {locs.get(loc, {}).get("name", "")}'.strip()
        parts.append(f'{label}: {clean_num(ex):,} over par')
        if len(parts) == MAX_TRANSFER_SOURCES:
            break
    return '; '.join(parts)


# ---------------------------------------------------------------------------
# Writing helpers
# ---------------------------------------------------------------------------
def title_bar(ws, ref, value, size=14):
    ws.merge_cells(ref)
    c = ws[ref.split(':')[0]]
    c.value = value
    c.font = Font(bold=True, size=size, color='FFFFFF')
    c.fill = fill(C_TITLE)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[c.row].height = 26


def section_bar(ws, row, first_col, last_col, value):
    for col in range(first_col, last_col + 1):
        ws.cell(row, col).fill = fill(C_SECTION)
    c = ws.cell(row, first_col, value)
    c.font = Font(bold=True, color='FFFFFF', size=11)


def header_row(ws, row, headers, first_col=1, height=32):
    for j, h in enumerate(headers):
        c = ws.cell(row, first_col + j, h)
        c.font = Font(bold=True, color='FFFFFF', size=10)
        c.fill = fill(C_HEAD)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = height


def link_formula(sheet, label):
    return f'=HYPERLINK("#\'{sheet}\'!A1","{label}")'


def page_setup(ws, landscape=True, title_rows=None):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'landscape' if landscape else 'portrait'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    if title_rows:
        ws.print_title_rows = title_rows
    ws.oddFooter.center.text = f'(c) 2026 Alain Yon | OE Tracker v{VERSION} | Page &P of &N'
    ws.oddFooter.center.size = 8


# ---------------------------------------------------------------------------
# Tracking workbook
# ---------------------------------------------------------------------------
def write_location_sheet(wb, loc, rows, locs, sources_line):
    ws = wb.create_sheet(loc)
    name = locs.get(loc, {}).get('name', '')
    last_col = len(COLS)
    title_bar(ws, f'A1:{get_column_letter(last_col)}1',
              f'{loc}  {name}  |  Operational Equipment'.replace('    ', '  '))
    ws['A2'] = link_formula('SUMMARY', '<< Back to Summary')
    ws['A2'].font = Font(underline='single', color=C_LINK, size=9)
    ws['C2'] = sources_line
    ws['C2'].font = Font(italic=True, size=9, color=C_GREY)
    n_first, n_last = FIRST_ROW, FIRST_ROW + len(rows) - 1

    # Totals above the header (SUBTOTAL so they follow any filter)
    ws.cell(3, 1, 'TOTALS').font = Font(bold=True)
    for col in range(1, last_col + 1):
        ws.cell(3, col).fill = fill(C_TOTAL)
        ws.cell(3, col).border = BORDER
    for col, fmt in ((4, F_QTY), (5, F_MONEY), (6, F_QTY), (7, F_MONEY),
                     (9, F_QTY), (10, F_MONEY)):
        L = get_column_letter(col)
        c = ws.cell(3, col, f'=SUBTOTAL(9,{L}{n_first}:{L}{n_last})')
        c.number_format, c.font = fmt, Font(bold=True)
    c = ws.cell(3, 11, f'=COUNTIF(K{n_first}:K{n_last},"HIGH")&" HIGH"')
    c.font = Font(bold=True, color=C_HIGH_FONT)
    c.alignment = Alignment(horizontal='center')
    header_row(ws, HDR_ROW, COLS)

    band = fill(C_BAND)
    inp = fill(C_INPUT)
    for idx, it in enumerate(rows):
        r = FIRST_ROW + idx
        has_scrap = bool(it['scrap_qty'] or it['scrap_val'])
        vals = [it['mat'], it['desc'], it['price'], clean_num(it['stock']),
                round(it['hold'], 2),
                clean_num(it['scrap_qty']) if has_scrap else None,
                round(it['scrap_val'], 2) if has_scrap else None,
                clean_num(it['par'])]
        for col, v in enumerate(vals, 1):
            ws.cell(r, col, v)
        if it['dnu']:
            ws.cell(r, 9, 0)
            ws.cell(r, 10, 0)
            ws.cell(r, 11, '')
        else:
            # stock basis: own stock, or the whole group's stock on an order-code line
            basis = (f'IF(AND(O{r}<>"",N(O{r})=A{r}),'
                     f'SUMIF($O${n_first}:$O${n_last},A{r},$D${n_first}:$D${n_last}),D{r})')
            ws.cell(r, 9, f'=IF(H{r}="","",ROUNDUP(ROUND(MAX(H{r}-{basis},0),4),0))')
            ws.cell(r, 10, f'=IF(I{r}="","",I{r}*C{r})')
            ws.cell(r, 11, f'=IF(OR(H{r}="",N(I{r})=0),"",'
                           f'IF({basis}<H{r}*{HIGH_STOCK_RATIO},"HIGH","MEDIUM"))')
        ws.cell(r, 12, it['flag'] or None)
        ws.cell(r, 13, it['note'] or None)
        ws.cell(r, 14, it['first_seen'] or None)
        ws.cell(r, 15, it['group'] or None)
        fmts = {3: F_MONEY, 4: qty_format(it['stock']), 5: F_MONEY,
                6: qty_format(it['scrap_qty']), 7: F_MONEY,
                8: qty_format(it['par']), 9: F_QTY, 10: F_MONEY}
        for col in range(1, last_col + 1):
            c = ws.cell(r, col)
            c.border = BORDER
            if col in fmts:
                c.number_format = fmts[col]
            if col in (8, 13, 15):
                c.fill = inp
            elif idx % 2:
                c.fill = band
        ws.cell(r, 1).alignment = Alignment(horizontal='center')
        ws.cell(r, 11).alignment = Alignment(horizontal='center')
        ws.cell(r, 12).font = Font(size=9, color='C00000')
        ws.cell(r, 14).font = Font(size=9, color=C_GREY)

    rng = f'K{n_first}:K{n_last}'
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"HIGH"'],
                                  fill=fill(C_HIGH_FILL), font=Font(bold=True, color=C_HIGH_FONT)))
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"MEDIUM"'],
                                  fill=fill(C_MED_FILL), font=Font(color=C_MED_FONT)))
    dv = DataValidation(type='decimal', operator='greaterThanOrEqual', formula1='0',
                        allow_blank=True, showErrorMessage=True,
                        errorTitle='Par Level', error='Enter a number of 0 or more, or leave blank.')
    dv.add(f'H{n_first}:H{n_last + 200}')
    ws.add_data_validation(dv)
    dg = DataValidation(type='whole', operator='greaterThan', formula1='0', allow_blank=True,
                        showErrorMessage=True, errorTitle='Par Group',
                        error='Enter the SAP number of the line that carries the group par, or leave blank.')
    dg.add(f'O{n_first}:O{n_last + 200}')
    ws.add_data_validation(dg)

    widths = [10, 42, 12, 11, 14, 10, 12, 11, 10, 13, 10, 26, 30, 10, 13]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = f'C{FIRST_ROW}'
    ws.auto_filter.ref = f'A{HDR_ROW}:{get_column_letter(last_col)}{n_last}'
    page_setup(ws, title_rows=f'{HDR_ROW}:{HDR_ROW}')
    return (n_first, n_last)


def write_locations_sheet(wb, locs, prev_order):
    ws = wb.create_sheet('LOCATIONS')
    title_bar(ws, 'A1:G1', 'LOCATIONS  |  names and approval settings')
    ws['A2'] = ('Yellow cells are yours to edit and are carried into next month. '
                'On Approval = N keeps that location off every approval file '
                '(its tab still calculates orders). Approval File blank = the F&B Manager file. '
                'A name in Approval File (e.g. Linen Main) gives those locations their own file, '
                'signed by Approver 1 then Approver 2 (write "Title: Name").')
    ws['A2'].font = Font(italic=True, size=9)
    ws['A2'].alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells('A2:G2')
    ws.row_dimensions[2].height = 42
    header_row(ws, 4, ['Code', 'Name', 'On Approval (Y/N)', 'Notes', 'Approval File',
                       'Approver 1', 'Approver 2'])
    order = [c for c in prev_order if c in locs] + sorted(c for c in locs if c not in prev_order)
    inp = fill(C_INPUT)
    for i, code in enumerate(order):
        r = 5 + i
        d = locs[code]
        ws.cell(r, 1, code).font = Font(bold=True)
        ws.cell(r, 2, d['name'] or None)
        ws.cell(r, 3, d['appr'])
        ws.cell(r, 4, d['notes'] or None)
        ws.cell(r, 5, d.get('file') or None)
        ws.cell(r, 6, d.get('a1') or None)
        ws.cell(r, 7, d.get('a2') or None)
        for col in range(1, 8):
            ws.cell(r, col).border = BORDER
            if col > 1:
                ws.cell(r, col).fill = inp
        ws.cell(r, 3).alignment = Alignment(horizontal='center')
        ws.cell(r, 4).alignment = Alignment(wrap_text=True, vertical='top')
    last = 4 + len(order)
    dv = DataValidation(type='list', formula1='"Y,N"', allow_blank=False, showErrorMessage=True,
                        errorTitle='On Approval', error='Choose Y or N')
    dv.add(f'C5:C{last + 50}')
    ws.add_data_validation(dv)
    for col, w in zip('ABCDEFG', (8, 26, 12, 60, 14, 34, 34)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A5'
    page_setup(ws, landscape=False, title_rows='4:4')


def write_checks_sheet(wb, checks):
    ws = wb.create_sheet('CHECKS')
    title_bar(ws, 'A1:E1', 'DATA CHECKS  |  read before sending the approval file')
    counts = defaultdict(int)
    for c in checks:
        counts[c[0]] += 1
    r = 3
    ws.cell(r, 1, 'Summary').font = Font(bold=True)
    for t in sorted(counts):
        r += 1
        ws.cell(r, 1, t)
        ws.cell(r, 2, counts[t])
    r += 2
    header_row(ws, r, ['Check', 'Location', 'Material', 'Description', 'Detail'], height=20)
    hdr = r
    for c in sorted(checks, key=lambda x: (x[0], x[1], x[2] or 0)):
        r += 1
        for col, v in enumerate(c, 1):
            cell = ws.cell(r, col, v if v not in ('', None) else None)
            cell.border = BORDER
            cell.alignment = Alignment(vertical='top', wrap_text=(col == 5))
    if not checks:
        ws.cell(hdr + 1, 1, 'No issues found')
    for col, w in zip('ABCDE', (32, 9, 10, 42, 80)):
        ws.column_dimensions[col].width = w
    ws.auto_filter.ref = f'A{hdr}:E{max(r, hdr + 1)}'
    ws.freeze_panes = f'A{hdr + 1}'
    page_setup(ws, title_rows=f'{hdr}:{hdr}')


def write_summary_sheet(ws, period, built, sources, controls, loc_codes, ranges,
                        approval_info, n_checks):
    title_bar(ws, 'A1:K1', f'{PROPERTY_NAME.upper()}  |  OE TRACKING')
    ws['A2'] = f'Period: {period}    |    Built {built:%d %b %Y %H:%M}    |    oe_tracker.py v{VERSION}'
    ws['A2'].font = Font(italic=True, size=9, color=C_GREY)

    section_bar(ws, 4, 1, 11, 'SOURCE FILES')
    rows = [('Stock on hand (MB52)', sources.get('MB52')),
            ('Prices (MM60)', sources.get('MM60')),
            ('Scrap (MB51)', sources.get('MB51'))]
    for i, (label, info) in enumerate(rows):
        r = 5 + i
        ws.cell(r, 1, label).font = Font(bold=True)
        if info:
            ws.cell(r, 3, info['name'])
            ws.cell(r, 7, 'File date')
            ws.cell(r, 8, info['mtime'].strftime('%d %b %Y %H:%M'))
            if info.get('extra'):
                ws.cell(r, 10, info['extra'])
        else:
            ws.cell(r, 3, 'not supplied - scrap shows as 0').font = Font(color='C00000')

    T0, T1 = 21, 21 + len(loc_codes) - 1     # location table rows
    TT = T1 + 1                              # table total row
    section_bar(ws, 9, 1, 11, 'KEY FIGURES')
    kpis = [
        (10, 'Hold value (MB52)', f'=D{TT}', F_RAND,
         'MB52 control total', controls.get('MB52')),
        (11, 'Scrap value (MB51)', f'=E{TT}', F_RAND,
         'MB51 control total', controls.get('MB51')),
        (12, 'Order value - all locations (live)', f'=I{TT}', F_RAND, None, None),
        (13, 'Order value - locations On Approval = Y (live)',
         f'=SUMIFS(I{T0}:I{T1},J{T0}:J{T1},"Y")', F_RAND, None, None),
        (14, 'Order value on the approval files (as built)', approval_info['value'], F_RAND,
         None, None),
        (15, 'Lines to order (live)', f'=G{TT}', F_QTY, 'HIGH urgency lines', f'=H{TT}'),
        (16, 'Par levels set', f'=F{TT}', F_QTY, 'Items tracked', f'=C{TT}'),
        (17, 'Data checks to review', n_checks, F_QTY, None, None),
    ]
    for r, label, val, fmt, label2, val2 in kpis:
        ws.cell(r, 1, label).font = Font(bold=True)
        c = ws.cell(r, 5, val)
        c.number_format, c.font = fmt, Font(bold=True, size=11, color=C_HEAD)
        for col in range(1, 6):
            ws.cell(r, col).fill = fill(C_TOTAL)
        if label2:
            ws.cell(r, 7, label2)
            c2 = ws.cell(r, 9, val2)
            c2.number_format = fmt
            if label2.endswith('control total') and val2 is not None:
                d = ws.cell(r, 10, f'=E{r}-I{r}')
                d.number_format = F_RAND
                ws.cell(r, 11, f'=IF(ABS(J{r})<0.05,"ties","CHECK")').font = Font(bold=True)
            elif label2.endswith('control total'):
                ws.cell(r, 9, 'no total line in file').font = Font(italic=True, color=C_GREY)
    ws.cell(14, 7, approval_info['text']).font = Font(italic=True, size=9)
    ws.cell(17, 7, link_formula('CHECKS', 'Open CHECKS >>')).font = Font(underline='single',
                                                                           color=C_LINK)

    section_bar(ws, 19, 1, 11, 'LOCATIONS')
    header_row(ws, 20, ['Code', 'Name', 'Items', 'Hold Value', 'Scrap Value', 'Pars Set',
                        'Lines to Order', 'HIGH', 'Order Value', 'On Approval', 'Open'])
    for i, loc in enumerate(loc_codes):
        r = T0 + i
        a, b = ranges[loc]
        q = f"'{loc}'!"
        cells = [
            loc,
            f'=IFERROR(INDEX(LOCATIONS!$B:$B,MATCH($A{r},LOCATIONS!$A:$A,0))&"","")',
            f'=COUNTA({q}A{a}:A{b})',
            f'=SUM({q}E{a}:E{b})',
            f'=SUM({q}G{a}:G{b})',
            f'=COUNT({q}H{a}:H{b})',
            f'=COUNTIF({q}I{a}:I{b},">0")',
            f'=COUNTIF({q}K{a}:K{b},"HIGH")',
            f'=SUM({q}J{a}:J{b})',
            f'=IFERROR(INDEX(LOCATIONS!$C:$C,MATCH($A{r},LOCATIONS!$A:$A,0))&"","Y")',
            link_formula(loc, f'{loc} >>'),
        ]
        for col, v in enumerate(cells, 1):
            c = ws.cell(r, col, v)
            c.border = BORDER
            if i % 2:
                c.fill = fill(C_BAND)
        for col, fmt in ((3, F_QTY), (4, F_MONEY), (5, F_MONEY), (6, F_QTY), (7, F_QTY),
                         (8, F_QTY), (9, F_MONEY)):
            ws.cell(r, col).number_format = fmt
        ws.cell(r, 1).font = Font(bold=True)
        ws.cell(r, 10).alignment = Alignment(horizontal='center')
        ws.cell(r, 11).font = Font(underline='single', color=C_LINK)
    ws.cell(TT, 1, 'TOTAL').font = Font(bold=True)
    for col in range(1, 12):
        ws.cell(TT, col).fill = fill(C_TOTAL)
        ws.cell(TT, col).border = BORDER
    for col, fmt in ((3, F_QTY), (4, F_MONEY), (5, F_MONEY), (6, F_QTY), (7, F_QTY),
                     (8, F_QTY), (9, F_MONEY)):
        L = get_column_letter(col)
        c = ws.cell(TT, col, f'=SUM({L}{T0}:{L}{T1})')
        c.number_format, c.font = fmt, Font(bold=True)
    ws.conditional_formatting.add(f'K10:K11', CellIsRule(operator='equal', formula=['"CHECK"'],
                                  fill=fill(C_HIGH_FILL), font=Font(bold=True, color=C_HIGH_FONT)))
    for col, w in zip('ABCDEFGHIJK', (14, 24, 9, 16, 14, 9, 10, 8, 15, 10, 10)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A3'
    page_setup(ws, landscape=False)


def write_tracking_workbook(path, period, built, sources, controls, items, locs,
                            loc_order, checks, approval_info, scrap_range):
    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = 'SUMMARY'
    write_locations_sheet(wb, locs, loc_order)
    write_checks_sheet(wb, checks)
    by_loc = defaultdict(list)
    for it in items:
        by_loc[it['loc']].append(it)
    loc_codes = sorted(by_loc)
    src = f'Stock: {sources["MB52"]["name"]}  |  Prices: {sources["MM60"]["name"]}'
    if sources.get('MB51'):
        src += f'  |  Scrap: {sources["MB51"]["name"]}'
        if scrap_range:
            src += f' ({scrap_range[0]:%d %b} - {scrap_range[1]:%d %b %Y})'
    ranges = {}
    for loc in loc_codes:
        ranges[loc] = write_location_sheet(wb, loc, by_loc[loc], locs, src)
    write_summary_sheet(ws_sum, period, built, sources, controls, loc_codes, ranges,
                        approval_info, len(checks))
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    return loc_codes


# ---------------------------------------------------------------------------
# Approval file
# ---------------------------------------------------------------------------
APPROVAL_COLS = ['Location', 'Location Name', 'Material', 'Description', 'Stock', 'Par Level',
                 'Suggested Qty', 'Unit Price', 'Suggested Value', 'Urgency',
                 'Transfer Option (over par elsewhere)', 'Approve (Y/N)', 'Approved Qty',
                 'Approved Value', 'Notes']
APPROVAL_HDR = 8


def approval_lines(items, locs):
    """Lines per approval file. Key '' is the F&B Manager file; any other key is
    the Approval File name from the LOCATIONS tab (e.g. 'Linen Main')."""
    excess = transfer_sources(items, locs)
    files, excluded = defaultdict(list), defaultdict(lambda: [0, 0.0])
    for it in items:
        if it['order'] <= 0 or it['dnu']:
            continue
        loc = locs.get(it['loc'], {})
        if loc.get('appr', 'Y') != 'Y':
            excluded[it['loc']][0] += 1
            excluded[it['loc']][1] += it['order_val']
            continue
        files[loc.get('file', '')].append(dict(it, transfer=transfer_text(it, excess, locs),
                                               loc_name=loc.get('name', '')))
    rank = {'HIGH': 0, 'MEDIUM': 1, '': 2}
    for lines in files.values():
        lines.sort(key=lambda d: (d['loc'], rank.get(d['urgency'], 2), -d['order_val'], d['desc']))
    files.setdefault('', [])
    return dict(files), dict(excluded)


def approval_file_name(label, period):
    if label:
        return f'OE_Approval_{period_slug(label)}_{period_slug(period)}.xlsx'
    return f'OE_Approval_{period_slug(period)}.xlsx'


def write_approval_file(path, period, built, sources, lines, excluded, locs, separate=None):
    wb = Workbook()
    ws = wb.active
    ws.title = 'APPROVAL'
    nc = len(APPROVAL_COLS)
    title_bar(ws, f'A1:{get_column_letter(nc)}1', f'OE ORDER APPROVAL  |  {period}  |  {PROPERTY_NAME}')
    ws['A2'] = (f'Prepared by Cost Control on {built:%d %b %Y}.  Stock: {sources["MB52"]["name"]}.  '
                f'Prices: {sources["MM60"]["name"]}.')
    ws['A2'].font = Font(italic=True, size=9, color=C_GREY)
    ws['A3'] = ('HOW TO APPROVE:  1) Choose Y or N in "Approve (Y/N)".   2) "Approved Qty" is filled '
                'with the suggested quantity; change it if you want fewer or more.   3) Notes are '
                'optional.   4) Sign at the bottom, save, and send the file back.  Where a Transfer '
                'Option is shown, that stock could be moved instead of bought.')
    ws['A3'].alignment = Alignment(wrap_text=True, vertical='top')
    ws['A3'].font = Font(size=10)
    ws.merge_cells(f'A3:{get_column_letter(nc)}3')
    ws.row_dimensions[3].height = 42
    n0 = APPROVAL_HDR + 1
    n1 = APPROVAL_HDR + max(len(lines), 1)

    ws['A5'] = 'Lines'
    ws['D5'] = len(lines)
    ws['A6'] = 'Suggested value'
    ws['D6'] = f'=SUM(I{n0}:I{n1})'
    ws['E5'] = 'Decided (Y or N)'
    ws['H5'] = f'=COUNTIF(L{n0}:L{n1},"Y")+COUNTIF(L{n0}:L{n1},"N")'
    ws['E6'] = 'Approved value'
    ws['H6'] = f'=SUM(N{n0}:N{n1})'
    for ref in ('A5', 'A6', 'E5', 'E6'):
        ws[ref].font = Font(bold=True)
    for ref in ('D5', 'H5'):
        ws[ref].number_format = F_QTY
    for ref in ('D6', 'H6'):
        ws[ref].number_format = F_RAND
    for ref in ('D5', 'D6', 'H5', 'H6'):
        ws[ref].font = Font(bold=True, color=C_HEAD, size=11)
        ws[ref].alignment = Alignment(horizontal='left')
    ws.merge_cells('H5:I5')
    ws.merge_cells('H6:I6')
    notes_txt = []
    if excluded:
        notes_txt.append('Not on this file (On Approval = N): ' + '; '.join(
            f'{l} {locs.get(l, {}).get("name", "")} ({n} lines, R{v:,.2f})'
            for l, (n, v) in sorted(excluded.items())))
    if separate:
        notes_txt.append('On their own approval file: ' + '; '.join(
            f'{lab} ({", ".join(sorted(codes))}: {n} lines, R{v:,.2f})'
            for lab, (codes, n, v) in sorted(separate.items())))
    if notes_txt:
        txt = '   '.join(notes_txt)
        ws['K5'] = txt
        ws['K5'].font = Font(italic=True, size=9, color=C_GREY)
        ws['K5'].alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(f'K5:{get_column_letter(nc)}6')

    header_row(ws, APPROVAL_HDR, APPROVAL_COLS, height=34)
    inp = fill(C_INPUT)
    for i, d in enumerate(lines):
        r = n0 + i
        vals = [d['loc'], d['loc_name'], d['mat'], d['desc'], clean_num(d['stock']),
                clean_num(d['par']), d['order'], d['price'], f'=G{r}*H{r}', d['urgency'],
                d['transfer'] or None, None, d['order'],
                f'=IF(UPPER(TRIM(L{r}))="Y",N(M{r})*H{r},0)', None]
        for col, v in enumerate(vals, 1):
            c = ws.cell(r, col, v)
            c.border = BORDER
            if col in (12, 13, 15):
                c.fill = inp
            elif i % 2:
                c.fill = fill(C_BAND)
        for col, fmt in ((5, qty_format(d['stock'])), (6, qty_format(d['par'])), (7, F_QTY),
                         (8, F_MONEY), (9, F_MONEY), (13, F_QTY), (14, F_MONEY)):
            ws.cell(r, col).number_format = fmt
        for col in (1, 3, 10, 12):
            ws.cell(r, col).alignment = Alignment(horizontal='center')
        ws.cell(r, 11).font = Font(size=9, color='375623')
        ws.cell(r, 11).alignment = Alignment(wrap_text=True, vertical='top')
    if not lines:
        ws.cell(n0, 1, 'No orders this month')

    rng = f'J{n0}:J{n1}'
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"HIGH"'],
                                  fill=fill(C_HIGH_FILL), font=Font(bold=True, color=C_HIGH_FONT)))
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"MEDIUM"'],
                                  fill=fill(C_MED_FILL), font=Font(color=C_MED_FONT)))
    dv = DataValidation(type='list', formula1='"Y,N"', allow_blank=True, showErrorMessage=True,
                        errorTitle='Approve', error='Choose Y or N (or leave blank if undecided)')
    dv.add(f'L{n0}:L{n1}')
    ws.add_data_validation(dv)
    dq = DataValidation(type='whole', operator='greaterThanOrEqual', formula1='0',
                        allow_blank=True, showErrorMessage=True,
                        errorTitle='Approved Qty', error='Enter a whole number of 0 or more')
    dq.add(f'M{n0}:M{n1}')
    ws.add_data_validation(dq)

    s = n1 + 3
    for k, who in enumerate(('Approved by (F&B Manager)', 'Prepared by (Cost Controller)')):
        r = s + k * 2
        ws.cell(r, 1, who).font = Font(bold=True)
        for col, lab in ((4, 'Name:'), (8, 'Signature:'), (12, 'Date:')):
            ws.cell(r, col, lab).alignment = Alignment(horizontal='right')
            for cc in range(col + 1, col + 3):
                ws.cell(r, cc).border = Border(bottom=Side(style='thin', color='000000'))
        ws.row_dimensions[r].height = 24

    widths = [8, 18, 9, 40, 8, 8, 10, 11, 13, 9, 30, 10, 10, 13, 26]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = f'E{n0}'
    ws.auto_filter.ref = f'A{APPROVAL_HDR}:{get_column_letter(nc)}{n1}'
    page_setup(ws, title_rows=f'{APPROVAL_HDR}:{APPROVAL_HDR}')

    # By-location view (live)
    wl = wb.create_sheet('BY LOCATION')
    title_bar(wl, 'A1:F1', f'APPROVAL BY LOCATION  |  {period}')
    header_row(wl, 3, ['Location', 'Name', 'Lines', 'Suggested Value', 'Approved Lines',
                       'Approved Value'])
    codes = sorted({d['loc'] for d in lines})
    for i, loc in enumerate(codes):
        r = 4 + i
        rng_a = f'APPROVAL!$A${n0}:$A${n1}'
        wl.cell(r, 1, loc).font = Font(bold=True)
        wl.cell(r, 2, locs.get(loc, {}).get('name', ''))
        wl.cell(r, 3, f'=COUNTIF({rng_a},A{r})')
        wl.cell(r, 4, f'=SUMIFS(APPROVAL!$I${n0}:$I${n1},{rng_a},A{r})')
        wl.cell(r, 5, f'=COUNTIFS({rng_a},A{r},APPROVAL!$L${n0}:$L${n1},"Y")')
        wl.cell(r, 6, f'=SUMIFS(APPROVAL!$N${n0}:$N${n1},{rng_a},A{r})')
        for col, fmt in ((3, F_QTY), (4, F_MONEY), (5, F_QTY), (6, F_MONEY)):
            wl.cell(r, col).number_format = fmt
        for col in range(1, 7):
            wl.cell(r, col).border = BORDER
    tr = 4 + len(codes)
    wl.cell(tr, 1, 'TOTAL').font = Font(bold=True)
    for col, fmt in ((3, F_QTY), (4, F_MONEY), (5, F_QTY), (6, F_MONEY)):
        L = get_column_letter(col)
        c = wl.cell(tr, col, f'=SUM({L}4:{L}{tr - 1})' if codes else 0)
        c.number_format, c.font = fmt, Font(bold=True)
    for col in range(1, 7):
        wl.cell(tr, col).fill = fill(C_TOTAL)
    for col, w in zip('ABCDEF', (10, 24, 8, 16, 14, 16)):
        wl.column_dimensions[col].width = w
    page_setup(wl, landscape=False)

    info = wb.create_sheet('_INFO')
    for r, (k, v) in enumerate((('File', 'OE_APPROVAL'), ('Version', VERSION), ('Period', period),
                                ('Built', built.strftime('%Y-%m-%d %H:%M:%S')),
                                ('Lines', len(lines)),
                                ('MB52', sources['MB52']['name']),
                                ('MM60', sources['MM60']['name'])), 1):
        info.cell(r, 1, k)
        info.cell(r, 2, v)
    info.sheet_state = 'hidden'
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


TWO_STEP_COLS = ['Location', 'Location Name', 'Material', 'Description', 'Stock', 'Par Level',
                 'Suggested Qty', 'Unit Price', 'Suggested Value', 'Urgency',
                 'Transfer Option (over par elsewhere)', 'Tracker Notes',
                 'Step 1 Approve (Y/N)', 'Step 1 Qty', 'Step 2 Approve (Y/N)', 'Approved Qty',
                 'Approved Value', 'Notes']


def write_two_step_approval(path, period, built, sources, lines, label, a1, a2, locs):
    """Approval file for one Approval File group (e.g. Linen Main): Approver 1 decides
    first (Step 1), Approver 2 confirms (Step 2). A line is ordered only with Y in both."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'APPROVAL'
    nc = len(TWO_STEP_COLS)
    title_bar(ws, f'A1:{get_column_letter(nc)}1',
              f'OE ORDER APPROVAL  |  {label}  |  {period}  |  {PROPERTY_NAME}')
    ws['A2'] = (f'Prepared by Cost Control on {built:%d %b %Y}.  Stock: {sources["MB52"]["name"]}.  '
                f'Prices: {sources["MM60"]["name"]}.  Step 1: {a1}.  Step 2: {a2}.')
    ws['A2'].font = Font(italic=True, size=9, color=C_GREY)
    ws['A3'] = (f'HOW TO APPROVE:  1) {a1} chooses Y or N in "Step 1 Approve" and can change '
                f'"Step 1 Qty" (filled with the suggested qty), then passes the file on.   '
                f'2) {a2} chooses Y or N in "Step 2 Approve". "Approved Qty" follows Step 1 Qty; '
                f'type over it to change it.   3) Only lines with Y in BOTH steps are ordered.   '
                f'4) Both sign at the bottom; save and send the file back to Cost Control.  '
                f'Stock on a group line is the combined stock of every code in the group.')
    ws['A3'].alignment = Alignment(wrap_text=True, vertical='top')
    ws['A3'].font = Font(size=10)
    ws.merge_cells(f'A3:{get_column_letter(nc)}3')
    ws.row_dimensions[3].height = 56
    hdr = APPROVAL_HDR
    n0, n1 = hdr + 1, hdr + max(len(lines), 1)

    ws['A5'] = 'Lines'
    ws['D5'] = len(lines)
    ws['A6'] = 'Suggested value'
    ws['D6'] = f'=SUM(I{n0}:I{n1})'
    ws['E5'] = 'Decided: Step 1 / Step 2'
    ws['H5'] = (f'=(COUNTIF(M{n0}:M{n1},"Y")+COUNTIF(M{n0}:M{n1},"N"))&" / "&'
                f'(COUNTIF(O{n0}:O{n1},"Y")+COUNTIF(O{n0}:O{n1},"N"))')
    ws['E6'] = 'Approved value (Y in both)'
    ws['H6'] = f'=SUM(Q{n0}:Q{n1})'
    for ref in ('A5', 'A6', 'E5', 'E6'):
        ws[ref].font = Font(bold=True)
    ws['D5'].number_format = F_QTY
    for ref in ('D6', 'H6'):
        ws[ref].number_format = F_RAND
    for ref in ('D5', 'D6', 'H5', 'H6'):
        ws[ref].font = Font(bold=True, color=C_HEAD, size=11)
        ws[ref].alignment = Alignment(horizontal='left')
    ws.merge_cells('H5:I5')
    ws.merge_cells('H6:I6')
    # who signs which step, above the decision columns
    for first, last, who in ((13, 14, f'STEP 1: {a1}'), (15, 16, f'STEP 2: {a2}')):
        ws.merge_cells(start_row=hdr - 1, start_column=first, end_row=hdr - 1, end_column=last)
        c = ws.cell(hdr - 1, first, who)
        c.font = Font(bold=True, color='FFFFFF', size=9)
        c.fill = fill(C_SECTION)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[hdr - 1].height = 28

    header_row(ws, hdr, TWO_STEP_COLS, height=34)
    inp = fill(C_INPUT)
    for i, d in enumerate(lines):
        r = n0 + i
        tnote = d['note'] or ''
        if d.get('group_text') and ' + ' in d['group_text']:
            tnote = f'Group stock: {d["group_text"]}.' + (f' {tnote}' if tnote else '')
        vals = [d['loc'], d['loc_name'], d['mat'], d['desc'], clean_num(d['basis']),
                clean_num(d['par']), d['order'], d['price'], f'=G{r}*H{r}', d['urgency'],
                d['transfer'] or None, tnote or None, None, d['order'], None,
                f'=IF(N{r}="","",N{r})',
                (f'=IF(AND(OR(UPPER(TRIM(M{r}))="Y",UPPER(TRIM(M{r}))="YES"),'
                 f'OR(UPPER(TRIM(O{r}))="Y",UPPER(TRIM(O{r}))="YES")),N(P{r})*H{r},0)'), None]
        for col, v in enumerate(vals, 1):
            c = ws.cell(r, col, v)
            c.border = BORDER
            if col in (13, 14, 15, 16, 18):
                c.fill = inp
            elif i % 2:
                c.fill = fill(C_BAND)
        for col, fmt in ((5, qty_format(d['basis'])), (6, qty_format(d['par'])), (7, F_QTY),
                         (8, F_MONEY), (9, F_MONEY), (14, F_QTY), (16, F_QTY), (17, F_MONEY)):
            ws.cell(r, col).number_format = fmt
        for col in (1, 3, 10, 13, 15):
            ws.cell(r, col).alignment = Alignment(horizontal='center')
        ws.cell(r, 11).font = Font(size=9, color='375623')
        ws.cell(r, 11).alignment = Alignment(wrap_text=True, vertical='top')
        ws.cell(r, 12).font = Font(size=9)
        ws.cell(r, 12).alignment = Alignment(wrap_text=True, vertical='top')
    if not lines:
        ws.cell(n0, 1, 'No orders this month')

    rng = f'J{n0}:J{n1}'
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"HIGH"'],
                                  fill=fill(C_HIGH_FILL), font=Font(bold=True, color=C_HIGH_FONT)))
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"MEDIUM"'],
                                  fill=fill(C_MED_FILL), font=Font(color=C_MED_FONT)))
    dv = DataValidation(type='list', formula1='"Y,N"', allow_blank=True, showErrorMessage=True,
                        errorTitle='Approve', error='Choose Y or N (or leave blank if undecided)')
    dv.add(f'M{n0}:M{n1}')
    dv.add(f'O{n0}:O{n1}')
    ws.add_data_validation(dv)
    dq = DataValidation(type='whole', operator='greaterThanOrEqual', formula1='0',
                        allow_blank=True, showErrorMessage=True,
                        errorTitle='Quantity', error='Enter a whole number of 0 or more')
    dq.add(f'N{n0}:N{n1}')
    dq.add(f'P{n0}:P{n1}')
    ws.add_data_validation(dq)

    s = n1 + 3
    for k, who in enumerate((f'Step 1 approved by ({a1})', f'Step 2 approved by ({a2})',
                             'Prepared by (Cost Controller)')):
        r = s + k * 2
        ws.cell(r, 1, who).font = Font(bold=True)
        for col, lab in ((5, 'Name:'), (9, 'Signature:'), (13, 'Date:')):
            ws.cell(r, col, lab).alignment = Alignment(horizontal='right')
            for cc in range(col + 1, col + 3):
                ws.cell(r, cc).border = Border(bottom=Side(style='thin', color='000000'))
        ws.row_dimensions[r].height = 24

    widths = [8, 16, 9, 36, 8, 8, 9, 10, 12, 9, 22, 34, 9, 9, 9, 10, 12, 24]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = f'E{n0}'
    ws.auto_filter.ref = f'A{hdr}:{get_column_letter(nc)}{n1}'
    page_setup(ws, title_rows=f'{hdr}:{hdr}')

    info = wb.create_sheet('_INFO')
    for r, (k, v) in enumerate((('File', 'OE_APPROVAL'), ('Version', VERSION), ('Period', period),
                                ('Built', built.strftime('%Y-%m-%d %H:%M:%S')),
                                ('Lines', len(lines)),
                                ('MB52', sources['MB52']['name']),
                                ('MM60', sources['MM60']['name']),
                                ('Label', label), ('Approver 1', a1), ('Approver 2', a2),
                                ('Steps', 2)), 1):
        info.cell(r, 1, k)
        info.cell(r, 2, v)
    info.sheet_state = 'hidden'
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


# ---------------------------------------------------------------------------
# OE transfer form (for operations)
# ---------------------------------------------------------------------------
TF_FIRST, TF_LINES = 10, 20          # item lines start on row 10, 20 lines
TF_FONT = 'Arial'
TF_PASSWORD = 'DEMO'                # sheet and workbook-structure password
TF_BAD_QTY = 'Qty not valid'


def tf_label(desc, mat):
    """Drop-down text for an item. Excel treats * ? ~ as wildcards in MATCH and
    SUMIFS, and a leading = would turn the cell into a formula, so those are
    replaced or removed."""
    lab = re.sub(r'[*?~]', '-', (desc or '').strip()).lstrip('= ').strip()
    return lab or str(mat)


def write_transfer_form(path, built, sources, items, locs):
    """One-page transfer form. The item list follows the From outlet (MB52);
    prices come from MM60. MB52 stock is kept only in the locked, very hidden
    ITEMS sheet (used for the 'Above stock' flag); it is never shown on the form."""
    from openpyxl.workbook.defined_name import DefinedName

    def add_name(name, ref):
        dn = DefinedName(name, attr_text=ref)
        try:
            wb.defined_names[name] = dn         # openpyxl 3.1 and later
        except (AttributeError, TypeError):
            wb.defined_names.append(dn)         # openpyxl 3.0
    wb = Workbook()
    ws = wb.active
    ws.title = 'TRANSFER'
    data = wb.create_sheet('ITEMS')

    # ---- hidden ITEMS sheet: key | location | SAP no | label | unit price ----
    by_loc = defaultdict(list)
    for it in items:
        if it['stock'] and it['stock'] > 0 and not it['dnu']:
            by_loc[it['loc']].append(it)
    for c, h in enumerate(['Key', 'Location', 'SAP No', 'Item', 'Unit Price', '',
                           'Message', '', 'From list', 'To list', 'MB52 stock'], 1):
        data.cell(1, c, h).font = Font(bold=True, name=TF_FONT)
    r = 2
    ranges = {}
    for loc in sorted(by_loc):
        rows = sorted(by_loc[loc], key=lambda d: (d['desc'].lower(), d['mat']))
        seen = defaultdict(int)
        for it in rows:
            seen[tf_label(it['desc'], it['mat']).lower()] += 1
        start = r
        for it in rows:
            label = tf_label(it['desc'], it['mat'])
            if seen[label.lower()] > 1:
                label = f'{label} ({it["mat"]})'
            price = it['price']
            if not price and it['stock']:
                price = round(it['hold'] / it['stock'], 2) if it['hold'] else None
            data.cell(r, 1, f'{loc}|{label}')
            data.cell(r, 2, loc)
            data.cell(r, 3, it['mat'])
            data.cell(r, 4, label)
            data.cell(r, 5, round(price, 2) if price else 0)
            data.cell(r, 11, it['stock'])
            r += 1
        ranges[loc] = (start, r - 1)
        add_name(f'L_{loc}', f'ITEMS!$D${start}:$D${r - 1}')
    data.cell(2, 7, 'Choose the From outlet first')
    add_name('L_NONE', 'ITEMS!$G$2')
    from_codes = sorted(by_loc)
    to_codes = sorted(locs)
    label_of = lambda c: f'{c} {locs.get(c, {}).get("name", "")}'.strip()
    for i, c in enumerate(from_codes):
        data.cell(2 + i, 9, label_of(c))
    for i, c in enumerate(to_codes):
        data.cell(2 + i, 10, label_of(c))
    add_name('FROM_LIST', f'ITEMS!$I$2:$I${1 + max(len(from_codes), 1)}')
    add_name('TO_LIST', f'ITEMS!$J$2:$J${1 + max(len(to_codes), 1)}')
    for col, w in zip('ABCDEFGHIJK', (30, 9, 9, 42, 11, 2, 28, 2, 26, 26, 12)):
        data.column_dimensions[col].width = w
    data.sheet_state = 'veryHidden'     # cannot be unhidden from the Excel menu
    data.protection.sheet = True
    data.protection.password = TF_PASSWORD

    # ---- the form ----
    f_reg = Font(name=TF_FONT, size=10)
    f_bold = Font(name=TF_FONT, size=10, bold=True)
    f_small = Font(name=TF_FONT, size=8, color=C_GREY)
    inp = fill(C_INPUT)
    unlocked = Protection(locked=False)
    line = Side(style='thin', color='000000')
    box = Border(left=line, right=line, top=line, bottom=line)
    under = Border(bottom=line)
    ws.sheet_view.showGridLines = False
    for col, w in zip('ABCDEFG', (5, 11, 45, 10, 12, 15, 13)):
        ws.column_dimensions[col].width = w

    ws.merge_cells('A1:F1')
    ws['A1'] = f'OE TRANSFER FORM  |  {PROPERTY_NAME}'
    ws['A1'].font = Font(name=TF_FONT, size=14, bold=True, color='FFFFFF')
    ws['A1'].fill = fill(C_TITLE)
    ws['A1'].alignment = Alignment(vertical='center', indent=1)
    ws.row_dimensions[1].height = 28
    ws.merge_cells('A2:F2')
    ws['A2'] = (f'Operating equipment moved between storage locations.  Item lists: MB52 '
                f'{sources["MB52"]["mtime"]:%d %b %Y}.  Prices: MM60 {sources["MM60"]["mtime"]:%d %b %Y}.  '
                f'Form built {built:%d %b %Y}.')
    ws['A2'].font = f_small

    def label(ref, value, merge=None):
        if merge:
            ws.merge_cells(merge)
        ws[ref] = value
        ws[ref].font = f_bold
        ws[ref].alignment = Alignment(vertical='center')

    def input_cell(ref, merge=None, fmt=None, align='left'):
        if merge:
            ws.merge_cells(merge)
            cells = [c for row in ws[merge] for c in row]
        else:
            cells = [ws[ref]]
        for c in cells:
            c.fill = inp
            c.border = box
            c.protection = unlocked
            c.font = f_reg
            c.alignment = Alignment(horizontal=align, vertical='center')
            if fmt:
                c.number_format = fmt

    label('A4', 'TRACKING NO.', 'A4:B4')
    ws.merge_cells('C4:F4')
    ws['C4'] = ('=IF(OR($C$5="",$C$6="",$E$5=""),"Fills in when From, To and Date are entered",'
                'IF(LEFT($C$5,4)=LEFT($C$6,4),"From and To must be different",'
                'IFERROR("OET-"&LEFT($C$5,4)&"-"&LEFT($C$6,4)&"-"&RIGHT(YEAR($E$5),2)'
                '&TEXT(MONTH($E$5),"00")&TEXT(DAY($E$5),"00")&"-"&TEXT(MAX(1,N($E$6)),"00"),'
                '"Check the date")))')
    ws['C4'].font = Font(name=TF_FONT, size=13, bold=True, color=C_HEAD)
    ws['C4'].alignment = Alignment(vertical='center')
    for c in ws['C4:F4'][0]:
        c.border = box
    ws.row_dimensions[4].height = 24

    label('A5', 'From outlet', 'A5:B5')
    input_cell('C5')
    label('D5', 'Date')
    input_cell('E5', 'E5:F5', fmt='dd/mm/yyyy')
    label('A6', 'To outlet', 'A6:B6')
    input_cell('C6')
    label('D6', 'Seq.')
    input_cell('E6', fmt='00', align='center')
    ws['E6'] = 1
    ws['F6'] = '2 = 2nd same day'
    ws['F6'].font = f_small
    label('A7', 'Reason', 'A7:B7')
    input_cell('C7', 'C7:F7')
    for rr in (5, 6, 7):
        ws.row_dimensions[rr].height = 20

    ws.merge_cells('A8:F8')
    last = TF_FIRST + TF_LINES - 1
    ws['A8'] = (f'=IF($C$5="","Choose the From outlet first: the item list depends on it.",'
                f'IF(AND($C$6<>"",LEFT($C$5,4)=LEFT($C$6,4)),'
                f'"From and To are the same outlet: choose a different To outlet.",'
                f'IF(COUNTIF($B${TF_FIRST}:$B${last},"Check")>0,'
                f'"An item marked Check is not held by the From outlet: pick it again.",'
                f'IF(SUMPRODUCT(($C${TF_FIRST}:$C${last}="")*($D${TF_FIRST}:$D${last}<>""))>0,'
                f'"A Qty has been typed on a line with no item.",'
                f'IF(COUNTIF($G${TF_FIRST}:$G${last},"{TF_BAD_QTY}")>0,'
                f'"A Qty is not a whole number of 1 or more: retype it.",'
                f'IF(COUNTIF($G${TF_FIRST}:$G${last},"Above stock")>0,'
                f'"Qty is above the last MB52 stock on "&COUNTIF($G${TF_FIRST}:$G${last},"Above stock")'
                f'&" line(s). Check the physical count before moving it.",""))))))')
    ws['A8'].font = Font(name=TF_FONT, size=9, bold=True, color='C00000')

    hdr = TF_FIRST - 1
    for c, h in enumerate(['#', 'SAP No', 'Description  (pick from list)', 'Qty',
                           'Unit Price', 'Value'], 1):
        cell = ws.cell(hdr, c, h)
        cell.font = Font(name=TF_FONT, size=10, bold=True, color='FFFFFF')
        cell.fill = fill(C_HEAD)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = box
    ws.row_dimensions[hdr].height = 20
    ws.cell(hdr, 7, 'Flag').font = Font(name=TF_FONT, size=9, bold=True, color=C_GREY)
    ws.cell(hdr, 7).alignment = Alignment(vertical='center')

    key = f'LEFT($C$5,4)&"|"&$C{{r}}'
    for i in range(TF_LINES):
        r = TF_FIRST + i
        k = key.format(r=r)
        ws.cell(r, 1, i + 1)
        ws.cell(r, 2, f'=IF($C{r}="","",IFERROR(INDEX(ITEMS!$C:$C,MATCH({k},ITEMS!$A:$A,0)),"Check"))')
        ws.cell(r, 5, f'=IF(OR($B{r}="",$B{r}="Check"),"",INDEX(ITEMS!$E:$E,MATCH({k},ITEMS!$A:$A,0)))')
        ws.cell(r, 6, f'=IF(OR($E{r}="",$D{r}="",$G{r}="{TF_BAD_QTY}"),"",$D{r}*$E{r})')
        ws.cell(r, 7, f'=IF(OR($B{r}="",$B{r}="Check",$D{r}=""),"",'
                      f'IF(NOT(ISNUMBER($D{r})),"{TF_BAD_QTY}",'
                      f'IF(OR($D{r}<1,$D{r}<>INT($D{r})),"{TF_BAD_QTY}",'
                      f'IF(SUMIFS($D${TF_FIRST}:$D${last},$C${TF_FIRST}:$C${last},"="&$C{r},'
                      f'$D${TF_FIRST}:$D${last},">0")'
                      f'>INDEX(ITEMS!$K:$K,MATCH({k},ITEMS!$A:$A,0)),"Above stock",""))))')
        ws.cell(r, 7).font = Font(name=TF_FONT, size=9, bold=True, color='C00000')
        ws.cell(r, 7).alignment = Alignment(vertical='center')
        for c in range(1, 7):
            cell = ws.cell(r, c)
            cell.border = box
            cell.font = f_reg
            cell.alignment = Alignment(vertical='center',
                                       horizontal='center' if c in (1, 2, 4) else None)
        for c in (3, 4):
            ws.cell(r, c).fill = inp
            ws.cell(r, c).protection = unlocked
        ws.cell(r, 4).number_format = '#,##0'
        ws.cell(r, 5).number_format = F_MONEY
        ws.cell(r, 6).number_format = F_MONEY
        ws.row_dimensions[r].height = 18
    tot = last + 1
    ws.cell(tot, 3, 'TOTAL').alignment = Alignment(horizontal='right')
    ws.cell(tot, 4, f'=SUMIFS(D{TF_FIRST}:D{last},B{TF_FIRST}:B{last},"<>Check",'
                    f'C{TF_FIRST}:C{last},"<>",G{TF_FIRST}:G{last},"<>{TF_BAD_QTY}")'
                    ).number_format = '#,##0'
    ws.cell(tot, 6, f'=SUM(F{TF_FIRST}:F{last})').number_format = F_RAND
    for c in range(1, 7):
        ws.cell(tot, c).font = f_bold
        ws.cell(tot, c).fill = fill(C_TOTAL)
        ws.cell(tot, c).border = box
    ws.cell(tot, 4).alignment = Alignment(horizontal='center')

    # signatures
    s = tot + 2
    for ref, txt in ((f'A{s}', 'ISSUED BY  (From outlet)'), (f'D{s}', 'RECEIVED BY  (To outlet)')):
        ws[ref] = txt
        ws[ref].font = Font(name=TF_FONT, size=10, bold=True, color='FFFFFF')
    ws.merge_cells(f'A{s}:C{s}')
    ws.merge_cells(f'D{s}:F{s}')
    for c in range(1, 7):
        ws.cell(s, c).fill = fill(C_HEAD)
    for j, lab in enumerate(('Name', 'Signature', 'Date')):
        rr = s + 1 + j
        ws.merge_cells(f'A{rr}:B{rr}')
        ws[f'A{rr}'] = lab
        ws[f'A{rr}'].font = f_reg
        ws[f'A{rr}'].alignment = Alignment(vertical='bottom')
        ws[f'C{rr}'].border = under
        ws[f'C{rr}'].protection = unlocked
        ws[f'D{rr}'] = lab
        ws[f'D{rr}'].font = f_reg
        ws[f'D{rr}'].alignment = Alignment(vertical='bottom')
        ws.merge_cells(f'E{rr}:F{rr}')
        for c in ('E', 'F'):
            ws[f'{c}{rr}'].border = under
            ws[f'{c}{rr}'].protection = unlocked
        ws.row_dimensions[rr].height = 28 if lab == 'Signature' else 22

    # cost control strip
    cc = s + 5
    ws.merge_cells(f'A{cc}:F{cc}')
    ws[f'A{cc}'] = 'COST CONTROL USE ONLY'
    ws[f'A{cc}'].font = Font(name=TF_FONT, size=9, bold=True, color=C_GREY)
    ws.merge_cells(f'A{cc + 1}:B{cc + 1}')
    ws[f'A{cc + 1}'] = 'SAP doc (311)'
    ws[f'D{cc + 1}'] = 'Posted by'
    ws.merge_cells(f'E{cc + 1}:F{cc + 1}')
    for ref in (f'A{cc + 1}', f'D{cc + 1}'):
        ws[ref].font = Font(name=TF_FONT, size=9, color=C_GREY)
    for ref in (f'C{cc + 1}', f'E{cc + 1}', f'F{cc + 1}'):
        ws[ref].border = Border(bottom=Side(style='thin', color=C_GREY))
        ws[ref].protection = unlocked
    ws.row_dimensions[cc + 1].height = 22

    # how to use (legend + example)
    h = cc + 3
    ws.merge_cells(f'A{h}:F{h + 4}')
    ws[f'A{h}'] = (
        'HOW TO USE  (fill in the yellow cells only)\n'
        '1) Choose the From and To outlet.  2) Type the date, e.g. 14/10/2026.  '
        '3) Pick each item from the Description list (it only lists what the From outlet holds) '
        'and type the Qty. SAP No, price and value fill in by themselves.  '
        '4) Print 2 copies, both outlets sign, send one to Cost Control.  '
        '5) A second transfer on the same route on the same day: set Seq. to 2.\n'
        'Example: From L095 CROCKERY, To L228 CROCKERY TW, Description "Cappuccino Cup '
        '220ml", Qty 34.   Next transfer: close without saving, or clear the yellow cells.')
    ws[f'A{h}'].font = Font(name=TF_FONT, size=8)
    ws[f'A{h}'].alignment = Alignment(wrap_text=True, vertical='top')

    # validations
    dv_from = DataValidation(type='list', formula1='FROM_LIST', allow_blank=True,
                             showErrorMessage=True, errorTitle='From outlet',
                             error='Choose an outlet from the list')
    dv_from.add('C5')
    dv_to = DataValidation(type='list', formula1='TO_LIST', allow_blank=True,
                           showErrorMessage=True, errorTitle='To outlet',
                           error='Choose an outlet from the list')
    dv_to.add('C6')
    dv_item = DataValidation(type='list',
                             formula1='INDIRECT(IF($C$5="","L_NONE","L_"&LEFT($C$5,4)))',
                             allow_blank=True, showErrorMessage=True, errorTitle='Item',
                             error='Pick the item from the list. Only items the From outlet '
                                   'holds are listed.')
    dv_item.add(f'C{TF_FIRST}:C{last}')
    dv_qty = DataValidation(type='whole', operator='greaterThanOrEqual', formula1='1',
                            allow_blank=True, showErrorMessage=True, errorTitle='Qty',
                            error='Enter a whole number of 1 or more')
    dv_qty.add(f'D{TF_FIRST}:D{last}')
    dv_date = DataValidation(type='date', operator='between', formula1='43831', formula2='73050',
                             allow_blank=True, showErrorMessage=True, errorTitle='Date',
                             error='Type a date, e.g. 14/10/2026')
    dv_date.add('E5')
    dv_seq = DataValidation(type='whole', operator='between', formula1='1', formula2='99',
                            allow_blank=True, showErrorMessage=True, errorTitle='Seq.',
                            error='Enter 1 to 99')
    dv_seq.add('E6')
    for dv in (dv_from, dv_to, dv_item, dv_qty, dv_date, dv_seq):
        ws.add_data_validation(dv)
    ws.conditional_formatting.add(
        f'D{TF_FIRST}:D{last}',
        FormulaRule(formula=[f'$G{TF_FIRST}<>""'], fill=fill(C_HIGH_FILL),
                    font=Font(bold=True, color=C_HIGH_FONT)))
    ws.conditional_formatting.add(
        f'B{TF_FIRST}:B{last}',
        CellIsRule(operator='equal', formula=['"Check"'], fill=fill(C_HIGH_FILL),
                   font=Font(bold=True, color=C_HIGH_FONT)))

    ws.protection.sheet = True          # formulas locked; yellow cells open
    ws.protection.password = TF_PASSWORD
    from openpyxl.workbook.protection import WorkbookProtection
    wb.security = WorkbookProtection(workbookPassword=TF_PASSWORD, lockStructure=True)
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.5
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.print_area = f'A1:F{h + 4}'
    ws.oddFooter.center.text = f'(c) 2026 Alain Yon | OE Tracker v{VERSION}'
    ws.oddFooter.center.size = 7
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    return len(from_codes), sum(len(v) for v in by_loc.values())


def approval_has_decisions(path):
    """True if any data row of an existing approval file has a Y/N entered."""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return False
    try:
        if 'APPROVAL' not in wb.sheetnames:
            return False
        col = None
        for r in wb['APPROVAL'].iter_rows(values_only=True):
            r = list(r or [])
            if col is None:
                hn = [norm(x) for x in r]
                cols = [i for i, h in enumerate(hn) if h.endswith('approve (y/n)')]
                if cols:
                    col = cols
                continue
            if not r or not LOC_RE.match(text(r[0]).upper()):
                break           # end of the data rows (signature block follows)
            if any(len(r) > c and text(r[c]) for c in col):
                return True
        return False
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main(argv):
    folder = os.path.dirname(os.path.abspath(__file__))
    explicit, period = [], None
    args = list(argv)
    while args:
        a = args.pop(0)
        if a in ('-h', '--help'):
            print(__doc__)
            return 0
        if a == '--period' and args:
            period = args.pop(0)
        elif a.lower().endswith('.xlsx'):
            explicit.append(os.path.abspath(a))
        elif a.strip():
            print(f'  ! Ignored argument: {a}')
    built = datetime.now()
    period = period or built.strftime('%b %Y')
    wb_path = os.path.join(folder, WORKBOOK_NAME)

    print('=' * 70)
    print(f'  OE TRACKER v{VERSION}  |  monthly build  |  period {period}')
    print('=' * 70)

    if file_is_locked(wb_path):
        print(f'\nERROR: {WORKBOOK_NAME} is open in Excel. Close it and run again.')
        return 1

    checks = []
    files, notes = find_sap_files(folder, explicit)
    for n in notes:
        print(f'  ! {n}')
        checks.append(('Source files', '', None, '', n))
    missing = [k for k in ('MB52', 'MM60') if k not in files]
    if missing:
        print(f'\nERROR: no {" or ".join(missing)} export found in {folder}')
        print('       Put the SAP exports (.xlsx) in this folder, or drag them onto the .bat.')
        return 1

    sources, controls = {}, {}
    print('\n[1] Reading SAP exports')
    stock, controls['MB52'] = load_mb52(files['MB52'], checks)
    print(f'    MB52  {os.path.basename(files["MB52"])}: {len(stock)} lines, '
          f'{len({k[0] for k in stock})} locations')
    prices = load_mm60(files['MM60'], checks)
    print(f'    MM60  {os.path.basename(files["MM60"])}: {len(prices)} materials')
    scrap, scrap_range = {}, None
    if 'MB51' in files:
        scrap, controls['MB51'], scrap_range = load_mb51(files['MB51'], checks)
        rng = f'{scrap_range[0]:%d %b} - {scrap_range[1]:%d %b %Y}' if scrap_range else 'no dates'
        print(f'    MB51  {os.path.basename(files["MB51"])}: {len(scrap)} item/location lines, '
              f'posted {rng}')
    else:
        print('    MB51  not found - scrap will show as 0')
        checks.append(('Source files', '', None, '', 'No MB51 scrapping export found; scrap shows as 0'))
    for k, p in files.items():
        sources[k] = {'name': os.path.basename(p),
                      'mtime': datetime.fromtimestamp(os.path.getmtime(p))}
    if scrap_range:
        sources['MB51']['extra'] = f'Posted {scrap_range[0]:%d %b} - {scrap_range[1]:%d %b %Y}'

    print('\n[2] Carrying forward par levels, notes and locations')
    prev_path = find_previous_workbook(folder)
    prev = None
    if prev_path:
        prev = read_previous_workbook(prev_path, checks)
        n_pos = sum(1 for v in prev['par'].values() if v > 0)
        print(f'    From {os.path.relpath(prev_path, folder)}: {len(prev["par"])} par levels '
              f'({n_pos} above 0), {len(prev["notes"])} notes, '
              f'{len(prev["locations"])} location settings')
    else:
        print('    No previous workbook found - starting with no par levels')
        checks.append(('Previous workbook', '', None, '', 'None found; no par levels carried forward'))

    items = build_items(stock, prices, scrap, prev, period, checks)
    locs = build_locations(items, prev, checks)
    loc_order = prev['loc_order'] if prev else []

    # Control-total reconciliation
    hold_total = sum(i['hold'] for i in items)
    scrap_total = sum(i['scrap_val'] for i in items)
    if controls.get('MB52') is not None and abs(hold_total - controls['MB52']) > 0.05:
        checks.append(('Control total', '', None, '',
                       f'Hold value R{hold_total:,.2f} vs MB52 total line R{controls["MB52"]:,.2f}'))
    if controls.get('MB51') is not None and abs(scrap_total + controls['MB51']) > 0.05:
        checks.append(('Control total', '', None, '',
                       f'Scrap value R{scrap_total:,.2f} vs MB51 total line R{-controls["MB51"]:,.2f}'))
    if controls.get('MB51') is not None:
        controls['MB51'] = -controls['MB51']   # show scrap as a positive figure

    print('\n[3] Building approval lists')
    files_lines, excluded = approval_lines(items, locs)
    plan, info_parts, appr_value = [], [], 0.0
    for label in sorted(files_lines, key=lambda k: (k != '', k)):
        lines = files_lines[label]
        codes = sorted({d['loc'] for d in lines} |
                       {c for c, d in locs.items() if d.get('file', '') == label and d['appr'] == 'Y'})
        a1 = a2 = ''
        if label:
            a1 = next((locs[c]['a1'] for c in codes if locs[c].get('a1')), '')
            a2 = next((locs[c]['a2'] for c in codes if locs[c].get('a2')), '')
            if not a1:
                checks.append(('Approval file has no approver', '', None, '',
                               f'"{label}" has no Approver 1 on the LOCATIONS tab'))
        name = approval_file_name(label, period)
        path = os.path.join(folder, name)
        if os.path.exists(path) and (file_is_locked(path) or approval_has_decisions(path)):
            why = 'is open in Excel' if file_is_locked(path) else 'already has Y/N decisions in it'
            name = name.replace('.xlsx', f'_rebuilt_{built:%Y%m%d_%H%M}.xlsx')
            path = os.path.join(folder, name)
            print(f'    ! The existing approval file {why}; it was left alone. New file: {name}')
            checks.append(('Source files', '', None, '',
                           f'Existing approval file {why}; new one saved as {name}'))
        val = sum(d['order_val'] for d in lines)
        appr_value += val
        wt = sum(1 for d in lines if d['transfer'])
        plan.append((label, lines, codes, a1, a2, name, path, val))
        info_parts.append(f'{name}: {len(lines)} lines, {wt} with a transfer option')
        who = f'{a1} then {a2}' if label and a2 else (a1 or 'F&B Manager')
        print(f'    {name}: {len(lines)} lines, R{val:,.2f} '
              f'({sum(1 for d in lines if d["urgency"] == "HIGH")} HIGH, {wt} with a transfer option) '
              f'- {who}')
    separate = {lab: (codes, len(lines), val) for lab, lines, codes, a1, a2, n, p, val in plan if lab}
    approval_info = {'value': round(appr_value, 2), 'text': ';  '.join(info_parts)}
    lines = files_lines['']
    for l, (n, v) in sorted(excluded.items()):
        print(f'    Not on any approval file (On Approval = N): {l} {locs[l]["name"]}: '
              f'{n} lines, R{v:,.2f}')

    print('\n[4] Writing files')
    tmp = wb_path + '.tmp'
    loc_codes = write_tracking_workbook(tmp, period, built, sources, controls, items, locs,
                                        loc_order, checks, approval_info, scrap_range)
    if os.path.exists(wb_path):
        os.makedirs(os.path.join(folder, ARCHIVE_DIR), exist_ok=True)
        arch = os.path.join(folder, ARCHIVE_DIR,
                            f'OE_Tracking_Workbook_{built:%Y-%m-%d_%H%M%S}.xlsx')
        shutil.copy2(wb_path, arch)
        print(f'    Previous workbook archived: {os.path.relpath(arch, folder)}')
    try:
        os.replace(tmp, wb_path)
    except PermissionError:
        print(f'\nERROR: could not replace {WORKBOOK_NAME} (open in Excel?). '
              f'The new workbook is saved as {os.path.basename(tmp)}')
        return 1
    print(f'    {WORKBOOK_NAME}: {len(loc_codes)} location tabs, {len(items)} items')
    for label, lns, codes, a1, a2, name, path, val in plan:
        if not label:
            write_approval_file(path, period, built, sources, lns, excluded, locs, separate)
            print(f'    {name}: send to the F&B Manager')
        elif a2:
            write_two_step_approval(path, period, built, sources, lns, label, a1, a2, locs)
            print(f'    {name}: send to {a1}, then {a2}')
        else:
            write_two_step_approval(path, period, built, sources, lns, label, a1 or 'Approver 1',
                                    'Cost Control (confirm)', locs)
            print(f'    {name}: send to {a1 or "the approver"} (no Approver 2 set)')
    tf_name = 'OE_Transfer_Form.xlsx'
    if file_is_locked(os.path.join(folder, tf_name)):
        tf_name = f'OE_Transfer_Form_{built:%Y%m%d_%H%M}.xlsx'
        print('    ! OE_Transfer_Form.xlsx is open in Excel; saved the new one as ' + tf_name)
    n_from, n_lines = write_transfer_form(os.path.join(folder, tf_name), built, sources, items, locs)
    print(f'    {tf_name}: {n_from} outlets, {n_lines} items listed, for operations')

    print('\n' + '-' * 70)
    print(f'  Hold value (MB52)      R{hold_total:>16,.2f}')
    print(f'  Scrap value (MB51)     R{scrap_total:>16,.2f}')
    print(f'  On approval files      R{appr_value:>16,.2f}  '
          f'({sum(len(p[1]) for p in plan)} lines in {len(plan)} files)')
    print(f'  Data checks            {len(checks):>17}  (see CHECKS tab)')
    print('-' * 70)
    print('  If you change par levels, save the workbook and run this again;')
    print('  the approval file is rebuilt from the saved par levels.')
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main(sys.argv[1:]))
    except Exception as exc:  # last-resort message instead of a raw traceback
        import traceback
        traceback.print_exc()
        print(f'\nERROR: {exc}')
        sys.exit(1)
