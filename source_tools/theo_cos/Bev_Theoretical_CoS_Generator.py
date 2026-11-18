#!/usr/bin/env python3
"""
Beverage Theoretical Cost of Sales (CoS) Generator v5
Micros/Simphony Sales Mix + SAP MM60 -> theoretical beverage CoS

Copyright (c) 2026 Alain Yon
All rights reserved.

What changed from v4
- No more guessing. Every POS item is priced through beverage_item_map.xlsx:
  POS item -> SAP material number -> how it is sold (whole, house pour, ml, divisor).
- Prices come from MM60 by MATERIAL NUMBER, so a new month's MM60 updates costs
  without touching the map.
- House rules (tot, wine glass, draught pour, VAT...) live in the map's
  "House Rules" sheet. Change a pour size once, every line follows.
- One POS name can have a glass row and a bottle row ("Min Avg Price" column),
  because Micros cuts names at 32 characters.
- New POS items are asked about on screen (pick from a list) and saved to the map.
- Cocktails, shooters and coffees are costed per drink from their recipe sheet in
  Cocktail_Costing.xlsx (Treatment 'Recipe' + 'Recipe Sheet' in the map), using
  this period's MM60 prices.
- Lines whose CoS is confirmed as right (rebate stock, sommelier pricing) can be
  marked 'Skip CoS Flag' = Yes in the map so they stop appearing on the Review sheet.
  Anything not answered is left out of the theoretical and listed on the Review sheet.
- Summary covers every category, including Rum and Cocktails, on ex-VAT sales.

Usage:
    python Bev_Theoretical_CoS_Generator.py <sales_mix.xlsx> [output.xlsx]
        [--map beverage_item_map.xlsx] [--no-prompt]
"""

import os
import re
import sys
import shutil
from datetime import datetime
from difflib import SequenceMatcher

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

VERSION = '5.0'
FONT = 'Arial'
HEAD_FILL = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
REVIEW_FILL = PatternFill(start_color='F8D7DA', end_color='F8D7DA', fill_type='solid')
FLAG_FILL = PatternFill(start_color='FFF3CD', end_color='FFF3CD', fill_type='solid')
INPUT_FONT = Font(name=FONT, color='0000FF')

# Map sheet column order (1-based). Keep in sync with build/append code.
MAP_COLS = ['POS Item', 'Min Avg Price (R)', 'Treatment', 'SAP Material', 'SAP Description',
            'Pack (ml)', 'Serving', 'House Rule', 'Pour (ml)', 'Divisor', 'Manual Cost (R)',
            'POS Family', 'Notes', 'Added', 'Skip CoS Flag', 'Recipe Sheet']
TREATMENTS = ['SAP item', 'Recipe', 'Manual cost', 'Exclude', 'Cocktail rule', 'Table water rule']
SAFE_ARITH = re.compile(r'^[\d\.\s\+\-\*/\(\)]+$')
SERVINGS = ['Whole', 'House', 'Pour ml', 'Divisor']
HOUSE_RULES = {  # map value -> House Rules setting
    'Tot': 'Tot (ml)', 'Wine glass': 'Wine glass (ml)', 'Bubbly glass': 'Bubbly glass (ml)',
    'Draught': 'Draught (ml)', 'Cordial': 'Cordial (ml)', 'Juice glass': 'Juice glass (ml)',
    'Port glass': 'Port glass (ml)', 'Dessert glass': 'Dessert glass (ml)',
}
CATEGORIES = ['Spirits', 'Liqueurs & Aperitifs', 'Wine & Bubbly', 'Beer', 'Coolers & Ciders',
              'Minerals', 'Cocktails & Mixed', 'Other']
FLAG_LOW, FLAG_HIGH = 0.05, 0.80   # defaults; House Rules 'Flag CoS below/above' override


# ----------------------------------------------------------------------------- helpers
def norm(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip().lower()


def to_float(v, default=0.0):
    if v is None or v == '':
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def mat_key(v):
    """SAP material numbers arrive as int, float or text. Normalise to text digits."""
    if v is None or v == '':
        return ''
    try:
        return str(int(float(v)))
    except (TypeError, ValueError):
        return str(v).strip()


def pack_ml(desc):
    """Pack size in ml from an MM60 description (750ml, 1lt, 30L, 50Lt, Magnum...)."""
    d = str(desc or '').lower()
    if 'magnum' in d:
        return 1500.0
    m = re.search(r'(\d+(?:\.\d+)?)\s*(ml|m)\b', d)
    if m:
        return float(m.group(1))
    m = re.search(r'(\d+(?:\.\d+)?)\s*(lt|l)\b', d)
    if m:
        return float(m.group(1)) * 1000
    return None


def clean_family(fam):
    f = str(fam or '').strip()
    for p in ('ent - ', 'prop - '):
        if f.lower().startswith(p):
            return f[len(p):].strip()
    return f


def category(fam):
    f = fam.lower()
    if any(x in f for x in ('cocktail', 'shooter', 'alcohol coffee')):
        return 'Cocktails & Mixed'
    if any(x in f for x in ('whisk', 'white spirit', 'brandy', 'cognac', 'rum', 'vodka', 'gin', 'tequila')):
        return 'Spirits'
    if any(x in f for x in ('liqueur', 'aperitif')):
        return 'Liqueurs & Aperitifs'
    if any(x in f for x in ('wine', 'port', 'sparkling', 'champagne', 'mcc')):
        return 'Wine & Bubbly'
    if 'beer' in f:
        return 'Beer'
    if 'cooler' in f or 'cider' in f:
        return 'Coolers & Ciders'
    if any(x in f for x in ('mineral', 'water', 'juice', 'soft', 'cordial', 'energy')):
        return 'Minerals'
    return 'Other'


def fnum(v):
    return ('%g' % v) if v is not None else ''


# ----------------------------------------------------------------------------- loaders
def find_header(ws, first_label, max_scan=30):
    for r in range(1, max_scan + 1):
        if norm(ws.cell(r, 1).value) == norm(first_label):
            return r, {norm(ws.cell(r, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(r, c).value}
    return None, {}


def load_pos(path, sheet='Report'):
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]
    meta = {}
    for r in range(1, 8):
        label = norm(ws.cell(r, 1).value)
        if label:
            meta[label] = ws.cell(r, 2).value
    hr, cols = find_header(ws, 'Item')
    if not hr:
        raise SystemExit(f'Could not find the "Item" header row on sheet {sheet}.')

    def col(name, fallback):
        return cols.get(norm(name), fallback)
    c_item, c_fam = col('Item', 1), col('Family Group', 2)
    c_net, c_qty, c_avg = col('Sales Less Item Disc', 6), col('Qty Sold', 8), col('Average Price', 10)
    items = []
    for r in range(hr + 1, ws.max_row + 1):
        name = ws.cell(r, c_item).value
        if not name or str(name).strip().lower().startswith('total'):
            continue
        qty = to_float(ws.cell(r, c_qty).value)
        net = to_float(ws.cell(r, c_net).value)
        avg = to_float(ws.cell(r, c_avg).value) or (net / qty if qty else 0.0)
        items.append(dict(item=str(name).strip(), fam=clean_family(ws.cell(r, c_fam).value),
                          qty=qty, net=net, avg=avg))
    wb.close()
    return meta, items


def load_mm60(path, sheet='MM60'):
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]
    hr, cols = find_header(ws, 'Plant')
    hr = hr or 1
    c_mat = cols.get('material', 2)
    c_desc = cols.get('material description', 3)
    c_grp = cols.get('material group', 4)
    c_uom = cols.get('base unit of measure', 5)
    c_price = cols.get('price', 6)
    out = {}
    for r in range(hr + 1, ws.max_row + 1):
        m = mat_key(ws.cell(r, c_mat).value)
        if not m:
            continue
        out[m] = dict(desc=str(ws.cell(r, c_desc).value or '').strip(), group=str(ws.cell(r, c_grp).value or ''),
                      uom=str(ws.cell(r, c_uom).value or ''), price=to_float(ws.cell(r, c_price).value))
    wb.close()
    return out


def load_map(path):
    if not os.path.exists(path):
        raise SystemExit(f'Item map not found: {path}\nPut beverage_item_map.xlsx next to this script, or pass --map.')
    wb = load_workbook(path, data_only=True)
    house = {}
    hs = wb['House Rules']
    for r in range(2, hs.max_row + 1):
        k = hs.cell(r, 1).value
        if k:
            house[norm(k)] = hs.cell(r, 2).value
    ms = wb['Item Map']
    head = [norm(ms.cell(1, c).value) for c in range(1, ms.max_column + 1)]
    idx = {h: i for i, h in enumerate(head)}
    rows = {}
    for r in range(2, ms.max_row + 1):
        vals = [ms.cell(r, c).value for c in range(1, ms.max_column + 1)]
        get = lambda name: vals[idx[norm(name)]] if norm(name) in idx else None
        item = get('POS Item')
        if not item:
            continue
        rows.setdefault(norm(item), []).append(dict(
            row=r, item=str(item).strip(), min_avg=to_float(get('Min Avg Price (R)')),
            treatment=str(get('Treatment') or 'SAP item').strip(), material=mat_key(get('SAP Material')),
            pack=to_float(get('Pack (ml)'), None), serving=str(get('Serving') or 'Whole').strip(),
            rule=str(get('House Rule') or '').strip(), pour=to_float(get('Pour (ml)'), None),
            divisor=to_float(get('Divisor'), None), cost=to_float(get('Manual Cost (R)'), None),
            noflag=str(get('Skip CoS Flag') or '').strip().lower().startswith('y'),
            recipe=str(get('Recipe Sheet') or '').strip()))
    for k in rows:
        rows[k].sort(key=lambda x: x['min_avg'], reverse=True)
    wb.close()
    return house, rows


def cell_number(v):
    """Recipe quantities are numbers or simple formulas like =50/750."""
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().lstrip('=').lstrip('+')
    if SAFE_ARITH.match(s):
        try:
            return float(eval(s, {'__builtins__': {}}, {}))
        except Exception:
            return None
    return to_float(s, None)


def load_recipes(path):
    """Read every 'Menu N' sheet of the cocktail costing workbook: name (C2) and lines (SAP Nr in B, qty in D)."""
    if not path or not os.path.exists(path):
        return None
    wb = load_workbook(path)
    out = {}
    for ws in wb.worksheets:
        if not re.match(r'^menu \d+$', ws.title.strip().lower()):
            continue
        name = ws['C2'].value
        if isinstance(name, str) and name.startswith('='):
            name = None
        lines, bad = [], []
        for r in range(5, 40):
            b = ws.cell(r, 2).value
            if isinstance(b, str) and not b.strip().isdigit():
                break  # reached the totals block ("Menu item cost")
            sap = mat_key(b)
            if not sap:
                continue
            q = cell_number(ws.cell(r, 4).value)
            if q is None:
                bad.append(sap)
                continue
            lines.append((sap, q))
        out[norm(ws.title)] = dict(sheet=ws.title, name=str(name or '').strip(), lines=lines, bad=bad)
    wb.close()
    return out


def recipe_cost(rec, mm60):
    total, missing = 0.0, list(rec['bad'])
    for sap, q in rec['lines']:
        s = mm60.get(sap)
        if not s or not s['price']:
            missing.append(sap)
            continue
        total += q * s['price']
    return total, missing


def recipe_suggestions(item, recipes, n=6):
    if not recipes:
        return []
    il = norm(re.sub(r'\.[a-z]{3}$', '', item, flags=re.I))
    scored = sorted(((SequenceMatcher(None, il, norm(r['name'])).ratio(), k) for k, r in recipes.items() if r['name']),
                    reverse=True)
    return [k for _, k in scored[:n]]


def house_value(house, key, default):
    v = house.get(norm(key))
    return default if v in (None, '') else v


# ----------------------------------------------------------------------------- pricing
def pick_map_row(map_rows, item, avg):
    for m in map_rows.get(norm(item), []):
        if avg >= m['min_avg'] - 0.005:
            return m
    return None


def divisor_for(m, pack, house):
    s = m['serving'].lower()
    if s == 'whole':
        return 1.0, 'Whole'
    if s == 'divisor':
        return (m['divisor'] or None), f"÷{fnum(m['divisor'])}"
    if s == 'pour ml':
        if pack and m['pour']:
            return pack / m['pour'], f"{fnum(m['pour'])}ml pour"
        return None, 'Pour ml (pack size unknown)'
    if s == 'house':
        setting = HOUSE_RULES.get(m['rule'])
        ml = to_float(house_value(house, setting, 0)) if setting else 0
        if pack and ml:
            return pack / ml, f"House {m['rule'].lower()} {fnum(ml)}ml"
        return None, f"House {m['rule'] or '?'} (pack or rule missing)"
    return None, f"Unknown serving '{m['serving']}'"


def price_line(line, m, mm60, house, recipes=None):
    """Returns dict with treatment, material, desc, price, divisor, serving text, included, status, flat."""
    out = dict(treatment='', material='', desc='', price=None, div=None, serving='', included=False,
               status='', flat=False, noflag=bool(m and m.get('noflag')))
    if m is None:
        out['status'] = 'REVIEW: not in item map'
        return out
    t = m['treatment']
    out['treatment'] = t
    if t == 'Exclude':
        out['status'] = 'OK (left out by choice)'
        return out
    if t == 'Cocktail rule':
        rule = str(house_value(house, 'Cocktail rule', 'Exclude')).strip().lower()
        if rule.startswith('flat'):
            out.update(serving='Standard CoS %', included=True, flat=True, status='OK')
        elif rule.startswith('recipe'):
            if m['cost']:
                out.update(price=m['cost'], div=1.0, serving='Recipe cost', included=True, status='OK')
            else:
                out['status'] = 'REVIEW: no recipe cost in item map'
        else:
            out['status'] = 'OK (cocktail rule: left out)'
        return out
    if t == 'Recipe':
        if recipes is None:
            out['status'] = 'REVIEW: cocktail costing workbook not found (House Rules: Cocktail costing file)'
            return out
        rec = recipes.get(norm(m['recipe']))
        if not rec:
            out['status'] = f"REVIEW: recipe sheet '{m['recipe']}' not found in the costing workbook"
            return out
        if not rec['lines']:
            out['status'] = f"REVIEW: {rec['sheet']} has no ingredients"
            return out
        cost, missing = recipe_cost(rec, mm60)
        if missing:
            out['status'] = f"REVIEW: {rec['sheet']} material(s) {', '.join(missing)} not in MM60 or no price"
            return out
        out.update(price=cost, div=1.0, desc=f"{rec['sheet']}: {rec['name']}", serving='Recipe (per drink)',
                   included=True, status='OK')
        return out
    if t == 'Table water rule':
        out.update(price=to_float(house_value(house, 'Table water cost per bottle (R)', 0)), div=1.0,
                   serving='Table water rule', included=True, status='OK')
        return out
    if t == 'Manual cost':
        if not m['cost']:
            out['status'] = 'REVIEW: manual cost missing in item map'
            return out
        out.update(price=m['cost'], desc='(manual cost)')
        pack = m['pack']
    else:  # SAP item
        sap = mm60.get(m['material'])
        out['material'] = m['material']
        if not sap:
            out['status'] = f"REVIEW: material {m['material']} not in MM60"
            return out
        out['desc'] = sap['desc']
        if not sap['price']:
            out['status'] = 'REVIEW: no price in MM60'
            return out
        out['price'] = sap['price']
        pack = m['pack'] or pack_ml(sap['desc'])
    div, text = divisor_for(m, pack, house)
    out['serving'] = text
    if not div:
        out['status'] = f'REVIEW: {text}'
        return out
    out.update(div=div, included=True, status='OK')
    return out


# ----------------------------------------------------------------------------- prompting
def suggestions(item, mm60, n=6):
    """Beverage-only candidates (BEV groups + cold-pressed juices). Never Do Not Use."""
    il = norm(re.sub(r'\b(gls|gl|g)\b$', '', item, flags=re.I))
    scored = []
    for k, s in mm60.items():
        if not (s['group'].startswith('BEV') or s['group'] == 'FOO001800'):
            continue
        d = s['desc']
        if d == 'Do Not Use' or d.upper().startswith('DNU'):
            continue
        dl = d.lower()
        score = SequenceMatcher(None, il, dl).ratio()
        toks = set(re.findall(r'[a-z]{3,}', il))
        if toks:
            score += 0.5 * len(toks & set(re.findall(r'[a-z]{3,}', dl))) / len(toks)
        scored.append((score, k))
    scored.sort(reverse=True)
    return [k for _, k in scored[:n]]


def ask(prompt, valid):
    while True:
        try:
            a = input(prompt).strip().upper()
        except EOFError:
            return 'Q'
        if a in valid:
            return a
        print('   Please type one of: ' + ', '.join(sorted(valid)))


def ask_number(prompt):
    while True:
        try:
            a = input(prompt).strip().replace(',', '.').replace('R', '').replace('r', '')
        except EOFError:
            return None
        try:
            v = float(a)
            if v > 0:
                return v
        except ValueError:
            pass
        print('   Please type a number above 0.')


def prompt_serving(pack, house):
    opts = [('1', 'Whole bottle / unit (÷1)', ('Whole', '', None, None))]
    i = 2
    for rule, setting in HOUSE_RULES.items():
        ml = to_float(house_value(house, setting, 0))
        div = f', ÷{pack / ml:.2f}' if pack and ml else ''
        opts.append((str(i), f'House {rule.lower()} ({fnum(ml)}ml{div})', ('House', rule, None, None)))
        i += 1
    for k, label, val in opts:
        print(f'     [{k}] {label}')
    print('     [P] Pour size in ml     [D] Custom divisor')
    a = ask('   How is it sold? ', {k for k, _, _ in opts} | {'P', 'D'})
    if a == 'P':
        return ('Pour ml', '', ask_number('   Pour size in ml: '), None)
    if a == 'D':
        return ('Divisor', '', None, ask_number('   Divide the cost by: '))
    return dict((k, v) for k, _, v in opts)[a]


def prompt_new_items(unmapped, mm60, house, recipes=None):
    """Ask about each unmapped POS line. Returns list of new map rows (dicts)."""
    new_rows = []
    print('\n' + '=' * 64)
    print(f'  {len(unmapped)} POS item(s) are not in the item map yet.')
    print('  Pick an answer for each. [S] skips one, [Q] skips the rest.')
    print('  Skipped items are left out of the theoretical and listed on the Review sheet.')
    print('=' * 64)
    for n, line in enumerate(unmapped, 1):
        print(f"\n[{n}/{len(unmapped)}] {line['item']}  ({line['fam']})  sold {fnum(line['qty'])} @ avg R{line['avg']:,.2f}")
        sug = suggestions(line['item'], mm60)
        for i, k in enumerate(sug, 1):
            s = mm60[k]
            print(f"     [{i}] {s['desc']}  R{s['price']:,.2f} {s['uom']}  #{k}")
        rsug = recipe_suggestions(line['item'], recipes)
        if rsug:
            print('     Recipes: ' + '   '.join(f"[R{i}] {recipes[k]['name']} ({recipes[k]['sheet']})" for i, k in enumerate(rsug, 1)))
        print('     [N] Type a SAP material number   [M] Not in SAP, type a cost')
        print('     [C] Cocktail rule   [T] Table water rule   [X] Leave out   [S] Skip   [Q] Skip all')
        a = ask('   Which SAP item or recipe? ', {str(i) for i in range(1, len(sug) + 1)} | set('NMCTXSQ')
                | {f'R{i}' for i in range(1, len(rsug) + 1)})
        if a == 'Q':
            break
        if a == 'S':
            continue
        row = dict(item=line['item'], min_avg=0.0, treatment='SAP item', material='', desc='', pack=None,
                   serving='Whole', rule='', pour=None, divisor=None, cost=None, fam=line['fam'], recipe='')
        if a.startswith('R'):
            rec = recipes[rsug[int(a[1:]) - 1]]
            row.update(treatment='Recipe', recipe=rec['sheet'], desc=rec['name'])
            new_rows.append(row)
            print('   Saved.')
            continue
        if a in ('C', 'T', 'X'):
            row['treatment'] = {'C': 'Cocktail rule', 'T': 'Table water rule', 'X': 'Exclude'}[a]
            if a == 'C' and str(house_value(house, 'Cocktail rule', '')).lower().startswith('recipe'):
                row['cost'] = ask_number('   Recipe cost per drink R: ')
            new_rows.append(row)
            continue
        if a == 'M':
            row['treatment'] = 'Manual cost'
            row['cost'] = ask_number('   Cost per bottle / unit R: ')
            row['pack'] = ask_number('   Bottle / unit size in ml (e.g. 750): ')
        else:
            if a == 'N':
                while True:
                    k = mat_key(input('   SAP material number: ').strip())
                    if k in mm60:
                        break
                    print('   Not found in this MM60. Try again.')
            else:
                k = sug[int(a) - 1]
            row['material'], row['desc'] = k, mm60[k]['desc']
            row['pack'] = pack_ml(mm60[k]['desc'])
            if row['pack'] is None:
                row['pack'] = ask_number('   Pack size in ml (not in the description): ')
        row['serving'], row['rule'], row['pour'], row['divisor'] = prompt_serving(row['pack'], house)
        new_rows.append(row)
        print('   Saved.')
    return new_rows


def append_to_map(map_path, new_rows):
    stamp = datetime.now().strftime('%Y%m%d_%H%M')
    backup = map_path.replace('.xlsx', f'_backup_{stamp}.xlsx')
    shutil.copy2(map_path, backup)
    wb = load_workbook(map_path)
    ws = wb['Item Map']
    r = ws.max_row + 1
    while r > 2 and not ws.cell(r - 1, 1).value:
        r -= 1
    for x in new_rows:
        vals = [x['item'], x['min_avg'], x['treatment'], int(x['material']) if x['material'].isdigit() else x['material'],
                x['desc'], x['pack'], x['serving'], x['rule'], x['pour'], x['divisor'], x['cost'], x['fam'],
                '', f"Added {datetime.now():%Y-%m-%d} by generator", '', x.get('recipe', '')]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c)
            cell.value = v
            cell.font = Font(name=FONT)
        r += 1
    wb.save(map_path)
    print(f'\n  Item map updated ({len(new_rows)} new row(s)). Backup: {os.path.basename(backup)}')


# ----------------------------------------------------------------------------- report
def write_report(out_path, meta, lines, house, source, map_path):
    wb = Workbook()
    wb.remove(wb['Sheet'])
    ws = wb.create_sheet('Summary')
    wd = wb.create_sheet('Detail')
    wr = wb.create_sheet('Review')
    bold = Font(name=FONT, bold=True)
    head = Font(name=FONT, bold=True, color='FFFFFF')
    base = Font(name=FONT)

    # ---- Summary: header + assumptions
    ws['A1'] = 'Beverage Theoretical CoS Report'
    ws['A1'].font = Font(name=FONT, bold=True, size=14)
    ws['A2'] = f"Period: {meta.get('business dates', '')}"
    ws['A3'] = f"Location: {meta.get('locations', '')}"
    ws['A4'] = f"Source: {os.path.basename(source)}  |  Item map: {os.path.basename(map_path)}  |  v{VERSION} run {datetime.now():%Y-%m-%d %H:%M}"
    for c in ('A2', 'A3', 'A4'):
        ws[c].font = base
    vat_incl = str(house_value(house, 'Sales are VAT-inclusive', 'Yes')).strip().lower().startswith('y')
    rule = str(house_value(house, 'Cocktail rule', 'Exclude'))
    pct = to_float(house_value(house, 'Cocktail standard CoS %', 0.22))
    if pct > 1:
        pct = pct / 100
    assumptions = [('Sales are VAT-inclusive', 'Yes' if vat_incl else 'No', None),
                   ('VAT rate', to_float(house_value(house, 'VAT rate', 0.15)), '0%'),
                   ('Cocktail rule', rule, None),
                   ('Cocktail standard CoS %', pct, '0.0%')]
    for i, (k, v, fmt) in enumerate(assumptions, 6):
        ws.cell(i, 1, k).font = base
        c = ws.cell(i, 2, v)
        c.font = INPUT_FONT
        if fmt:
            c.number_format = fmt
    ws['C6'] = 'Blue cells come from the House Rules sheet of the item map. Change them there for the next run.'
    ws['C6'].font = Font(name=FONT, italic=True, color='666666')
    VAT_FLAG, VAT_RATE, PCT = 'Summary!$B$6', 'Summary!$B$7', 'Summary!$B$9'

    # ---- Detail
    dcols = ['Item', 'POS Family', 'Category', 'Treatment', 'SAP Material', 'SAP Description', 'Cost / Unit (R)',
             'Serving', 'Divisor', 'Unit Cost (R)', 'Qty Sold', 'Theoretical Cost (R)', 'Net Sales (POS)',
             'Net Sales ex VAT', 'CoS %', 'In Theoretical', 'Status']
    widths = [34, 22, 20, 14, 12, 40, 14, 22, 9, 13, 10, 16, 15, 15, 9, 13, 46]
    for c, (h, w) in enumerate(zip(dcols, widths), 1):
        cell = wd.cell(1, c, h)
        cell.font, cell.fill = head, HEAD_FILL
        cell.alignment = Alignment(wrap_text=True, vertical='center')
        wd.column_dimensions[get_column_letter(c)].width = w
    wd.freeze_panes = 'B2'
    review = []
    flag_low = to_float(house_value(house, 'Flag CoS below', FLAG_LOW))
    flag_high = to_float(house_value(house, 'Flag CoS above', FLAG_HIGH))
    for i, ln in enumerate(lines, 2):
        p = ln['priced']
        status = p['status']
        if (p['included'] and status == 'OK' and ln['qty'] > 0 and ln['net'] > 0
                and p['treatment'] in ('SAP item', 'Manual cost', 'Recipe') and not p['noflag']):
            ex = ln['net'] / (1 + to_float(house_value(house, 'VAT rate', 0.15))) if vat_incl else ln['net']
            cos = (p['price'] / p['div'] * ln['qty']) / ex if ex else 0
            if cos < flag_low:
                status = f'FLAG: CoS {cos:.1%} looks low. Check the SAP price or pour'
            elif cos > flag_high:
                status = f'FLAG: CoS {cos:.0%} looks high. Check the SAP price or menu price'
        vals = [ln['item'], ln['fam'], ln['cat'], p['treatment'] or '-',
                int(p['material']) if str(p['material']).isdigit() else (p['material'] or ''),
                p['desc'], p['price'], p['serving'], p['div']]
        for c, v in enumerate(vals, 1):
            wd.cell(i, c, v).font = base
        if p['flat']:
            wd.cell(i, 10, f'=IF(K{i}=0,0,N{i}/K{i}*{PCT})')
        else:
            wd.cell(i, 10, f'=IF(OR(G{i}="",I{i}="",I{i}=0),0,G{i}/I{i})')
        wd.cell(i, 11, ln['qty'])
        wd.cell(i, 12, f'=J{i}*K{i}')
        wd.cell(i, 13, ln['net'])
        wd.cell(i, 14, f'=IF({VAT_FLAG}="Yes",M{i}/(1+{VAT_RATE}),M{i})')
        wd.cell(i, 15, f'=IF(N{i}=0,0,L{i}/N{i})')
        wd.cell(i, 16, 'Yes' if p['included'] else 'No')
        wd.cell(i, 17, status)
        for c in range(10, 18):
            wd.cell(i, c).font = base
        for c, fmt in ((7, '#,##0.00'), (9, '0.00'), (10, '#,##0.00'), (12, '#,##0.00'), (13, '#,##0.00'),
                       (14, '#,##0.00'), (15, '0.0%')):
            wd.cell(i, c).number_format = fmt
        if status.startswith('REVIEW'):
            wd.cell(i, 17).fill = REVIEW_FILL
        elif status.startswith('FLAG'):
            wd.cell(i, 17).fill = FLAG_FILL
        if status.startswith(('REVIEW', 'FLAG')) and ln['qty'] != 0:
            review.append((ln, status))
    last = len(lines) + 1
    wd.auto_filter.ref = f'A1:Q{last}'

    # ---- Summary: category table
    hdr_row = 11
    for c, h in enumerate(['Category', 'Sales ex VAT', 'Theoretical Cost', 'CoS %', 'Sales left out (ex VAT)', 'Lines'], 1):
        cell = ws.cell(hdr_row, c, h)
        cell.font, cell.fill = head, HEAD_FILL
    used = [c for c in CATEGORIES if any(ln['cat'] == c for ln in lines)]
    r0 = hdr_row + 1
    rng = lambda col: f'Detail!${col}$2:${col}${last}'
    for j, cat in enumerate(used):
        r = r0 + j
        ws.cell(r, 1, cat).font = base
        ws.cell(r, 2, f'=SUMIFS({rng("N")},{rng("C")},A{r},{rng("P")},"Yes")')
        ws.cell(r, 3, f'=SUMIFS({rng("L")},{rng("C")},A{r},{rng("P")},"Yes")')
        ws.cell(r, 4, f'=IF(B{r}=0,0,C{r}/B{r})')
        ws.cell(r, 5, f'=SUMIFS({rng("N")},{rng("C")},A{r},{rng("P")},"No")')
        ws.cell(r, 6, f'=COUNTIFS({rng("C")},A{r},{rng("K")},"<>0")')
    rl = r0 + len(used) - 1
    tr = rl + 2
    ws.cell(tr, 1, 'TOTAL').font = bold
    for c, col in ((2, 'B'), (3, 'C'), (5, 'E'), (6, 'F')):
        ws.cell(tr, c, f'=SUM({col}{r0}:{col}{rl})').font = bold
    ws.cell(tr, 4, f'=IF(B{tr}=0,0,C{tr}/B{tr})').font = bold
    for r in range(r0, tr + 1):
        for c, fmt in ((2, '#,##0.00'), (3, '#,##0.00'), (4, '0.0%'), (5, '#,##0.00'), (6, '0')):
            ws.cell(r, c).number_format = fmt
    ws.cell(tr + 2, 1, 'Theoretical CoS (weighted by sales mix)').font = bold
    ws.cell(tr + 2, 4, f'=D{tr}').font = bold
    ws.cell(tr + 2, 4).number_format = '0.0%'
    ws.cell(tr + 3, 1, 'Lines needing review (left out until fixed)').font = base
    ws.cell(tr + 3, 4, f'=COUNTIFS({rng("Q")},"REVIEW*",{rng("K")},"<>0")')
    ws.cell(tr + 4, 1, f'Lines flagged (CoS under {flag_low:.0%} or over {flag_high:.0%})').font = base
    ws.cell(tr + 4, 4, f'=COUNTIF({rng("Q")},"FLAG*")')
    ws.cell(tr + 5, 1, '"Sales left out" are lines the item map excludes (e.g. cocktails) or that need review. See the Review sheet.').font = Font(name=FONT, italic=True, color='666666')
    for col, w in zip('ABCDEF', (42, 16, 16, 10, 22, 8)):
        ws.column_dimensions[col].width = w

    # ---- Review
    rcols = ['Item', 'POS Family', 'Qty Sold', 'Avg Price (R)', 'Net Sales (POS)', 'Issue', 'What to do']
    for c, (h, w) in enumerate(zip(rcols, (34, 22, 10, 13, 15, 60, 70)), 1):
        cell = wr.cell(1, c, h)
        cell.font, cell.fill = head, HEAD_FILL
        wr.column_dimensions[get_column_letter(c)].width = w
    for i, (ln, status) in enumerate(review, 2):
        if 'not in item map' in status:
            todo = 'Run again and answer the prompt, or add a row to the Item Map. Suggestions: ' + '; '.join(ln['suggest'])
        elif status.startswith('FLAG'):
            todo = 'If the SAP price is wrong, fix it in SAP (MM60 feeds every run). If it is right, no action.'
        else:
            todo = 'Fix the row for this item on the Item Map sheet.'
        for c, v in enumerate([ln['item'], ln['fam'], ln['qty'], round(ln['avg'], 2), ln['net'], status, todo], 1):
            wr.cell(i, c, v).font = base
        wr.cell(i, 6).fill = REVIEW_FILL if status.startswith('REVIEW') else FLAG_FILL
    if not review:
        wr['A2'] = 'Nothing to review.'
        wr['A2'].font = base

    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    return len(review)


# ----------------------------------------------------------------------------- main
def find_costing_file(house, map_path, here):
    name = str(house_value(house, 'Cocktail costing file', 'Cocktail_Costing.xlsx')).strip()
    if os.path.isabs(name):
        return name
    for base in (os.path.dirname(os.path.abspath(map_path)), here, os.getcwd()):
        p = os.path.join(base, name)
        if os.path.exists(p):
            return p
    return None


def create_theoretical_report(file_path, output_path=None, map_path=None, prompt=True):
    here = os.path.dirname(os.path.abspath(__file__))
    map_path = map_path or os.path.join(here, 'beverage_item_map.xlsx')
    meta, items = load_pos(file_path)
    mm60 = load_mm60(file_path)
    house, map_rows = load_map(map_path)
    recipes_path = find_costing_file(house, map_path, here)
    recipes = load_recipes(recipes_path)
    if recipes is None:
        print('  Note: cocktail costing workbook not found; recipe lines will be listed for review.')

    unmapped = [ln for ln in items if ln['qty'] != 0 and pick_map_row(map_rows, ln['item'], ln['avg']) is None]
    if unmapped and prompt:
        new_rows = prompt_new_items(unmapped, mm60, house, recipes)
        if new_rows:
            append_to_map(map_path, new_rows)
            house, map_rows = load_map(map_path)

    for ln in items:
        m = pick_map_row(map_rows, ln['item'], ln['avg'])
        ln['priced'] = price_line(ln, m, mm60, house, recipes)
        ln['cat'] = category(ln['fam'])
        ln['suggest'] = []
        if m is None:
            ln['suggest'] = [f"{mm60[k]['desc']} #{k}" for k in suggestions(ln['item'], mm60, 3)]

    if output_path is None:
        base, ext = os.path.splitext(file_path)
        output_path = base + '_CoS_Summary.xlsx'
    n_review = write_report(output_path, meta, items, house, file_path, map_path)

    sold = [ln for ln in items if ln['qty'] != 0]
    priced = sum(1 for ln in sold if ln['priced']['included'])
    print(f'\n  Lines sold: {len(sold)}  |  in theoretical: {priced}  |  left out by choice: '
          f"{sum(1 for ln in sold if ln['priced']['status'].startswith('OK (')) }  |  to review/flagged: {n_review}")
    print(f'  Report: {output_path}')
    print('  Sheets: Summary (formulas), Detail (one row per POS line), Review (what needs a look)')
    return output_path


if __name__ == '__main__':
    args = sys.argv[1:]
    prompt = '--no-prompt' not in args and sys.stdin.isatty()
    args = [a for a in args if a != '--no-prompt']
    map_path = None
    if '--map' in args:
        i = args.index('--map')
        map_path = args[i + 1]
        del args[i:i + 2]
    if not args:
        print(__doc__)
        sys.exit(1)
    create_theoretical_report(args[0], args[1] if len(args) > 1 else None, map_path, prompt)
