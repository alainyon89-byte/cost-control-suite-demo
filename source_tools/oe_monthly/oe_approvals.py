#!/usr/bin/env python3
"""
=============================================================================
COST CONTROL SUITE  |  OE Approvals v3.1
(c) 2026 Alain Yon. All rights reserved.
Version: 3.1 | Last updated: 6 November 2026
=============================================================================
Reads the approval file the F&B Manager sent back (OE_Approval_<Mon>_<YYYY>.xlsx)
and writes OE_Approved_Orders_<Mon>_<YYYY>.xlsx with four tabs. A two-step
file (e.g. OE_Approval_Linen_Main_<Mon>_<YYYY>.xlsx, Executive Housekeeper then
Rooms Division Manager) gives OE_Approved_Orders_Linen_Main_<Mon>_<YYYY>.xlsx;
a line counts as approved only with Y in both steps.

    SUMMARY        totals, by location, and anything that needs a follow-up
    APPROVED       Y lines with the approved quantity, grouped by location
    NOT APPROVED   N lines
    PENDING        lines with no decision yet

Usage
    Drag the signed approval file onto Run_OE_Approvals.bat, or double-click
    the .bat to use the newest OE_Approval_*.xlsx in this folder.

Requires Python 3.8+ and openpyxl  (pip install openpyxl)
"""

import os
import re
import sys
import glob
import warnings
from datetime import datetime
from collections import defaultdict

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print('ERROR: openpyxl is not installed. Run:  python -m pip install openpyxl')
    sys.exit(1)

VERSION = '3.1'
LOC_RE = re.compile(r'^L\d{3}$')
C_TITLE, C_HEAD, C_TOTAL, C_SUB = 'FF6B35', '004E89', 'F0E5CF', 'DDEBF7'
F_MONEY, F_RAND, F_QTY = '#,##0.00', '"R "#,##0.00', '#,##0'
THIN = Side(style='thin', color='D9D9D9')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

NEEDED = {'location': 'loc', 'location name': 'name', 'material': 'mat',
          'description': 'desc', 'suggested qty': 'sugg', 'unit price': 'price',
          'urgency': 'urg', 'transfer option (over par elsewhere)': 'transfer',
          'approve (y/n)': 'dec', 'approved qty': 'qty', 'notes': 'notes',
          'step 1 approve (y/n)': 'dec1', 'step 1 qty': 'qty1', 'step 2 approve (y/n)': 'dec2'}


def fill(c):
    return PatternFill(start_color=c, end_color=c, fill_type='solid')


def text(v):
    return '' if v is None else str(v).strip()


def norm(v):
    return re.sub(r'\s+', ' ', text(v)).lower()


def num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = text(v).replace(',', '').replace(' ', '')
    try:
        return float(s) if s else None
    except ValueError:
        return None


def clean(x):
    if x is None:
        return None
    return int(x) if float(x).is_integer() else round(x, 4)


def file_is_locked(path):
    if not os.path.exists(path):
        return False
    try:
        with open(path, 'r+b'):
            return False
    except PermissionError:
        return True
    except OSError:
        return False


def find_approval_file(folder):
    files = [f for f in glob.glob(os.path.join(folder, 'OE_Approval_*.xlsx'))
             if not os.path.basename(f).startswith('~$')]
    return max(files, key=os.path.getmtime) if files else None


def read_approval(path):
    wb = load_workbook(path, data_only=True)
    if 'APPROVAL' not in wb.sheetnames:
        raise ValueError(f'{os.path.basename(path)} has no APPROVAL tab - is this an OE approval file?')
    period = None
    meta = {}
    if '_INFO' in wb.sheetnames:
        for k, v in wb['_INFO'].iter_rows(min_row=1, max_col=2, values_only=True):
            meta[text(k)] = text(v)
            if text(k) == 'Period':
                period = text(v)
    if not period:
        m = re.search(r'OE_Approval_([A-Za-z]{3})_(\d{4})', os.path.basename(path))
        period = f'{m.group(1)} {m.group(2)}' if m else datetime.now().strftime('%b %Y')
    ws = wb['APPROVAL']
    hdr, cmap = None, {}
    for r in range(1, 21):
        vals = [norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if ('approve (y/n)' in vals or 'step 1 approve (y/n)' in vals) and 'material' in vals:
            hdr = r
            for c, v in enumerate(vals, 1):
                if v in NEEDED:
                    cmap[NEEDED[v]] = c
            break
    if hdr is None:
        raise ValueError('Header row with "Approve (Y/N)" not found on the APPROVAL tab')
    two_step = 'dec1' in cmap
    need = ('loc', 'mat', 'sugg', 'price', 'dec1', 'dec2', 'qty') if two_step else \
        ('loc', 'mat', 'sugg', 'price', 'dec', 'qty')
    missing = [k for k in need if k not in cmap]
    if missing:
        raise ValueError(f'APPROVAL tab is missing columns: {missing}')

    lines = []
    for r in range(hdr + 1, ws.max_row + 1):
        loc = text(ws.cell(r, cmap['loc']).value).upper()
        if not LOC_RE.match(loc):
            break       # end of data (blank row, then signature block)
        get = lambda k: ws.cell(r, cmap[k]).value if k in cmap else None
        lines.append({'row': r, 'loc': loc, 'name': text(get('name')), 'mat': get('mat'),
                      'desc': text(get('desc')), 'sugg': num(get('sugg')) or 0,
                      'price': num(get('price')) or 0.0, 'urg': text(get('urg')),
                      'transfer': text(get('transfer')), 'dec_raw': text(get('dec')),
                      'dec1_raw': text(get('dec1')), 'dec2_raw': text(get('dec2')),
                      'qty1_raw': get('qty1'),
                      'qty_raw': get('qty'), 'notes': text(get('notes'))})
    wb.close()
    meta['two_step'] = two_step
    return period, lines, meta


YN = {'Y': 'Y', 'YES': 'Y', 'N': 'N', 'NO': 'N'}


def decide(lines, meta=None):
    meta = meta or {}
    two = meta.get('two_step')
    s1, s2 = meta.get('Approver 1') or 'Step 1', meta.get('Approver 2') or 'Step 2'
    issues = []
    for d in lines:
        if two:
            d1, d2 = YN.get(d['dec1_raw'].upper(), ''), YN.get(d['dec2_raw'].upper(), '')
            for raw, ok, who in ((d['dec1_raw'], d1, s1), (d['dec2_raw'], d2, s2)):
                if raw and not ok:
                    issues.append((d, f'{who} entered "{raw}" - not Y or N, read as blank'))
            if 'N' in (d1, d2):
                d['decision'] = 'N'
                d['dec_raw'] = f'{d1 or "-"}/{d2 or "-"}'
            elif d1 == 'Y' and d2 == 'Y':
                d['decision'] = 'Y'
                d['dec_raw'] = 'Y/Y'
            else:
                d['decision'] = ''
                d['dec_raw'] = f'{d1 or "-"}/{d2 or "-"}'
                if d1 == 'Y':
                    issues.append((d, f'Waiting on {s2}'))
                elif d2 == 'Y':
                    issues.append((d, f'{s2} said Y but {s1} has not decided'))
            if d['decision'] == 'Y' and num(d['qty_raw']) is None and num(d['qty1_raw']) is not None:
                d['qty_raw'] = d['qty1_raw']    # Approved Qty formula not calculated yet
        else:
            raw = d['dec_raw'].upper()
            d['decision'] = YN.get(raw, '')
            if raw and not d['decision']:
                issues.append((d, f'Approve column says "{d["dec_raw"]}" - not Y or N, left as pending'))
        q = num(d['qty_raw'])
        if d['decision'] == 'Y':
            if q is None:
                q = d['sugg']
                issues.append((d, 'Approved Qty blank - suggested qty used'))
            elif q <= 0:
                d['decision'] = 'N'
                issues.append((d, 'Approved with qty 0 - treated as not approved'))
            elif q > d['sugg']:
                issues.append((d, f'Approved {clean(q)}, more than the {clean(d["sugg"])} suggested'))
            if q is not None and q != int(q):
                issues.append((d, f'Approved Qty {q} is not a whole number'))
        d['qty'] = q if d['decision'] == 'Y' else 0
        d['value'] = round((d['qty'] or 0) * d['price'], 2)
        d['sugg_value'] = round(d['sugg'] * d['price'], 2)
    return issues


# ---------------------------------------------------------------------------
def title(ws, ref, value):
    ws.merge_cells(ref)
    c = ws[ref.split(':')[0]]
    c.value = value
    c.font = Font(bold=True, size=14, color='FFFFFF')
    c.fill = fill(C_TITLE)
    c.alignment = Alignment(vertical='center', indent=1)
    ws.row_dimensions[c.row].height = 26


def header(ws, row, cols):
    for j, h in enumerate(cols, 1):
        c = ws.cell(row, j, h)
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = fill(C_HEAD)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 30


def setup(ws, widths, title_row):
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f'{title_row}:{title_row}'
    ws.oddFooter.center.text = f'(c) 2026 Alain Yon | OE Approvals v{VERSION} | Page &P of &N'
    ws.oddFooter.center.size = 8
    ws.freeze_panes = f'A{title_row + 1}'


def write_approved(ws, period, lines):
    cols = ['Location', 'Location Name', 'Material', 'Description', 'Approved Qty',
            'Unit Price', 'Value', 'Urgency', 'Transfer Option', 'Notes']
    title(ws, 'A1:J1', f'APPROVED OE ORDERS  |  {period}')
    header(ws, 3, cols)
    r = 4
    first_data = r
    by_loc = defaultdict(list)
    for d in lines:
        by_loc[d['loc']].append(d)
    for loc in sorted(by_loc):
        start = r
        for d in by_loc[loc]:
            vals = [d['loc'], d['name'], d['mat'], d['desc'], clean(d['qty']), d['price'],
                    f'=E{r}*F{r}', d['urg'], d['transfer'] or None, d['notes'] or None]
            for j, v in enumerate(vals, 1):
                c = ws.cell(r, j, v)
                c.border = BORDER
            ws.cell(r, 5).number_format = F_QTY
            ws.cell(r, 6).number_format = F_MONEY
            ws.cell(r, 7).number_format = F_MONEY
            r += 1
        ws.cell(r, 1, loc).font = Font(bold=True)
        ws.cell(r, 2, f'{by_loc[loc][0]["name"]} total').font = Font(bold=True)
        ws.cell(r, 5, f'=SUBTOTAL(9,E{start}:E{r - 1})').number_format = F_QTY
        ws.cell(r, 7, f'=SUBTOTAL(9,G{start}:G{r - 1})').number_format = F_MONEY
        for j in range(1, 11):
            ws.cell(r, j).fill = fill(C_SUB)
            ws.cell(r, j).font = Font(bold=True)
            ws.cell(r, j).border = BORDER
        r += 2
    ws.cell(r, 1, 'GRAND TOTAL').font = Font(bold=True)
    if lines:
        ws.cell(r, 7, f'=SUBTOTAL(9,G{first_data}:G{r - 1})').number_format = F_RAND
    else:
        ws.cell(r, 7, 0).number_format = F_RAND
        ws.cell(first_data, 1, 'No approved lines')
    for j in range(1, 11):
        ws.cell(r, j).fill = fill(C_TOTAL)
        ws.cell(r, j).font = Font(bold=True)
    setup(ws, [9, 18, 10, 42, 10, 11, 13, 9, 30, 30], 3)


def write_list(ws, heading, lines, decision_col=True):
    cols = ['Location', 'Location Name', 'Material', 'Description', 'Suggested Qty',
            'Unit Price', 'Suggested Value', 'Urgency', 'Entered', 'Notes']
    title(ws, 'A1:J1', heading)
    header(ws, 3, cols)
    for i, d in enumerate(lines):
        r = 4 + i
        vals = [d['loc'], d['name'], d['mat'], d['desc'], clean(d['sugg']), d['price'],
                d['sugg_value'], d['urg'], d['dec_raw'] or None, d['notes'] or None]
        for j, v in enumerate(vals, 1):
            ws.cell(r, j, v).border = BORDER
        ws.cell(r, 5).number_format = F_QTY
        ws.cell(r, 6).number_format = F_MONEY
        ws.cell(r, 7).number_format = F_MONEY
    if not lines:
        ws.cell(4, 1, 'None')
    setup(ws, [9, 18, 10, 42, 10, 11, 13, 9, 9, 40], 3)


def write_summary(ws, period, src, built, groups, issues, names, meta=None):
    meta = meta or {}
    label = meta.get('Label', '')
    title(ws, 'A1:G1', f'OE APPROVALS  |  {label + "  |  " if label else ""}{period}')
    ws['A2'] = f'From {src}  |  processed {built:%d %b %Y %H:%M}  |  oe_approvals.py v{VERSION}'
    ws['A2'].font = Font(italic=True, size=9, color='808080')
    if meta.get('two_step'):
        ws['A3'] = (f'Two-step sign-off: {meta.get("Approver 1", "Step 1")} then '
                    f'{meta.get("Approver 2", "Step 2")}. Approved = Y in both steps.')
        ws['A3'].font = Font(italic=True, size=9)
    header(ws, 4, ['Status', 'Lines', 'Value', '', '', '', ''])
    labels = (('Approved', 'Y', 'value'), ('Not approved', 'N', 'sugg_value'),
              ('Pending (no decision)', '', 'sugg_value'))
    for i, (lab, key, vk) in enumerate(labels):
        r = 5 + i
        ws.cell(r, 1, lab).font = Font(bold=True)
        ws.cell(r, 2, len(groups[key]))
        c = ws.cell(r, 3, round(sum(d[vk] for d in groups[key]), 2))
        c.number_format = F_RAND
    ws.cell(8, 1, 'Not approved and pending show the suggested value.').font = \
        Font(italic=True, size=9, color='808080')

    header(ws, 10, ['Location', 'Name', 'Approved Lines', 'Approved Value',
                    'Not Approved', 'Pending', ''])
    locs = sorted({d['loc'] for g in groups.values() for d in g})
    for i, loc in enumerate(locs):
        r = 11 + i
        ws.cell(r, 1, loc).font = Font(bold=True)
        ws.cell(r, 2, names.get(loc, ''))
        ws.cell(r, 3, sum(1 for d in groups['Y'] if d['loc'] == loc))
        c = ws.cell(r, 4, round(sum(d['value'] for d in groups['Y'] if d['loc'] == loc), 2))
        c.number_format = F_MONEY
        ws.cell(r, 5, sum(1 for d in groups['N'] if d['loc'] == loc))
        ws.cell(r, 6, sum(1 for d in groups[''] if d['loc'] == loc))
        for j in range(1, 7):
            ws.cell(r, j).border = BORDER
    r = 11 + len(locs)
    ws.cell(r, 1, 'TOTAL').font = Font(bold=True)
    for j, L in ((3, 'C'), (4, 'D'), (5, 'E'), (6, 'F')):
        c = ws.cell(r, j, f'=SUM({L}11:{L}{r - 1})' if locs else 0)
        c.font = Font(bold=True)
        c.number_format = F_MONEY if j == 4 else F_QTY
    for j in range(1, 7):
        ws.cell(r, j).fill = fill(C_TOTAL)

    r += 2
    ws.cell(r, 1, 'FOLLOW UP').font = Font(bold=True, color='C00000')
    r += 1
    if not issues:
        ws.cell(r, 1, 'Nothing to follow up')
    for d, msg in issues:
        ws.cell(r, 1, d['loc'])
        ws.cell(r, 2, d['mat'])
        ws.cell(r, 3, d['desc'])
        ws.cell(r, 5, msg)
        r += 1
    for col, w in zip('ABCDEFG', (22, 22, 16, 16, 14, 10, 4)):
        ws.column_dimensions[col].width = w


def main(argv):
    folder = os.path.dirname(os.path.abspath(__file__))
    path = next((os.path.abspath(a) for a in argv if a.lower().endswith('.xlsx')), None)
    print('=' * 70)
    print(f'  OE APPROVALS v{VERSION}')
    print('=' * 70)
    if path is None:
        path = find_approval_file(folder)
        if path is None:
            print('\nERROR: no OE_Approval_*.xlsx found. Drag the signed file onto the .bat.')
            return 1
        print(f'  Using newest approval file: {os.path.basename(path)}')
    if not os.path.exists(path):
        print(f'\nERROR: file not found: {path}')
        return 1

    period, lines, meta = read_approval(path)
    issues = decide(lines, meta)
    groups = {'Y': [], 'N': [], '': []}
    for d in lines:
        groups[d['decision']].append(d)
    names = {d['loc']: d['name'] for d in lines}
    built = datetime.now()

    slug = re.sub(r'[^A-Za-z0-9]+', '_', period).strip('_')
    label = re.sub(r'[^A-Za-z0-9]+', '_', meta.get('Label', '')).strip('_')
    out = os.path.join(os.path.dirname(path),
                       f'OE_Approved_Orders_{label + "_" if label else ""}{slug}.xlsx')
    if file_is_locked(out):
        out = out.replace('.xlsx', f'_{built:%Y%m%d_%H%M}.xlsx')
    wb = Workbook()
    write_summary(wb.active, period, os.path.basename(path), built, groups, issues, names, meta)
    wb.active.title = 'SUMMARY'
    write_approved(wb.create_sheet('APPROVED'), period, groups['Y'])
    write_list(wb.create_sheet('NOT APPROVED'), f'NOT APPROVED  |  {period}', groups['N'])
    write_list(wb.create_sheet('PENDING'), f'PENDING - NO DECISION  |  {period}', groups[''])
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)

    print(f'\n  Lines read      {len(lines)}')
    print(f'  Approved        {len(groups["Y"]):>4}   R{sum(d["value"] for d in groups["Y"]):>14,.2f}')
    print(f'  Not approved    {len(groups["N"]):>4}')
    print(f'  Pending         {len(groups[""]):>4}')
    if issues:
        print(f'\n  Follow up ({len(issues)}):')
        for d, msg in issues[:15]:
            print(f'    {d["loc"]} {d["mat"]} {d["desc"][:30]}: {msg}')
        if len(issues) > 15:
            print(f'    ... and {len(issues) - 15} more on the SUMMARY tab')
    if groups['']:
        who = 'the approvers' if meta.get('two_step') else 'the F&B Manager'
        print(f'\n  ! Some lines have no decision. Ask {who}, then run this again.')
    print(f'\n  Saved: {os.path.basename(out)}')
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main(sys.argv[1:]))
    except Exception as exc:
        import traceback
        traceback.print_exc()
        print(f'\nERROR: {exc}')
        sys.exit(1)
