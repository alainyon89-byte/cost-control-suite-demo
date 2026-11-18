# (c) 2026 Alain Yon
"""Compare two .xlsx files and list every difference.

Fails on: sheet names / order / visibility, cell values and formula text,
number formats, cell styles (font, fill, border, alignment, lock), merged
ranges, data validations, conditional formats, hyperlinks, comments, sheet and
workbook protection, defined names.
Reports only (never fails): column widths, row heights, print setup.

Ignore rules (from a tool's "ignore_in_regression" list in tools.json):
  {"sheet": "Summary", "cell": "A4", "why": "run date"}   one cell (sheet may be "*")
  {"mask": "run \\d{4}-\\d\\d-\\d\\d \\d\\d:\\d\\d", "why": "run date"}
        text matching the regex is blanked on both sides before comparing
Everything ignored is listed in the result.

Usage:  python compare_xlsx.py a.xlsx b.xlsx
"""
import fnmatch
import re
import sys
import warnings

from openpyxl import load_workbook

warnings.filterwarnings("ignore")


def _mask(value, masks):
    if isinstance(value, str):
        for rx in masks:
            value = rx.sub("<ignored>", value)
    return value


def _style_key(cell):
    if not cell.has_style:
        return None
    f, fl, b, a, p = cell.font, cell.fill, cell.border, cell.alignment, cell.protection

    def col(c):
        if c is None:
            return None
        return (c.type, c.rgb if c.type == "rgb" else c.theme if c.type == "theme" else c.indexed, c.tint)

    def side(s):
        return (s.style, col(s.color)) if s is not None else None

    return (
        (f.name, f.sz, f.b, f.i, f.u, f.strike, col(f.color), f.vertAlign),
        (fl.fill_type, col(getattr(fl, "fgColor", None)), col(getattr(fl, "bgColor", None))),
        (side(b.left), side(b.right), side(b.top), side(b.bottom)),
        (a.horizontal, a.vertical, a.wrap_text, a.text_rotation, a.indent, a.shrink_to_fit),
        (p.locked, p.hidden),
    )


def _dv_key(dv):
    return (str(dv.sqref), dv.type, dv.operator, dv.formula1, dv.formula2, dv.allow_blank,
            dv.showErrorMessage, dv.showInputMessage, dv.errorTitle, dv.error, dv.promptTitle, dv.prompt)


def _cf_key(ws):
    out = []
    for rng, rules in ws.conditional_formatting._cf_rules.items():
        for r in rules:
            dxf = r.dxf
            fill = getattr(getattr(dxf, "fill", None), "bgColor", None) if dxf else None
            font = getattr(getattr(dxf, "font", None), "color", None) if dxf else None
            out.append((str(rng.sqref), r.type, r.operator, tuple(r.formula or ()), r.priority, r.stopIfTrue,
                        getattr(fill, "rgb", None), getattr(font, "rgb", None)))
    return sorted(out, key=repr)


def _protection_key(ws):
    p = ws.protection
    return tuple((k, getattr(p, k)) for k in (
        "sheet", "password", "objects", "scenarios", "formatCells", "formatColumns", "formatRows",
        "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns", "deleteRows",
        "selectLockedCells", "selectUnlockedCells", "sort", "autoFilter", "pivotTables"))


def _wb_protection_key(wb):
    s = wb.security
    if s is None:
        return None
    return (s.lockStructure, s.lockWindows, s.workbookPassword, s.revisionsPassword)


def _defined_names(wb):
    out = {}
    try:
        for name, dn in wb.defined_names.items():
            out[name] = dn.attr_text
    except AttributeError:
        for dn in wb.defined_names.definedName:
            out[dn.name] = dn.attr_text
    return out


def _print_key(ws):
    ps, pm = ws.page_setup, ws.page_margins
    return (ps.orientation, ps.paperSize, ps.fitToWidth, ps.fitToHeight, ps.scale,
            ws.print_area, ws.print_title_rows, ws.print_title_cols,
            pm.left, pm.right, pm.top, pm.bottom,
            ws.oddHeader.center.text if ws.oddHeader else None,
            ws.oddFooter.center.text if ws.oddFooter else None,
            str(ws.freeze_panes), ws.sheet_properties.tabColor.rgb if ws.sheet_properties.tabColor else None)


def compare(path_a, path_b, ignore=None, max_per_kind=40):
    """Returns dict(diffs=[...], notes=[...], ignored=[...]).  Identical when diffs is empty."""
    ignore = ignore or []
    masks = [re.compile(r["mask"]) for r in ignore if "mask" in r]
    cell_rules = [(r.get("sheet", "*"), r["cell"].upper(), r.get("why", "")) for r in ignore if "cell" in r]
    diffs, notes, ignored = [], [], []
    counts = {}

    def add(kind, msg):
        counts[kind] = counts.get(kind, 0) + 1
        if counts[kind] <= max_per_kind:
            diffs.append(f"{kind}: {msg}")
        elif counts[kind] == max_per_kind + 1:
            diffs.append(f"{kind}: ... more of the same not listed")

    def cell_ignored(sheet, coord):
        for sh, c, why in cell_rules:
            if fnmatch.fnmatch(sheet, sh) and fnmatch.fnmatch(coord, c):
                return why or "listed in ignore_in_regression"
        return None

    wa = load_workbook(path_a)
    wb = load_workbook(path_b)

    if wa.sheetnames != wb.sheetnames:
        add("sheets", f"names/order differ: {wa.sheetnames} vs {wb.sheetnames}")
    if _wb_protection_key(wa) != _wb_protection_key(wb):
        add("workbook protection", f"{_wb_protection_key(wa)} vs {_wb_protection_key(wb)}")
    da, db = _defined_names(wa), _defined_names(wb)
    for n in sorted(set(da) | set(db)):
        if da.get(n) != db.get(n):
            add("defined name", f"{n}: {da.get(n)!r} vs {db.get(n)!r}")
    ca = getattr(wa.calculation, "fullCalcOnLoad", None)
    cb = getattr(wb.calculation, "fullCalcOnLoad", None)
    if ca != cb:
        add("workbook", f"fullCalcOnLoad {ca} vs {cb}")

    for name in wa.sheetnames:
        if name not in wb.sheetnames:
            continue
        sa, sb = wa[name], wb[name]
        if sa.sheet_state != sb.sheet_state:
            add("visibility", f"{name}: {sa.sheet_state} vs {sb.sheet_state}")
        if not hasattr(sa, "_cells") or not hasattr(sb, "_cells"):
            notes.append(f"{name}: chart sheet, not compared cell by cell")
            continue
        if sa.dimensions != sb.dimensions:
            add("size", f"{name}: used range {sa.dimensions} vs {sb.dimensions}")
        keys = set(sa._cells) | set(sb._cells)
        for key in sorted(keys):
            a, b = sa._cells.get(key), sb._cells.get(key)
            coord = (a or b).coordinate
            va = _mask(a.value, masks) if a is not None else None
            vb = _mask(b.value, masks) if b is not None else None
            raw_differs = (a.value if a is not None else None) != (b.value if b is not None else None)
            if va != vb or type(va) is not type(vb):
                why = cell_ignored(name, coord)
                if why:
                    ignored.append(f"{name}!{coord}: {why}")
                else:
                    add("value", f"{name}!{coord}: {va!r} vs {vb!r}")
            elif raw_differs:
                ignored.append(f"{name}!{coord}: masked text differs ({a.value!r} vs {b.value!r})")
            if a is not None and b is not None:
                if a.number_format != b.number_format:
                    add("number format", f"{name}!{coord}: {a.number_format!r} vs {b.number_format!r}")
                if a.style_id != b.style_id or wa is not wb:
                    if _style_key(a) != _style_key(b):
                        add("style", f"{name}!{coord}: font/fill/border/alignment differ")
                ha = a.hyperlink.target or a.hyperlink.location if a.hyperlink else None
                hb = b.hyperlink.target or b.hyperlink.location if b.hyperlink else None
                if ha != hb:
                    add("hyperlink", f"{name}!{coord}: {ha!r} vs {hb!r}")
                ta = a.comment.text if a.comment else None
                tb = b.comment.text if b.comment else None
                if ta != tb:
                    add("comment", f"{name}!{coord}")
        ma = sorted(str(r) for r in sa.merged_cells.ranges)
        mb = sorted(str(r) for r in sb.merged_cells.ranges)
        if ma != mb:
            add("merged ranges", f"{name}: {len(ma)} vs {len(mb)} ranges; first difference "
                                 f"{sorted(set(ma) ^ set(mb))[:5]}")
        va_ = sorted((_dv_key(d) for d in sa.data_validations.dataValidation), key=repr)
        vb_ = sorted((_dv_key(d) for d in sb.data_validations.dataValidation), key=repr)
        if va_ != vb_:
            add("data validation", f"{name}: {len(va_)} vs {len(vb_)} rules or their settings differ")
        if _cf_key(sa) != _cf_key(sb):
            add("conditional format", f"{name}: rules differ")
        if _protection_key(sa) != _protection_key(sb):
            add("sheet protection", f"{name}: protection flags or password differ")
        if (sa.auto_filter.ref or None) != (sb.auto_filter.ref or None):
            add("autofilter", f"{name}: {sa.auto_filter.ref} vs {sb.auto_filter.ref}")
        ta_ = sorted((t.name, t.ref) for t in sa.tables.values())
        tb_ = sorted((t.name, t.ref) for t in sb.tables.values())
        if ta_ != tb_:
            add("table", f"{name}: {ta_} vs {tb_}")
        if (len(sa._charts), len(sa._images)) != (len(sb._charts), len(sb._images)):
            add("charts/images", f"{name}: {len(sa._charts)}/{len(sa._images)} vs {len(sb._charts)}/{len(sb._images)}")
        # report only
        wda = {k: v.width for k, v in sa.column_dimensions.items() if v.width}
        wdb = {k: v.width for k, v in sb.column_dimensions.items() if v.width}
        if wda != wdb:
            notes.append(f"{name}: column widths differ (report only)")
        hta = {k: v.height for k, v in sa.row_dimensions.items() if v.height}
        htb = {k: v.height for k, v in sb.row_dimensions.items() if v.height}
        if hta != htb:
            notes.append(f"{name}: row heights differ (report only)")
        if _print_key(sa) != _print_key(sb):
            notes.append(f"{name}: print setup / freeze panes / tab colour differ (report only)")
    wa.close()
    wb.close()
    # collapse the ignored list
    if len(ignored) > 30:
        ignored = ignored[:30] + [f"... and {len(ignored) - 30} more ignored cells"]
    return {"diffs": diffs, "notes": notes, "ignored": ignored}


def main(argv):
    if len(argv) != 2:
        print(__doc__)
        return 2
    res = compare(argv[0], argv[1])
    for d in res["diffs"]:
        print("DIFF  ", d)
    for n in res["notes"]:
        print("NOTE  ", n)
    print("IDENTICAL" if not res["diffs"] else f"{len(res['diffs'])} difference line(s)")
    return 0 if not res["diffs"] else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
