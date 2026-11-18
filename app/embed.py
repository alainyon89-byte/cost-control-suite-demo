# (c) 2026 Alain Yon
"""'embedded_sheet' delivery of the shared MM60 (spec 6A).

Some scripts read MM60 as a sheet inside another input (Theoretical CoS reads
the 'MM60' sheet inside the Micros Sales Mix workbook). The app makes a temp
copy of the user's file, writes the shared MM60 in as that sheet, and runs the
script on the copy. The user's original file is never changed.
"""
import warnings

warnings.filterwarnings("ignore")


class EmbedError(Exception):
    pass


def has_sheet(path, sheet):
    from openpyxl import load_workbook
    try:
        wb = load_workbook(path, read_only=True)
        try:
            return sheet in wb.sheetnames
        finally:
            wb.close()
    except Exception:
        return False


def embed_sheet(user_file, shared_file, dest_path, sheet="MM60"):
    """Copy user_file to dest_path with `sheet` replaced by the first sheet of shared_file (values only)."""
    from openpyxl import load_workbook
    wb = load_workbook(user_file)
    for ws in wb.worksheets:
        if ws.title == sheet:
            continue
        for row in ws.iter_rows():
            for c in row:
                if c.data_type == "f":
                    raise EmbedError(
                        f"{ws.title}!{c.coordinate} holds a formula. Saving a copy would lose its calculated value, "
                        "so the shared MM60 was not written into this file.")
    if sheet in wb.sheetnames:
        del wb[sheet]
    target = wb.create_sheet(sheet)
    src = load_workbook(shared_file, read_only=True, data_only=True)
    try:
        for row in src.worksheets[0].iter_rows(values_only=True):
            target.append(list(row))
    finally:
        src.close()
    wb.save(dest_path)
    return dest_path
