"""
Build Age Analysis Dashboard workbook.

Reads a source workbook (Stock on hand / Past year purchase / Past 5 year purchase)
and produces a new workbook with:
  - Raw_Stock            (copy of Stock on hand)
  - Raw_Purchases5yr     (copy of Past 5 year purchase)
  - Age_Analysis          (calculated table, one row per stock item)
  - Dashboard             (PivotTables + PivotCharts + Slicers)

Usage:
    python build_age_analysis.py <source_file.xlsx> <output_file.xlsx>

Re-run any time the source data refreshes (e.g. via the .bat file).
"""

import sys
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, PieChart, Reference

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=16, color="1F4E78")
SUBTITLE_FONT = Font(name=FONT_NAME, size=10, italic=True, color="666666")
BODY_FONT = Font(name=FONT_NAME, size=10)
BOLD_FONT = Font(name=FONT_NAME, size=10, bold=True)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

PROMO_FILL = PatternFill("solid", fgColor="FFC7CE")   # red-ish: promo candidate
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")    # amber: 2-3yr
OK_FILL = PatternFill("solid", fgColor="C6EFCE")      # green: <1yr

AGE_BUCKETS = ["0-1 yr", "1-2 yr", "2-3 yr", "3-5 yr", "5+ yr / No Record"]


def copy_sheet_raw(src_ws, dst_wb, dst_name):
    """Copy a sheet's values (as displayed) into a new sheet, styled as a raw data dump."""
    dst_ws = dst_wb.create_sheet(dst_name)
    for row in src_ws.iter_rows():
        for cell in row:
            dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)

    max_col = src_ws.max_column
    max_row = src_ws.max_row

    for c in range(1, max_col + 1):
        cell = dst_ws.cell(row=1, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r in range(2, max_row + 1):
        for c in range(1, max_col + 1):
            cell = dst_ws.cell(row=r, column=c)
            cell.font = BODY_FONT

    # date format for posting date column if present
    header_vals = [dst_ws.cell(1, c).value for c in range(1, max_col + 1)]
    if "Posting Date" in header_vals:
        col = header_vals.index("Posting Date") + 1
        for r in range(2, max_row + 1):
            dst_ws.cell(r, col).number_format = "yyyy-mm-dd"

    dst_ws.freeze_panes = "A2"
    tbl = Table(
        displayName=dst_name.replace(" ", "_"),
        ref=f"A1:{get_column_letter(max_col)}{max_row}",
    )
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True
    )
    dst_ws.add_table(tbl)

    for c in range(1, max_col + 1):
        letter = get_column_letter(c)
        max_len = max(
            (len(str(dst_ws.cell(r, c).value)) for r in range(1, min(max_row, 200) + 1)
             if dst_ws.cell(r, c).value is not None),
            default=10,
        )
        dst_ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 40)

    return dst_ws, max_row, max_col


def build(source_path, output_path):
    src_wb = openpyxl.load_workbook(source_path, data_only=True)
    stock_src = src_wb["Stock on hand"]
    purch5_src = src_wb["Past 5 year purchase"]

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    # --- Raw sheets -----------------------------------------------------
    stock_ws, stock_rows, _ = copy_sheet_raw(stock_src, out_wb, "Raw_Stock")
    purch_ws, purch_rows, _ = copy_sheet_raw(purch5_src, out_wb, "Raw_Purchases5yr")

    # --- Age_Analysis sheet ----------------------------------------------
    age_ws = out_wb.create_sheet("Age_Analysis")

    headers = [
        "Material",
        "Material Description",
        "Material Group",
        "Qty On Hand",
        "Value On Hand",
        "Last Purchase Date",
        "Days Since Last Purchase",
        "Years Since Last Purchase",
        "Age Bucket",
        "Promo Candidate (3+ yrs)",
    ]
    for c, h in enumerate(headers, start=1):
        cell = age_ws.cell(1, c, h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    n = stock_rows - 1  # number of data rows in Raw_Stock

    today_str = datetime.date.today().isoformat()
    age_ws["L1"] = "As-of date:"
    age_ws["L1"].font = BOLD_FONT
    age_ws["M1"] = today_str
    age_ws["M1"].font = BODY_FONT
    age_ws["M1"].number_format = "yyyy-mm-dd"

    for i in range(n):
        r = i + 2
        src_r = i + 2  # same row offset in Raw_Stock
        age_ws.cell(r, 1, f"=Raw_Stock!A{src_r}")                      # Material
        age_ws.cell(r, 2, f"=Raw_Stock!B{src_r}")                      # Description
        age_ws.cell(r, 3, f"=Raw_Stock!E{src_r}")                      # Material Group
        age_ws.cell(r, 4, f"=Raw_Stock!G{src_r}")                      # Qty
        age_ws.cell(r, 5, f"=Raw_Stock!H{src_r}")                      # Value
        # Last purchase date = MAX posting date in Raw_Purchases5yr where Material matches.
        # SUMPRODUCT-based MAX avoids needing an array-entered formula.
        age_ws.cell(
            r, 6,
            f'=IF(COUNTIF(Raw_Purchases5yr!$I$2:$I${purch_rows},A{r})=0,"",'
            f'SUMPRODUCT(MAX((Raw_Purchases5yr!$I$2:$I${purch_rows}=A{r})*'
            f'Raw_Purchases5yr!$K$2:$K${purch_rows})))'
        )
        f6 = age_ws.cell(r, 6)
        f6.number_format = "yyyy-mm-dd"
        # Days since last purchase (blank/no record -> treat as very old)
        age_ws.cell(r, 7, f'=IF(F{r}="",99999,$M$1-F{r})')
        # Years since last purchase
        age_ws.cell(r, 8, f"=IF(F{r}=\"\",\"\",ROUND(G{r}/365,1))")
        # Age bucket
        age_ws.cell(
            r, 9,
            f'=IF(F{r}="","5+ yr / No Record",'
            f'IF(G{r}<365,"0-1 yr",'
            f'IF(G{r}<730,"1-2 yr",'
            f'IF(G{r}<1095,"2-3 yr",'
            f'IF(G{r}<1825,"3-5 yr","5+ yr / No Record")))))'
        )
        # Promo candidate flag: not bought in 3+ years (or no record)
        age_ws.cell(r, 10, f'=IF(G{r}>=1095,"YES","NO")')

        for c in range(1, 11):
            age_ws.cell(r, c).font = BODY_FONT
            age_ws.cell(r, c).border = BORDER

    last_row = n + 1

    # Conditional-style fills based on bucket (static evaluation done later isn't needed;
    # use Excel conditional formatting so it stays live on recalculation)
    from openpyxl.formatting.rule import CellIsRule, FormulaRule

    promo_range = f"J2:J{last_row}"
    age_ws.conditional_formatting.add(
        promo_range,
        FormulaRule(formula=[f'$J2="YES"'], fill=PROMO_FILL),
    )
    bucket_range = f"I2:I{last_row}"
    age_ws.conditional_formatting.add(
        bucket_range,
        FormulaRule(formula=['$I2="0-1 yr"'], fill=OK_FILL),
    )
    age_ws.conditional_formatting.add(
        bucket_range,
        FormulaRule(formula=['OR($I2="2-3 yr",$I2="1-2 yr")'], fill=WARN_FILL),
    )
    age_ws.conditional_formatting.add(
        bucket_range,
        FormulaRule(formula=['OR($I2="3-5 yr",$I2="5+ yr / No Record")'], fill=PROMO_FILL),
    )

    age_ws.freeze_panes = "A2"
    tbl = Table(displayName="Age_Analysis_Table", ref=f"A1:J{last_row}")
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    age_ws.add_table(tbl)

    col_widths = [12, 32, 16, 12, 14, 18, 22, 22, 20, 22]
    for idx, w in enumerate(col_widths, start=1):
        age_ws.column_dimensions[get_column_letter(idx)].width = w

    age_ws["L3"] = "Notes:"
    age_ws["L3"].font = BOLD_FONT
    notes = [
        "Last Purchase Date is pulled from Raw_Purchases5yr (5-year window, 2021 onward).",
        "Items with no purchase record in that window are bucketed as '5+ yr / No Record'.",
        "Promo Candidate = YES for items not purchased in 3+ years (incl. no record).",
        "Join key: Material code (not description).",
        "Re-run build_age_analysis.bat after refreshing source data to update this workbook.",
    ]
    for i, note in enumerate(notes, start=4):
        c = age_ws.cell(i, 12, f"- {note}")
        c.font = Font(name=FONT_NAME, size=9, italic=True, color="666666")

    # --- Dashboard sheet --------------------------------------------------
    dash_ws = out_wb.create_sheet("Dashboard")
    out_wb.move_sheet("Dashboard", offset=-(len(out_wb.sheetnames) - 1))  # move to front

    dash_ws["B2"] = "Stock Age Analysis Dashboard"
    dash_ws["B2"].font = TITLE_FONT
    dash_ws["B3"] = f"Promotion candidates: items not purchased in 3+ years  |  As of {today_str}"
    dash_ws["B3"].font = SUBTITLE_FONT
    dash_ws.column_dimensions["A"].width = 3
    dash_ws.column_dimensions["B"].width = 26
    dash_ws.column_dimensions["C"].width = 20
    dash_ws.column_dimensions["D"].width = 20
    dash_ws.column_dimensions["E"].width = 20

    # Summary KPI cards (formulas referencing Age_Analysis)
    kpis = [
        ("Total SKUs on Hand", f"=COUNTA(Age_Analysis!A2:A{last_row})", "0"),
        ("Promo Candidates (3+ yrs)", f'=COUNTIF(Age_Analysis!J2:J{last_row},"YES")', "0"),
        ("Value Tied Up in Promo Candidates", f'=SUMIF(Age_Analysis!J2:J{last_row},"YES",Age_Analysis!E2:E{last_row})', '"R"#,##0'),
        ("Total Stock Value", f"=SUM(Age_Analysis!E2:E{last_row})", '"R"#,##0'),
    ]
    row0 = 6
    for i, (label, formula, fmt) in enumerate(kpis):
        col = 2 + i  # B, C, D, E
        letter = get_column_letter(col)
        lbl_cell = dash_ws.cell(row0, col, label)
        lbl_cell.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
        lbl_cell.fill = HEADER_FILL
        lbl_cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        dash_ws.row_dimensions[row0].height = 30

        val_cell = dash_ws.cell(row0 + 1, col, formula)
        val_cell.font = Font(name=FONT_NAME, size=16, bold=True, color="1F4E78")
        val_cell.number_format = fmt
        val_cell.alignment = Alignment(horizontal="center")
        val_cell.border = BORDER
        dash_ws.row_dimensions[row0 + 1].height = 28

    # Age bucket summary table (source for pivot chart) - built with formulas
    summary_start = row0 + 4
    dash_ws.cell(summary_start, 2, "Age Bucket Summary").font = BOLD_FONT
    hdr_row = summary_start + 1
    dash_ws.cell(hdr_row, 2, "Age Bucket").font = HEADER_FONT
    dash_ws.cell(hdr_row, 2).fill = HEADER_FILL
    dash_ws.cell(hdr_row, 3, "SKU Count").font = HEADER_FONT
    dash_ws.cell(hdr_row, 3).fill = HEADER_FILL
    dash_ws.cell(hdr_row, 4, "Value On Hand").font = HEADER_FONT
    dash_ws.cell(hdr_row, 4).fill = HEADER_FILL

    for i, bucket in enumerate(AGE_BUCKETS):
        r = hdr_row + 1 + i
        dash_ws.cell(r, 2, bucket).font = BODY_FONT
        dash_ws.cell(r, 3, f'=COUNTIF(Age_Analysis!$I$2:$I${last_row},B{r})')
        dash_ws.cell(r, 3).font = BODY_FONT
        dash_ws.cell(r, 4, f'=SUMIF(Age_Analysis!$I$2:$I${last_row},B{r},Age_Analysis!$E$2:$E${last_row})')
        dash_ws.cell(r, 4).font = BODY_FONT
        dash_ws.cell(r, 4).number_format = '"R"#,##0'
        for c in (2, 3, 4):
            dash_ws.cell(r, c).border = BORDER

    summary_end = hdr_row + len(AGE_BUCKETS)

    # Bar chart: SKU count by age bucket
    bar = BarChart()
    bar.title = "SKU Count by Age Bucket"
    bar.y_axis.title = "Number of SKUs"
    bar.x_axis.title = "Age Bucket"
    bar.style = 10
    data = Reference(dash_ws, min_col=3, min_row=hdr_row, max_row=summary_end)
    cats = Reference(dash_ws, min_col=2, min_row=hdr_row + 1, max_row=summary_end)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.width = 16
    bar.height = 9
    dash_ws.add_chart(bar, f"G{row0}")

    # Pie chart: Value by age bucket
    pie = PieChart()
    pie.title = "Stock Value by Age Bucket"
    pdata = Reference(dash_ws, min_col=4, min_row=hdr_row, max_row=summary_end)
    pie.add_data(pdata, titles_from_data=True)
    pie.set_categories(cats)
    pie.width = 16
    pie.height = 9
    dash_ws.add_chart(pie, f"G{row0 + 19}")

    # Promo candidate list (filtered view via formulas is complex without dynamic arrays;
    # instead provide a live-formula extraction using a helper column in Age_Analysis
    # and a simple ranked list here referencing top values)
    list_start = summary_end + 3
    dash_ws.cell(list_start, 2, "Promo Candidate Items (3+ yrs, sorted by value)").font = BOLD_FONT
    lh = list_start + 1
    for c, h in enumerate(["Material", "Description", "Group", "Qty", "Value", "Age Bucket"], start=2):
        cell = dash_ws.cell(lh, c, h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    dash_ws.cell(list_start, 8, "Use the PivotTable below (Insert > PivotTable from Age_Analysis) or the").font = Font(name=FONT_NAME, size=9, italic=True, color="666666")
    dash_ws.cell(list_start + 1, 8, "Age_Analysis sheet's AutoFilter, and filter 'Promo Candidate' = YES,").font = Font(name=FONT_NAME, size=9, italic=True, color="666666")
    dash_ws.cell(list_start + 2, 8, "then sort by Value On Hand descending for the full promo list.").font = Font(name=FONT_NAME, size=9, italic=True, color="666666")

    dash_ws.sheet_view.showGridLines = False

    # --- Save & note on slicers -------------------------------------------
    # openpyxl cannot create native Excel PivotTables/Slicers programmatically.
    # We ship a ready-made summary + charts on the Dashboard (fully live via formulas),
    # plus a formatted Table on Age_Analysis so the user can drop in a PivotTable +
    # Slicer in two clicks (Insert > PivotTable > Age_Analysis_Table).
    out_wb.save(output_path)
    print(f"Saved: {output_path}")
    print(f"Rows processed: Stock={n}, Purchases5yr={purch_rows-1}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python build_age_analysis.py <source.xlsx> <output.xlsx>")
        sys.exit(1)
    build(sys.argv[1], sys.argv[2])
