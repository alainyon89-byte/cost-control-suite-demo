"""
Generator for the year-structured Top & Bottom Sellers workbook.
Builds: 12x Input <Mon>, 12x Food <Mon>, 12x Bev Gross <Mon>, 12x Bev Vol <Mon>, Dashboard.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(name=FONT, bold=True, size=14)
SUB_FONT = Font(name=FONT, size=10, italic=True, color="555555")
INSTR_FILL = PatternFill("solid", fgColor="FFFF00")
INSTR_FONT = Font(name=FONT, bold=True, size=10)
NORMAL_FONT = Font(name=FONT, size=10)
BOLD = Font(name=FONT, bold=True, size=10)
SECTION_FONT = Font(name=FONT, bold=True, size=11)
TOP_FILL = PatternFill("solid", fgColor="E2EFDA")
BOT_FILL = PatternFill("solid", fgColor="FCE4E4")
TOTAL_FILL = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

COLS = ["Item","Family Group","Major Group","Gross Sales","Item Discounts",
        "Sales Less Item Disc","% Sales","Qty Sold","% Qty Sold","Average Price"]
NUM_FMT = {
    "Gross Sales": '#,##0.00',
    "Item Discounts": '#,##0.00;(#,##0.00);-',
    "Sales Less Item Disc": '#,##0.00',
    "% Sales": '0.00%',
    "Qty Sold": '#,##0',
    "% Qty Sold": '0.00%',
    "Average Price": '#,##0.00',
}

INPUT_FIRST = 8
INPUT_LAST = 1000

def L(c): return get_column_letter(c)

ITEM_C, FAMILY_C, MAJOR_C, GROSS_C, DISC_C, NET_C, PCTS_C, QTY_C, PCTQ_C, AVG_C = range(1,11)
HELP_START = 13
c_foodflag, c_bevflag, c_foodgrossrank, c_bevgrossrank, c_bevqtyrank = range(HELP_START, HELP_START+5)
c_foodflag_nz, c_bevflag_gross_nz, c_bevflag_qty_nz = HELP_START+5, HELP_START+6, HELP_START+7
c_foodgrossrank_nz, c_bevgrossrank_nz, c_bevqtyrank_nz = HELP_START+8, HELP_START+9, HELP_START+10


def input_sheet_name(mon):
    return f"Input {mon}"


def build_input_tab(wb, mon, is_first):
    ws = wb.create_sheet(input_sheet_name(mon))
    ws.sheet_view.showGridLines = False

    ws["A1"] = f"Sales Mix Items Summary — {mon} — PASTE MICROS EXPORT HERE"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:J1")

    instr = [
        "INSTRUCTIONS:  (Also fill in Business Dates + Location in cells L2 and L3, on the right →)",
        f"1. Open the raw Micros 'Sales Mix Items Summary' export for {mon}.",
        "2. Copy the ENTIRE data block, including the header row (Item, Family Group, Major Group, Gross Sales...).",
        "3. Paste it starting at cell A7 below. This tab is dedicated to this month only — do not reuse it for another month.",
        "4. Do not change the header text or column order — the Food / Bev Gross / Bev Vol tabs for this month read this sheet by column position.",
    ]
    for i, line in enumerate(instr, start=2):
        ws.cell(i,1,line)
        ws.cell(i,1).font = INSTR_FONT if i==2 else NORMAL_FONT
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    for i in range(2,7):
        for c in range(1,11):
            ws.cell(i,c).fill = INSTR_FILL

    for c, h in enumerate(COLS, start=1):
        cell = ws.cell(7, c, h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = BORDER

    for r in range(INPUT_FIRST, INPUT_LAST+1):
        for c, h in enumerate(COLS, start=1):
            cell = ws.cell(r,c)
            cell.font = NORMAL_FONT
            if h in NUM_FMT:
                cell.number_format = NUM_FMT[h]

    widths = [34,22,20,13,13,16,10,10,10,13]
    for i,w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A8"

    # metadata cells
    ws["K2"] = "Business Dates:"; ws["K2"].font = INSTR_FONT; ws["K2"].fill = INSTR_FILL
    ws["L2"].font = NORMAL_FONT; ws["L2"].fill = INSTR_FILL
    ws["K3"] = "Location:"; ws["K3"].font = INSTR_FONT; ws["K3"].fill = INSTR_FILL
    ws["L3"].font = NORMAL_FONT; ws["L3"].fill = INSTR_FILL
    ws.column_dimensions["K"].width = 16
    ws.column_dimensions["L"].width = 20

    # helper header labels
    helper_headers = ["FoodFlag","BevFlag","FoodGrossRank","BevGrossRank","BevQtyRank",
                       "FoodFlagNZ","BevFlagGrossNZ","BevFlagQtyNZ",
                       "FoodGrossRankNZ","BevGrossRankNZ","BevQtyRankNZ"]
    for i,h in enumerate(helper_headers):
        cell = ws.cell(7, HELP_START+i, h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    GROSS_RNG = f"${L(GROSS_C)}${INPUT_FIRST}:${L(GROSS_C)}${INPUT_LAST}"
    QTY_RNG = f"${L(QTY_C)}${INPUT_FIRST}:${L(QTY_C)}${INPUT_LAST}"

    for r in range(INPUT_FIRST, INPUT_LAST+1):
        item_cell = f"${L(ITEM_C)}{r}"
        major_cell = f"${L(MAJOR_C)}{r}"
        gross_cell = f"${L(GROSS_C)}{r}"
        qty_cell = f"${L(QTY_C)}{r}"

        ws.cell(r, c_foodflag, f'=IF(AND({major_cell}<>"",{major_cell}="Food"),1,0)')
        ws.cell(r, c_bevflag, f'=IF(AND({major_cell}<>"",OR({major_cell}="Beverage Enterprise",{major_cell}="Beverage Property")),1,0)')

        ws.cell(r, c_foodgrossrank,
            f'=IF(${L(c_foodflag)}{r}=0,"",'
            f'SUMPRODUCT((${L(c_foodflag)}$8:${L(c_foodflag)}${INPUT_LAST}=1)*'
            f'(({GROSS_RNG}>{gross_cell})+(({GROSS_RNG}={gross_cell})*(ROW({GROSS_RNG})<ROW({gross_cell})))))+1)')
        ws.cell(r, c_bevgrossrank,
            f'=IF(${L(c_bevflag)}{r}=0,"",'
            f'SUMPRODUCT((${L(c_bevflag)}$8:${L(c_bevflag)}${INPUT_LAST}=1)*'
            f'(({GROSS_RNG}>{gross_cell})+(({GROSS_RNG}={gross_cell})*(ROW({GROSS_RNG})<ROW({gross_cell})))))+1)')
        ws.cell(r, c_bevqtyrank,
            f'=IF(${L(c_bevflag)}{r}=0,"",'
            f'SUMPRODUCT((${L(c_bevflag)}$8:${L(c_bevflag)}${INPUT_LAST}=1)*'
            f'(({QTY_RNG}>{qty_cell})+(({QTY_RNG}={qty_cell})*(ROW({QTY_RNG})<ROW({qty_cell})))))+1)')

        foodflag_cell = f"${L(c_foodflag)}{r}"
        bevflag_cell = f"${L(c_bevflag)}{r}"
        ws.cell(r, c_foodflag_nz, f'=IF(AND({foodflag_cell}=1,{gross_cell}>0),1,0)')
        ws.cell(r, c_bevflag_gross_nz, f'=IF(AND({bevflag_cell}=1,{gross_cell}>0),1,0)')
        ws.cell(r, c_bevflag_qty_nz, f'=IF(AND({bevflag_cell}=1,{qty_cell}>0),1,0)')

        fnz_cell = f"${L(c_foodflag_nz)}{r}"
        bgnz_cell = f"${L(c_bevflag_gross_nz)}{r}"
        bqnz_cell = f"${L(c_bevflag_qty_nz)}{r}"

        ws.cell(r, c_foodgrossrank_nz,
            f'=IF({fnz_cell}=0,"",'
            f'SUMPRODUCT((${L(c_foodflag_nz)}$8:${L(c_foodflag_nz)}${INPUT_LAST}=1)*'
            f'(({GROSS_RNG}<{gross_cell})+(({GROSS_RNG}={gross_cell})*(ROW({GROSS_RNG})<ROW({gross_cell})))))+1)')
        ws.cell(r, c_bevgrossrank_nz,
            f'=IF({bgnz_cell}=0,"",'
            f'SUMPRODUCT((${L(c_bevflag_gross_nz)}$8:${L(c_bevflag_gross_nz)}${INPUT_LAST}=1)*'
            f'(({GROSS_RNG}<{gross_cell})+(({GROSS_RNG}={gross_cell})*(ROW({GROSS_RNG})<ROW({gross_cell})))))+1)')
        ws.cell(r, c_bevqtyrank_nz,
            f'=IF({bqnz_cell}=0,"",'
            f'SUMPRODUCT((${L(c_bevflag_qty_nz)}$8:${L(c_bevflag_qty_nz)}${INPUT_LAST}=1)*'
            f'(({QTY_RNG}<{qty_cell})+(({QTY_RNG}={qty_cell})*(ROW({QTY_RNG})<ROW({qty_cell})))))+1)')

    for i in range(11):
        ws.column_dimensions[L(HELP_START+i)].hidden = True

    return ws


def build_output_tab(wb, sheet_name, mon, kind):
    """kind: 'Food', 'Bev Gross', 'Bev Vol' """
    input_name = input_sheet_name(mon)
    IQ = f"'{input_name}'!"

    def rng(col_idx):
        return f"{IQ}${L(col_idx)}${INPUT_FIRST}:${L(col_idx)}${INPUT_LAST}"

    ITEM_RNG = rng(ITEM_C); FAMILY_RNG = rng(FAMILY_C); MAJOR_RNG = rng(MAJOR_C)
    GROSS_RNG = rng(GROSS_C); DISC_RNG = rng(DISC_C); NET_RNG = rng(NET_C)
    PCTS_RNG = rng(PCTS_C); QTY_RNG = rng(QTY_C); PCTQ_RNG = rng(PCTQ_C); AVG_RNG = rng(AVG_C)
    DATA_RNGS = [ITEM_RNG, FAMILY_RNG, MAJOR_RNG, GROSS_RNG, DISC_RNG, NET_RNG, PCTS_RNG, QTY_RNG, PCTQ_RNG, AVG_RNG]

    if kind == "Food":
        top_flag_rng = rng(c_foodflag); top_rank_rng = rng(c_foodgrossrank)
        bottom_flag_nz_rng = rng(c_foodflag_nz); bottom_rank_nz_rng = rng(c_foodgrossrank_nz)
    elif kind == "Bev Gross":
        top_flag_rng = rng(c_bevflag); top_rank_rng = rng(c_bevgrossrank)
        bottom_flag_nz_rng = rng(c_bevflag_gross_nz); bottom_rank_nz_rng = rng(c_bevgrossrank_nz)
    elif kind == "Bev Vol":
        top_flag_rng = rng(c_bevflag); top_rank_rng = rng(c_bevqtyrank)
        bottom_flag_nz_rng = rng(c_bevflag_qty_nz); bottom_rank_nz_rng = rng(c_bevqtyrank_nz)
    else:
        raise ValueError(kind)

    top_n = bottom_n = 20
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    ws["A1"] = f"{kind} {mon} — Top {top_n} & Bottom {bottom_n} Sellers"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:K1")
    ws["A2"] = f'="Business Dates: "&\'{input_name}\'!L2&"   |   Location: "&\'{input_name}\'!L3'
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:K2")

    widths = [6,34,22,20,13,13,16,10,10,10,13]
    for i,w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 4

    def write_header_row(r):
        headers = ["Rank"] + COLS
        for c,h in enumerate(headers, start=1):
            cell = ws.cell(r,c,h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = BORDER

    ws.cell(row,1,f"TOP {top_n} SELLERS").font = SECTION_FONT
    ws.cell(row,1).fill = TOP_FILL
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=11)
    row += 1
    write_header_row(row)
    row += 1
    for i in range(top_n):
        r = row
        rank_val = i+1
        ws.cell(r,1, rank_val).alignment = Alignment(horizontal="center")
        ws.cell(r,1).font = NORMAL_FONT
        ws.cell(r,1).border = BORDER
        match_formula = f'MATCH({rank_val},{top_rank_rng},0)'
        for ci, data_rng in enumerate(DATA_RNGS, start=2):
            colname = COLS[ci-2]
            if colname == "Item":
                formula = f'=IFERROR(IF(INDEX({data_rng},{match_formula})="","(Blank Item Name)",INDEX({data_rng},{match_formula})),"")'
            else:
                formula = f'=IFERROR(INDEX({data_rng},{match_formula}),"")'
            cell = ws.cell(r, ci, formula)
            cell.font = NORMAL_FONT
            cell.border = BORDER
            if colname in NUM_FMT:
                cell.number_format = NUM_FMT[colname]
        row += 1

    row += 1
    ws.cell(row,1,f"BOTTOM {bottom_n} SELLERS (excluding items with zero sales)").font = SECTION_FONT
    ws.cell(row,1).fill = BOT_FILL
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=11)
    row += 1
    write_header_row(row)
    row += 1
    for i in range(bottom_n):
        r = row
        display_rank = i+1
        target_rank_expr = f'{i+1}'
        ws.cell(r,1, display_rank).alignment = Alignment(horizontal="center")
        ws.cell(r,1).font = NORMAL_FONT
        ws.cell(r,1).border = BORDER
        match_formula = f'MATCH({target_rank_expr},{bottom_rank_nz_rng},0)'
        for ci, data_rng in enumerate(DATA_RNGS, start=2):
            colname = COLS[ci-2]
            if colname == "Item":
                formula = f'=IFERROR(IF(INDEX({data_rng},{match_formula})="","(Blank Item Name)",INDEX({data_rng},{match_formula})),"")'
            else:
                formula = f'=IFERROR(INDEX({data_rng},{match_formula}),"")'
            cell = ws.cell(r, ci, formula)
            cell.font = NORMAL_FONT
            cell.border = BORDER
            if colname in NUM_FMT:
                cell.number_format = NUM_FMT[colname]
        row += 1

    row += 1
    ws.cell(row,1,"TOTAL (full month, all matching items):").font = BOLD
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    for c in range(1,4):
        ws.cell(row,c).fill = TOTAL_FILL
    total_formulas = {
        4: f'=SUMPRODUCT(({top_flag_rng}=1)*{GROSS_RNG})',
        5: f'=SUMPRODUCT(({top_flag_rng}=1)*{DISC_RNG})',
        6: f'=SUMPRODUCT(({top_flag_rng}=1)*{NET_RNG})',
        7: f'=SUMPRODUCT(({top_flag_rng}=1)*{PCTS_RNG})',
        8: f'=SUMPRODUCT(({top_flag_rng}=1)*{QTY_RNG})',
        9: f'=SUMPRODUCT(({top_flag_rng}=1)*{PCTQ_RNG})',
        10:f'=IFERROR(SUMPRODUCT(({top_flag_rng}=1)*{GROSS_RNG})/SUMPRODUCT(({top_flag_rng}=1)*{QTY_RNG}),"")',
    }
    for c,f in total_formulas.items():
        cell = ws.cell(row,c,f)
        cell.font = BOLD
        cell.fill = TOTAL_FILL
        colname = COLS[c-1]
        if colname in NUM_FMT:
            cell.number_format = NUM_FMT[colname]
    ws.freeze_panes = "A5"
    return ws



def build_dashboard(wb, food_seed_items=None, bev_seed_items=None, master_cap=600):
    """
    food_seed_items / bev_seed_items: lists of distinct item name strings to seed the
    master lists with (plain values, written once by this generator). Rankings and YTD
    totals against these items are live formulas that update automatically whenever you
    paste new month data. The ITEM LIST itself is fixed until you re-run this generator
    with an updated seed list (e.g. when new menu items appear) -- this tradeoff exists
    because a fully formula-driven "auto-detect brand new items across 12 months" design
    was tested and found to silently overload LibreOffice's recalculation at this row
    count (see project notes). Re-seeding is a deliberate, cheap, low-frequency operation.
    """
    food_seed_items = food_seed_items or []
    bev_seed_items = bev_seed_items or []

    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False

    ws["A1"] = "Executive Dashboard — Year-to-Date Top Sellers"
    ws["A1"].font = Font(name=FONT, bold=True, size=16)
    ws.merge_cells("A1:F1")
    ws["A2"] = "Live view aggregated across all populated Input tabs for the year."
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:F2")

    ws["A4"] = "YEAR-TO-DATE SUMMARY"
    ws["A4"].font = SECTION_FONT
    ws["A4"].fill = PatternFill("solid", fgColor="D9E1F2")
    ws.merge_cells("A4:F4")

    def sum_expr(flag_col, val_col):
        terms = []
        for m in MONTHS:
            iq = f"'Input {m}'!"
            flag_rng = f"{iq}${L(flag_col)}${INPUT_FIRST}:${L(flag_col)}${INPUT_LAST}"
            val_rng = f"{iq}${L(val_col)}${INPUT_FIRST}:${L(val_col)}${INPUT_LAST}"
            terms.append(f"SUMPRODUCT(({flag_rng}=1)*{val_rng})")
        return "+".join(terms)

    labels = ["Food Gross Sales (YTD)", "Food Qty Sold (YTD)",
              "Beverage Gross Sales (YTD)", "Beverage Qty Sold (YTD)"]
    formulas = [
        f"={sum_expr(c_foodflag, GROSS_C)}",
        f"={sum_expr(c_foodflag, QTY_C)}",
        f"={sum_expr(c_bevflag, GROSS_C)}",
        f"={sum_expr(c_bevflag, QTY_C)}",
    ]
    r = 5
    for label, formula in zip(labels, formulas):
        ws.cell(r,1,label).font = BOLD
        cell = ws.cell(r,2,formula)
        cell.font = Font(name=FONT, size=12, bold=True, color="1F4E78")
        cell.number_format = '#,##0.00' if 'Gross' in label else '#,##0'
        r += 1

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 18
    for c in "CDEFGHIJK":
        ws.column_dimensions[c].width = 14

    # ---- Master item lists (fixed values, written once; re-run generator to add new items) ----
    MASTER_START_COL = 27  # AA
    FOOD_MASTER_COL = MASTER_START_COL        # AA
    FOOD_GROSS_COL = MASTER_START_COL+1       # AB
    FOOD_QTY_COL = MASTER_START_COL+2         # AC
    FOOD_RANK_COL = MASTER_START_COL+3        # AD

    BEV_MASTER_COL = MASTER_START_COL+5       # AF
    BEV_GROSS_COL = MASTER_START_COL+6        # AG
    BEV_QTY_COL = MASTER_START_COL+7          # AH
    BEV_RANK_COL = MASTER_START_COL+8         # AI

    for col, label in [(FOOD_MASTER_COL,"FoodMasterItem"), (FOOD_GROSS_COL,"FoodYTDGross"),
                        (FOOD_QTY_COL,"FoodYTDQty"), (FOOD_RANK_COL,"FoodRank"),
                        (BEV_MASTER_COL,"BevMasterItem"), (BEV_GROSS_COL,"BevYTDGross"),
                        (BEV_QTY_COL,"BevYTDQty"), (BEV_RANK_COL,"BevRank")]:
        cell = ws.cell(7, col, label)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    for col in [FOOD_MASTER_COL, FOOD_GROSS_COL, FOOD_QTY_COL, FOOD_RANK_COL,
                BEV_MASTER_COL, BEV_GROSS_COL, BEV_QTY_COL, BEV_RANK_COL]:
        ws.column_dimensions[L(col)].hidden = True

    MASTER_FIRST_ROW = 8

    def sumif_across_months(item_cell, flag_col, val_col):
        terms = []
        for m in MONTHS:
            iq = f"'Input {m}'!"
            item_rng = f"{iq}${L(ITEM_C)}${INPUT_FIRST}:${L(ITEM_C)}${INPUT_LAST}"
            flag_rng = f"{iq}${L(flag_col)}${INPUT_FIRST}:${L(flag_col)}${INPUT_LAST}"
            val_rng = f"{iq}${L(val_col)}${INPUT_FIRST}:${L(val_col)}${INPUT_LAST}"
            terms.append(f"SUMPRODUCT(({item_rng}={item_cell})*({flag_rng}=1)*{val_rng})")
        return "+".join(terms)

    def write_master_block(seed_items, master_col, gross_col, qty_col, rank_col, flag_col):
        n = len(seed_items)
        last_row = MASTER_FIRST_ROW + n - 1
        for i, name in enumerate(seed_items):
            r = MASTER_FIRST_ROW + i
            ws.cell(r, master_col, name)
            item_cell = f"${L(master_col)}{r}"
            ws.cell(r, gross_col, f"={sumif_across_months(item_cell, flag_col, GROSS_C)}")
            ws.cell(r, qty_col, f"={sumif_across_months(item_cell, flag_col, QTY_C)}")
        gross_rng = f"${L(gross_col)}${MASTER_FIRST_ROW}:${L(gross_col)}{last_row}"
        for i in range(n):
            r = MASTER_FIRST_ROW + i
            gross_cell = f"${L(gross_col)}{r}"
            # tie-broken descending rank within this master list only (n x n cost, cheap at n<=600)
            ws.cell(r, rank_col,
                f'=SUMPRODUCT(({gross_rng}>{gross_cell})+(({gross_rng}={gross_cell})*(ROW({gross_rng})<ROW({gross_cell}))))+1')
        return last_row

    food_last_row = write_master_block(food_seed_items, FOOD_MASTER_COL, FOOD_GROSS_COL, FOOD_QTY_COL, FOOD_RANK_COL, c_foodflag)
    bev_last_row = write_master_block(bev_seed_items, BEV_MASTER_COL, BEV_GROSS_COL, BEV_QTY_COL, BEV_RANK_COL, c_bevflag)

    FOOD_RANK_RNG = f"${L(FOOD_RANK_COL)}${MASTER_FIRST_ROW}:${L(FOOD_RANK_COL)}{food_last_row}"
    FOOD_MASTER_RNG = f"${L(FOOD_MASTER_COL)}${MASTER_FIRST_ROW}:${L(FOOD_MASTER_COL)}{food_last_row}"
    FOOD_GROSS_RNG = f"${L(FOOD_GROSS_COL)}${MASTER_FIRST_ROW}:${L(FOOD_GROSS_COL)}{food_last_row}"
    FOOD_QTY_RNG = f"${L(FOOD_QTY_COL)}${MASTER_FIRST_ROW}:${L(FOOD_QTY_COL)}{food_last_row}"

    BEV_RANK_RNG = f"${L(BEV_RANK_COL)}${MASTER_FIRST_ROW}:${L(BEV_RANK_COL)}{bev_last_row}"
    BEV_MASTER_RNG = f"${L(BEV_MASTER_COL)}${MASTER_FIRST_ROW}:${L(BEV_MASTER_COL)}{bev_last_row}"
    BEV_GROSS_RNG = f"${L(BEV_GROSS_COL)}${MASTER_FIRST_ROW}:${L(BEV_GROSS_COL)}{bev_last_row}"
    BEV_QTY_RNG = f"${L(BEV_QTY_COL)}${MASTER_FIRST_ROW}:${L(BEV_QTY_COL)}{bev_last_row}"

    def write_top20_table(start_row, title, rank_rng, item_rng, gross_rng, qty_rng, fill):
        r = start_row
        ws.cell(r,1,title).font = SECTION_FONT
        ws.cell(r,1).fill = fill
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        r += 1
        headers = ["Rank","Item","YTD Gross Sales","YTD Qty Sold"]
        for c,h in enumerate(headers, start=1):
            cell = ws.cell(r,c,h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.border = BORDER
        r += 1
        for i in range(20):
            rank_val = i+1
            match_formula = f'MATCH({rank_val},{rank_rng},0)'
            ws.cell(r,1,rank_val).font = NORMAL_FONT
            ws.cell(r,1).alignment = Alignment(horizontal="center")
            ws.cell(r,1).border = BORDER
            item_formula = f'=IFERROR(TRIM(INDEX({item_rng},{match_formula})),"")'
            gross_formula = f'=IFERROR(INDEX({gross_rng},{match_formula}),"")'
            qty_formula = f'=IFERROR(INDEX({qty_rng},{match_formula}),"")'
            c_item = ws.cell(r,2,item_formula); c_item.font = NORMAL_FONT; c_item.border = BORDER
            c_gross = ws.cell(r,3,gross_formula); c_gross.font = NORMAL_FONT; c_gross.border = BORDER
            c_gross.number_format = '#,##0.00'
            c_qty = ws.cell(r,4,qty_formula); c_qty.font = NORMAL_FONT; c_qty.border = BORDER
            c_qty.number_format = '#,##0'
            r += 1
        return r + 1

    next_row = write_top20_table(11, "TOP 20 FOOD ITEMS (Year-to-Date, by Gross Sales)",
                                  FOOD_RANK_RNG, FOOD_MASTER_RNG, FOOD_GROSS_RNG, FOOD_QTY_RNG, TOP_FILL)
    write_top20_table(next_row, "TOP 20 BEVERAGE ITEMS (Year-to-Date, by Gross Sales)",
                       BEV_RANK_RNG, BEV_MASTER_RNG, BEV_GROSS_RNG, BEV_QTY_RNG, TOP_FILL)

    note_row = next_row + 24
    ws.cell(note_row, 1,
        "Note: the item list above is fixed (captures items known at the time this workbook "
        "was last generated/refreshed). Gross Sales, Qty Sold, and Rank update live as you "
        "paste new months. If a genuinely new menu item appears, ask for the workbook to be "
        "regenerated to add it to this list.")
    ws.cell(note_row, 1).font = Font(name=FONT, size=9, italic=True, color="777777")
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=6)

    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 14
    ws.freeze_panes = "A5"

    return ws


def main(food_seed_items=None, bev_seed_items=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for i, mon in enumerate(MONTHS):
        build_input_tab(wb, mon, is_first=(i==0))

    for mon in MONTHS:
        build_output_tab(wb, f"Food {mon}", mon, "Food")
        build_output_tab(wb, f"Bev Gross {mon}", mon, "Bev Gross")
        build_output_tab(wb, f"Bev Vol {mon}", mon, "Bev Vol")

    build_dashboard(wb, food_seed_items=food_seed_items, bev_seed_items=bev_seed_items)
    order = ["Dashboard"]
    for m in MONTHS:
        order.append(f"Input {m}")
    for m in MONTHS:
        order.append(f"Food {m}")
        order.append(f"Bev Gross {m}")
        order.append(f"Bev Vol {m}")
    wb._sheets = [wb[name] for name in order]

    wb.save("Top_Bottom_Sellers_Year.xlsx")
    print("Full year workbook built:", len(wb.sheetnames), "sheets")


if __name__ == "__main__":
    import json, os
    seed_path = "seed_items.json"
    food_seed, bev_seed = None, None
    if os.path.exists(seed_path):
        seed = json.load(open(seed_path))
        food_seed = seed.get("food")
        bev_seed = seed.get("bev")
        print(f"Loaded seed list: {len(food_seed or [])} food items, {len(bev_seed or [])} bev items")
    main(food_seed_items=food_seed, bev_seed_items=bev_seed)
