# (c) 2026 Alain Yon
"""Regression cases.  One case = one sample input set for one tool (demo data from sample_data/).

Each run happens in a throwaway root laid out like the installed data folder:
    <root>/tools/<folder>/   the tool's scripts (source_tools/<folder>) + its tool_files
    <root>/inbox/            the files the user would drop on the tool
    <root>/sap_data/         shared MM60 / kitchen MI20 (only for the shared cases)

place     (file under sample_data/<section>/<folder>/inputs, destination relative to <root>)
a_steps   the original flow, exactly as the .bat runs Python (folder A)
          args may use {root}; cwd is the tool folder unless "cwd" is given
b         the same job asked of the app (folder B): inputs, options, answers
outputs   golden name -> glob relative to <root>
data      data files (relative to <root>) that must end up the same in A and B
"""
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def _sections():
    with open(os.path.join(ROOT, "app", "tools.json"), encoding="utf-8") as f:
        return {t.get("folder", t["id"]): t["group"] for t in json.load(f)["tools"]}


SECTIONS = _sections()


def sample_dir(folder):
    """sample_data/<section>/<folder>: tool_files/, inputs/, answers_*.txt and expected/ live here."""
    return os.path.join(ROOT, "sample_data", SECTIONS[folder], folder)


T = "tools/"
I = "{root}/inbox/"
CC_SALES = "CC_Bar_Sales_Oct_2026.xlsx"
CELLAR = "Cellar_Consumption_10_November_2026.xlsx"
CC_MICROS = "CC_Function_Sales_10_Nov_2026.xlsx"
VAR, BAR, CT = "L123_Variance_Oct_2026.xlsx", "Pool_Bar_Sales_Oct_2026.xlsx", "Cocktail_Tracker.xlsx"
MI20 = "MI20_Kitchen_Oct_2026.xlsx"
COFFEE_QTY = ["446", "12", "30", "500", "200", "15", "10", "5", "40", "25"]

CASES = [
    dict(id="theo_cos", tool="theo_cos", folder="theo_cos",
         place=[(CC_SALES, "inbox/" + CC_SALES)],
         a_steps=[dict(script="Bev_Theoretical_CoS_Generator.py", args=[I + CC_SALES], tty=True, answers=[])],
         b=dict(inputs={"file": ["inbox/" + CC_SALES]}, answers=[]),
         outputs={"CoS_Summary.xlsx": "inbox/CC_Bar_Sales_Oct_2026_CoS_Summary.xlsx"},
         data=[T + "theo_cos/beverage_item_map.xlsx"]),

    # an older item map: six POS items are new, answered on screen and added to the map
    dict(id="theo_cos_new_items", tool="theo_cos", folder="theo_cos",
         place=[(CC_SALES, "inbox/" + CC_SALES), ("beverage_item_map_OLD.xlsx", T + "theo_cos/beverage_item_map.xlsx")],
         a_steps=[dict(script="Bev_Theoretical_CoS_Generator.py", args=[I + CC_SALES], tty=True, answers="@answers")],
         b=dict(inputs={"file": ["inbox/" + CC_SALES]}, answers="@answers"),
         outputs={"CoS_Summary.xlsx": "inbox/CC_Bar_Sales_Oct_2026_CoS_Summary.xlsx"},
         data=[T + "theo_cos/beverage_item_map.xlsx", T + "theo_cos/beverage_item_map_backup_*.xlsx"]),

    # Cellar: the sheet is picked from the report's Business Date; the item map knows every renamed item
    dict(id="cellar_consumption", tool="cellar_consumption", folder="cellar_consumption",
         place=[(CC_MICROS, "inbox/" + CC_MICROS), (CELLAR, "inbox/" + CELLAR)],
         a_steps=[dict(script="update_cellar_consumption.py", tty=True, answers=[],
                       args=["--files", I + CELLAR, I + CC_MICROS])],
         b=dict(inputs={"files": ["inbox/" + CELLAR, "inbox/" + CC_MICROS]}, answers=[]),
         outputs={"Cellar_UPDATED.xlsx": "inbox/Cellar_Consumption_10_November_2026_UPDATED.xlsx",
                  "Cellar_UPDATE_LOG.txt": "inbox/Cellar_Consumption_10_November_2026_UPDATE_LOG.txt"},
         data=[T + "cellar_consumption/cellar_item_map.xlsx"]),

    # --no-prompt with an older item map (so there ARE unknown names) and the sheet typed
    dict(id="cellar_no_prompt", tool="cellar_consumption", folder="cellar_consumption",
         place=[(CC_MICROS, "inbox/" + CC_MICROS), (CELLAR, "inbox/" + CELLAR),
                ("cellar_item_map_OLD.xlsx", T + "cellar_consumption/cellar_item_map.xlsx")],
         a_steps=[dict(script="update_cellar_consumption.py", tty=True, answers=[],
                       args=["--files", I + CELLAR, I + CC_MICROS, "--sheets", "10.11.2026", "--no-prompt"])],
         b=dict(inputs={"files": ["inbox/" + CELLAR, "inbox/" + CC_MICROS]},
                options={"sheets": "10.11.2026", "no_prompt": True}, answers=[]),
         outputs={"Cellar_UPDATED.xlsx": "inbox/Cellar_Consumption_10_November_2026_UPDATED.xlsx",
                  "Cellar_UPDATE_LOG.txt": "inbox/Cellar_Consumption_10_November_2026_UPDATE_LOG.txt"},
         data=[T + "cellar_consumption/cellar_item_map.xlsx"]),

    # the older item map with the questions answered (one answer rejected and asked again)
    dict(id="cellar_prompts", tool="cellar_consumption", folder="cellar_consumption",
         place=[(CC_MICROS, "inbox/" + CC_MICROS), (CELLAR, "inbox/" + CELLAR),
                ("cellar_item_map_OLD.xlsx", T + "cellar_consumption/cellar_item_map.xlsx")],
         a_steps=[dict(script="update_cellar_consumption.py", tty=True, answers="@answers",
                       args=["--files", I + CELLAR, I + CC_MICROS])],
         b=dict(inputs={"files": ["inbox/" + CELLAR, "inbox/" + CC_MICROS]}, answers="@answers"),
         outputs={"Cellar_UPDATED.xlsx": "inbox/Cellar_Consumption_10_November_2026_UPDATED.xlsx",
                  "Cellar_UPDATE_LOG.txt": "inbox/Cellar_Consumption_10_November_2026_UPDATE_LOG.txt"},
         data=[T + "cellar_consumption/cellar_item_map.xlsx", T + "cellar_consumption/cellar_item_map_backup_*.xlsx"]),

    dict(id="recount_gate", tool="recount_selection", folder="recount_selection",
         place=[(VAR, "inbox/" + VAR), (BAR, "inbox/" + BAR), (CT, "inbox/" + CT)],
         a_steps=[dict(script="select_recount_items.py", args=[I + VAR, I + BAR, I + CT], rc=2),
                  dict(script="select_recount_items.py",
                       args=[I + VAR, I + BAR, I + CT, "--approve", "--limit", "25", "--approved-by", "Bar Manager"])],
         # .bat menu: Enter = build; gate: approver name, then line count
         b=dict(inputs={"files": ["inbox/" + VAR, "inbox/" + BAR, "inbox/" + CT]},
                options={"mode": "build"}, answers=["Bar Manager", "25"]),
         outputs={"Recount_Selection.xlsx": "inbox/Recount_Selection_*.xlsx"},
         data=[T + "recount_selection/item_aliases.csv"]),

    dict(id="recount_export_review", tool="recount_selection", folder="recount_selection",
         place=[(VAR, "inbox/" + VAR), (BAR, "inbox/" + BAR), (CT, "inbox/" + CT)],
         a_steps=[dict(script="select_recount_items.py", args=[I + VAR, I + BAR, I + CT, "--export-review"])],
         b=dict(inputs={"files": ["inbox/" + VAR, "inbox/" + BAR, "inbox/" + CT]},
                options={"mode": "review"}, answers=[]),
         outputs={"Sales_Item_Review.xlsx": "inbox/Sales_Item_Review_*.xlsx"},
         data=[T + "recount_selection/item_aliases.csv"]),

    dict(id="top_bottom_sellers", tool="top_bottom_sellers", folder="top_bottom_sellers",
         place=[],
         a_steps=[dict(script="generate_year_workbook.py", args=[])],
         b=dict(action="build_year", inputs={}, options={"year": "2099"}, answers=[]),
         outputs={"Top_Bottom_Sellers_Year.xlsx": T + "top_bottom_sellers/Top_Bottom_Sellers_Year.xlsx"},
         b_outputs={"Top_Bottom_Sellers_Year.xlsx": T + "top_bottom_sellers/Top_Bottom_Sellers_2099.xlsx"}),

    dict(id="food_sales_mix_main", tool="food_sales_mix", folder="food_sales_mix",
         place=[("Main_Hotel_Food_Oct_2026.xlsx", T + "food_sales_mix/Main_Hotel_Food_Oct_2026.xlsx")],
         a_steps=[dict(script="update_sales_mix.py", args=["Food_Sales_Mix_Main.xlsx", "Main_Hotel_Food_Oct_2026.xlsx"])],
         b=dict(inputs={"file": [T + "food_sales_mix/Main_Hotel_Food_Oct_2026.xlsx"]}, options={"outlet": "main"},
                answers=[]),
         outputs={"Food_Sales_Mix_Main.xlsx": T + "food_sales_mix/Food_Sales_Mix_Main.xlsx",
                  "update_report.txt": T + "food_sales_mix/update_report_*.txt"},
         data=[T + "food_sales_mix/item_additions.json", T + "food_sales_mix/item_aliases.json",
               T + "food_sales_mix/backups/Food_Sales_Mix_Main_backup_*.xlsx"]),

    dict(id="food_sales_mix_tower", tool="food_sales_mix", folder="food_sales_mix",
         place=[("Tower_Wing_Food_Oct_2026.xlsx", T + "food_sales_mix/Tower_Wing_Food_Oct_2026.xlsx")],
         a_steps=[dict(script="update_sales_mix_tower.py",
                       args=["Food_Sales_Mix_Tower.xlsx", "Tower_Wing_Food_Oct_2026.xlsx"])],
         b=dict(inputs={"file": [T + "food_sales_mix/Tower_Wing_Food_Oct_2026.xlsx"]}, options={"outlet": "tower"},
                answers=[]),
         outputs={"Food_Sales_Mix_Tower.xlsx": T + "food_sales_mix/Food_Sales_Mix_Tower.xlsx",
                  "update_report.txt": T + "food_sales_mix/update_report_*.txt"},
         data=[T + "food_sales_mix/item_additions_tower.json",
               T + "food_sales_mix/backups/Food_Sales_Mix_Tower_backup_*.xlsx"]),

    dict(id="cost_comparator", tool="cost_comparator", folder="cost_comparator",
         place=[("MM60_D100_Sep_2026.xlsx", T + "cost_comparator/MM60_D100_Sep_2026.xlsx"),
                ("MM60_D100_Oct_2026.xlsx", T + "cost_comparator/MM60_D100_Oct_2026.xlsx")],
         a_steps=[dict(script="mm60_comparator.py", args=[],
                       answers=["MM60_D100_Sep_2026.xlsx", "MM60_D100_Oct_2026.xlsx", "September 2026", "October 2026", ""])],
         b=dict(inputs={"old": [T + "cost_comparator/MM60_D100_Sep_2026.xlsx"],
                        "new": [T + "cost_comparator/MM60_D100_Oct_2026.xlsx"]},
                options={"label_old": "September 2026", "label_new": "October 2026"}, answers=[]),
         outputs={"MM60_Comparative.xlsx": T + "cost_comparator/MM60_Comparative_September_2026_vs_October_2026.xlsx"}),

    dict(id="stock_age", tool="stock_age", folder="stock_age",
         place=[("Stock_Age_Source.xlsx", T + "stock_age/Stock_Age_Source.xlsx")],
         a_steps=[dict(script="build_age_analysis.py", args=["Stock_Age_Source.xlsx", "Age_Analysis_Dashboard.xlsx"])],
         b=dict(inputs={"file": [T + "stock_age/Stock_Age_Source.xlsx"]}, answers=[]),
         outputs={"Age_Analysis_Dashboard.xlsx": T + "stock_age/Age_Analysis_Dashboard.xlsx"}),

    dict(id="coffee_loss", tool="coffee_loss", folder="coffee_loss",
         place=[(MI20, T + "coffee_loss/Mi20_files/" + MI20)],
         a_steps=[dict(script="coffee_loss_tracker_v2.py", args=[], answers=["October 2026"] + COFFEE_QTY)],
         b=dict(inputs={"mi20": [T + "coffee_loss/Mi20_files/" + MI20]},
                options=dict({"month": "October 2026"}, **{f"qty_{i}": q for i, q in enumerate(COFFEE_QTY)}),
                answers=[]),
         outputs={"Coffee_Loss_Tracker.xlsx": T + "coffee_loss/Coffee_Loss_Tracker_October_2026.xlsx"}),

    dict(id="breakfast_tracker", tool="breakfast_tracker", folder="breakfast_tracker",
         place=[(MI20, T + "breakfast_tracker/Mi20_kitchen.xlsx")],
         a_steps=[dict(script="breakfast_rebuild.py", args=[], answers=["October 2026"])],
         b=dict(inputs={"mi20": [T + "breakfast_tracker/Mi20_kitchen.xlsx"]}, options={"month": "October 2026"},
                answers=[]),
         outputs={"Breakfast_MonthEnd_Tool.xlsx": T + "breakfast_tracker/Breakfast_MonthEnd_Tool.xlsx"}),

    dict(id="canteen_loss", tool="canteen_loss", folder="canteen_loss",
         place=[(MI20, T + "canteen_loss/MI20_kitchen.xlsx")],
         a_steps=[dict(script="canteen_rebuild.py", args=[], answers=["October 2026"])],
         b=dict(inputs={"mi20": [T + "canteen_loss/MI20_kitchen.xlsx"]}, options={"month": "October 2026"},
                answers=[]),
         outputs={"Canteen_MonthEnd_Tool.xlsx": T + "canteen_loss/Canteen_MonthEnd_Tool.xlsx"}),

    dict(id="oe_monthly", tool="oe_monthly", folder="oe_monthly",
         place=[("MB52_OE_Oct_2026.xlsx", "inbox/MB52_OE_Oct_2026.xlsx"),
                ("MM60_D100_Oct_2026.xlsx", "inbox/MM60_D100_Oct_2026.xlsx"),
                ("MB51_Oct_2026.xlsx", "inbox/MB51_Oct_2026.xlsx")],
         a_steps=[dict(script="oe_tracker.py",
                       args=[I + "MB52_OE_Oct_2026.xlsx", I + "MM60_D100_Oct_2026.xlsx", I + "MB51_Oct_2026.xlsx"])],
         b=dict(inputs={"files": ["inbox/MB52_OE_Oct_2026.xlsx", "inbox/MB51_Oct_2026.xlsx"],
                        "mm60": ["inbox/MM60_D100_Oct_2026.xlsx"]}, answers=[]),
         outputs={"OE_Tracking_Workbook.xlsx": T + "oe_monthly/OE_Tracking_Workbook.xlsx",
                  "OE_Approval.xlsx": T + "oe_monthly/OE_Approval_[A-Z][a-z][a-z]_20[0-9][0-9].xlsx",
                  "OE_Approval_Linen_Main.xlsx": T + "oe_monthly/OE_Approval_Linen_Main_???_20??.xlsx",
                  "OE_Approval_Linen_Tower.xlsx": T + "oe_monthly/OE_Approval_Linen_Tower_???_20??.xlsx",
                  "OE_Transfer_Form.xlsx": T + "oe_monthly/OE_Transfer_Form.xlsx"},
         data=[T + "oe_monthly/Archive/OE_Tracking_Workbook_*.xlsx"]),

    dict(id="oe_approvals", tool="oe_approvals", folder="oe_monthly",
         place=[("OE_Approval_Oct_2026.xlsx", "inbox/OE_Approval_Oct_2026.xlsx")],
         a_steps=[dict(script="oe_approvals.py", args=[I + "OE_Approval_Oct_2026.xlsx"])],
         b=dict(inputs={"file": ["inbox/OE_Approval_Oct_2026.xlsx"]}, answers=[]),
         outputs={"OE_Approved_Orders.xlsx": "inbox/OE_Approved_Orders_*.xlsx"}),
]

# Tools with no sample inputs supplied: listed so the report says so instead of staying silent.
NO_SAMPLES = {
}
