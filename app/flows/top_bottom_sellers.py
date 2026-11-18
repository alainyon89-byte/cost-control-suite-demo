# (c) 2026 Alain Yon
"""Top & Bottom Sellers: 'Build a new year workbook'.

There is no .bat for this tool. generate_year_workbook.py is run from its own
folder (it reads seed_items.json and writes Top_Bottom_Sellers_Year.xlsx there),
and the workbook is then given the year in its name, as Alain does by hand
today (Top_Bottom_Sellers_2026.xlsx). A year workbook that already exists is
never overwritten: it holds the pasted months.
"""
import os


def run(ctx):
    year = str(ctx.options.get("year", "")).strip()
    rc = ctx.run_script()
    if rc != 0:
        return {"status": "failed", "exit_code": rc}
    built = os.path.join(ctx.tool_dir, "Top_Bottom_Sellers_Year.xlsx")
    target = os.path.join(ctx.tool_dir, f"Top_Bottom_Sellers_{year}.xlsx")
    if not os.path.exists(built):
        ctx.log("The script finished but Top_Bottom_Sellers_Year.xlsx was not found.", "error")
        return {"status": "failed", "exit_code": rc}
    if os.path.exists(target):
        ctx.log(f"Top_Bottom_Sellers_{year}.xlsx already exists and was left alone (it holds your pasted months).")
        ctx.log("The new empty workbook is saved as Top_Bottom_Sellers_Year.xlsx in the tool folder.")
        return {"status": "finished", "exit_code": 0,
                "headline": f"Finished. Top_Bottom_Sellers_{year}.xlsx already exists and was left alone"}
    os.replace(built, target)
    ctx.log(f"Saved as Top_Bottom_Sellers_{year}.xlsx")
    return {"status": "finished", "exit_code": 0}
