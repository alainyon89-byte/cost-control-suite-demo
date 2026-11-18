# (c) 2026 Alain Yon
"""Recount Selection Builder: the flow of Run_Recount_Selection.bat, step for step.

.bat                                              here
-----------------------------------------------  ---------------------------------------------
menu [1] build / [2] export review (Enter = 1)   the 'What do you want to do?' option on the page
[2]: python script <files> --export-review       same, then the same closing instructions
[1]: python script <files>                       same (preview only, nothing is written)
  exit 0  -> done                                 finished
  exit 2  -> APPROVAL REQUIRED: ask the           Question card: approver name (blank = cancel),
            approver's name (blank cancels),      then how many lines (blank = 20), then
            ask line count (blank = 20), rerun    python script <files> --approve --limit N
            with --approve --limit N               --approved-by "NAME"
            --approved-by "NAME"
  anything else -> 'The run did not complete'     failed
"""


def run(ctx):
    files = ctx.expand_args()
    if ctx.options.get("mode") == "review":
        rc = ctx.run_script(args=files + ["--export-review"])
        if rc != 0:
            return {"status": "failed", "exit_code": rc}
        ctx.log("")
        ctx.log("Fill in the CONFIRM Material No column in the sheet above, then drop it on this tool")
        ctx.log("together with the same three files (rename it if you like - it is recognised by its")
        ctx.log("contents, not its name).")
        return {"status": "finished", "exit_code": 0, "headline": "Sales Item Review sheet written"}

    rc = ctx.run_script(args=files)
    if rc == 0:
        return {"status": "finished", "exit_code": 0}
    if rc != 2:
        return {"status": "failed", "exit_code": rc}

    ctx.log("")
    ctx.log("------------------------------------------------------------------")
    ctx.log(" APPROVAL REQUIRED")
    ctx.log(" More lines are unexplained by sales than the standard limit of 20.")
    ctx.log(" Nothing has been written yet.")
    ctx.log("------------------------------------------------------------------")
    approver = ctx.ask("Type the approver's name to sign off",
                       "More lines are unexplained by sales than the standard limit of 20. Nothing has been "
                       "written yet. Leave the box empty and confirm to cancel.", card_id="approver")
    if approver == "":
        ctx.log("Cancelled. No recount sheet was created.")
        return {"status": "cancelled", "exit_code": 2, "headline": "Cancelled. No recount sheet was created."}
    count = ctx.ask("How many lines should be recounted?", "Leave empty for 20.",
                    options=[{"value": "20", "label": "20 lines (standard)"}], card_id="count")
    if count == "":
        count = "20"
    rc = ctx.run_script(args=files + ["--approve", "--limit", count, "--approved-by", approver])
    if rc != 0:
        return {"status": "failed", "exit_code": rc}
    ctx.log("")
    ctx.log("Done. The recount selection sheet is in the same folder as the variance report.")
    return {"status": "finished", "exit_code": 0}
