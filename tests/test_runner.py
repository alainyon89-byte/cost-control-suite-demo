# (c) 2026 Alain Yon
"""Tests the subprocess runner with a fake script that prompts, exits 2 and writes an xlsx.

python test_runner.py           run the tests
(the same file, run with --fake, is the fake tool script)
"""
import os
import sys
import tempfile
import threading
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def fake(argv):
    """The fake tool: behaves like the real scripts' awkward corners."""
    mode = argv[0] if argv else "ask"
    print("Fake tool starting")
    if mode == "tty":
        print("console=" + str(sys.stdin.isatty()))
        return 0
    if mode == "ask":
        name = input("\nEnter month label (e.g., 'June 2026'): ").strip()
        print(f"  Month: {name}")
        print("     [1] First choice  R8.83 BT  #55000")
        print("     [2] Second choice")
        print("     [N] Type a number   [M] Manual cost")
        pick = input("   Which SAP item or recipe? ").strip().upper()
        print("   Saved.", pick)
        accent = input("  Café Latte: ")
        print("unicode ok ÷ ─ →", accent)
        from openpyxl import Workbook
        wb = Workbook()
        wb.active["A1"] = f"{name}|{pick}|{accent}"
        wb.save(os.path.join(os.getcwd(), "fake_output.xlsx"))
        sys.stderr.write("a warning on stderr\n")
        print("Saved: fake_output.xlsx")
        return 0
    if mode == "exit2":
        print("** APPROVAL REQUIRED - NOTHING WRITTEN **")
        return 2
    if mode == "noprompt":
        sys.stdout.write("working")          # partial line, then more output: must NOT stay a question
        sys.stdout.flush()
        time.sleep(0.9)
        print(" ... done")
        return 0
    if mode == "hang":
        input("waiting forever: ")
        return 0
    if mode == "crash":
        raise RuntimeError("boom")
    return 0


def run_tests():
    sys.path.insert(0, ROOT)
    from app.runner import ScriptProcess, build_command
    failures = []

    def check(name, cond, extra=""):
        print(("ok    " if cond else "FAIL  ") + name + (f"  {extra}" if extra and not cond else ""))
        if not cond:
            failures.append(name)

    def launch(mode, answers=None, console=False, cwd=None):
        lines, prompts, cleared = [], [], []
        answers = list(answers or [])
        holder = {}

        def on_line(text, is_answer=False):
            lines.append(text)

        def on_prompt(text):
            prompts.append(text)
            if answers:
                threading.Thread(target=lambda a=answers.pop(0): holder["p"].send(a), daemon=True).start()

        p = ScriptProcess(build_command(os.path.abspath(__file__), ["--fake", mode], console), cwd or HERE,
                          on_line, on_prompt, lambda: cleared.append(1))
        holder["p"] = p
        p.start()
        return p, lines, prompts, cleared

    work = tempfile.mkdtemp(prefix="ccs_runner_")
    p, lines, prompts, cleared = launch("ask", ["June 2026", "n", "ünï"], cwd=work)
    rc = p.wait()
    check("exit code 0", rc == 0, str(rc))
    check("three prompts detected", len(prompts) == 3, repr(prompts))
    check("prompt text arrives without its newline", prompts and prompts[0].startswith("Enter month label"), repr(prompts[:1]))
    check("answers reach stdin as typed", "  Month: June 2026" in lines and any("Saved. N" in ln for ln in lines), repr(lines))
    check("answer is echoed on the prompt line", any(ln.startswith("Enter month label") and ln.endswith("June 2026") for ln in lines))
    check("utf-8 both ways", any("÷ ─ → ünï" in ln for ln in lines), repr(lines[-4:]))
    check("stderr is streamed too", "a warning on stderr" in lines)
    check("xlsx written in the cwd", os.path.exists(os.path.join(work, "fake_output.xlsx")))
    from openpyxl import load_workbook
    check("xlsx holds the answers", load_workbook(os.path.join(work, "fake_output.xlsx")).active["A1"].value == "June 2026|N|ünï")

    p, lines, prompts, cleared = launch("exit2")
    check("exit code 2 is passed through", p.wait() == 2)

    p, lines, prompts, cleared = launch("noprompt")
    rc = p.wait()
    check("a partial line followed by more output is withdrawn as a question", len(prompts) == 1 and len(cleared) == 1,
          f"prompts={prompts} cleared={cleared}")
    check("partial line is joined up in the log", "working ... done" in lines, repr(lines))

    p, lines, prompts, cleared = launch("tty", console=False)
    p.wait()
    check("plain run: stdin is a pipe", "console=False" in lines, repr(lines))
    p, lines, prompts, cleared = launch("tty", console=True)
    p.wait()
    check("console shim: stdin.isatty() is True", "console=True" in lines, repr(lines))

    p, lines, prompts, cleared = launch("crash")
    rc = p.wait()
    check("a crash gives exit code 1 and the traceback in the log", rc == 1 and any("RuntimeError: boom" in ln for ln in lines))

    p, lines, prompts, cleared = launch("hang")
    t0 = time.time()
    while not prompts and time.time() - t0 < 10:
        time.sleep(0.05)
    p.kill()
    rc = p.wait()
    check("stop run ends the process", rc != 0 and time.time() - t0 < 15, str(rc))

    print("\n" + ("ALL PASSED" if not failures else f"{len(failures)} FAILED"))
    return 0 if not failures else 1


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--fake":
        sys.exit(fake(sys.argv[2:]))
    sys.exit(run_tests())
