# (c) 2026 Alain Yon
"""Regression: prove the app produces the same files as the original .bat flow.

For each case in cases.py:
  folder A  the original flow: python <script> <args as the .bat passes them>,
            answers fed to stdin in order
  folder B  the same job through the app's runner, answers given through the
            prompt API
  then      A vs B, and A vs the expected outputs, with compare_xlsx.py

Inputs, tool working files, typed answers and expected outputs come from
sample_data/<section>/<tool>/ (generated demo data, see tools/generate_dummy_data.py).

  python run_regression.py                 everything
  python run_regression.py --only theo_cos,oe_monthly
  python run_regression.py --a-only        folder A against the expected outputs only
  python run_regression.py --quick         every tool once through the runner: exits cleanly, writes its outputs
  python run_regression.py --make-golden   (re)create sample_data/.../expected from folder A
  python run_regression.py --source <dir>  use this source_tools folder
"""
import filecmp
import fnmatch
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, ROOT)

from compare_xlsx import compare  # noqa: E402
import cases as cases_mod  # noqa: E402

SHIM = os.path.join(ROOT, "app", "tty_shim.py")
TEXT_MASKS_DEFAULT = [r"\d{4}-\d\d-\d\d \d\d:\d\d(:\d\d)?", r"\d{8}_\d{4,6}"]


def find_source():
    for cand in (os.path.join(ROOT, "source_tools"),):
        if os.path.isdir(cand):
            return cand
    raise SystemExit("source_tools folder not found")


def load_manifest_ignores():
    path = os.path.join(ROOT, "app", "tools.json")
    out = {}
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            for t in json.load(f)["tools"]:
                out[t["id"]] = t.get("ignore_in_regression", [])
    return out


def answers_for(case, value, source):
    if value == "@answers":
        p = os.path.join(cases_mod.sample_dir(case["folder"]), f"answers_{case['id']}.txt")
        with open(p, encoding="utf-8") as f:
            return [ln.rstrip("\r\n") for ln in f.readlines()]
    return list(value or [])


def prepare(case, source, root):
    """Lay out <root> like the installed data folder: scripts, the tool's working files, then the inputs."""
    src = os.path.join(source, case["folder"])
    dst = os.path.join(root, "tools", case["folder"])
    shutil.copytree(src, dst, ignore=shutil.ignore_patterns("samples", "__pycache__"))
    sd = cases_mod.sample_dir(case["folder"])
    if os.path.isdir(os.path.join(sd, "tool_files")):
        shutil.copytree(os.path.join(sd, "tool_files"), dst, dirs_exist_ok=True)
    os.makedirs(os.path.join(root, "inbox"), exist_ok=True)
    for s, d in case.get("place", []):
        target = os.path.join(root, d.replace("/", os.sep))
        os.makedirs(os.path.dirname(target), exist_ok=True)
        shutil.copy2(os.path.join(sd, "inputs", s.replace("/", os.sep)), target)


def run_a(case, source, root, log):
    """The original flow, as the .bat runs it."""
    tool_dir = os.path.join(root, "tools", case["folder"])
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
    for step in case["a_steps"]:
        cmd = [sys.executable, "-u", "-X", "utf8"]
        if step.get("tty"):
            cmd.append(SHIM)
        cmd.append(step["script"])
        cmd += [a.replace("{root}", root).replace("/", os.sep) if "{root}" in a else a for a in step["args"]]
        answers = answers_for(case, step.get("answers"), source)
        p = subprocess.run(cmd, cwd=tool_dir, input="".join(a + "\n" for a in answers), capture_output=True,
                           text=True, encoding="utf-8", errors="replace", env=env)
        log.append(p.stdout + ("\n[stderr]\n" + p.stderr if p.stderr.strip() else ""))
        want = step.get("rc", 0)
        if p.returncode != want:
            return f"exit code {p.returncode}, expected {want}\n{(p.stdout + p.stderr)[-1500:]}"
    return None


def run_b(case, root, log):
    """The same job through the app's runner."""
    from app.engine import Engine
    eng = Engine(data_root=root, app_root=os.path.join(root, "_appdata"))
    b = case["b"]
    inputs = {k: [os.path.join(root, p.replace("/", os.sep)) for p in v] for k, v in b.get("inputs", {}).items()}
    source = find_source()
    answers = answers_for(case, b.get("answers"), source)
    run = eng.start_run(case["tool"], action=b.get("action"), inputs=inputs, options=b.get("options", {}),
                        skip_preflight=False)
    if run.get("error"):
        return "could not start: " + run["error"]
    rid, cursor, used = run["run_id"], 0, 0
    deadline = time.time() + 900
    while time.time() < deadline:
        res = eng.poll(rid, cursor)
        cursor = res["cursor"]
        for ev in res["events"]:
            if ev["type"] == "line":
                log.append(ev["text"])
            elif ev["type"] == "prompt":
                log.append("[prompt] " + ev["text"])
                if used >= len(answers):
                    eng.stop(rid)
                    return f"the app asked a question the case has no answer for: {ev['text']!r}"
                eng.answer(rid, answers[used])
                log.append("[answer] " + answers[used])
                used += 1
        if res["done"]:
            final = res["result"]
            log.append(f"[result] {final['status']}  {final.get('headline', '')}")
            if final["status"] != "finished":
                return f"app run ended as {final['status']}: {final.get('headline', '')}"
            if used != len(answers):
                return f"only {used} of {len(answers)} answers were asked for"
            return None
        time.sleep(0.05)
    eng.stop(rid)
    return "timed out"


def resolve_outputs(case, root, which="outputs"):
    found, missing = {}, []
    spec = case.get(which) or case["outputs"]
    for name, pattern in spec.items():
        hits = sorted(glob.glob(os.path.join(root, pattern.replace("/", os.sep))))
        hits = [h for h in hits if "_backup" not in os.path.basename(h) or "_backup" in pattern]
        if hits:
            found[name] = hits[-1]
        else:
            missing.append(pattern)
    return found, missing


def resolve_data(case, root):
    out = {}
    for pattern in case.get("data", []):
        hits = sorted(glob.glob(os.path.join(root, pattern.replace("/", os.sep))))
        key = re.sub(r"[*?\[\]]", "_", pattern)
        out[key] = hits[-1] if hits else None
    return out


ROOT_RX = re.compile(r"[A-Za-z]:\\[^\r\n\"']*?ccs_reg_[^\\]+\\[AB]")


def scrub_paths(txt):
    """Local paths out of text: the temp run folder, this checkout and the user's home folder.
    Expected outputs are published with the demo, so they must not carry a Windows user name."""
    txt = ROOT_RX.sub("<root>", txt)
    for real, token in ((ROOT, "<repo>"), (os.path.expanduser("~"), "%USERPROFILE%")):
        txt = txt.replace(real, token).replace(real.replace("\\", "/"), token)
    return txt


def compare_files(a, b, ignore):
    """Returns (diff lines, ignored lines, notes)."""
    if a is None or b is None:
        if a is None and b is None:
            return [], [], []
        return [f"file exists on one side only ({'A' if a else 'B'})"], [], []
    if a.lower().endswith((".xlsx", ".xlsm")):
        res = compare(a, b, ignore)
        return res["diffs"], res["ignored"], res["notes"]
    if filecmp.cmp(a, b, shallow=False):
        return [], [], []
    masks = [re.compile(m) for m in TEXT_MASKS_DEFAULT] + [re.compile(r["mask"]) for r in ignore if "mask" in r]

    def norm(path):
        with open(path, encoding="utf-8", errors="replace") as f:
            txt = scrub_paths(f.read())
        for m in masks:
            txt = m.sub("<ignored>", txt)
        return txt.splitlines()
    la, lb = norm(a), norm(b)
    if la == lb:
        return [], ["text file: dates, times and the temp folder path masked"], []
    diffs = [f"line {i + 1}: {x!r} vs {y!r}" for i, (x, y) in enumerate(zip(la, lb)) if x != y][:10]
    if len(la) != len(lb):
        diffs.append(f"{len(la)} lines vs {len(lb)} lines")
    return diffs, [], []


def run_case(case, source, ignores, mode, keep):
    result = dict(id=case["id"], tool=case["tool"], ok=True, lines=[], ignored=set(), notes=set())

    def fail(msg):
        result["ok"] = False
        result["lines"].append("FAIL  " + msg)

    base = tempfile.mkdtemp(prefix="ccs_reg_")
    ignore = ignores.get(case["tool"], []) + case.get("ignore", [])
    golden_dir = os.path.join(cases_mod.sample_dir(case["folder"]), "expected", case["id"])
    try:
        root_a = os.path.join(base, "A")
        log_a = []
        if mode != "quick":
            prepare(case, source, root_a)
            err = run_a(case, source, root_a, log_a)
            if err:
                fail("folder A (original flow): " + err)
                return result
            out_a, missing = resolve_outputs(case, root_a)
            for m in missing:
                fail(f"folder A wrote no file matching {m}")
            if mode == "make-golden":
                os.makedirs(golden_dir, exist_ok=True)
                for name, path in out_a.items():
                    dest = os.path.join(golden_dir, name)
                    if name.lower().endswith(".txt"):           # text logs: local paths scrubbed
                        with open(path, encoding="utf-8", errors="replace") as f:
                            txt = scrub_paths(f.read())
                        with open(dest, "w", encoding="utf-8", newline="") as f:
                            f.write(txt)
                    else:
                        shutil.copy2(path, dest)
                with open(os.path.join(golden_dir, "_console.txt"), "w", encoding="utf-8") as f:
                    f.write(scrub_paths("\n".join(log_a)))
                result["lines"].append(f"golden written: {', '.join(out_a)}")
                return result
            # A vs golden
            if not os.path.isdir(golden_dir):
                fail("no expected outputs for this case (run --make-golden once)")
            else:
                for name, path in out_a.items():
                    g = os.path.join(golden_dir, name)
                    if not os.path.exists(g):
                        fail(f"golden file missing: {name}")
                        continue
                    diffs, ign, notes = compare_files(g, path, ignore)
                    result["ignored"].update(ign)
                    result["notes"].update(notes)
                    for d in diffs:
                        fail(f"A vs golden  {name}: {d}")
        if mode in ("full", "quick"):
            root_b = os.path.join(base, "B")
            prepare(case, source, root_b)
            log_b = []
            err = run_b(case, root_b, log_b)
            if err:
                fail("folder B (through the app): " + err)
                result["lines"].append("      last app log lines: " + " | ".join(log_b[-6:]))
                return result
            out_b, missing = resolve_outputs(case, root_b, "b_outputs")
            for m in missing:
                fail(f"folder B wrote no file matching {m}")
            if mode == "full":
                for name, path in out_a.items():
                    if name not in out_b:
                        continue
                    diffs, ign, notes = compare_files(path, out_b[name], ignore)
                    result["ignored"].update(ign)
                    result["notes"].update(notes)
                    for d in diffs:
                        fail(f"A vs B  {name}: {d}")
                da, db = resolve_data(case, root_a), resolve_data(case, root_b)
                for key in da:
                    diffs, ign, notes = compare_files(da[key], db.get(key), ignore)
                    result["ignored"].update(ign)
                    for d in diffs:
                        fail(f"data file A vs B  {key}: {d}")
        return result
    except Exception as exc:  # a broken case must not hide the others
        import traceback
        fail(f"harness error: {exc}\n{traceback.format_exc()[-1200:]}")
        return result
    finally:
        if keep or not result["ok"]:
            result["lines"].append(f"      kept for inspection: {base}")
        else:
            shutil.rmtree(base, ignore_errors=True)


def main(argv):
    mode = "full"
    only, keep = None, False
    source = None
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--a-only":
            mode = "a-only"
        elif a == "--quick":
            mode = "quick"
        elif a == "--make-golden":
            mode = "make-golden"
        elif a == "--keep":
            keep = True
        elif a == "--only":
            only = argv[i + 1].split(",")
            i += 1
        elif a == "--source":
            source = argv[i + 1]
            i += 1
        else:
            print(__doc__)
            return 2
        i += 1
    source = source or find_source()
    # oe_tracker.py builds some "Group par" notes from a Python set, so their word order changes from
    # run to run even with the original .bat. A fixed hash seed (tests only) makes A, B and golden comparable.
    os.environ["PYTHONHASHSEED"] = "0"
    ignores = load_manifest_ignores()
    todo = [c for c in cases_mod.CASES if not only or c["id"] in only or c["tool"] in only]
    if mode == "quick":
        seen, quick = set(), []
        for c in todo:
            if c["tool"] not in seen:
                seen.add(c["tool"])
                quick.append(c)
        todo = quick
    print(f"Regression ({mode}) with Python {sys.version.split()[0]}  -  {len(todo)} case(s)\n")
    failed = 0
    for case in todo:
        if not os.path.isdir(os.path.join(source, case["folder"])):
            print(f"SKIP  {case['id']:<26} tool folder not supplied")
            continue
        t0 = time.time()
        r = run_case(case, source, ignores, mode, keep)
        status = "PASS" if r["ok"] else "FAIL"
        failed += 0 if r["ok"] else 1
        print(f"{status}  {r['id']:<26} {time.time() - t0:5.1f}s")
        for ln in r["lines"]:
            print("      " + ln)
        for ig in sorted(r["ignored"])[:8]:
            print("      ignored: " + ig)
        for n in sorted(r["notes"])[:4]:
            print("      note: " + n)
    if not only:
        for name, why in cases_mod.NO_SAMPLES.items():
            print(f"NOT TESTED  {name}: {why}")
    print(f"\n{'ALL PASSED' if not failed else str(failed) + ' FAILED'}")
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
