# (c) 2026 Alain Yon
"""Scratch helper: run one tool script in a throwaway copy of its folder and
show what it printed and which files it wrote.  Not part of the regression.

python explore.py <tool_id> [--tty] [--place src=dst ...] [--answers "a|b|c"] -- script.py args...
  The work folder gets source_tools/<tool> plus the tool's sample working files
  (sample_data/<section>/<tool>/tool_files). src is relative to sample_data/<section>/<tool>/inputs,
  dst relative to the work folder.
"""
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def snapshot(folder):
    out = {}
    for d, _, files in os.walk(folder):
        for f in files:
            p = os.path.join(d, f)
            st = os.stat(p)
            out[os.path.relpath(p, folder)] = (st.st_mtime_ns, st.st_size)
    return out


def main(argv):
    tool = argv[0]
    rest = argv[1:]
    split = rest.index("--")
    opts, cmd = rest[:split], rest[split + 1:]
    tty = "--tty" in opts
    answers = ""
    places = []
    i = 0
    while i < len(opts):
        if opts[i] == "--answers":
            answers = opts[i + 1].replace("|", "\n") + "\n"
            i += 1
        elif opts[i] == "--place":
            places.append(opts[i + 1].split("=", 1))
            i += 1
        i += 1
    sys.path.insert(0, HERE)
    from cases import sample_dir
    src = os.path.join(ROOT, "source_tools", tool)
    sample = sample_dir(tool)
    work = os.path.join(tempfile.mkdtemp(prefix="ccs_explore_"), tool)
    shutil.copytree(src, work, ignore=shutil.ignore_patterns("__pycache__"))
    if os.path.isdir(os.path.join(sample, "tool_files")):
        shutil.copytree(os.path.join(sample, "tool_files"), work, dirs_exist_ok=True)
    for s, d in places:
        dst = os.path.join(work, d)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(os.path.join(sample, "inputs", s), dst)
    before = snapshot(work)
    full = [sys.executable, "-u", "-X", "utf8"]
    if tty:
        full.append(os.path.join(ROOT, "app", "tty_shim.py"))
    full += cmd
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
    p = subprocess.run(full, cwd=work, input=answers, capture_output=True, text=True, encoding="utf-8",
                       errors="replace", env=env)
    out = p.stdout[-6000:] if "--full" not in opts else p.stdout
    print(out)
    if p.stderr.strip():
        print("--- stderr ---\n" + p.stderr[-3000:])
    print(f"--- exit code {p.returncode}   work folder {work}")
    after = snapshot(work)
    for k in sorted(after):
        if k not in before:
            print("  NEW     ", k)
        elif before[k] != after[k]:
            print("  CHANGED ", k)


if __name__ == "__main__":
    main(sys.argv[1:])
