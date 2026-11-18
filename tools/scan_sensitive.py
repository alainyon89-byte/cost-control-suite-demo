# (c) 2026 Alain Yon
"""Scan a folder for text that must not appear in the public demo.

    python tools/scan_sensitive.py <folder> [--json out.json] [--max-per-term N]
                                            [--no-gitignore] [--no-allowlist]

The search terms are NOT kept in this file: they are read from
tools/sensitive_terms.local.txt (git-ignored, so the list of names it looks
for is never published). Each line is `category | regex`; lines starting with
# are comments. Without that file only the generic checks run (paths, e-mail,
phone, secrets).

Text files are searched line by line. Every .xlsx/.xlsm is opened with openpyxl:
string cells, sheet names and the document properties (creator,
lastModifiedBy, company, title) are checked. For workbooks only the cell
reference and the matched term are printed, never the cell's value. File and
folder names are checked too.

Files the folder's .gitignore leaves out are skipped (they are never
published); --no-gitignore scans them as well.

tools/scan_allowlist.txt lists reviewed false positives (the word "token" in
the app's session-token code, for example). Each line is
`path | matched text | text on the line (optional) | reason`. Allowlisted hits
are counted and listed but do not fail the scan; --no-allowlist turns it off.
Allowlist entries that no longer match anything are reported, so the list
cannot go stale unnoticed.

Exit code 1 if anything not allowlisted matches, 0 if clean.
"""
import argparse
import fnmatch
import json
import os
import re
import sys
import warnings

warnings.filterwarnings("ignore")

SKIP_DIRS = {".git", "build", "dist", "__pycache__", ".venv", "venv", "node_modules"}
TEXT_EXT = {".py", ".csv", ".json", ".txt", ".md", ".ini", ".toml", ".spec", ".bat", ".cmd",
            ".js", ".html", ".css", ".cfg", ".yml", ".yaml", ".xml", ".ps1", ".svg"}
TEXT_NAMES = {".gitignore", ".gitattributes", "LICENSE", "LICENCE"}
XLSX_EXT = {".xlsx", ".xlsm"}
HERE = os.path.dirname(os.path.abspath(__file__))
TERMS_FILE = os.path.join(HERE, "sensitive_terms.local.txt")
ALLOW_FILE = os.path.join(HERE, "scan_allowlist.txt")
SELF_FILES = {os.path.abspath(__file__), os.path.abspath(TERMS_FILE), os.path.abspath(ALLOW_FILE)}

# generic checks that are safe to publish
GENERIC = [
    ("Paths", r"[A-Za-z]:\\+Users\\+"),
    ("Paths", r"/Users/"),
    ("Paths", r"OneDrive"),
    ("Paths", r"Desktop\\"),
    ("Paths", r"(?<![\\\w])\\\\[A-Za-z][\w.-]*\\[\w$.-]+"),          # \\server\share
    ("Contact details", r"[\w.-]+@[\w-]+\.[\w.-]+"),
    ("Contact details", r"(?<!\d)0\d{2}\s?\d{3}\s?\d{4}(?!\d)"),
    ("Secrets", r"api[_-]?key|passw(or)?d|\btoken\b|secret|connection[_ ]?string|Server=[^;]+;"),
]


def load_terms():
    rules = list(GENERIC)
    if os.path.isfile(TERMS_FILE):
        with open(TERMS_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "|" not in line:
                    continue
                cat, rx = line.split("|", 1)
                rules.append((cat.strip(), rx.strip()))
    return [(c, r, re.compile(r, re.IGNORECASE)) for c, r in rules]


def load_allowlist():
    """[(path, matched text lower, line text or '', reason)] from scan_allowlist.txt."""
    out = []
    if not os.path.isfile(ALLOW_FILE):
        return out
    with open(ALLOW_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            if not line.strip() or line.lstrip().startswith("#"):
                continue
            parts = [p.strip() for p in line.split("|")]
            if len(parts) < 4:
                print(f"note: allowlist line ignored (needs 4 fields): {line}")
                continue
            path, term, ctx, reason = parts[0], parts[1], parts[2], "|".join(parts[3:])
            out.append((path.replace("\\", "/"), term.lower(), ctx, reason))
    return out


def load_gitignore(root):
    """Plain patterns from <root>/.gitignore (no negation support, none is used)."""
    pats = []
    p = os.path.join(root, ".gitignore")
    if os.path.isfile(p):
        with open(p, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and not line.startswith("!"):
                    pats.append(line)
    return pats


def ignored(rel, is_dir, pats):
    """True if git would ignore this path (rel uses /). Covers the pattern forms used here."""
    name = rel.rsplit("/", 1)[-1]
    for pat in pats:
        dir_only = pat.endswith("/")
        pt = pat.rstrip("/")
        if dir_only and not is_dir:
            continue
        if pt.startswith("/") or "/" in pt:
            if fnmatch.fnmatch(rel, pt.lstrip("/")):
                return True
        elif fnmatch.fnmatch(name, pt):
            return True
    return False


def walk(root, pats, skipped):
    for dp, dns, fns in os.walk(root):
        rel_dp = os.path.relpath(dp, root).replace("\\", "/")
        rel_dp = "" if rel_dp == "." else rel_dp + "/"
        keep = []
        for d in dns:
            if d in SKIP_DIRS:
                continue
            if ignored(rel_dp + d, True, pats):
                skipped.append(rel_dp + d + "/")
                continue
            keep.append(d)
        dns[:] = keep
        for fn in fns:
            p = os.path.abspath(os.path.join(dp, fn))
            if p in SELF_FILES:
                continue
            if ignored(rel_dp + fn, False, pats):
                skipped.append(rel_dp + fn)
                continue
            yield p


def scan_name(rel, rules):
    for cat, raw, rx in rules:
        if cat == "Paths":
            continue
        m = rx.search(rel)
        if m:
            yield ("file name", cat, m.group(0), rel)


def scan_text(path, rules):
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            for n, line in enumerate(f, 1):
                for cat, raw, rx in rules:
                    m = rx.search(line)
                    if m:
                        yield (f"line {n}", cat, m.group(0), line.strip()[:160])
    except OSError as e:
        yield ("-", "Unreadable", type(e).__name__, "")


def scan_xlsx(path, rules):
    from openpyxl import load_workbook
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        yield ("-", "Unreadable", type(e).__name__, "")
        return
    try:
        props = wb.properties
        for attr in ("creator", "lastModifiedBy", "company", "title"):
            v = getattr(props, attr, None)
            if v:
                for cat, raw, rx in rules:
                    m = rx.search(str(v))
                    if m:
                        yield (f"property {attr}", cat, m.group(0), "")
        for ws in wb.worksheets:
            for cat, raw, rx in rules:          # sheet names too
                m = rx.search(ws.title)
                if m:
                    yield (f"sheet name '{ws.title}'", cat, m.group(0), "")
            try:
                for row in ws.iter_rows():
                    for c in row:
                        v = getattr(c, "value", None)
                        if isinstance(v, str) and v:
                            for cat, raw, rx in rules:
                                m = rx.search(v)
                                if m:
                                    yield (f"{ws.title}!{c.coordinate}", cat, m.group(0), "")
            except Exception:
                pass
    finally:
        wb.close()


def allowed(rel, hit, allow, used):
    loc, cat, term, ctx = hit
    for i, (path, aterm, actx, reason) in enumerate(allow):
        if path == rel and aterm == term.lower() and (not actx or actx in ctx):
            used.add(i)
            return reason
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--json", help="also write all hits to this JSON file")
    ap.add_argument("--max-per-term", type=int, default=5,
                    help="locations printed per file and term (the count is always complete)")
    ap.add_argument("--no-gitignore", action="store_true", help="also scan files .gitignore leaves out")
    ap.add_argument("--no-allowlist", action="store_true", help="report reviewed false positives as hits")
    a = ap.parse_args()
    rules = load_terms()
    if not os.path.isfile(TERMS_FILE):
        print(f"note: {TERMS_FILE} not found, only generic checks run")
    root = os.path.abspath(a.folder)
    pats = [] if a.no_gitignore else load_gitignore(root)
    allow = [] if a.no_allowlist else load_allowlist()
    used = set()
    skipped = []
    results, allowed_hits = {}, {}
    for p in walk(root, pats, skipped):
        rel = os.path.relpath(p, root).replace("\\", "/")
        hits = list(scan_name(rel, rules))
        ext = os.path.splitext(p)[1].lower()
        if ext in TEXT_EXT or os.path.basename(p) in TEXT_NAMES:
            hits += list(scan_text(p, rules))
        elif ext in XLSX_EXT:
            hits += list(scan_xlsx(p, rules))
        real = []
        for h in hits:
            reason = allowed(rel, h, allow, used)
            if reason:
                allowed_hits.setdefault(rel, []).append((h, reason))
            else:
                real.append(h)
        if real:
            results[rel] = real
    if skipped:
        print("Skipped (left out by .gitignore, never published):")
        for s in skipped:
            print(f"  {s}")
    total = 0
    for rel, hits in sorted(results.items()):
        print(f"\n{rel}")
        groups = {}
        for loc, cat, term, ctx in hits:
            groups.setdefault((cat, term.lower()), []).append((loc, term, ctx))
        for (cat, _), items in sorted(groups.items()):
            total += len(items)
            print(f"  [{cat}] '{items[0][1]}' x{len(items)}")
            for loc, term, ctx in items[:a.max_per_term]:
                print(f"      {loc}" + (f": {ctx}" if ctx else ""))
            if len(items) > a.max_per_term:
                print(f"      ... and {len(items) - a.max_per_term} more")
    n_allowed = sum(len(v) for v in allowed_hits.values())
    if allowed_hits:
        print(f"\nAllowlisted (reviewed false positives, tools/scan_allowlist.txt): {n_allowed}")
        for rel, items in sorted(allowed_hits.items()):
            by_reason = {}
            for (loc, cat, term, ctx), reason in items:
                by_reason.setdefault((term.lower(), reason), []).append(loc)
            for (term, reason), locs in by_reason.items():
                print(f"  {rel}: '{term}' x{len(locs)} ({', '.join(locs[:6])}) - {reason}")
    stale = [allow[i] for i in range(len(allow)) if i not in used]
    if stale:
        print("\nAllowlist entries that matched nothing (remove them):")
        for path, term, ctx, reason in stale:
            print(f"  {path} | {term} | {ctx}")
    if a.json:
        with open(a.json, "w", encoding="utf-8") as f:
            json.dump({k: [list(h) for h in v] for k, v in results.items()}, f, indent=1)
    print(f"\n{total} hit(s) in {len(results)} file(s); {n_allowed} allowlisted; {len(skipped)} path(s) skipped")
    sys.exit(1 if results else 0)


if __name__ == "__main__":
    main()
