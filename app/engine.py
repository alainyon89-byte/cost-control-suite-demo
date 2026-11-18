# (c) 2026 Alain Yon
"""The back end the page talks to: tools, pre-flight, runs, history, SAP data.

Everything here returns plain JSON-able data, so the same object serves the
pywebview js_api, the Edge fallback server and the regression tests.
"""
import datetime
import filecmp
import glob
import importlib
import importlib.metadata
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
import uuid

from . import branding, detect, embed, paths, preflight
from .history import History
from .manifest import Manifest, fill_tokens
from .runner import ScriptProcess, build_command, python_exe
from .sapdata import KINDS as SAP_KINDS, PLANT_NAMES, SapData, mm60_kind
from .settings import Settings

OUTPUT_EXTS = (".xlsx", ".xlsm", ".txt", ".csv")
SKIP_DIRS = {"__pycache__", "_backup"}
BACKUP_HINTS = ("backup", "archive", "_old")


class StopRun(Exception):
    pass


def _now():
    return datetime.datetime.now().isoformat(timespec="seconds")


def _same_file(a, b):
    try:
        return os.path.samefile(a, b)
    except OSError:
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


class Run:
    def __init__(self, tool, inputs, options):
        self.id = uuid.uuid4().hex[:12]
        self.tool = tool
        self.inputs = inputs
        self.options = options
        self.events = []
        self.lock = threading.Lock()
        self.done = False
        self.result = None
        self.started = time.time()
        self.started_iso = _now()
        self.proc = None
        self.stop_requested = False
        self.question = None            # card currently waiting for an answer
        self._flow_answer = None
        self._flow_event = threading.Event()
        self.log_lines = []

    def emit(self, type_, **kw):
        with self.lock:
            kw["type"] = type_
            self.events.append(kw)

    def line(self, text, kind="out"):
        with self.lock:
            self.log_lines.append(text)
            self.events.append({"type": "line", "text": text, "kind": kind, "n": len(self.log_lines) - 1})


class RunContext:
    """What a flow (app/flows/<tool_id>.py) gets to work with."""

    def __init__(self, engine, run):
        self.engine, self.run = engine, run
        self.tool = run.tool
        self.inputs = {k: list(v) for k, v in run.inputs.items()}
        self.options = dict(run.options)
        self.tool_dir = engine.tool_dir(self.tool)
        self.tmp_dir = os.path.join(engine.app_root, "run_tmp", run.id)
        self.staged = set()             # files the app put in place (never reported as outputs)
        self.relocate = None            # (temp folder, real folder) for embedded-sheet runs
        self.block = []                 # output lines since the last answer (for question cards)
        self.last_cards = {}
        self.exit_codes = []
        self._auto_counts = {}

    # ---- logging / questions -------------------------------------------------
    def log(self, text, kind="app"):
        self.run.line(text, kind)

    def ask(self, title, help_text="", options=None, allow_text=True, card_id="flow"):
        """A question from the app's own flow (e.g. the recount approval gate). Returns the answer text."""
        card = {"id": card_id, "kind": "choice" if options else "text", "title": title, "help": help_text,
                "options": options or [], "allow_text": allow_text, "source": "flow"}
        self.run._flow_event.clear()
        self.run.question = card
        self.run.emit("prompt", text=title, card=card)
        while not self.run._flow_event.wait(0.1):
            if self.run.stop_requested:
                raise StopRun()
        self.run.question = None
        answer = self.run._flow_answer or ""
        self.log(f"{title} {answer}", "answer")
        return answer

    # ---- templates -------------------------------------------------------------
    def _fill(self, text):
        def rep(m):
            kind, key = m.group(1), m.group(2)
            if kind == "opt":
                return str(self.options.get(key, ""))
            vals = self.inputs.get(key, [])
            if kind == "in_name":
                return os.path.basename(vals[0]) if vals else ""
            return vals[0] if vals else ""
        return re.sub(r"\{(opt|in|in_name):([A-Za-z0-9_]+)\}", rep, fill_tokens(text))

    def expand_args(self, template=None):
        out = []
        for item in (self.tool.get("args", []) if template is None else template):
            if isinstance(item, dict):
                cond = item.get("when", "")
                val = self.options.get(cond[4:]) if cond.startswith("opt:") else self.inputs.get(cond[3:])
                if val not in (None, "", False, []):
                    out.extend(self.expand_args(item.get("args", [])))
                continue
            m = re.fullmatch(r"\{in:([A-Za-z0-9_]+)\}", item)
            if m:
                out.extend(p for p in self.inputs.get(m.group(1), []) if not str(p).startswith("@"))
            elif item == "{flags}":
                for opt in self.tool.get("options", []):
                    if opt.get("type") == "checkbox" and opt.get("arg") and self.options.get(opt["key"]):
                        out.append(opt["arg"])
            else:
                out.append(self._fill(item))
        return out

    # ---- running a script --------------------------------------------------------
    def run_script(self, script=None, args=None, auto_answers=None):
        """Run the tool's script (or another script in the tool folder). Returns the exit code."""
        tool, run = self.tool, self.run
        script = script or tool["script"]
        args = self.expand_args() if args is None else args
        auto = tool.get("auto_answers", []) if auto_answers is None else auto_answers
        cmd = build_command(script, args, tool.get("needs_console", False))
        shown = " ".join(f'"{a}"' if " " in a else a for a in ["python", script] + args)
        self.log(f"> {shown}", "cmd")
        self.block = []

        def on_line(text, is_answer=False):
            run.line(text, "answer" if is_answer else "out")
            if is_answer:
                self.block = []
            else:
                self.block.append(text)
                if len(self.block) > 80:
                    del self.block[:-80]
                for i, p in enumerate(tool.get("progress", [])):
                    if re.search(p["regex"], text):
                        run.emit("progress", index=i)

        def on_prompt(text):
            for i, rule in enumerate(auto):
                if re.search(rule["match"], text):
                    n = self._auto_counts.get((i, text), 0) + 1
                    self._auto_counts[(i, text)] = n
                    if n > 3:
                        break       # the script keeps rejecting the answer: hand the question to the user
                    proc.send(self._fill(rule["answer"]))
                    return
            card = self._card_for(text)
            run.question = card
            run.emit("prompt", text=text, card=card)

        def on_cleared():
            run.question = None
            run.emit("prompt_cleared")

        proc = ScriptProcess(cmd, self.tool_dir, on_line, on_prompt, on_cleared)
        run.proc = proc
        try:
            proc.start()
        except OSError as e:
            self.log(f"Could not start Python: {e}", "error")
            return 1
        rc = proc.wait()
        run.proc = None
        run.question = None
        self.exit_codes.append(rc)
        if run.stop_requested:
            raise StopRun()
        return rc

    def _card_for(self, text):
        """Build the Question card from the tool's prompt parsers. No parser match = plain Reply box."""
        for p in self.tool.get("prompts", []):
            if not re.search(p["match"], text):
                continue
            block = list(self.block)
            title = p.get("title") or ""
            if p.get("title_from_prompt"):
                title = text.strip().rstrip(":").strip()
            if p.get("title_regex"):
                for i in range(len(block) - 1, -1, -1):
                    m = re.search(p["title_regex"], block[i])
                    if m:
                        title = p.get("title_prefix", "") + m.group(1).strip()
                        block = block[i + 1:]
                        break
            options = []
            if p.get("options_regex"):
                for ln in block:
                    for m in re.finditer(p["options_regex"], ln):
                        if m.group(1) != "":
                            options.append({"value": m.group(1), "label": re.sub(r"\s{3,}", "   ", m.group(2).strip())})
            if not options and p["id"] in self.last_cards:
                options = self.last_cards[p["id"]]["options"]      # asked again after a wrong key
                title = title or self.last_cards[p["id"]]["title"]
            options = options + [o for o in p.get("extra_options", [])
                                 if o["value"] not in {x["value"] for x in options}]
            card = {"id": p["id"], "kind": p.get("kind", "text"), "title": title or text.strip(),
                    "help": p.get("help", ""), "options": options, "allow_text": True, "source": "script",
                    "prompt": text.strip()}
            self.last_cards[p["id"]] = card
            return card
        return {"id": "reply", "kind": "reply", "title": text.strip(), "help": "", "options": [],
                "allow_text": True, "source": "script", "prompt": text.strip()}

    def status_for(self, rc):
        codes = self.tool.get("exit_codes", {})
        return codes.get(str(rc), "finished" if rc == 0 else "failed")


class Engine:
    def __init__(self, data_root=None, app_root=None, manifest_path=None):
        self.data_root = data_root or paths.default_data_root()
        self.app_root = app_root or paths.default_app_root()
        os.makedirs(self.data_root, exist_ok=True)
        os.makedirs(os.path.join(self.app_root, "logs", "runs"), exist_ok=True)
        self.manifest = Manifest(manifest_path)
        self.settings = Settings(self.data_root)
        self.history = History(self.data_root)
        self.sap = SapData(self.data_root, self.settings)
        self.runs = {}
        self.active = {}                # tool id -> run id
        self._lock = threading.Lock()

    # ---- folders ---------------------------------------------------------------
    def tool_dir(self, tool):
        folder = tool.get("folder", tool["id"])
        custom = (self.settings.get("data_folders") or {}).get(folder)
        if custom and os.path.isdir(custom):
            return custom
        return os.path.join(self.data_root, "tools", folder)

    def shared_kind(self, spec, inputs, options):
        """Which shared file an input slot means. MM60s are kept per plant (D100 Hotel, D200 Conference Centre)."""
        shared = spec.get("shared", "")
        if not shared.startswith("mm60"):
            return shared
        src = spec.get("plant_from", "")
        if src.startswith("opt:"):
            return mm60_kind((options or {}).get(src[4:]))
        if src.startswith("file:"):
            files = [p for p in (inputs or {}).get(src[5:], []) if not str(p).startswith("@")]
            if files and os.path.isfile(files[0]):
                return mm60_kind(detect.sniff(files[0]).get("plant"))
        return "mm60"

    def _script_ok(self, tool, script=None):
        script = script or tool.get("script")
        return bool(script) and os.path.isfile(os.path.join(self.tool_dir(tool), script))

    # ---- what the page shows -----------------------------------------------------
    def tool_view(self, tool_id, brief=False):
        t = self.manifest.get(tool_id)
        tdir = self.tool_dir(t)
        variants = []
        for key, v in (t.get("variants") or {}).items():
            variants.append({"value": key, "label": v.get("label", key),
                             "installed": self._script_ok(t, v.get("script", t.get("script")))})
        installed = any(v["installed"] for v in variants) if variants else self._script_ok(t)
        last = self.history.last_run(tool_id)
        view = {"id": t["id"], "name": t["name"], "group": t["group"], "version": t.get("version", ""),
                "what": t["what"], "inputs_line": t.get("inputs_line", ""), "installed": installed,
                "last_run": {"time": last["time"], "status": last["status"]} if last else None,
                "running": self.active.get(tool_id)}
        if brief:
            return view
        inputs = []
        for spec in t.get("inputs", []):
            s = {k: spec.get(k) for k in ("key", "label", "accept", "multiple", "required", "shared", "empty_note")}
            s["detect"] = spec.get("detect", [])
            s["deliver"] = (spec.get("deliver") or {}).get("mode")
            if spec.get("shared"):
                def info_for(kind, spec=spec):
                    if spec["shared"] == "mm60_previous":
                        prev = self.sap.previous(kind)
                        return {"loaded": bool(prev), "label": "Previous " + SAP_KINDS[kind]["label"],
                                "original_name": self.sap.original_name_of(prev) if prev else None, "path": prev}
                    return self.sap.info(kind)
                if spec["shared"].startswith("mm60"):
                    s["plant_from"] = spec.get("plant_from", "")
                    s["by_plant"] = {branding.HOTEL_PLANT: info_for("mm60"), branding.CC_PLANT: info_for("mm60_cc")}
                    s["shared_info"] = s["by_plant"][branding.HOTEL_PLANT]
                else:
                    s["shared_info"] = info_for(spec["shared"])
            inputs.append(s)
        options = []
        for o in t.get("options", []):
            o = dict(o)
            o["default"] = fill_tokens(o.get("default", ""))
            options.append(o)
        open_files = []
        merged = list(t.get("open_files", []))
        for v in (t.get("variants") or {}).values():
            merged += v.get("open_files", [])
        for f in merged:
            p = os.path.join(tdir, fill_tokens(f["file"]))
            open_files.append({"label": f["label"], "path": p, "exists": os.path.exists(p)})
        view.update(inputs=inputs, options=options, variants=variants, variant_option=t.get("variant_option"),
                    variant_label=t.get("variant_label"), help=t.get("help", ""), open_files=open_files,
                    tool_dir=tdir, tool_dir_exists=os.path.isdir(tdir), run_label=t.get("run_label", "Run"),
                    page=t.get("page"), progress=[p["label"] for p in t.get("progress", [])],
                    recent=self.history.query(tool=tool_id, limit=5))
        return view

    def state(self):
        month = datetime.date.today().strftime("%Y-%m")
        done = self.history.finished_this_month(month)
        checklist = []
        for tid in self.settings.get("month_end_checklist") or []:
            if tid in self.manifest.ids():
                checklist.append({"id": tid, "name": self.manifest.get(tid)["name"], "done": tid in done})
        return {"app": {"name": paths.APP_NAME, "version": paths.APP_VERSION, "user": paths.windows_user(),
                        "first_name": paths.first_name(), "month": datetime.date.today().strftime("%B %Y"),
                        "data_root": self.data_root, "app_root": self.app_root,
                        "property_name": branding.property_name(),
                        "plants": {"hotel": branding.HOTEL_PLANT, "cc": branding.CC_PLANT,
                                   "names": dict(branding.SETTINGS["plants"])}},
                "groups": self.manifest.groups,
                "tools": [self.tool_view(t, brief=True) for t in self.manifest.ids()],
                "sap": self.sap.status(), "checklist": checklist, "settings": self.settings.data}

    def sniff(self, tool_id, file_paths):
        """Label dropped files for one tool (labels only; the script still decides for itself)."""
        t = self.manifest.get(tool_id) if tool_id else None
        out = []
        for p in file_paths:
            s = detect.sniff(p)
            s["recognised"] = s["kind"] != "unknown"
            if t:
                wanted = [d for spec in t.get("inputs", []) for d in spec.get("detect", [])]
                hit = next((d for d in wanted if detect.matches(s["kind"], d["kind"])), None)
                s["fits"] = bool(hit) or not wanted
                if hit:
                    s["label"] = hit["label"]
            out.append(s)
        return out

    def suggest(self, file_paths):
        res = detect.suggest_tools(file_paths)
        res["tools"] = [{"id": t, "name": "SAP data" if t == "sap_data" else self.manifest.get(t)["name"]}
                        for t in res["tools"] if t == "sap_data" or t in self.manifest.ids()]
        return res

    def preflight(self, tool_id, inputs, options):
        tool = self.manifest.resolve(tool_id, options)
        return preflight.run_checks(tool, self.tool_dir(tool), inputs or {}, options or {}, self.sap,
                                    self._script_ok(tool), lambda spec: self.shared_kind(spec, inputs, options))

    # ---- runs --------------------------------------------------------------------
    def start_run(self, tool_id, action=None, inputs=None, options=None, skip_preflight=False):
        inputs = {k: [p for p in v if p] for k, v in (inputs or {}).items()}
        options = dict(options or {})
        base = self.manifest.get(tool_id)
        if action and base.get("variant_option"):
            options[base["variant_option"]] = action
        for o in base.get("options", []):
            options.setdefault(o["key"], fill_tokens(o.get("default", "")))
        try:
            tool = self.manifest.resolve(tool_id, options)
        except Exception as e:
            return {"error": str(e)}
        with self._lock:
            if tool_id in self.active:
                return {"error": f"{tool['name']} is already running. Wait for it to finish or stop it first."}
            if not skip_preflight:
                bad = preflight.blocking(preflight.run_checks(
                    tool, self.tool_dir(tool), inputs, options, self.sap, self._script_ok(tool),
                    lambda spec: self.shared_kind(spec, inputs, options)))
                if bad:
                    return {"error": bad[0]["message"] or bad[0]["text"], "preflight": bad}
            run = Run(tool, inputs, options)
            self.runs[run.id] = run
            self.active[tool_id] = run.id
        threading.Thread(target=self._execute, args=(run,), daemon=True).start()
        return {"run_id": run.id}

    def poll(self, run_id, cursor=0):
        run = self.runs.get(run_id)
        if not run:
            return {"events": [], "cursor": cursor, "done": True, "result": {"status": "failed", "headline": "That run is no longer available."}}
        with run.lock:
            events = run.events[cursor:]
            return {"events": events, "cursor": cursor + len(events), "done": run.done, "result": run.result,
                    "question": run.question, "tool": run.tool["id"]}

    def answer(self, run_id, text):
        run = self.runs.get(run_id)
        if not run or run.done:
            return {"ok": False}
        text = "" if text is None else str(text)
        q = run.question
        if q and q.get("source") == "flow":
            run._flow_answer = text
            run._flow_event.set()
            return {"ok": True}
        if run.proc:
            run.question = None
            return {"ok": run.proc.send(text)}
        return {"ok": False}

    def stop(self, run_id):
        run = self.runs.get(run_id)
        if run and not run.done:
            run.stop_requested = True
            if run.proc:
                run.proc.kill()
        return {"ok": True}

    # ---- one run, start to finish ---------------------------------------------------
    def _execute(self, run):
        ctx = RunContext(self, run)
        tool = run.tool
        outcome = {"status": "failed", "exit_code": None}
        before = {}
        try:
            run.emit("status", text="Getting ready")
            self._deliver_inputs(ctx)
            watch = self._watch_dirs(ctx)
            before = self._snapshot(watch)
            run.emit("status", text="Running")
            if tool.get("flow"):
                flow = importlib.import_module(f"{__package__}.flows.{tool['flow']}")
                outcome = flow.run(ctx) or {}
            else:
                rc = ctx.run_script()
                outcome = {"status": ctx.status_for(rc), "exit_code": rc}
        except StopRun:
            outcome = {"status": "stopped", "exit_code": None}
        except Exception as e:
            ctx.log(f"The app hit a problem before or after the script ran: {e}", "error")
            for ln in traceback.format_exc().splitlines():
                ctx.log(ln, "error")
            outcome = {"status": "failed", "exit_code": None, "headline": f"The app could not complete the run: {e}"}
        try:
            self._finish(ctx, outcome, before)
        except Exception as e:
            run.result = {"status": "failed", "headline": f"The run ended but the app could not summarise it: {e}",
                          "outputs": [], "summary": [], "review": [], "notices": []}
        finally:
            with self._lock:
                self.active.pop(tool["id"], None)
            run.done = True
            run.emit("done")
            shutil.rmtree(ctx.tmp_dir, ignore_errors=True)

    def _deliver_inputs(self, ctx):
        """Shared SAP files: put each one where the script expects it (spec 6A). No script edits."""
        tool = ctx.tool
        for spec in tool.get("inputs", []):
            if not spec.get("shared"):
                continue
            key, mode = spec["key"], (spec.get("deliver") or {}).get("mode", "pass_as_file")
            given = ctx.inputs.get(key, [])
            override = next((p for p in given if not str(p).startswith("@")), None)
            from_store = None
            kind = self.shared_kind(spec, ctx.run.inputs, ctx.options)
            if not override and "@in_file" not in given:
                if spec["shared"] == "mm60_previous":
                    from_store = self.sap.previous(kind)
                else:
                    from_store = self.sap.path(kind)
            source = override or from_store
            if mode == "embedded_sheet":
                target_key, sheet = spec["deliver"]["into"], spec["deliver"].get("sheet", "MM60")
                user_file = (ctx.inputs.get(target_key) or [None])[0]
                ctx.inputs[key] = []
                if not user_file:
                    continue
                in_file = embed.has_sheet(user_file, sheet)
                # Default (spec 6A): the shared MM60 for this file's plant (D100 or D200) is used when it is
                # loaded; otherwise, or when the user asks, the MM60 sheet inside the file.
                want_shared = bool(override) or "@in_file" not in given
                if not want_shared or not source:
                    if not in_file:
                        ctx.log(f"Note: {os.path.basename(user_file)} has no {sheet} sheet and the {SAP_KINDS[kind]['label']} "
                                "is not loaded on the SAP data page.", "app")
                    else:
                        ctx.log(f"Using the {sheet} sheet inside {os.path.basename(user_file)}.", "app")
                    continue
                os.makedirs(ctx.tmp_dir, exist_ok=True)
                temp = os.path.join(ctx.tmp_dir, os.path.basename(user_file))
                try:
                    embed.embed_sheet(user_file, source, temp, sheet)
                except embed.EmbedError as e:
                    ctx.log(f"Note: {e} The {sheet} sheet inside the file is used instead.", "app")
                    continue
                ctx.inputs[target_key] = [temp]
                ctx.staged.add(os.path.normcase(temp))
                ctx.relocate = (ctx.tmp_dir, os.path.dirname(os.path.abspath(user_file)))
                ctx.log(f"{SAP_KINDS[kind]['label'] if from_store else 'MM60'} from {self.sap.original_name_of(source)} written into a temporary copy of "
                        f"{os.path.basename(user_file)}. Your original file is not changed.", "app")
                continue
            if not source:
                ctx.inputs[key] = []
                continue
            if mode == "copy_to_tool_folder":
                d = spec["deliver"]
                orig = self.sap.original_name_of(source) if from_store else os.path.basename(source)
                if d.get("to_dir"):
                    target = os.path.join(ctx.tool_dir, d["to_dir"], orig if d.get("keep_name") else d["to"])
                else:
                    target = os.path.join(ctx.tool_dir, d["to"])
                os.makedirs(os.path.dirname(target), exist_ok=True)
                stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
                backup_dir = os.path.join(ctx.tool_dir, "_backup", stamp)
                stale = []
                for pattern in d.get("clear", []):
                    stale += glob.glob(os.path.join(ctx.tool_dir, pattern.replace("/", os.sep)))
                if os.path.exists(target):
                    stale.append(target)
                for p in {os.path.normcase(os.path.abspath(s)): s for s in stale}.values():
                    if _same_file(p, source) or (os.path.isfile(p) and _same_file(p, target)
                                                  and filecmp.cmp(p, source, shallow=False)):
                        continue        # already the right file in the right place
                    os.makedirs(backup_dir, exist_ok=True)
                    shutil.move(p, os.path.join(backup_dir, os.path.basename(p)))
                    ctx.log(f"Moved the old {os.path.basename(p)} to _backup\\{stamp} so the script cannot pick a stale file.", "app")
                if not os.path.exists(target):
                    shutil.copy2(source, target)
                    ctx.log(f"{spec['label']}: {orig} put in the tool folder as {os.path.relpath(target, ctx.tool_dir)}.", "app")
                ctx.staged.add(os.path.normcase(os.path.abspath(target)))
                ctx.inputs[key] = [target]
            else:                       # pass_as_file
                if from_store:
                    # give the script the file under its original export name and date, as the .bat run would see it
                    os.makedirs(ctx.tmp_dir, exist_ok=True)
                    staged = os.path.join(ctx.tmp_dir, self.sap.original_name_of(source))
                    shutil.copy2(source, staged)
                    ctx.staged.add(os.path.normcase(staged))
                    ctx.inputs[key] = [staged]
                    ctx.log(f"{spec['label']}: {os.path.basename(staged)} from SAP data.", "app")
                else:
                    ctx.inputs[key] = [source]

    def _watch_dirs(self, ctx):
        dirs = {os.path.normcase(os.path.abspath(ctx.tool_dir)): (ctx.tool_dir, True)}
        for vals in ctx.inputs.values():
            for p in vals:
                if str(p).startswith("@"):
                    continue
                d = os.path.dirname(os.path.abspath(p))
                dirs.setdefault(os.path.normcase(d), (d, False))
        for vals in ctx.run.inputs.values():
            for p in vals:
                if not str(p).startswith("@"):
                    d = os.path.dirname(os.path.abspath(p))
                    dirs.setdefault(os.path.normcase(d), (d, False))
        ctx.watch = list(dirs.values())
        return ctx.watch

    def _snapshot(self, watch):
        snap = {}
        for folder, recursive in watch:
            if not os.path.isdir(folder):
                continue
            if recursive:
                for d, subdirs, files in os.walk(folder):
                    subdirs[:] = [s for s in subdirs if s not in SKIP_DIRS]
                    for f in files:
                        self._snap_one(snap, os.path.join(d, f))
            else:
                for f in os.listdir(folder):
                    self._snap_one(snap, os.path.join(folder, f))
        return snap

    @staticmethod
    def _snap_one(snap, p):
        if p.lower().endswith(OUTPUT_EXTS) and not os.path.basename(p).startswith("~$"):
            try:
                st = os.stat(p)
                snap[os.path.normcase(p)] = (p, st.st_mtime_ns, st.st_size)
            except OSError:
                pass

    def _finish(self, ctx, outcome, before):
        run, tool = ctx.run, ctx.tool
        # embedded-sheet runs wrote next to the temp copy: move those files next to the user's own file
        moved = {}
        if ctx.relocate:
            tmp, real = ctx.relocate
            if os.path.isdir(tmp):
                for f in os.listdir(tmp):
                    src = os.path.join(tmp, f)
                    if os.path.normcase(src) in ctx.staged or not os.path.isfile(src):
                        continue
                    dst = os.path.join(real, f)
                    try:
                        shutil.move(src, dst) if not os.path.exists(dst) else shutil.copy2(src, dst)
                        moved[os.path.normcase(dst)] = dst
                    except OSError as e:
                        ctx.log(f"Could not move {f} next to your file: {e}", "error")
        after = self._snapshot(ctx.watch)
        outputs = []
        data_names = {n.lower() for n in tool.get("data_files", [])}
        originals = {os.path.normcase(os.path.abspath(p)) for v in run.inputs.values() for p in v
                     if not str(p).startswith("@")}
        for key, (p, mtime, size) in sorted(after.items()):
            if key in ctx.staged or os.path.normcase(os.path.dirname(p)) == os.path.normcase(ctx.tmp_dir):
                continue
            state = "new" if key not in before else ("changed" if before[key][1:] != (mtime, size) else None)
            if key in moved:
                state = "new"
            if not state:
                continue
            rel = os.path.relpath(p, ctx.tool_dir) if p.lower().startswith(ctx.tool_dir.lower()) else p
            low = rel.lower()
            is_backup = any(h in low for h in BACKUP_HINTS)
            is_data = os.path.basename(p).lower() in data_names and state == "changed" and key not in originals
            kind = "backup" if is_backup else "log" if low.endswith(".txt") else "data" if (
                is_data and not self._is_declared_output(tool, p)) else "output"
            outputs.append({"path": p, "name": os.path.basename(p), "folder": os.path.dirname(p), "state": state,
                            "kind": kind})
        outputs.sort(key=lambda o: ("output", "log", "data", "backup").index(o["kind"]))
        lines = run.log_lines
        status = outcome.get("status", "failed")
        if status == "stopped":
            wrote = [o for o in outputs if o["kind"] in ("output", "data")]
            headline = ("Stopped, nothing written" if tool.get("atomic") and not wrote
                        else "Stopped, check the output folder")
        elif status == "finished":
            headline = outcome.get("headline") or "Finished"
        elif status == "cancelled":
            headline = outcome.get("headline") or "Cancelled"
        else:
            headline = outcome.get("headline") or ("The script stopped: " + self._last_error_line(lines))
        summary = []
        for rule in tool.get("summary_from_log", []):
            val = None
            for ln in lines:
                m = re.search(rule["regex"], ln)
                if m:
                    val = (m.group(1) if m.groups() else m.group(0)).strip()
            if val is not None:
                summary.append({"label": rule["label"], "value": val})
        review, seen = [], set()
        for needle in tool.get("review_links_from_log", []):
            for n, ln in enumerate(lines):
                if needle in ln and n not in seen:
                    seen.add(n)
                    review.append({"text": ln.strip()[:160], "line": n})
                    if sum(1 for r in review if needle in r["text"]) >= 3:
                        break
        review.sort(key=lambda r: r["line"])
        notices = [ln.strip() for ln in lines if any(k in ln for k in tool.get("important_lines", []))]
        duration = round(time.time() - run.started, 1)
        log_path = os.path.join(self.app_root, "logs", "runs",
                                f"{datetime.datetime.now():%Y%m%d_%H%M%S}_{tool['id']}.log")
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(f"{tool['name']} {tool.get('version', '')}  |  {run.started_iso}  |  {paths.windows_user()}  |  {status}\n")
                f.write("\n".join(lines) + "\n")
        except OSError:
            log_path = None
        entry = {"time": run.started_iso, "tool": tool["id"], "tool_name": tool["name"], "version": tool.get("version", ""),
                 "variant": tool.get("variant"), "user": paths.windows_user(),
                 "inputs": [os.path.basename(p) for v in run.inputs.values() for p in v if not str(p).startswith("@")],
                 "options": run.options, "status": status, "exit_codes": ctx.exit_codes,
                 "outputs": [o["path"] for o in outputs if o["kind"] in ("output", "log")],
                 "duration": duration, "log": log_path, "headline": headline}
        self.history.add(entry)
        run.result = {"status": status, "headline": headline, "exit_code": outcome.get("exit_code"),
                      "outputs": outputs, "summary": summary, "review": review[:12], "notices": notices[:6],
                      "log_path": log_path, "duration": duration, "tool": tool["id"], "tool_name": tool["name"],
                      "version": tool.get("version", ""), "help": tool.get("help", ""),
                      "details": self._details(tool, run, status, lines)}

    @staticmethod
    def _is_declared_output(tool, path):
        import fnmatch
        return any(fnmatch.fnmatch(os.path.basename(path), o["glob"]) for o in tool.get("outputs", []))

    @staticmethod
    def _last_error_line(lines):
        tail = [ln.strip() for ln in lines if ln.strip()]
        for ln in reversed(tail[-40:]):
            if re.search(r"\bERROR\b|^\w*(Error|Exception):|^Traceback|\[ERROR\]", ln) and not ln.startswith("Traceback"):
                return ln[:240]
        return tail[-1][:240] if tail else "it ended without printing anything"

    @staticmethod
    def _details(tool, run, status, lines):
        """Text for 'Copy details for Alain'."""
        ins = [f"  {k}: {', '.join(os.path.basename(p) for p in v)}" for k, v in run.inputs.items() if v]
        return "\n".join([
            f"Cost Control Suite {paths.APP_VERSION}", f"Tool: {tool['name']} {tool.get('version', '')} ({tool['id']})",
            f"When: {run.started_iso}   User: {paths.windows_user()}   Result: {status}",
            f"Python: {sys.version.split()[0]} ({python_exe()})", "Inputs:", *(ins or ["  (none)"]),
            f"Options: {run.options}", "Log:", *lines])

    # ---- small services for the page ----------------------------------------------
    def history_rows(self, tool=None, month=None):
        return self.history.query(tool=tool or None, month=month or None)

    def read_log(self, path):
        logs = os.path.join(self.app_root, "logs")
        if not path or not os.path.abspath(path).lower().startswith(os.path.abspath(logs).lower()):
            return {"ok": False, "text": ""}
        try:
            with open(path, encoding="utf-8", errors="replace") as f:
                return {"ok": True, "text": f.read()}
        except OSError:
            return {"ok": False, "text": ""}

    def open_path(self, path):
        if not path or not os.path.exists(path):
            return {"ok": False, "reason": "That file or folder is no longer there."}
        try:
            os.startfile(path)
            return {"ok": True}
        except OSError as e:
            return {"ok": False, "reason": str(e)}

    def reveal(self, path):
        if not path or not os.path.exists(path):
            return {"ok": False, "reason": "That file is no longer there."}
        explorer = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "explorer.exe")
        subprocess.Popen([explorer, "/select,", os.path.normpath(path)])
        return {"ok": True}

    def open_tool_folder(self, tool_id):
        d = self.tool_dir(self.manifest.get(tool_id))
        return self.open_path(d)

    def open_logs(self):
        return self.open_path(os.path.join(self.app_root, "logs"))

    # ---- SAP data ---------------------------------------------------------------------
    def sap_status(self):
        s = self.sap.status()
        for k in s:
            s[k]["history"] = self.sap.history(k)
        return s

    def sap_inspect(self, kind, path):
        if kind not in SAP_KINDS:
            return {"ok": False, "reason": "Unknown SAP file type."}
        return self.sap.inspect(kind, path)

    def sap_replace(self, kind, path):
        if kind not in SAP_KINDS:
            return {"ok": False, "reason": "Unknown SAP file type."}
        return self.sap.replace(kind, path, paths.windows_user())

    # ---- settings ------------------------------------------------------------------------
    def save_settings(self, changes):
        return self.settings.update(changes)

    def check_setup(self):
        rows = [{"what": "Python", "value": f"{sys.version.split()[0]} at {python_exe()}", "ok": sys.version_info[:2] >= (3, 10)}]
        for pkg in ("openpyxl", "pandas", "numpy", "pdfplumber", "pywebview"):
            try:
                rows.append({"what": pkg, "value": importlib.metadata.version(pkg), "ok": True})
            except importlib.metadata.PackageNotFoundError:
                rows.append({"what": pkg, "value": "not installed", "ok": False})
        wv = webview2_version()
        rows.append({"what": "WebView2", "value": wv or "not found (the Edge window is used instead)", "ok": True})
        folders = [("App folder", self.app_root), ("Data folder", self.data_root)]
        for tid in self.manifest.ids():
            t = self.manifest.get(tid)
            folders.append((f"{t['name']} folder", self.tool_dir(t)))
        seen = set()
        for label, folder in folders:
            if folder in seen:
                continue
            seen.add(folder)
            ok, value = False, "folder not there yet"
            if os.path.isdir(folder):
                try:
                    probe = os.path.join(folder, f".write_test_{uuid.uuid4().hex[:6]}")
                    with open(probe, "w") as f:
                        f.write("x")
                    os.remove(probe)
                    ok, value = True, folder
                except OSError:
                    value = f"cannot write to {folder}"
            rows.append({"what": label, "value": value, "ok": ok or "folder not there" in value,
                         "missing": not os.path.isdir(folder)})
        return rows


def webview2_version():
    try:
        import winreg
    except ImportError:
        return None
    key = r"Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
    for hive, base in ((winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\WOW6432Node\\"), (winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\"),
                       (winreg.HKEY_CURRENT_USER, "Software\\")):
        try:
            with winreg.OpenKey(hive, base + key) as k:
                v = winreg.QueryValueEx(k, "pv")[0]
                if v and v != "0.0.0.0":
                    return v
        except OSError:
            continue
    return None
