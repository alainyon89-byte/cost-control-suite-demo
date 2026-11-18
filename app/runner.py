# (c) 2026 Alain Yon
"""Runs one tool script as a subprocess, never imported (spec 4.2, 4.3).

    <private python>\\python.exe -u -X utf8 <script.py> <args exactly as the .bat passes them>

* stdout and stderr are streamed line by line
* stdin stays open so the app can answer the script's questions
* a question is detected when the last output has no newline and nothing new
  arrives for 400 ms (input() prints its prompt without a trailing newline)
"""
import codecs
import os
import subprocess
import sys
import threading
import time

HERE = os.path.dirname(os.path.abspath(__file__))
SHIM = os.path.join(HERE, "tty_shim.py")
PROMPT_IDLE_SECONDS = 0.4
CREATE_NO_WINDOW = 0x08000000


def python_exe():
    """The interpreter running the app (the private Python). pythonw has no console, so use python.exe."""
    exe = sys.executable
    if os.path.basename(exe).lower() == "pythonw.exe":
        cand = os.path.join(os.path.dirname(exe), "python.exe")
        if os.path.exists(cand):
            return cand
    return exe


def build_command(script, args, needs_console=False):
    cmd = [python_exe(), "-u", "-X", "utf8"]
    if needs_console:
        cmd.append(SHIM)
    cmd.append(script)
    cmd.extend(args)
    return cmd


class ScriptProcess:
    """One running script. Callbacks are called from the reader threads."""

    def __init__(self, cmd, cwd, on_line, on_prompt, on_prompt_cleared=None, env=None,
                 idle=PROMPT_IDLE_SECONDS):
        self.cmd, self.cwd = cmd, cwd
        self.on_line, self.on_prompt, self.on_prompt_cleared = on_line, on_prompt, on_prompt_cleared
        self.idle = idle
        self.env = dict(os.environ)
        self.env.update({"PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"})
        from . import branding           # report titles in the tools read the property name from here
        self.env.setdefault("CCS_PROPERTY_NAME", branding.property_name())
        self.env.update(env or {})
        self.proc = None
        self._buf = ""                  # output since the last newline
        self._last_data = 0.0
        self._announced = None          # prompt text already reported
        self._lock = threading.RLock()
        self._eof = threading.Event()
        self.killed = False

    def start(self):
        self.proc = subprocess.Popen(
            self.cmd, cwd=self.cwd, env=self.env, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT, bufsize=0,
            creationflags=CREATE_NO_WINDOW if os.name == "nt" else 0)
        threading.Thread(target=self._read, daemon=True).start()
        threading.Thread(target=self._watch, daemon=True).start()
        return self

    def _read(self):
        decoder = codecs.getincrementaldecoder("utf-8")(errors="replace")
        fd = self.proc.stdout.fileno()
        while True:
            try:
                chunk = os.read(fd, 8192)
            except OSError:
                chunk = b""
            if not chunk:
                break
            text = decoder.decode(chunk)
            if not text:
                continue
            with self._lock:
                if self._announced is not None:
                    # more output arrived while a question was showing: it was not a question
                    self._announced = None
                    if self.on_prompt_cleared:
                        self.on_prompt_cleared()
                self._buf += text
                self._last_data = time.monotonic()
                while "\n" in self._buf:
                    line, self._buf = self._buf.split("\n", 1)
                    self.on_line(line.rstrip("\r"))
        with self._lock:
            if self._buf.strip():
                self.on_line(self._buf.rstrip("\r"))
            self._buf = ""
        self._eof.set()

    def _watch(self):
        while not self._eof.is_set():
            time.sleep(0.05)
            fire = None
            with self._lock:
                if (self._buf.strip() and self._announced is None
                        and time.monotonic() - self._last_data >= self.idle and self.proc.poll() is None):
                    self._announced = self._buf
                    fire = self._buf
            if fire is not None:
                self.on_prompt(fire)

    def pending_prompt(self):
        with self._lock:
            return self._announced

    def send(self, text):
        """Answer the script exactly as typed keys would: the text plus Enter."""
        with self._lock:
            shown = self._buf
            self._buf = ""
            self._announced = None
            self._last_data = time.monotonic()
        self.on_line(shown + text, True)
        try:
            self.proc.stdin.write((text + "\n").encode("utf-8"))
            self.proc.stdin.flush()
            return True
        except (OSError, ValueError):
            return False

    def wait(self):
        rc = self.proc.wait()
        self._eof.wait(timeout=10)
        try:
            self.proc.stdin.close()
        except Exception:
            pass
        return rc

    def kill(self):
        """Stop run: end the whole process tree."""
        self.killed = True
        if not self.proc or self.proc.poll() is not None:
            return
        try:
            taskkill = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "taskkill.exe")
            subprocess.run([taskkill, "/PID", str(self.proc.pid), "/T", "/F"], capture_output=True,
                           creationflags=CREATE_NO_WINDOW if os.name == "nt" else 0, timeout=15)
        except Exception:
            pass
        try:
            if self.proc.poll() is None:
                self.proc.kill()
        except Exception:
            pass
