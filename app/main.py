# (c) 2026 Alain Yon
"""Cost Control Suite: the window.

  pythonw.exe app\\main.py             native window (pywebview on Edge WebView2)
  python.exe  app\\main.py --edge      force the fallback: local page in an Edge app window
  python.exe  app\\main.py --selftest  headless check used by the installer (no window)

No Flask, no telemetry, no network calls. The fallback server listens on
127.0.0.1 only, on a random free port, and needs a one-time token.
"""
import json
import os
import secrets
import subprocess
import sys
import threading
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
if __package__ in (None, ""):
    sys.path.insert(0, os.path.dirname(HERE))
    __package__ = "app"

from . import paths                                    # noqa: E402
from .engine import Engine, webview2_version           # noqa: E402

WEB = os.path.join(HERE, "web")
ENGINE_METHODS = {"state", "tool_view", "sniff", "suggest", "preflight", "start_run", "poll", "answer", "stop",
                  "history_rows", "read_log", "open_path", "reveal", "open_tool_folder", "open_logs",
                  "sap_status", "sap_inspect", "sap_replace", "save_settings", "check_setup"}


def app_log(text):
    try:
        folder = os.path.join(paths.default_app_root(), "logs")
        os.makedirs(folder, exist_ok=True)
        with open(os.path.join(folder, "app.log"), "a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')}  {text}\n")
    except OSError:
        pass


class Api:
    """The one object the page talks to (pywebview js_api, or the fallback server)."""

    def __init__(self, engine, mode):
        self._engine = engine
        self._mode = mode
        self._window = None

    def call(self, name, args=None):
        args = args or []
        try:
            if name == "pick_files":
                return self._pick_files(*args)
            if name == "pick_folder":
                return self._pick_folder()
            if name == "mode":
                return {"mode": self._mode, "drop": self._mode == "window"}
            if name not in ENGINE_METHODS:
                return {"error": f"Unknown request: {name}"}
            return getattr(self._engine, name)(*args)
        except Exception as e:
            app_log(f"api {name} failed: {e}\n{traceback.format_exc()}")
            return {"error": f"{type(e).__name__}: {e}"}

    # ---- file pickers ---------------------------------------------------------
    def _pick_files(self, multiple=False, accept=None):
        accept = accept or [".xlsx"]
        if self._window is not None:
            import webview
            pattern = ";".join("*" + e for e in accept)
            types = (f"Files ({pattern})", "All files (*.*)")
            res = self._window.create_file_dialog(webview.FileDialog.OPEN, allow_multiple=bool(multiple),
                                                  file_types=types)
            return list(res) if res else []
        return _powershell_dialog(multiple, accept)

    def _pick_folder(self):
        if self._window is not None:
            import webview
            res = self._window.create_file_dialog(webview.FileDialog.FOLDER)
            return list(res)[:1] if res else []
        return _powershell_dialog(False, None, folder=True)


def _powershell_dialog(multiple, accept, folder=False):
    """File picker for the Edge fallback (a browser page cannot see real file paths)."""
    ps = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "WindowsPowerShell", "v1.0",
                      "powershell.exe")
    if folder:
        script = ("Add-Type -AssemblyName System.Windows.Forms; $d = New-Object System.Windows.Forms.FolderBrowserDialog; "
                  "if ($d.ShowDialog() -eq 'OK') { $d.SelectedPath }")
    else:
        flt = "Files|" + ";".join("*" + e for e in accept) + "|All files|*.*"
        script = ("Add-Type -AssemblyName System.Windows.Forms; $d = New-Object System.Windows.Forms.OpenFileDialog; "
                  f"$d.Filter = '{flt}'; $d.Multiselect = ${'true' if multiple else 'false'}; "
                  "if ($d.ShowDialog() -eq 'OK') { $d.FileNames }")
    try:
        out = subprocess.run([ps, "-NoProfile", "-STA", "-Command", script], capture_output=True, text=True,
                             timeout=600, creationflags=0x08000000).stdout
        return [ln.strip() for ln in out.splitlines() if ln.strip()]
    except Exception:
        return []


# ---- native window ---------------------------------------------------------------
def run_window(engine):
    import webview
    from webview.dom import DOMEventHandler
    api = Api(engine, "window")
    size = engine.settings.get("window") or {}
    window = webview.create_window(paths.APP_NAME, url=os.path.join(WEB, "index.html"), js_api=api,
                                   width=int(size.get("width", 1280)), height=int(size.get("height", 860)),
                                   min_size=(1024, 700), background_color="#F4F3EF", text_select=True)
    api._window = window

    def on_drop(event):
        files = (event.get("dataTransfer") or {}).get("files", [])
        found = [f.get("pywebviewFullPath") for f in files if f.get("pywebviewFullPath")]
        if found:
            window.evaluate_js(f"window.ccsDropped({json.dumps(found)})")

    def on_loaded():
        try:
            window.dom.document.events.drop += DOMEventHandler(on_drop, True, False)
        except Exception as e:
            app_log(f"drop handler not attached: {e}")

    def on_closing():
        try:
            engine.save_settings({"window": {"width": int(window.width), "height": int(window.height)}})
        except Exception:
            pass

    window.events.loaded += on_loaded
    window.events.closing += on_closing
    webview.start(private_mode=True)


# ---- fallback: 127.0.0.1 + Edge app window ----------------------------------------
def find_edge():
    for base in (os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles"), os.environ.get("LOCALAPPDATA")):
        if base:
            p = os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe")
            if os.path.exists(p):
                return p
    return None


def run_fallback(engine, open_browser=True, port=0):
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
    api = Api(engine, "edge")
    token = secrets.token_urlsafe(24)
    state = {"last_seen": time.time()}
    types = {".html": "text/html; charset=utf-8", ".css": "text/css; charset=utf-8",
             ".js": "application/javascript; charset=utf-8", ".woff2": "font/woff2", ".ico": "image/x-icon",
             ".svg": "image/svg+xml", ".txt": "text/plain; charset=utf-8"}

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _send(self, code, body, ctype="application/json; charset=utf-8"):
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            path = self.path.split("?", 1)[0]
            rel = "index.html" if path in ("/", "") else path.lstrip("/")
            full = os.path.normpath(os.path.join(WEB, rel))
            if not full.startswith(WEB) or not os.path.isfile(full):
                return self._send(404, b"not found", "text/plain")
            with open(full, "rb") as f:
                self._send(200, f.read(), types.get(os.path.splitext(full)[1].lower(), "application/octet-stream"))

        def do_POST(self):
            state["last_seen"] = time.time()
            if self.headers.get("X-CCS-Token") != token or self.client_address[0] != "127.0.0.1":
                return self._send(403, b'{"error":"forbidden"}')
            try:
                body = json.loads(self.rfile.read(int(self.headers.get("Content-Length") or 0)) or b"{}")
                result = api.call(body.get("name"), body.get("args") or [])
                self._send(200, json.dumps(result, ensure_ascii=False, default=str).encode("utf-8"))
            except Exception as e:
                self._send(500, json.dumps({"error": str(e)}).encode("utf-8"))

    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{server.server_address[1]}/#t={token}"
    threading.Thread(target=server.serve_forever, daemon=True).start()
    app_log(f"fallback server on {server.server_address[1]}")
    if not open_browser:
        return server, url
    edge = find_edge()
    if edge:
        subprocess.Popen([edge, f"--app={url}", "--window-size=1280,860"])
    else:
        os.startfile(url)
    # stay alive while the page is open (it calls in every few seconds) or a tool is running
    while time.time() - state["last_seen"] < 90 or engine.active:
        time.sleep(2)
    server.shutdown()


def selftest():
    """Headless: import every package, load the manifest, start the engine and the local server."""
    ok = True
    for mod in ("openpyxl", "pandas", "numpy", "pdfplumber", "webview"):
        try:
            __import__(mod)
            print(f"  PASS  import {mod}")
        except Exception as e:
            ok = False
            print(f"  FAIL  import {mod}: {e}")
    try:
        engine = Engine()
        st = engine.state()
        print(f"  PASS  app engine started, {len(st['tools'])} tools in tools.json")
        server, url = run_fallback(engine, open_browser=False)
        import urllib.request
        with urllib.request.urlopen(url.split("#")[0], timeout=10) as r:
            page_ok = r.status == 200 and b"Cost Control Suite" in r.read()
        server.shutdown()
        print("  PASS  app page served" if page_ok else "  FAIL  app page not served")
        ok = ok and page_ok
        for t in st["tools"]:
            print(f"  {'PASS' if t['installed'] else 'NOTE'}  {t['name']}: " +
                  ("script in place" if t["installed"] else "not installed yet (add its folder and run the installer again)"))
        print(f"  NOTE  WebView2: {webview2_version() or 'not found, the Edge window will be used'}")
    except Exception as e:
        ok = False
        print(f"  FAIL  app engine: {e}")
    return 0 if ok else 1


def main(argv):
    if "--selftest" in argv:
        return selftest()
    engine = Engine()
    if "--edge" not in argv and webview2_version():
        try:
            run_window(engine)
            return 0
        except Exception as e:
            app_log(f"native window failed, using Edge instead: {e}\n{traceback.format_exc()}")
    run_fallback(engine)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
