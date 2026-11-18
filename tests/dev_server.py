# (c) 2026 Alain Yon
"""Developer helper: serve the app page on 127.0.0.1 against a throwaway data folder.

python dev_server.py [port]     prints the address to open (the same page the Edge fallback shows)
The data folder is laid out by the installer's own copy step (scripts, the demo's tool files and
sample inputs) in a temp folder, so nothing outside it is touched.
"""
import os
import shutil
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)


def make_icon():
    """Draws app/web/icon.ico (navy tile, brass CC)."""
    from PIL import Image, ImageDraw, ImageFont
    size = 256
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((8, 8, size - 8, size - 8), radius=52, fill="#13213C")
    try:
        font = ImageFont.truetype("segoeuib.ttf", 132)
    except OSError:
        font = ImageFont.load_default()
    box = d.textbbox((0, 0), "CC", font=font)
    d.text(((size - (box[2] - box[0])) / 2 - box[0], (size - (box[3] - box[1])) / 2 - box[1]), "CC", font=font, fill="#E2C275")
    img.save(os.path.join(ROOT, "app", "web", "icon.ico"), sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (256, 256)])


def main():
    if "--icon" in sys.argv:
        make_icon()
        print("icon written")
        return
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
    base = os.path.join(tempfile.gettempdir(), "ccs_demo_dev")
    data, appdata = os.path.join(base, "data"), os.path.join(base, "app")
    if not os.path.isdir(os.path.join(data, "tools")):
        sys.path.insert(0, os.path.join(ROOT, "installer"))
        import install_steps
        install_steps.copy_tools(ROOT, data)
    os.environ["CCS_DATA_ROOT"], os.environ["CCS_APP_ROOT"] = data, appdata
    from app.engine import Engine
    from app.main import run_fallback
    server, url = run_fallback(Engine(), open_browser=False, port=port)
    print("OPEN " + url, flush=True)
    print("sample inputs to browse to: " + os.path.join(data, "Sample inputs"), flush=True)
    while True:
        time.sleep(5)


if __name__ == "__main__":
    main()
