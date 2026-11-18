# (c) 2026 Alain Yon
"""Demo branding, read from config/app_settings.json (next to the app folder).

The property name is the only place the property is named. The app shows it in the
sidebar, and passes it to every tool script as CCS_PROPERTY_NAME so report titles
use it too. Plant codes (D100, D200) are the demo's SAP plants.
"""
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULTS = {
    "property_name": "Demo Hotel",
    "currency": "ZAR",
    "user_display_name": "Demo User",
    "units": {"main": "Main Hotel", "tower": "Tower Wing"},
    "plants": {"D100": "Hotel", "D200": "Conference Centre"},
}
HOTEL_PLANT, CC_PLANT = "D100", "D200"


def settings_path():
    return os.environ.get("CCS_SETTINGS_FILE") or os.path.join(os.path.dirname(HERE), "config", "app_settings.json")


def load():
    out = json.loads(json.dumps(DEFAULTS))
    try:
        with open(settings_path(), encoding="utf-8") as f:
            data = json.load(f)
        for k, v in data.items():
            if k.startswith("_"):
                continue
            if isinstance(v, dict) and isinstance(out.get(k), dict):
                out[k].update(v)
            else:
                out[k] = v
    except (OSError, ValueError):
        pass
    return out


SETTINGS = load()


def property_name():
    return str(SETTINGS.get("property_name") or DEFAULTS["property_name"])


def user_display_name():
    return str(SETTINGS.get("user_display_name") or DEFAULTS["user_display_name"])


def plant_name(code):
    return SETTINGS["plants"].get(code, code)


def unit_name(key):
    return SETTINGS["units"].get(key, key)
