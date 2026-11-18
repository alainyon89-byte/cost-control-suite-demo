# (c) 2026 Alain Yon
"""Where things live on the PC (spec section 6).

  code  : %LOCALAPPDATA%\\CostControlSuiteDemo     (replaced on update)
  data  : Documents\\Cost Control Suite Demo       (never replaced)
Both can be overridden with CCS_APP_ROOT / CCS_DATA_ROOT (used by the tests).
The demo uses its own folders, so it never touches another install on the same PC.
"""
import ctypes
import os
import uuid

from . import branding

APP_VERSION = "1.0"
APP_NAME = "Cost Control Suite"
DATA_FOLDER = "Cost Control Suite Demo"
APP_FOLDER = "CostControlSuiteDemo"
HERE = os.path.dirname(os.path.abspath(__file__))


def _documents():
    try:
        class GUID(ctypes.Structure):
            _fields_ = [("d1", ctypes.c_ulong), ("d2", ctypes.c_ushort), ("d3", ctypes.c_ushort),
                        ("d4", ctypes.c_ubyte * 8)]
        u = uuid.UUID("{FDD39AD0-238F-46AF-ADB4-6C85480369C7}")  # FOLDERID_Documents
        g = GUID(u.time_low, u.time_mid, u.time_hi_version, (ctypes.c_ubyte * 8)(*u.bytes[8:]))
        out = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(g), 0, None, ctypes.byref(out)) == 0:
            path = out.value
            ctypes.windll.ole32.CoTaskMemFree(out)
            if path:
                return path
    except Exception:
        pass
    return os.path.join(os.environ.get("USERPROFILE", os.path.expanduser("~")), "Documents")


def default_data_root():
    return os.environ.get("CCS_DATA_ROOT") or os.path.join(_documents(), DATA_FOLDER)


def default_app_root():
    if os.environ.get("CCS_APP_ROOT"):
        return os.environ["CCS_APP_ROOT"]
    base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local")
    return os.path.join(base, APP_FOLDER)


def windows_user():
    """The demo records the configured display name (user_display_name), never the Windows login."""
    return branding.user_display_name()


def first_name():
    """First word of the configured display name, for the greeting on Home."""
    return (branding.user_display_name().split() or ["there"])[0]
