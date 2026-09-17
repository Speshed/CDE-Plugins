from __future__ import annotations

import sys
from pathlib import Path


def application_root() -> Path:
    """Return the runtime root containing assets/plugins/shared.

    Source mode:
        repository root.

    Normal frozen launcher with external release folder:
        folder containing Larix_CDE.exe.

    Portable one-file launcher:
        PyInstaller's temporary extraction directory (sys._MEIPASS), where
        assets, plugin descriptors and child executables are bundled.
    """
    if getattr(sys, "frozen", False):
        bundle_root = getattr(sys, "_MEIPASS", None)
        if bundle_root:
            candidate = Path(bundle_root).resolve()
            if (candidate / "plugins").is_dir() and (candidate / "assets").is_dir():
                return candidate
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def assets_dir() -> Path:
    return application_root() / "assets"


def plugins_dir() -> Path:
    return application_root() / "plugins"
