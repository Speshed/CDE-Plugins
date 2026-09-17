# Reproducible single-file build.
# Keep source code in the Python archive and place only true runtime resources
# into the extracted bundle. This avoids accidentally shipping raw source files
# while making newly-added JSON/XLSX/image resources portable automatically.
from pathlib import Path

from PyInstaller.building.build_main import Analysis, PYZ, EXE
from PyInstaller.utils.hooks import collect_data_files
import PySide6

ROOT = Path(SPECPATH).resolve().parent / "app"

PLUGIN_NAMES = ("larix", "vitrocad", "signal", "projectpoint")
RUNTIME_SUFFIXES = {
    ".xlsx", ".xlsm", ".json",
    ".png", ".jpg", ".jpeg", ".ico", ".svg",
    ".html", ".htm", ".css", ".js",
    ".pak", ".dat",
}
SKIP_DIR_NAMES = {"__pycache__", "tests", "docs", "dist", "build"}
# User-editable/cache files must not be frozen as defaults. The application may
# create them beside the executable or in the user's profile at run time.
SKIP_FILE_NAMES = {"connection_profiles.json"}


def add_runtime_tree(datas: list[tuple[str, str]], root: Path) -> None:
    if not root.is_dir():
        return
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        relative = path.relative_to(ROOT)
        if any(part.lower() in SKIP_DIR_NAMES for part in relative.parts[:-1]):
            continue
        if path.name.lower() in SKIP_FILE_NAMES:
            continue
        if path.suffix.lower() not in RUNTIME_SUFFIXES:
            continue
        destination = relative.parent.as_posix()
        datas.append((str(path), destination))


# Shared visual assets are a real data directory and are copied recursively.
datas = [(str(ROOT / "assets"), "assets")]

# plugin.py descriptors must be physical files because the launcher discovers
# them dynamically. Other Python modules stay compiled in PyInstaller's PYZ.
for plugin_name in PLUGIN_NAMES:
    plugin = ROOT / "plugins" / plugin_name
    descriptor = plugin / "plugin.py"
    if descriptor.is_file():
        datas.append((str(descriptor), f"plugins/{plugin_name}"))
    add_runtime_tree(datas, plugin)

# Keep Qt's release data, but retain only the Russian Chromium locale pack and
# drop debug-only WebEngine payloads. PyInstaller's standard PySide6 hooks still
# collect the matching Qt binaries/plugins.
for source, destination in collect_data_files("PySide6"):
    normalized = destination.replace("\\", "/")
    name = Path(normalized).name
    if "qtwebengine_locales/" in normalized and name != "ru.pak":
        continue
    if name.endswith(".debug.pak") or "snapshot.debug" in name:
        continue
    datas.append((source, destination))

hiddenimports = [
    "plugins.larix.entrypoint",
    "plugins.vitrocad.entrypoint",
    "plugins.signal.entrypoint",
    "plugins.projectpoint.entrypoint",
]
pathex = [str(ROOT)] + [str(ROOT / "plugins" / name) for name in PLUGIN_NAMES]

a = Analysis(
    [str(ROOT / "launcher.py")],
    pathex=pathex,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)


def keep_runtime_data(item: tuple[str, str, str]) -> bool:
    destination = item[0].replace("\\", "/").lower()
    if "/qtwebengine_locales/" in f"/{destination}" and not destination.endswith("/ru.pak"):
        return False
    if "debug.pak" in destination or "snapshot.debug" in destination:
        return False
    return True


# Hooks may append WebEngine resources after our explicit data collection, so
# filter the final Analysis TOCs as well.
a.datas[:] = [item for item in a.datas if keep_runtime_data(item)]
a.binaries[:] = [item for item in a.binaries if keep_runtime_data(item)]

ru_locale = Path(PySide6.__file__).resolve().parent / "translations/qtwebengine_locales/ru.pak"
if ru_locale.is_file():
    wanted = "PySide6/translations/qtwebengine_locales/ru.pak"
    if not any(item[0].replace("\\", "/") == wanted for item in a.datas):
        a.datas.append((wanted, str(ru_locale), "DATA"))

pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Larix_CDE",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPX makes little practical difference for this Qt/WebEngine bundle and
    # can increase false positives/quarantine events on another PC.
    upx=False,
    console=False,
    icon=str(ROOT / "assets/icon.ico"),
)
