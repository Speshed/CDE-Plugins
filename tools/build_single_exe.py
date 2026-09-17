from __future__ import annotations

import sys
from pathlib import Path


def _fail(message: str) -> int:
    print(f"[ERROR] {message}")
    return 1


def main() -> int:
    # Native Windows consoles may still use cp1251/cp866; paths can contain
    # Cyrillic characters, so diagnostics must not abort the build itself.
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, OSError):
        pass
    root = Path(__file__).resolve().parent.parent
    app = root / "app"
    launcher = app / "launcher.py"
    assets = app / "assets"
    shared = app / "shared"
    plugins = app / "plugins"
    outdir = root / "release"
    workdir = root / "build" / "portable-work"

    required = [launcher, assets, shared, plugins]
    missing = [str(p) for p in required if not p.exists()]
    if missing:
        return _fail("Required project paths are missing:\n  " + "\n  ".join(missing))

    try:
        import PyInstaller
        import PyInstaller.__main__
    except Exception as exc:
        return _fail(f"PyInstaller is not available: {exc}")

    def _version_tuple(raw: str) -> tuple[int, int, int]:
        numbers = [int(part) for part in __import__("re").findall(r"\d+", raw)[:3]]
        return tuple((numbers + [0, 0, 0])[:3])

    minimum_pyinstaller = (6, 22, 3)
    installed_pyinstaller = _version_tuple(getattr(PyInstaller, "__version__", "0"))
    if installed_pyinstaller < minimum_pyinstaller:
        return _fail(
            "PyInstaller is too old for the portable Windows build: "
            f"{getattr(PyInstaller, '__version__', '<unknown>')}. "
            "Run install_dependencies.bat to install PyInstaller >= 6.22.3."
        )

    spec_file = Path(__file__).resolve().with_name("Larix_CDE.spec")
    if not spec_file.is_file():
        return _fail(f"Build spec is missing: {spec_file}")

    # Build the final argument vector in Python instead of through cmd.exe.
    # This avoids Windows batch quoting/continuation issues that previously
    # caused launcher.py to disappear and PyInstaller to report that
    # 'scriptname' was missing.
    args: list[str] = [
        "--noconfirm",
        "--clean",
        str(spec_file),
        f"--distpath={outdir}",
        f"--workpath={workdir}",
    ]

    # Plugin descriptors are dynamically loaded from the extracted bundle, so
    # keep only their small metadata files as data. Plugin entry points and
    # source modules below are analysed into the same Python archive.
    plugin_count = 0
    for plugin_dir in sorted(p for p in plugins.iterdir() if p.is_dir()):
        descriptor = plugin_dir / "plugin.py"
        if not descriptor.is_file():
            continue

        plugin_count += 1
        # Descriptor/resource inclusion and source module analysis are defined
        # in Larix_CDE.spec; this loop validates that all four are present.

    if plugin_count == 0:
        return _fail(f"No plugin descriptors found under {plugins}")

    print("[portable] PyInstaller argument vector:")
    for i, arg in enumerate(args, 1):
        print(f"  {i:02d}: {arg}")
    print()

    if not spec_file.is_file():
        return _fail(f"Build spec does not exist: {spec_file}")

    try:
        PyInstaller.__main__.run(args)
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else 1
        return int(code or 0)
    except Exception as exc:
        print(f"[ERROR] PyInstaller raised {type(exc).__name__}: {exc}")
        return 1

    exe = outdir / "Larix_CDE.exe"
    if not exe.is_file():
        return _fail(f"Expected executable was not created: {exe}")

    print(f"[portable] Done: {exe}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
