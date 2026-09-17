from __future__ import annotations

import ast
from pathlib import Path
from typing import Iterable

MAX_EXE_BYTES = 500 * 1024 * 1024
EXPECTED_EXE = "Larix_CDE.exe"
NESTED_EXECUTABLES = {"larix_platform_plugin.exe", "vitrocad.exe", "projectpoint.exe", "sgnl_platform.exe"}


def _toc_strings(value: object) -> Iterable[str]:
    if isinstance(value, str):
        yield value
    elif isinstance(value, (tuple, list)):
        for item in value:
            yield from _toc_strings(item)
    elif isinstance(value, dict):
        for item in value.items():
            yield from _toc_strings(item)


def _toc_names(value: object) -> Iterable[str]:
    if isinstance(value, tuple) and len(value) >= 3 and isinstance(value[0], str) and isinstance(value[2], str):
        yield value[0]
        return
    if isinstance(value, (tuple, list)):
        for item in value:
            yield from _toc_names(item)


def _read_toc(path: Path) -> list[str]:
    try:
        return list(_toc_names(ast.literal_eval(path.read_text(encoding="utf-8"))))
    except (OSError, SyntaxError, ValueError) as exc:
        raise ValueError(f"Не удалось прочитать TOC {path}: {exc}") from exc


def verify_distribution(portable_dir: Path, build_dir: Path, *, max_bytes: int = MAX_EXE_BYTES) -> list[str]:
    """Return all distribution errors; an empty list means the build is valid."""
    errors: list[str] = []
    portable_dir, build_dir = Path(portable_dir), Path(build_dir)
    exe = portable_dir / EXPECTED_EXE
    if not exe.is_file():
        errors.append(f"Не найден обязательный файл {exe}")
    elif exe.stat().st_size > max_bytes:
        errors.append(f"Размер {exe.name} превышает лимит: {exe.stat().st_size / (1024 * 1024):.1f} MiB > {max_bytes / (1024 * 1024):.1f} MiB")
    if portable_dir.is_dir():
        files = sorted(item.name for item in portable_dir.iterdir())
        if files != [EXPECTED_EXE]:
            errors.append(f"release должен содержать только Larix_CDE.exe; найдено: {', '.join(files) or '<пусто>'}")

    toc_path = build_dir / "Larix_CDE" / "PKG-00.toc"
    if not toc_path.is_file():
        errors.append(f"Не найден финальный PyInstaller TOC: {toc_path}")
        return errors
    try:
        names = [name.replace("\\", "/").lower() for name in _read_toc(toc_path)]
    except ValueError as exc:
        errors.append(str(exc))
        return errors
    nested = sorted({name for name in names if name.rsplit("/", 1)[-1] in NESTED_EXECUTABLES})
    if nested:
        errors.append("В TOC найдены вложенные EXE инструментов: " + ", ".join(nested))
    locales = sorted(set(name for name in names if "/qtwebengine_locales/" in f"/{name}"))
    if locales != ["pyside6/translations/qtwebengine_locales/ru.pak"]:
        errors.append("В TOC должен присутствовать только qtwebengine_locales/ru.pak")
    if any("debug.pak" in name or "snapshot.debug" in name for name in names):
        errors.append("В TOC обнаружены debug-ресурсы WebEngine")
    for resource in ("pyside6/qtwebengineprocess.exe", "pyside6/resources/icudtl.dat", "pyside6/resources/qtwebengine_resources.pak"):
        if resource not in names:
            errors.append(f"В TOC отсутствует обязательный WebEngine ресурс: {resource}")

    required_runtime_entries = {
        "assets/icon.ico",
        "assets/right-arrow.png",
        "plugins/larix/plugin.py",
        "plugins/vitrocad/plugin.py",
        "plugins/signal/plugin.py",
        "plugins/projectpoint/plugin.py",
        "plugins/projectpoint/projectpoint/resources/builtin_role_permissions.json",
    }
    missing_runtime = sorted(entry for entry in required_runtime_entries if entry not in names)
    if missing_runtime:
        errors.append("В TOC отсутствуют runtime-ресурсы: " + ", ".join(missing_runtime))
    return errors


def main() -> int:
    root = Path(__file__).resolve().parent.parent
    errors = verify_distribution(root / "release", root / "build" / "portable-work")
    if errors:
        for error in errors:
            print(f"[ERROR] {error}")
        return 1
    exe = root / "release" / EXPECTED_EXE
    print("Release build: OK")
    print(f"File: {exe}")
    print(f"Size: {exe.stat().st_size / (1024 * 1024):.1f} MiB")
    print("Distribution: send only Larix_CDE.exe")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
