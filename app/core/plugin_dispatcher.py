"""Internal plugin mode dispatcher for the single-executable launcher.

The dispatcher deliberately imports only the selected plugin.  This keeps the
normal menu startup light and gives PyInstaller a single, explicit list of
entry points to analyse.
"""
from __future__ import annotations

import importlib
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Mapping, Sequence


INTERNAL_PLUGIN_ARGUMENT = "--plugin"


class PluginDispatchError(RuntimeError):
    """A user-facing error in an internal plugin invocation."""


class UnknownPluginError(PluginDispatchError):
    pass


@dataclass(frozen=True)
class PluginEntryPoint:
    module: str
    function: str = "main"
    source_root: str | None = None


PLUGIN_ENTRY_POINTS: Mapping[str, PluginEntryPoint] = {
    "larix": PluginEntryPoint("plugins.larix.entrypoint"),
    "vitrocad": PluginEntryPoint("plugins.vitrocad.entrypoint"),
    "signal": PluginEntryPoint("plugins.signal.entrypoint"),
    "projectpoint": PluginEntryPoint("plugins.projectpoint.entrypoint"),
}


def parse_plugin_key(
    argv: Sequence[str],
    *,
    argument: str = INTERNAL_PLUGIN_ARGUMENT,
) -> str | None:
    """Return the internal plugin key, or ``None`` for ordinary menu mode.

    Only the exact two-argument form is accepted.  A malformed internal
    invocation is rejected instead of accidentally opening the main menu.
    """
    args = list(argv)
    if not args:
        return None
    if args[0] != argument:
        raise PluginDispatchError(
            f"Неизвестный служебный аргумент: {args[0]}. "
            f"Ожидался запуск без аргументов или {argument} <ключ>."
        )
    if len(args) != 2 or not args[1].strip():
        raise PluginDispatchError(f"Формат запуска: {argument} <ключ>")
    return args[1].strip().casefold()


def clean_plugin_argv(argv: Sequence[str]) -> list[str]:
    """Remove the dispatch switch before a Qt application is created."""
    args = list(argv)
    if len(args) >= 3 and args[1] == INTERNAL_PLUGIN_ARGUMENT:
        return [args[0], *args[3:]]
    return args


def dispatch_plugin(
    key: str,
    *,
    registry: Mapping[str, PluginEntryPoint] = PLUGIN_ENTRY_POINTS,
    importer: Callable[[str], object] = importlib.import_module,
) -> int:
    """Lazily import and execute the selected plugin entry point."""
    normalized = key.strip().casefold()
    entry = registry.get(normalized)
    if entry is None:
        allowed = ", ".join(registry)
        raise UnknownPluginError(
            f"Неизвестный ключ плагина: {key!r}. Допустимые ключи: {allowed}."
        )

    if entry.source_root:
        root = Path(entry.source_root).resolve()
        if str(root) not in sys.path:
            sys.path.insert(0, str(root))

    try:
        module = importer(entry.module)
        callback = getattr(module, entry.function)
    except (ImportError, AttributeError) as exc:
        raise PluginDispatchError(
            f"Не удалось загрузить плагин {normalized!r}: {exc}"
        ) from exc
    if not callable(callback):
        raise PluginDispatchError(
            f"Entry point плагина {normalized!r} не является вызываемой функцией"
        )
    result = callback()
    return int(result or 0)
