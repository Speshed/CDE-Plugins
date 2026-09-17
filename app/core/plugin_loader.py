from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType
from typing import Iterable

from core.plugin import PluginSpec


class PluginLoadError(RuntimeError):
    pass


def _load_module(path: Path, module_name: str) -> ModuleType:
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise PluginLoadError(f"Не удалось загрузить описание плагина: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def discover_plugins(root: Path, *, frozen: bool | None = None) -> tuple[PluginSpec, ...]:
    """Discover all ``plugins/<name>/plugin.py`` descriptors.

    A broken plugin does not silently disappear: startup fails with a readable
    diagnostic so an incomplete installation is noticed immediately.
    """
    if frozen is None:
        frozen = bool(getattr(sys, "frozen", False) or getattr(sys, "_MEIPASS", None))

    plugins_dir = root / "plugins"
    if not plugins_dir.is_dir():
        raise PluginLoadError(f"Папка plugins не найдена: {plugins_dir}")

    found: list[PluginSpec] = []
    errors: list[str] = []
    keys: set[str] = set()

    for descriptor in sorted(plugins_dir.glob("*/plugin.py")):
        plugin_root = descriptor.parent
        module_name = f"larix_cde_plugin_{plugin_root.name}"
        try:
            module = _load_module(descriptor, module_name)
            factory = getattr(module, "get_plugin", None)
            if not callable(factory):
                raise PluginLoadError("нет функции get_plugin()")
            item = factory()
            if not isinstance(item, PluginSpec):
                raise PluginLoadError("get_plugin() должен вернуть PluginSpec")
            if item.key in keys:
                raise PluginLoadError(f"повторяющийся key: {item.key}")
            item = item.with_root(plugin_root)
            physical_entry_exists = item.source_path.is_file() or item.frozen_path.is_file()
            if not physical_entry_exists and not (
                frozen and _has_internal_entry_point(item.key)
            ):
                raise PluginLoadError(
                    f"не найден ни исходный entry point ({item.source_entry}), "
                    f"ни собранный ({item.frozen_entry})"
                )
            keys.add(item.key)
            found.append(item)
        except Exception as exc:
            errors.append(f"{plugin_root.name}: {exc}")

    if errors:
        raise PluginLoadError("Ошибки загрузки плагинов:\n" + "\n".join(f"• {e}" for e in errors))
    if not found:
        raise PluginLoadError("В папке plugins не найдено ни одного плагина")

    found.sort(key=lambda p: (p.order, p.title.casefold()))
    return tuple(found)


def _has_internal_entry_point(key: str) -> bool:
    """Return whether the one-file dispatcher contains a bundled plugin key."""
    # Import lazily: plugin_dispatcher imports no loader code, while keeping
    # discovery usable in source-only tooling that does not need dispatching.
    from core.plugin_dispatcher import PLUGIN_ENTRY_POINTS

    return key.strip().casefold() in PLUGIN_ENTRY_POINTS
