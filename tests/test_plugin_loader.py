from pathlib import Path

import pytest

from core.plugin_loader import PluginLoadError, discover_plugins


_DESCRIPTOR = '''
from pathlib import Path
from core.plugin import PluginSpec

def get_plugin():
    return PluginSpec(
        key={key!r}, title="Test", card_title="Test", card_subtitle="",
        image_light_name="light.png", image_dark_name="dark.png",
        source_entry=Path("entry.py"), frozen_entry=Path("dist/app.exe"),
    )
'''


def _plugin(root: Path, name: str, key: str, *, entry: bool = False) -> None:
    folder = root / "plugins" / name
    folder.mkdir(parents=True)
    (folder / "plugin.py").write_text(_DESCRIPTOR.format(key=key), encoding="utf-8")
    if entry:
        (folder / "entry.py").write_text("# source entry\n", encoding="utf-8")


def test_source_discovery_requires_physical_entrypoint(tmp_path):
    _plugin(tmp_path, "larix", "larix")

    with pytest.raises(PluginLoadError, match="не найден ни исходный entry point"):
        discover_plugins(tmp_path, frozen=False)


def test_frozen_discovery_accepts_registered_internal_entrypoint(tmp_path):
    _plugin(tmp_path, "larix", "larix")

    plugins = discover_plugins(tmp_path, frozen=True)

    assert [item.key for item in plugins] == ["larix"]


def test_frozen_discovery_still_rejects_unknown_key_without_physical_entrypoint(tmp_path):
    _plugin(tmp_path, "unknown", "unknown")

    with pytest.raises(PluginLoadError, match="не найден ни исходный entry point"):
        discover_plugins(tmp_path, frozen=True)


def test_source_discovery_accepts_physical_entrypoint(tmp_path):
    _plugin(tmp_path, "larix", "larix", entry=True)

    plugins = discover_plugins(tmp_path, frozen=False)

    assert [item.key for item in plugins] == ["larix"]


def test_malformed_descriptor_remains_an_error_in_frozen_mode(tmp_path):
    folder = tmp_path / "plugins" / "broken"
    folder.mkdir(parents=True)
    (folder / "plugin.py").write_text("answer = 42\n", encoding="utf-8")

    with pytest.raises(PluginLoadError, match="нет функции get_plugin"):
        discover_plugins(tmp_path, frozen=True)


def test_duplicate_keys_remain_an_error_in_frozen_mode(tmp_path):
    _plugin(tmp_path, "a", "larix")
    _plugin(tmp_path, "b", "larix")

    with pytest.raises(PluginLoadError, match="повторяющийся key"):
        discover_plugins(tmp_path, frozen=True)
