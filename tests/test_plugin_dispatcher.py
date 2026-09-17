from __future__ import annotations

import pytest

from core.plugin_dispatcher import (
    PluginDispatchError,
    PluginEntryPoint,
    UnknownPluginError,
    clean_plugin_argv,
    dispatch_plugin,
    parse_plugin_key,
)


def test_parse_menu_mode_without_arguments() -> None:
    assert parse_plugin_key([]) is None


def test_parse_plugin_key_is_case_insensitive() -> None:
    assert parse_plugin_key(["--plugin", "SIGNAL"]) == "signal"


def test_clean_plugin_argv_removes_internal_switch() -> None:
    assert clean_plugin_argv(["Larix_CDE.exe", "--plugin", "signal"]) == ["Larix_CDE.exe"]


@pytest.mark.parametrize("args", [["--plugin"], ["--plugin", "signal", "extra"], ["--other"]])
def test_parse_rejects_malformed_internal_arguments(args: list[str]) -> None:
    with pytest.raises(PluginDispatchError):
        parse_plugin_key(args)


def test_dispatch_imports_only_selected_entry_point() -> None:
    calls: list[str] = []

    class FakeModule:
        def main(self) -> int:
            calls.append("main")
            return 7

    def importer(name: str) -> FakeModule:
        calls.append(name)
        return FakeModule()

    result = dispatch_plugin(
        "VitroCAD",
        registry={"vitrocad": PluginEntryPoint("fake.vitrocad.entrypoint")},
        importer=importer,
    )
    assert result == 7
    assert calls == ["fake.vitrocad.entrypoint", "main"]


def test_dispatch_reports_unknown_key() -> None:
    with pytest.raises(UnknownPluginError, match="Неизвестный ключ плагина"):
        dispatch_plugin("missing", registry={"signal": PluginEntryPoint("fake.signal")})
