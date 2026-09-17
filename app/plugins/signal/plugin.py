from pathlib import Path
from core.plugin import PluginSpec


def get_plugin() -> PluginSpec:
    return PluginSpec(
        key="signal",
        title="SIGNAL Docs",
        card_title="Docs",
        card_subtitle="",
        image_light_name="sod_signal.png",
        image_dark_name="sod_signal.png",
        source_entry=Path("SIGNAL.py"),
        frozen_entry=Path("dist/SGNL_Platform/SGNL_Platform.exe"),
        order=30,
        logo_max_width=126,
        logo_max_height=68,
    )
