from pathlib import Path
from core.plugin import PluginSpec


def get_plugin() -> PluginSpec:
    return PluginSpec(
        key="larix",
        title="Larix Platform",
        card_title="Larix\nPlatform",
        card_subtitle="",
        image_light_name="sod_larix_light.png",
        image_dark_name="sod_larix_dark.png",
        source_entry=Path("Larix_User_Platform.py"),
        frozen_entry=Path("dist/Larix_Platform_Plugin.exe"),
        order=10,
        logo_max_width=120,
        logo_max_height=92,
    )
