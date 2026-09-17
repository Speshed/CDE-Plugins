from pathlib import Path
from core.plugin import PluginSpec


def get_plugin() -> PluginSpec:
    return PluginSpec(
        key="vitrocad",
        title="Vitro",
        card_title="Vitro",
        card_subtitle="",
        image_light_name="sod_vitrocad_light.png",
        image_dark_name="sod_vitrocad_dark.png",
        source_entry=Path("VitroCAD.py"),
        frozen_entry=Path("dist/VitroCAD.exe"),
        order=20,
        logo_max_width=126,
        logo_max_height=86,
    )
