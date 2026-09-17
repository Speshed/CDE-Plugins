from pathlib import Path
from core.plugin import PluginSpec


def get_plugin() -> PluginSpec:
    return PluginSpec(
        key="projectpoint",
        title="Project Point",
        card_title="Project Point",
        card_subtitle="Единая среда участников\nпроектов строительства",
        image_light_name="sod_projectpoint_light.png",
        image_dark_name="sod_projectpoint_dark.png",
        source_entry=Path("main.py"),
        frozen_entry=Path("dist/ProjectPoint.exe"),
        order=40,
        logo_max_width=92,
        logo_max_height=84,
    )
