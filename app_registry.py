from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Tuple


@dataclass(frozen=True)
class AppSpec:
    key: str
    title: str
    card_title: str
    card_subtitle: str
    image_light_name: str
    image_dark_name: str
    source_entry: Path
    frozen_entry: Path
    logo_max_width: int = 118
    logo_max_height: int = 86


APPS: Tuple[AppSpec, ...] = (
    AppSpec(
        key="larix",
        title="Larix Platform",
        card_title="Larix\nPlatform",
        card_subtitle="",
        image_light_name="sod_larix_light.png",
        image_dark_name="sod_larix_dark.png",
        source_entry=Path("apps/larix/Larix_User_Platform.py"),
        frozen_entry=Path("apps/larix/Larix_Platform_Plugin.exe"),
        logo_max_width=120,
        logo_max_height=92,
    ),
    AppSpec(
        key="vitrocad",
        title="Vitro",
        card_title="Vitro",
        card_subtitle="",
        image_light_name="sod_vitrocad_light.png",
        image_dark_name="sod_vitrocad_dark.png",
        source_entry=Path("apps/vitrocad/VitroCAD.py"),
        frozen_entry=Path("apps/vitrocad/VitroCAD.exe"),
        logo_max_width=126,
        logo_max_height=86,
    ),
    AppSpec(
        key="signal",
        title="Docs",
        card_title="Docs",
        card_subtitle="",
        image_light_name="sod_signal.png",
        image_dark_name="sod_signal.png",
        source_entry=Path("apps/signal/SIGNAL.py"),
        frozen_entry=Path("apps/signal/SGNL_Platform/SGNL_Platform.exe"),
        logo_max_width=126,
        logo_max_height=68,
    ),
    AppSpec(
        key="projectpoint",
        title="Project Point",
        card_title="Project Point",
        card_subtitle="Единая среда участников\nпроектов строительства",
        image_light_name="sod_projectpoint_light.png",
        image_dark_name="sod_projectpoint_dark.png",
        source_entry=Path("apps/projectpoint/main.py"),
        frozen_entry=Path("apps/projectpoint/ProjectPoint.exe"),
        logo_max_width=92,
        logo_max_height=84,
    ),
)
