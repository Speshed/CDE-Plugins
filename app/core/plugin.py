from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class PluginSpec:
    """Metadata contract for one CDE integration.

    The launcher discovers ``plugins/*/plugin.py`` and never needs a hard-coded
    list of CDE systems.  Business code stays inside each plugin folder.
    """

    key: str
    title: str
    card_title: str
    card_subtitle: str
    image_light_name: str
    image_dark_name: str
    source_entry: Path
    frozen_entry: Path
    order: int = 100
    logo_max_width: int = 118
    logo_max_height: int = 86
    plugin_root: Path | None = None

    def with_root(self, root: Path) -> "PluginSpec":
        return PluginSpec(
            key=self.key,
            title=self.title,
            card_title=self.card_title,
            card_subtitle=self.card_subtitle,
            image_light_name=self.image_light_name,
            image_dark_name=self.image_dark_name,
            source_entry=self.source_entry,
            frozen_entry=self.frozen_entry,
            order=self.order,
            logo_max_width=self.logo_max_width,
            logo_max_height=self.logo_max_height,
            plugin_root=root,
        )

    @property
    def source_path(self) -> Path:
        if self.plugin_root is None:
            raise RuntimeError(f"Plugin root is not bound for {self.key}")
        return self.plugin_root / self.source_entry

    @property
    def frozen_path(self) -> Path:
        if self.plugin_root is None:
            raise RuntimeError(f"Plugin root is not bound for {self.key}")
        return self.plugin_root / self.frozen_entry
