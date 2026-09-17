from __future__ import annotations

from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parent.parent
APP_ROOT = PROJECT_ROOT / "app"
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from core.plugin_loader import discover_plugins


def main() -> int:
    root = APP_ROOT
    errors: list[str] = []

    try:
        plugins = discover_plugins(root)
    except Exception as exc:
        print(f"[ERROR] {exc}")
        return 1

    assets = root / "assets"
    required_common = [
        "icon.ico",
        "brand_larix_cde_light.png",
        "brand_larix_cde_dark.png",
        "sun.png",
        "moon.png",
        "exit.png",
        "right-arrow.png",
    ]
    for name in required_common:
        if not (assets / name).is_file():
            errors.append(f"assets/{name} missing")

    for plugin in plugins:
        for name in {plugin.image_light_name, plugin.image_dark_name}:
            if not (assets / name).is_file():
                errors.append(f"{plugin.key}: assets/{name} missing")
        if not plugin.source_path.is_file():
            errors.append(f"{plugin.key}: source entry missing: {plugin.source_path}")

    print(f"Plugins found: {len(plugins)}")
    for plugin in plugins:
        print(f"  {plugin.order:02d}  {plugin.key:12s} -> {plugin.source_entry}")

    if errors:
        print("\nValidation errors:")
        for error in errors:
            print(f"  - {error}")
        return 1

    print("\nProject structure: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
