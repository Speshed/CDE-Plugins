from pathlib import Path
import sys

import pytest


ROOT = Path(__file__).resolve().parents[1]


def test_preview_asset_exists_and_is_used_for_preview_actions():
    assert (ROOT / "app/assets" / "preview.png").is_file()

    sources = {
        "projectpoint": ROOT / "app/plugins/projectpoint/projectpoint/ui/main_window.py",
        "signal": ROOT / "app/plugins/signal/SIGNAL.py",
        "larix": ROOT / "app/plugins/larix/Larix_User_Platform.py",
        "larix_comments": ROOT / "app/plugins/larix/coll/larix_comment_importer.py",
        "vitrocad": ROOT / "app/plugins/vitrocad/VitroCAD.py",
    }
    for name, path in sources.items():
        text = path.read_text(encoding="utf-8")
        assert "preview.png" in text, name

    projectpoint = sources["projectpoint"].read_text(encoding="utf-8")
    assert 'asset = "preview.png"' in projectpoint
    assert 'asset = "comparison.png"' not in projectpoint

    comments = sources["larix_comments"].read_text(encoding="utf-8")
    assert 'self.btn_open_preview.setIcon(qicon("preview.png"))' in comments


@pytest.mark.parametrize(
    "module_name,class_name",
    [
        ("folder_creator", "ModernSection"),
        ("permissions", "ModernSection"),
        ("schedule_sync", "MainWindow"),
    ],
)
def test_vitrocad_preview_widgets_construct_offscreen(module_name, class_name, monkeypatch):
    monkeypatch.setenv("QT_QPA_PLATFORM", "offscreen")
    from PySide6 import QtWidgets

    modules_dir = str(ROOT / "app/plugins/vitrocad/modules")
    if modules_dir not in sys.path:
        sys.path.insert(0, modules_dir)
    module = __import__(module_name)
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    if module_name == "schedule_sync":
        # The schedule window builds its preview card during normal setup.
        widget = getattr(module, class_name)()
        badge = next(label for label in widget.findChildren(QtWidgets.QLabel)
                     if label.property("iconAsset") == "preview.png")
    else:
        widget = getattr(module, class_name)("preview.png", "Предпросмотр", collapsible=False)
        badge = widget.findChild(QtWidgets.QLabel, "sectionIcon")

    assert badge is not None
    assert badge.property("iconAsset") == "preview.png"
    assert not badge.pixmap().isNull()
    widget.deleteLater()
    app.processEvents()
