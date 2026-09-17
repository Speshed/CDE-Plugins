import os
import sys
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "app/plugins/signal"))
import SIGNAL


@pytest.fixture(scope="module")
def qt_app():
    pytest.importorskip("PySide6")
    from PySide6 import QtWidgets

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    return app


def test_attribute_excel_controls_replace_manual_form_offscreen(tmp_path, qt_app, monkeypatch):
    monkeypatch.setattr(SIGNAL.QtWidgets.QMessageBox, "information", lambda *args, **kwargs: None)
    monkeypatch.setattr(SIGNAL.QtWidgets.QMessageBox, "warning", lambda *args, **kwargs: None)
    monkeypatch.setattr(SIGNAL.QtWidgets.QMessageBox, "critical", lambda *args, **kwargs: None)
    monkeypatch.setattr(SIGNAL.QtWidgets.QMessageBox, "question", lambda *args, **kwargs: SIGNAL.QtWidgets.QMessageBox.Yes)
    window = SIGNAL.MainWindow()
    try:
        assert not hasattr(window, "ed_attr_name")
        assert not hasattr(window, "cmb_attr_type")
        assert not hasattr(window, "btn_attr_add")
        assert not hasattr(window, "btn_attr_remove")
        assert window.btn_attr_preview.text() == "Предпросмотр"
        assert not window.btn_attr_preview.isEnabled()
        assert not window.btn_attr_create.isEnabled()
        assert window.tbl_attribute_queue.columnCount() == 8
        assert window.tbl_attribute_queue.horizontalHeaderItem(1).text() == "Строка"

        path = tmp_path / "attributes.xlsx"
        SIGNAL.build_excel_template("attributes", str(path))
        window.attribute_excel_path = str(path)
        window.lbl_attr_file.setText(path.name)
        window._invalidate_attribute_preview()
        assert window.btn_attr_preview.isEnabled()

        window.cmb_project.blockSignals(True)
        window.cmb_project.addItem("Demo", "project-id")
        window.cmb_project.setCurrentIndex(window.cmb_project.count() - 1)
        window.cmb_project.blockSignals(False)
        window.ed_project_id.setText("project-id")
        window.ed_token.setText("token")
        captured = []
        monkeypatch.setattr(window, "_run_worker", lambda worker: captured.append(worker))
        window._preview_attributes_excel()
        assert window.attribute_preview_result is not None, (window._current_project_id(), window.ed_token.text(), window.attribute_excel_path)
        assert not window.attribute_preview_result.errors
        assert window.tbl_attribute_queue.rowCount() == 4
        assert window.attribute_preview_pending
        assert not window.btn_attr_create.isEnabled()
        assert captured

        names = [draft["name"] for draft in window.attribute_preview_result.drafts]
        window._on_attribute_types_loaded({"attributes": [{"name": names[0], "deleted": False}]})
        assert len(window.attribute_drafts) == 3
        assert window.btn_attr_create.isEnabled()

        window._create_attributes()
        assert len(captured) == 2
        assert [draft["name"] for draft in captured[-1].drafts] == [draft["name"] for draft in window.attribute_drafts]
        window._set_busy(False)
        path.unlink()
        window._create_attributes()
        assert len(captured) == 2
        assert window.attribute_preview_result is None
        assert not window.btn_attr_create.isEnabled()

        SIGNAL.build_excel_template("attributes", str(path))
        window._preview_attributes_excel()
        window._on_attribute_types_loaded({"attributes": [{"name": names[0], "deleted": False}]})
        worker_count = len(captured)
        window.attribute_preview_project_id = "other-project"
        window._create_attributes()
        assert len(captured) == worker_count
        window.attribute_preview_project_id = "project-id"
        window._on_attributes_created({
            "attributes": [{"name": name, "deleted": False} for name in names],
            "stats": {"created_names": names, "created": len(names), "skipped": 0, "failed": 0},
        })
        assert not window.attribute_drafts
        assert not window.btn_attr_create.isEnabled()

        window.attribute_preview_result = SIGNAL.parse_attribute_excel(str(path))
        window.attribute_preview_pending = True
        window.attribute_preview_verified = False
        window.attribute_drafts = [dict(draft) for draft in window.attribute_preview_result.drafts]
        window._set_busy(True)
        assert not window.btn_attr_create.isEnabled()
        window._on_attribute_types_loaded({"attributes": [{"name": name, "deleted": False} for name in names]})
        assert not window.attribute_drafts
        assert not window.btn_attr_create.isEnabled()

        window.attribute_preview_result = None
        window.attribute_preview_pending = False
        window.attribute_preview_verified = False
        window.attribute_drafts = []
        window._set_busy(False)
        assert not window.btn_attr_create.isEnabled()
    finally:
        window.close()
        window.deleteLater()
        qt_app.processEvents()
