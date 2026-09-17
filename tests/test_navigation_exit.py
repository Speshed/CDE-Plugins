from shared.theme_core import RETURN_TO_MANAGER_CODE, classify_plugin_exit


def test_back_button_closes_window_and_returns_reserved_code(monkeypatch):
    from PySide6 import QtCore, QtWidgets
    from shared.ui_components import BackToManagerButton

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = QtWidgets.QWidget()
    closed = []

    def close_event(event):
        closed.append(True)
        event.accept()

    window.closeEvent = close_event
    button = BackToManagerButton(window)
    button.show()
    QtCore.QTimer.singleShot(0, button.click)
    QtCore.QTimer.singleShot(1000, app.quit)
    result = app.exec()
    assert closed
    assert result == RETURN_TO_MANAGER_CODE
    window.deleteLater()


def test_return_code_is_the_only_menu_return_signal() -> None:
    assert classify_plugin_exit(RETURN_TO_MANAGER_CODE) == "return_to_menu"
    assert classify_plugin_exit(0) == "closed"
    assert classify_plugin_exit(None) == "closed"
    assert classify_plugin_exit(1) == "crashed"
