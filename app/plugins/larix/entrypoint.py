from __future__ import annotations

import sys
import os
from pathlib import Path

_PLUGIN_ROOT = Path(__file__).resolve().parent
if str(_PLUGIN_ROOT) not in sys.path:
    sys.path.insert(0, str(_PLUGIN_ROOT))

from PySide6 import QtCore, QtGui, QtWidgets

from .Larix_User_Platform import MainWindow


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QtGui.QFont("Segoe UI", 9))
    window = MainWindow()
    window.show()
    if os.environ.get("CDE_SMOKE_TEST", "").strip() == "1":
        QtCore.QTimer.singleShot(250, app.quit)
    return app.exec()
