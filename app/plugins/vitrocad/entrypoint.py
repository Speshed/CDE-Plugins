import os
import sys
from pathlib import Path

_PLUGIN_ROOT = Path(__file__).resolve().parent
if str(_PLUGIN_ROOT) not in sys.path:
    sys.path.insert(0, str(_PLUGIN_ROOT))

from .VitroCAD import main

def main() -> int:
    from .VitroCAD import QtCore, QtGui, QtWidgets, ToolShell, APP_TITLE
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setStyle("Fusion")
    window = ToolShell()
    window.show()
    if os.environ.get("CDE_SMOKE_TEST", "").strip() == "1":
        QtCore.QTimer.singleShot(250, app.quit)
    return app.exec()

__all__ = ["main"]
