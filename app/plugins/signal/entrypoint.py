import os
import sys
from pathlib import Path

_PLUGIN_ROOT = Path(__file__).resolve().parent
if str(_PLUGIN_ROOT) not in sys.path:
    sys.path.insert(0, str(_PLUGIN_ROOT))

def main() -> int:
    from .SIGNAL import MainWindow, QtWidgets, QtCore, WEBENGINE_AVAILABLE
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    if os.environ.get("CDE_SMOKE_TEST", "").strip() == "1":
        if not WEBENGINE_AVAILABLE or getattr(window, "auth_view", None) is None:
            return 1
        QtCore.QTimer.singleShot(500, app.quit)
    return app.exec()

__all__ = ["main"]
