import os
import sys
from pathlib import Path

_PLUGIN_ROOT = Path(__file__).resolve().parent
if str(_PLUGIN_ROOT) not in sys.path:
    sys.path.insert(0, str(_PLUGIN_ROOT))

from .main import main

def main() -> int:
    from .main import QApplication, MainWindow
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    if os.environ.get("CDE_SMOKE_TEST", "").strip() == "1":
        from PySide6.QtCore import QTimer
        QTimer.singleShot(250, app.quit)
    return app.exec()

__all__ = ["main"]
