from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from projectpoint.ui import MainWindow


def main() -> int:
    """Start the ProjectPoint desktop application."""
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
