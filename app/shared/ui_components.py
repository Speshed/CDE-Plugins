from __future__ import annotations

from pathlib import Path
from typing import Optional

try:
    from PySide6 import QtCore, QtGui, QtWidgets
except ImportError:  # pragma: no cover - compatibility with older SGNL installs
    from PyQt5 import QtCore, QtGui, QtWidgets

from shared.theme_core import (
    apply_windows_titlebar_theme, initial_dark_theme, shared_asset_dir, themed_icon
)




class NoWheelTabBar(QtWidgets.QTabBar):
    """Tab bar that ignores the mouse wheel so scrolling never changes sections."""

    def wheelEvent(self, event: QtGui.QWheelEvent) -> None:
        event.accept()



class PreviewHeader(QtWidgets.QWidget):
    """Shared preview-card heading with the global preview icon."""

    def __init__(self, text: str, parent: Optional[QtWidgets.QWidget] = None, icon_size: int = 18):
        super().__init__(parent)
        self._asset_dir = shared_asset_dir(Path(__file__))
        self._icon_size = int(icon_size)
        self._dark_cache = None
        self.setObjectName("previewHeader")
        self.setAttribute(QtCore.Qt.WA_StyledBackground, False)

        row = QtWidgets.QHBoxLayout(self)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(8)

        self.icon_label = QtWidgets.QLabel()
        self.icon_label.setObjectName("inlineIcon")
        self.icon_label.setFixedSize(self._icon_size + 4, self._icon_size + 4)
        self.icon_label.setAlignment(QtCore.Qt.AlignCenter)
        row.addWidget(self.icon_label, 0, QtCore.Qt.AlignVCenter)

        self.title_label = QtWidgets.QLabel(text)
        self.title_label.setObjectName("cardTitle")
        row.addWidget(self.title_label, 0, QtCore.Qt.AlignVCenter)
        row.addStretch(1)
        self._refresh_icon(force=True)

    def _refresh_icon(self, force: bool = False) -> None:
        dark = bool(initial_dark_theme(False))
        if not force and dark == self._dark_cache:
            return
        self._dark_cache = dark
        icon = themed_icon(self._asset_dir, "preview.png", dark)
        self.icon_label.setPixmap(icon.pixmap(QtCore.QSize(self._icon_size, self._icon_size)))

    def paintEvent(self, event) -> None:
        self._refresh_icon()
        super().paintEvent(event)


def make_preview_header(text: str, parent: Optional[QtWidgets.QWidget] = None, icon_size: int = 18) -> PreviewHeader:
    return PreviewHeader(text, parent, icon_size)

def configure_preview_table(
    table: QtWidgets.QTableWidget,
    widths=None,
    *,
    min_height: int = 118,
    max_height: int = 225,
    row_height: int = 27,
) -> QtWidgets.QTableWidget:
    """Apply the compact Larix-style preview-table behaviour."""
    table.setObjectName("matrixTable")
    table.verticalHeader().setVisible(False)
    table.verticalHeader().setMinimumSectionSize(max(22, int(row_height) - 2))
    table.verticalHeader().setDefaultSectionSize(int(row_height))
    table.horizontalHeader().setMinimumHeight(30)
    table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
    table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
    table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
    table.setAlternatingRowColors(True)
    table.setWordWrap(False)
    table.setShowGrid(False)
    table.setHorizontalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
    table.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
    table.setMinimumHeight(int(min_height))
    table.setMaximumHeight(int(max_height))
    table.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
    if widths:
        for index, width in enumerate(list(widths)[: table.columnCount()]):
            table.setColumnWidth(index, int(width))
    return table


def fit_preview_height(
    table: QtWidgets.QTableWidget,
    *,
    visible_rows: int = 6,
    min_height: int = 118,
    max_height: int = 225,
) -> None:
    """Resize a compact preview without allowing it to take over the whole page."""
    count = min(max(int(table.rowCount()), 0), max(1, int(visible_rows)))
    header_h = table.horizontalHeader().height() or 30
    row_h = table.verticalHeader().defaultSectionSize() or 27
    target = header_h + max(2, count) * row_h + 10
    table.setMinimumHeight(max(int(min_height), min(target, int(max_height))))
    table.setMaximumHeight(int(max_height))


def _copy_table(source: QtWidgets.QTableWidget) -> QtWidgets.QTableWidget:
    rows = source.rowCount()
    columns = source.columnCount()
    table = QtWidgets.QTableWidget(rows, columns)
    table.setObjectName("matrixTable")
    table.verticalHeader().setVisible(False)
    table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
    table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
    table.setAlternatingRowColors(True)
    table.setWordWrap(False)
    table.setShowGrid(False)
    table.setHorizontalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
    table.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)

    headers = []
    for column in range(columns):
        header = source.horizontalHeaderItem(column)
        headers.append(header.text() if header is not None else str(column + 1))
    if headers:
        table.setHorizontalHeaderLabels(headers)

    for row in range(rows):
        for column in range(columns):
            item = source.item(row, column)
            if item is None:
                continue
            copy_item = item.clone()
            table.setItem(row, column, copy_item)

    header = table.horizontalHeader()
    header.setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
    header.setMinimumSectionSize(58)
    header.setStretchLastSection(False)
    for column in range(columns):
        width = source.columnWidth(column)
        table.setColumnWidth(column, max(90, min(width, 360)))
    return table


class TablePreviewDialog(QtWidgets.QDialog):
    """Read-only full-table view used by all SOD tools for «Подробнее»."""

    def __init__(
        self,
        parent: Optional[QtWidgets.QWidget],
        source_table: QtWidgets.QTableWidget,
        title: str,
        subtitle: str = "",
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(1080, 620)
        self.setMinimumSize(780, 460)
        if parent is not None:
            self.setStyleSheet(parent.styleSheet())
        QtCore.QTimer.singleShot(0, lambda: apply_windows_titlebar_theme(self, initial_dark_theme(False)))

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(16, 14, 16, 12)
        root.setSpacing(8)

        page_title = QtWidgets.QLabel(title)
        page_title.setObjectName("pageTitle")
        root.addWidget(page_title)

        info_text = subtitle.strip() or f"Полный предпросмотр: строк {source_table.rowCount()}, колонок {source_table.columnCount()}."
        info = QtWidgets.QLabel(info_text)
        info.setObjectName("rowSubtitle")
        info.setWordWrap(True)
        root.addWidget(info)

        self.table = _copy_table(source_table)
        root.addWidget(self.table, 1)

        footer = QtWidgets.QHBoxLayout()
        footer.addStretch(1)
        close_button = QtWidgets.QPushButton("Закрыть")
        close_button.setObjectName("modeSwitch")
        close_button.clicked.connect(self.accept)
        footer.addWidget(close_button)
        root.addLayout(footer)


def show_table_details(
    parent: Optional[QtWidgets.QWidget],
    source_table: QtWidgets.QTableWidget,
    title: str,
    subtitle: str = "",
) -> bool:
    if source_table.rowCount() <= 0:
        QtWidgets.QMessageBox.information(parent, "Подробнее", "Предпросмотр пока пуст. Сначала выполните проверку.")
        return False
    dialog = TablePreviewDialog(parent, source_table, title, subtitle)
    dialog.exec()
    return True


def make_details_button(text: str = "Подробнее") -> QtWidgets.QPushButton:
    button = QtWidgets.QPushButton(text)
    button.setObjectName("modeSwitch")
    button.setMinimumWidth(105)
    return button


__all__ = [
    "PreviewHeader",
    "make_preview_header",
    "TablePreviewDialog",
    "configure_preview_table",
    "fit_preview_height",
    "make_details_button",
    "show_table_details",
]


class BackToManagerButton(QtWidgets.QPushButton):
    """Shared top-left navigation button with the user-provided exit icon."""

    def __init__(self, window: QtWidgets.QWidget):
        super().__init__("Все СОД", window)
        self.setObjectName("backToManagerButton")
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setToolTip("Вернуться к выбору СОД")
        self.setIconSize(QtCore.QSize(18, 18))
        self._asset_dir = shared_asset_dir(Path(__file__))
        self._icon_dark = None
        self._refresh_icon()
        self._window = window
        self.clicked.connect(self._return_to_manager)

    def _return_to_manager(self) -> None:
        from shared.theme_core import RETURN_TO_MANAGER_CODE

        app = QtWidgets.QApplication.instance()
        if app is not None:
            quit_on_last = app.quitOnLastWindowClosed()
            app.setQuitOnLastWindowClosed(False)
            self._window.close()
            app.exit(RETURN_TO_MANAGER_CODE)
            app.setQuitOnLastWindowClosed(quit_on_last)
        else:
            self._window.close()

    def _refresh_icon(self) -> None:
        dark = bool(initial_dark_theme(False))
        if dark == self._icon_dark:
            return
        self._icon_dark = dark
        self.setIcon(themed_icon(self._asset_dir, "exit.png", dark))

    def paintEvent(self, event):
        # Theme changes are persisted globally; refresh lazily when Qt repaints.
        self._refresh_icon()
        super().paintEvent(event)


def make_back_to_manager_button(window: QtWidgets.QWidget) -> QtWidgets.QPushButton:
    """Create the same compact «Все СОД» navigation control for every tool."""
    from shared.theme_core import launched_from_manager
    button = BackToManagerButton(window)
    button.setVisible(launched_from_manager())
    return button


def install_status_bar(window: QtWidgets.QMainWindow, text: str = "Готово") -> QtWidgets.QStatusBar:
    """Install the common status strip and infer a calm semantic state from messages."""
    bar = window.statusBar()
    bar.setObjectName("unifiedStatusBar")
    bar.setSizeGripEnabled(False)

    def update_tone(message: str) -> None:
        value = str(message or "").casefold()
        if any(word in value for word in ("ошиб", "не удалось", "отмен")):
            tone = "danger"
        elif any(word in value for word in ("ожидан", "выполня", "загружа", "создаю", "провер", "свер")):
            tone = "busy"
        elif any(word in value for word in ("готов", "успеш", "загружен", "сохран", "выполнен", "создан")):
            tone = "success"
        else:
            tone = "neutral"
        bar.setProperty("tone", tone)
        bar.style().unpolish(bar)
        bar.style().polish(bar)

    bar.messageChanged.connect(update_tone)
    bar.showMessage(text)
    update_tone(text)
    return bar


def add_standard_header_controls(
    header_layout,
    window: QtWidgets.QWidget,
    theme_toggle: QtWidgets.QWidget,
) -> QtWidgets.QPushButton:
    """Place navigation at top-left and the theme control at top-right."""
    back = make_back_to_manager_button(window)
    # All tool headers already add their title first. Insert the return action
    # before it so navigation has one predictable position in every SOD.
    header_layout.insertWidget(0, back, 0, QtCore.Qt.AlignTop)
    header_layout.addWidget(theme_toggle, 0, QtCore.Qt.AlignTop)
    return back


__all__ += [
    "NoWheelTabBar",
    "make_back_to_manager_button",
    "install_status_bar",
    "add_standard_header_controls",
]


def mark_destructive_buttons(root: QtWidgets.QWidget) -> None:
    """Give delete/remove actions a single restrained destructive state."""
    keywords = ("удалить", "очистить", "сбросить")
    for button in root.findChildren(QtWidgets.QPushButton):
        label = button.text().strip().casefold()
        if any(word in label for word in keywords):
            button.setObjectName("destructiveButton")
            button.setIcon(QtGui.QIcon())


__all__ += ["mark_destructive_buttons"]
