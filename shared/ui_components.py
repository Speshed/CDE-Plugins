from __future__ import annotations

from typing import Optional

try:
    from PySide6 import QtCore, QtGui, QtWidgets
except ImportError:  # pragma: no cover - compatibility with older SGNL installs
    from PyQt5 import QtCore, QtGui, QtWidgets

from shared.theme_core import apply_windows_titlebar_theme, initial_dark_theme




class NoWheelTabBar(QtWidgets.QTabBar):
    """Tab bar that ignores the mouse wheel so scrolling never changes sections."""

    def wheelEvent(self, event: QtGui.QWheelEvent) -> None:
        event.accept()

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
    "TablePreviewDialog",
    "configure_preview_table",
    "fit_preview_height",
    "make_details_button",
    "show_table_details",
]


def make_back_to_manager_button(window: QtWidgets.QWidget) -> QtWidgets.QPushButton:
    """Create the same compact «Все СОД» navigation control for every tool."""
    button = QtWidgets.QPushButton("← Все СОД")
    button.setObjectName("backToManagerButton")
    button.setCursor(QtCore.Qt.PointingHandCursor)
    button.setToolTip("Вернуться к выбору СОД")
    from shared.theme_core import launched_from_manager
    button.setVisible(launched_from_manager())
    button.clicked.connect(window.close)
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
    """Append the common navigation + theme controls to an existing header row."""
    back = make_back_to_manager_button(window)
    header_layout.addWidget(back, 0, QtCore.Qt.AlignVCenter)
    header_layout.addWidget(theme_toggle, 0, QtCore.Qt.AlignVCenter)
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
