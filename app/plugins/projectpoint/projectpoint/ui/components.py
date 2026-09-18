from __future__ import annotations

from pathlib import Path
import shutil
from typing import Iterable, Sequence

import pandas as pd
from PySide6 import QtCore, QtGui, QtWidgets

from ..excel.common import read_excel_table
from ..styles import themed_icon, initial_dark_theme, apply_windows_titlebar_theme
from shared.ui_components import NoWheelTabBar, configure_preview_table, make_preview_header


PREVIEW_ROW_LIMIT = 250
PREVIEW_COLUMN_LIMIT = 60


def make_inline_icon(asset_dir: Path | str, asset_name: str, dark: bool = False, size: int = 16) -> QtWidgets.QLabel:
    label = QtWidgets.QLabel()
    label.setObjectName("inlineIcon")
    label.setFixedSize(size + 4, size + 4)
    label.setAlignment(QtCore.Qt.AlignCenter)
    icon = themed_icon(asset_dir, asset_name, dark)
    label.setPixmap(icon.pixmap(QtCore.QSize(size, size)))
    label.setProperty("iconAsset", asset_name)
    label.setProperty("iconSize", size)
    return label


def make_badge(asset_dir: Path | str, asset_name: str, dark: bool = False, tone: str = "info", size: int = 34) -> QtWidgets.QLabel:
    badge = QtWidgets.QLabel()
    badge.setObjectName("iconBadge")
    badge.setProperty("tone", tone)
    badge.setAlignment(QtCore.Qt.AlignCenter)
    badge.setFixedSize(size, size)
    icon = themed_icon(asset_dir, asset_name, dark)
    badge.setPixmap(icon.pixmap(QtCore.QSize(18, 18)))
    badge.setProperty("iconAsset", asset_name)
    badge.setProperty("iconSize", 18)
    return badge


def make_file_row(
    asset_dir: Path | str,
    dark: bool,
    title: str,
    subtitle: str,
    download_button: QtWidgets.QPushButton,
    upload_button: QtWidgets.QPushButton,
    *,
    is_last: bool = True,
) -> QtWidgets.QWidget:
    row_widget = QtWidgets.QWidget()
    row_widget.setObjectName("fileRowLast" if is_last else "fileRow")
    row = QtWidgets.QHBoxLayout(row_widget)
    row.setContentsMargins(0, 7, 0, 7)
    row.setSpacing(10)
    row.addWidget(make_badge(asset_dir, "Excel.png", dark, "info", 32))

    text_col = QtWidgets.QVBoxLayout()
    text_col.setContentsMargins(0, 0, 0, 0)
    text_col.setSpacing(1)
    title_label = QtWidgets.QLabel(title)
    title_label.setObjectName("rowTitle")
    subtitle_label = QtWidgets.QLabel(subtitle)
    subtitle_label.setObjectName("rowSubtitle")
    subtitle_label.setWordWrap(True)
    text_col.addWidget(title_label)
    text_col.addWidget(subtitle_label)
    row.addLayout(text_col, 1)

    download_button.setObjectName("downloadButton")
    download_button.setToolTip("Сохранить шаблон Excel")
    download_button.setMinimumWidth(150)
    download_button.setFixedHeight(34)
    upload_button.setObjectName("uploadButton")
    upload_button.setToolTip("Выбрать Excel-файл")
    upload_button.setProperty("fileLoaded", False)
    upload_button.setMinimumWidth(150)
    upload_button.setFixedHeight(34)
    row.addWidget(download_button)
    row.addWidget(upload_button)
    return row_widget


def make_file_path_row(
    file_edit: QtWidgets.QLineEdit,
    export_button: QtWidgets.QPushButton,
) -> QtWidgets.QWidget:
    row_widget = QtWidgets.QWidget()
    row = QtWidgets.QHBoxLayout(row_widget)
    row.setContentsMargins(0, 0, 0, 0)
    row.setSpacing(8)
    file_edit.setReadOnly(True)
    file_edit.setPlaceholderText("Excel-файл не выбран")
    export_button.setText("")
    export_button.setObjectName("downloadButton")
    export_button.setToolTip("Выгрузить текущий Excel-файл")
    export_button.setProperty("iconAsset", "close.png")
    export_button.setFixedSize(34, 34)
    export_button.setEnabled(False)
    row.addWidget(file_edit, 1)
    row.addWidget(export_button, 0)
    return row_widget


def copy_selected_excel_file(parent, source_path: str, dialog_title: str = "Выгрузить Excel-файл") -> bool:
    source = Path(str(source_path or "")).expanduser()
    if not source_path:
        QtWidgets.QMessageBox.warning(parent, "Выгрузка файла", "Сначала загрузите Excel-файл")
        return False
    if not source.exists() or not source.is_file():
        QtWidgets.QMessageBox.warning(parent, "Выгрузка файла", f"Файл не найден:\n{source}")
        return False

    suffix = source.suffix if source.suffix.lower() in {".xls", ".xlsx"} else ".xlsx"
    target_path, _ = QtWidgets.QFileDialog.getSaveFileName(
        parent,
        dialog_title,
        str(source.with_suffix(suffix).name),
        "Excel (*.xlsx *.xls)",
    )
    if not target_path:
        return False

    target = Path(target_path)
    if not target.suffix:
        target = target.with_suffix(suffix)

    try:
        if target.resolve() != source.resolve():
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, target)
    except Exception as exc:
        QtWidgets.QMessageBox.warning(parent, "Выгрузка файла", f"Не удалось выгрузить файл:\n{exc}")
        return False

    QtWidgets.QMessageBox.information(parent, "Выгрузка файла", f"Файл сохранён:\n{target}")
    return True


def make_info_banner(asset_dir: Path | str, dark: bool, text: str) -> QtWidgets.QFrame:
    banner = QtWidgets.QFrame()
    banner.setObjectName("infoBanner")
    row = QtWidgets.QHBoxLayout(banner)
    row.setContentsMargins(10, 7, 10, 7)
    row.setSpacing(8)
    row.addWidget(make_inline_icon(asset_dir, "information.png", dark, 16), 0, QtCore.Qt.AlignTop)
    label = QtWidgets.QLabel(text)
    label.setObjectName("infoText")
    label.setWordWrap(True)
    row.addWidget(label, 1)
    return banner


def make_status_chip(text: str, tone: str = "neutral") -> QtWidgets.QLabel:
    chip = QtWidgets.QLabel(text)
    chip.setObjectName("statusChip")
    chip.setProperty("tone", tone)
    chip.setAlignment(QtCore.Qt.AlignCenter)
    return chip


def mark_file_button(button: QtWidgets.QPushButton, loaded: bool) -> None:
    button.setProperty("fileLoaded", bool(loaded))
    button.style().unpolish(button)
    button.style().polish(button)
    button.update()


def _clean_preview_frame(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame()
    frame = df.copy()
    frame = frame.dropna(axis=0, how="all").dropna(axis=1, how="all")
    frame.columns = [str(col) for col in frame.columns]
    return frame


def read_preview_sheet(file_path: str, sheet_name, required_headers: Iterable[str] = ()) -> pd.DataFrame:
    required_headers = tuple(required_headers or ())
    if required_headers:
        df = read_excel_table(file_path, sheet_name=sheet_name, required_headers=required_headers)
    else:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    return _clean_preview_frame(df)


def load_preview_sheets(
    file_path: str,
    sheet_specs: Sequence[tuple[str, Iterable[str]]],
) -> list[tuple[str, pd.DataFrame]]:
    if not file_path:
        raise ValueError("Сначала выберите Excel-файл")
    sheets: list[tuple[str, pd.DataFrame]] = []
    for sheet_name, required_headers in sheet_specs:
        if not sheet_name:
            continue
        frame = read_preview_sheet(file_path, sheet_name, required_headers)
        sheets.append((str(sheet_name), frame))
    if not sheets:
        raise ValueError("Не выбран ни один лист для предпросмотра")
    return sheets


def populate_inline_preview(
    table: QtWidgets.QTableWidget,
    sheets: Sequence[tuple[str, pd.DataFrame]],
    *,
    row_limit: int = 35,
    column_limit: int = 12,
) -> tuple[int, int]:
    """Fill the compact on-page preview table in the same style as the reference app."""
    if not sheets:
        table.clear()
        table.setRowCount(0)
        table.setColumnCount(0)
        return 0, 0

    total_rows = sum(len(frame.index) for _, frame in sheets)
    if len(sheets) == 1:
        combined = sheets[0][1].copy()
    else:
        frames = []
        for sheet_name, frame in sheets:
            part = frame.copy()
            part.insert(0, "Лист", sheet_name)
            frames.append(part)
        combined = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()

    combined = _clean_preview_frame(combined)
    visible = combined.iloc[:row_limit, :column_limit]
    table.clear()
    table.setColumnCount(len(visible.columns))
    table.setRowCount(len(visible.index))
    if len(visible.columns):
        table.setHorizontalHeaderLabels([str(col) for col in visible.columns])
    configure_preview_table(table)
    for row_index, (_, row) in enumerate(visible.iterrows()):
        for column_index, value in enumerate(row.tolist()):
            text = "" if pd.isna(value) else str(value)
            item = QtWidgets.QTableWidgetItem(text)
            item.setToolTip(text)
            table.setItem(row_index, column_index, item)
    for column_index in range(table.columnCount()):
        table.setColumnWidth(column_index, 150)
    _fit_preview_height(table, len(visible.index))
    return total_rows, len(combined.columns)


def run_inline_excel_preview(
    parent,
    table: QtWidgets.QTableWidget,
    file_path: str,
    sheet_specs: Sequence[tuple[str, Iterable[str]]],
    summary_label: QtWidgets.QLabel | None = None,
) -> list[tuple[str, pd.DataFrame]] | None:
    try:
        sheets = load_preview_sheets(file_path, sheet_specs)
        total_rows, total_columns = populate_inline_preview(table, sheets)
    except Exception as exc:
        QtWidgets.QMessageBox.warning(parent, "Предпросмотр", f"Не удалось прочитать Excel:\n{exc}")
        return None
    if summary_label is not None:
        suffix = ""
        if total_rows > 35:
            suffix += " · показаны первые 35 строк"
        if total_columns > 12:
            suffix += " · первые 12 колонок"
        summary_label.setText(
            f"Предпросмотр готов · листов: {len(sheets)} · строк: {total_rows} · колонок: {total_columns}{suffix}"
        )
    return sheets


def _fit_preview_height(table: QtWidgets.QTableWidget, row_count: int) -> None:
    """Keep empty previews compact and expand only when rows are actually shown."""
    visible_rows = max(0, min(int(row_count), 6))
    header = 31 if table.horizontalHeader().isVisible() else 0
    target = header + max(2, visible_rows) * 27 + 8
    target = max(92, min(target, 205))
    table.setMinimumHeight(target)
    table.setMaximumHeight(235)


def clear_inline_preview(table: QtWidgets.QTableWidget) -> None:
    table.clear()
    table.setRowCount(0)
    table.setColumnCount(0)
    _fit_preview_height(table, 0)


def make_preview_table(min_height: int = 92, max_height: int = 235) -> QtWidgets.QTableWidget:
    table = QtWidgets.QTableWidget(0, 0)
    configure_preview_table(table)
    table.setMinimumHeight(min_height)
    table.setMaximumHeight(max_height)
    table.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
    return table


class ExcelPreviewDialog(QtWidgets.QDialog):
    """Larix-style read-only Excel preview. It never sends data to Project Point."""

    def __init__(self, parent, title: str, sheets: Sequence[tuple[str, pd.DataFrame]]):
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

        total_rows = sum(len(df.index) for _, df in sheets)
        subtitle = QtWidgets.QLabel(
            f"Локальный предпросмотр Excel: {len(sheets)} лист(а), {total_rows} строк. "
            "Никакие изменения на сервер не отправляются."
        )
        subtitle.setObjectName("rowSubtitle")
        subtitle.setWordWrap(True)
        root.addWidget(subtitle)

        if len(sheets) == 1:
            _, frame = sheets[0]
            root.addWidget(self._make_table(frame), 1)
        else:
            tabs = QtWidgets.QTabWidget()
            tabs.setTabBar(NoWheelTabBar())
            tabs.setObjectName("previewTabs")
            for sheet_name, frame in sheets:
                holder = QtWidgets.QWidget()
                holder_layout = QtWidgets.QVBoxLayout(holder)
                holder_layout.setContentsMargins(0, 8, 0, 0)
                holder_layout.setSpacing(6)
                summary = QtWidgets.QLabel(f"Строк: {len(frame.index)} · Колонок: {len(frame.columns)}")
                summary.setObjectName("rowSubtitle")
                holder_layout.addWidget(summary)
                holder_layout.addWidget(self._make_table(frame), 1)
                tabs.addTab(holder, str(sheet_name))
            root.addWidget(tabs, 1)

        footer = QtWidgets.QHBoxLayout()
        footer.addStretch(1)
        close_btn = QtWidgets.QPushButton("Закрыть")
        close_btn.setObjectName("modeSwitch")
        close_btn.clicked.connect(self.accept)
        footer.addWidget(close_btn)
        root.addLayout(footer)

    def _make_table(self, frame: pd.DataFrame) -> QtWidgets.QTableWidget:
        preview = frame.iloc[:PREVIEW_ROW_LIMIT, :PREVIEW_COLUMN_LIMIT]
        table = QtWidgets.QTableWidget(len(preview.index), len(preview.columns))
        table.setHorizontalHeaderLabels([str(col) for col in preview.columns])
        configure_preview_table(table)
        for row_index, (_, row) in enumerate(preview.iterrows()):
            for column_index, value in enumerate(row.tolist()):
                text = "" if pd.isna(value) else str(value)
                item = QtWidgets.QTableWidgetItem(text)
                item.setToolTip(text)
                table.setItem(row_index, column_index, item)
        for column_index in range(table.columnCount()):
            table.setColumnWidth(column_index, 160)
        return table


def show_excel_preview(
    parent,
    title: str,
    file_path: str,
    sheet_specs: Sequence[tuple[str, Iterable[str]]],
) -> bool:
    if not file_path:
        QtWidgets.QMessageBox.warning(parent, "Предпросмотр", "Сначала выберите Excel-файл")
        return False
    try:
        sheets = load_preview_sheets(file_path, sheet_specs)
    except Exception as exc:
        QtWidgets.QMessageBox.warning(parent, "Предпросмотр", f"Не удалось прочитать Excel:\n{exc}")
        return False
    dialog = ExcelPreviewDialog(parent, title, sheets)
    dialog.exec()
    return True


__all__ = [
    "ExcelPreviewDialog",
    "configure_preview_table",
    "make_badge",
    "make_file_row",
    "make_info_banner",
    "make_inline_icon",
    "make_status_chip",
    "load_preview_sheets",
    "make_file_path_row",
    "make_preview_table",
    "make_preview_header",
    "mark_file_button",
    "populate_inline_preview",
    "copy_selected_excel_file",
    "read_preview_sheet",
    "run_inline_excel_preview",
    "show_excel_preview",
]
