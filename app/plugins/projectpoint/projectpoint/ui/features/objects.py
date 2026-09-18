from __future__ import annotations

import os
import threading
import traceback

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox, QFileDialog, QFrame, QGridLayout, QGroupBox, QHBoxLayout,
    QLabel, QLineEdit, QMessageBox, QPushButton, QScrollArea, QVBoxLayout, QWidget,
)

from ...auth import *
from ...config import *
from ...api import *
from ...excel import *
from ..dialogs import LogDialog
from ..components import (
    clear_inline_preview,     make_badge, make_info_banner, make_preview_table, make_preview_header, mark_file_button,
    run_inline_excel_preview, show_excel_preview,
)


class ObjectsMixin:
    def _build_objects_tab(self):
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.setSpacing(0)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(2, 8, 2, 2)
        layout.setSpacing(8)

        scroll_area.setWidget(content)
        tab_layout.addWidget(scroll_area)

        file_card = QFrame()
        file_card.setObjectName("card")
        file_layout = QVBoxLayout(file_card)
        file_layout.setContentsMargins(14, 8, 14, 9)
        file_layout.setSpacing(6)

        file_header = QHBoxLayout()
        file_header.setSpacing(10)
        file_header.addWidget(make_badge(self.asset_dir, "Excel.png", self.is_dark_theme, "info", 32))
        file_text = QVBoxLayout()
        file_text.setSpacing(1)
        file_title = QLabel("Excel-файл объектов строительства")
        file_title.setObjectName("rowTitle")
        file_subtitle = QLabel("Загрузите таблицу Code / Title / ParentCode и проверьте её до создания")
        file_subtitle.setObjectName("rowSubtitle")
        file_text.addWidget(file_title)
        file_text.addWidget(file_subtitle)
        file_header.addLayout(file_text, 1)
        self.objects_template_btn = QPushButton("Скачать шаблон")
        self.objects_template_btn.setObjectName("secondaryBtn")
        self.objects_template_btn.setToolTip("Сохранить шаблон Excel")
        self.objects_template_btn.setMinimumWidth(150)
        self.objects_template_btn.setFixedHeight(34)
        self.objects_template_btn.clicked.connect(self._start_objects_template)
        file_header.addWidget(self.objects_template_btn)
        self.objects_file_btn = QPushButton("Загрузить файл")
        self.objects_file_btn.setObjectName("uploadButton")
        self.objects_file_btn.setToolTip("Выбрать Excel-файл")
        self.objects_file_btn.setProperty("fileLoaded", False)
        self.objects_file_btn.setMinimumWidth(150)
        self.objects_file_btn.setFixedHeight(34)
        self.objects_file_btn.clicked.connect(self._choose_objects_file)
        file_header.addWidget(self.objects_file_btn)
        file_layout.addLayout(file_header)

        self.objects_file_edit = QLineEdit()
        self.objects_file_edit.setReadOnly(True)
        self.objects_file_edit.setPlaceholderText("Excel-файл не выбран")
        file_layout.addWidget(self.objects_file_edit)
        sheet_row = QHBoxLayout()
        sheet_row.setSpacing(8)
        sheet_lbl = QLabel("Лист Excel")
        sheet_lbl.setObjectName("fieldLabel")
        sheet_row.addWidget(sheet_lbl)
        self.objects_sheet_combo = QComboBox()
        self.objects_sheet_combo.setEnabled(False)
        self.objects_sheet_combo.currentTextChanged.connect(self._invalidate_objects_preview)
        sheet_row.addWidget(self.objects_sheet_combo, 1)
        file_layout.addLayout(sheet_row)
        layout.addWidget(file_card)

        layout.addWidget(make_info_banner(
            self.asset_dir, self.is_dark_theme,
            "Предпросмотр показывает строки Excel локально. Родитель каждого объекта определяется выбранным ниже режимом.",
        ))

        parent_mode_group = QGroupBox("Режим определения родителя")
        parent_mode_layout = QGridLayout(parent_mode_group)
        parent_mode_layout.setContentsMargins(12, 12, 12, 12)
        parent_mode_layout.setHorizontalSpacing(10)
        parent_mode_layout.setVerticalSpacing(6)

        parent_mode_lbl = QLabel("Режим:")
        parent_mode_lbl.setObjectName("fieldLabel")
        parent_mode_layout.addWidget(parent_mode_lbl, 0, 0)
        self.objects_parent_mode_combo = QComboBox()
        self.objects_parent_mode_combo.setMinimumHeight(34)
        self.objects_parent_mode_combo.setMaximumWidth(360)
        self.objects_parent_mode_combo.addItem("Из Excel по ParentCode", "excel")
        self.objects_parent_mode_combo.addItem("Один базовый родитель для всех строк", "base")
        self.objects_parent_mode_combo.currentIndexChanged.connect(self._on_objects_parent_mode_changed)
        parent_mode_layout.addWidget(self.objects_parent_mode_combo, 0, 1)
        parent_mode_layout.setColumnStretch(2, 1)

        self.objects_parent_mode_hint = QLabel("")
        self.objects_parent_mode_hint.setWordWrap(True)
        self.objects_parent_mode_hint.setObjectName("rowSubtitle")
        parent_mode_layout.addWidget(self.objects_parent_mode_hint, 1, 0, 1, 3)

        layout.addWidget(parent_mode_group)

        columns_row = QHBoxLayout()
        columns_row.setSpacing(10)

        project_group = QGroupBox("Проект")
        project_layout = QGridLayout(project_group)
        project_layout.setContentsMargins(12, 12, 12, 12)
        project_layout.setHorizontalSpacing(10)
        project_layout.setVerticalSpacing(8)

        self.objects_projects_load_btn = QPushButton("Обновить")
        self.objects_projects_load_btn.setObjectName("secondaryBtn")
        self.objects_projects_load_btn.setMinimumHeight(30)
        self.objects_projects_load_btn.setMinimumWidth(110)
        self.objects_projects_load_btn.clicked.connect(self._start_objects_projects_load)
        project_layout.addWidget(self.objects_projects_load_btn, 0, 1)

        self.objects_project_search_edit = QLineEdit()
        self.objects_project_search_edit.setMinimumHeight(34)
        self.objects_project_search_edit.setPlaceholderText("Поиск по проекту...")
        self.objects_project_search_edit.textChanged.connect(self._filter_objects_project_combo)
        project_layout.addWidget(self.objects_project_search_edit, 1, 0, 1, 2)

        project_lbl = QLabel("Проект:")
        project_lbl.setObjectName("fieldLabel")
        self.objects_project_combo = QComboBox()
        self.objects_project_combo.setMinimumHeight(34)
        self.objects_project_combo.setEnabled(False)
        self.objects_project_combo.setMaxVisibleItems(20)
        self.objects_project_combo.view().setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.objects_project_combo.currentIndexChanged.connect(self._on_objects_project_changed)
        project_layout.addWidget(project_lbl, 2, 0)
        project_layout.addWidget(self.objects_project_combo, 2, 1)

        self.objects_project_info = QLabel("Id: -   |   Code: -")
        self.objects_project_info.setObjectName("rowSubtitle")
        project_layout.addWidget(self.objects_project_info, 3, 0)
        self.objects_project_details_btn = QPushButton("Подробнее")
        self.objects_project_details_btn.setObjectName("logBtn")
        self.objects_project_details_btn.setFixedHeight(28)
        self.objects_project_details_btn.setMinimumWidth(84)
        self.objects_project_details_btn.clicked.connect(self._show_selected_project_details)
        project_layout.addWidget(self.objects_project_details_btn, 3, 1)
        project_layout.setColumnStretch(0, 1)

        parent_group = QGroupBox("Родительский объект")
        self.objects_parent_group = parent_group
        parent_layout = QGridLayout(parent_group)
        parent_layout.setContentsMargins(12, 12, 12, 12)
        parent_layout.setHorizontalSpacing(10)
        parent_layout.setVerticalSpacing(8)

        self.objects_load_btn = QPushButton("Обновить")
        self.objects_load_btn.setObjectName("secondaryBtn")
        self.objects_load_btn.setMinimumHeight(30)
        self.objects_load_btn.setMinimumWidth(110)
        self.objects_load_btn.clicked.connect(self._start_objects_load)
        parent_layout.addWidget(self.objects_load_btn, 0, 1)

        self.objects_parent_search_edit = QLineEdit()
        self.objects_parent_search_edit.setMinimumHeight(34)
        self.objects_parent_search_edit.setPlaceholderText("Поиск по родителю...")
        self.objects_parent_search_edit.textChanged.connect(self._filter_objects_parent_combo)
        parent_layout.addWidget(self.objects_parent_search_edit, 1, 0, 1, 2)

        parent_lbl = QLabel("Базовый родитель:")
        parent_lbl.setObjectName("fieldLabel")
        self.objects_parent_combo = QComboBox()
        self.objects_parent_combo.setMinimumHeight(34)
        self.objects_parent_combo.setEnabled(False)
        self.objects_parent_combo.setMaxVisibleItems(20)
        self.objects_parent_combo.view().setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.objects_parent_combo.currentIndexChanged.connect(self._on_objects_parent_changed)
        parent_layout.addWidget(parent_lbl, 2, 0)
        parent_layout.addWidget(self.objects_parent_combo, 2, 1)

        self.objects_parent_info = QLabel("Id: -   |   Code: -")
        self.objects_parent_info.setObjectName("rowSubtitle")
        parent_layout.addWidget(self.objects_parent_info, 3, 0)
        self.objects_parent_details_btn = QPushButton("Подробнее")
        self.objects_parent_details_btn.setObjectName("logBtn")
        self.objects_parent_details_btn.setFixedHeight(28)
        self.objects_parent_details_btn.setMinimumWidth(84)
        self.objects_parent_details_btn.clicked.connect(self._show_selected_parent_details)
        parent_layout.addWidget(self.objects_parent_details_btn, 3, 1)
        parent_layout.setColumnStretch(0, 1)

        columns_row.addWidget(project_group, 1)
        columns_row.addWidget(parent_group, 1)
        layout.addLayout(columns_row)

        action_card = QFrame()
        action_card.setObjectName("card")
        action_layout = QVBoxLayout(action_card)
        action_layout.setContentsMargins(14, 9, 14, 10)
        action_layout.setSpacing(7)
        action_title = make_preview_header("Предпросмотр и создание объектов")
        action_layout.addWidget(action_title)
        self.objects_preview_table = make_preview_table(130, 190)
        action_layout.addWidget(self.objects_preview_table)
        self.objects_summary = QLabel("Загрузите Excel-файл для предпросмотра")
        self.objects_summary.setObjectName("rowSubtitle")
        self.objects_summary.setWordWrap(True)
        footer = QHBoxLayout()
        footer.setSpacing(8)
        footer.addWidget(self.objects_summary, 1)
        objects_log_btn = QPushButton("Log")
        objects_log_btn.setObjectName("logButton")
        objects_log_btn.clicked.connect(self._show_objects_log)
        footer.addWidget(objects_log_btn)
        self.objects_preview_details_btn = QPushButton("Подробнее")
        self.objects_preview_details_btn.setObjectName("modeSwitch")
        self.objects_preview_details_btn.setEnabled(False)
        self.objects_preview_details_btn.clicked.connect(self._show_objects_preview_details)
        footer.addWidget(self.objects_preview_details_btn)
        self.objects_preview_btn = QPushButton("Предпросмотр")
        self.objects_preview_btn.setObjectName("loginButton")
        self.objects_preview_btn.setEnabled(False)
        self.objects_preview_btn.clicked.connect(self._preview_objects_excel)
        footer.addWidget(self.objects_preview_btn)
        self.objects_create_btn = QPushButton("Создать объекты")
        self.objects_create_btn.setObjectName("greenAction")
        self.objects_create_btn.setMinimumWidth(170)
        self.objects_create_btn.clicked.connect(self._start_objects_create)
        footer.addWidget(self.objects_create_btn)
        action_layout.addLayout(footer)
        layout.addWidget(action_card)
        layout.addSpacing(2)

        self.objects_tab_index = self.tabs.addTab(tab, "Объекты строительства")
        self.objects_parent_mode_combo.setCurrentIndex(0)
        self._on_objects_parent_mode_changed(0)

    def _start_objects_template(self):
        self._save_template_file(
            "Сохранить шаблон объектов строительства",
            "Project Point Объекты строительства.xlsx",
            write_objects_import_template,
        )

    def _choose_objects_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл", "", "Excel (*.xlsx *.xls)"
        )
        if path:
            self.objects_file_edit.setText(path)
            mark_file_button(self.objects_file_btn, True)
            self._refresh_action_icons()
            try:
                xls = pd.ExcelFile(path)
                sheets = xls.sheet_names
                self.objects_sheet_combo.blockSignals(True)
                self.objects_sheet_combo.clear()
                self.objects_sheet_combo.addItems(sheets)
                self.objects_sheet_combo.setEnabled(True)
                self.objects_sheet_combo.blockSignals(False)
                self.objects_preview_btn.setEnabled(bool(sheets))
                self.objects_preview_details_btn.setEnabled(False)
                clear_inline_preview(self.objects_preview_table)
                self.objects_summary.setText(f"Файл выбран · листов: {len(sheets)} · запустите предпросмотр")
            except Exception as e:
                self.objects_sheet_combo.clear()
                self.objects_sheet_combo.setEnabled(False)
                self.objects_preview_btn.setEnabled(False)
                self.objects_preview_details_btn.setEnabled(False)
                self.objects_summary.setText("Не удалось прочитать Excel-файл")
                QMessageBox.warning(self, "Ошибка", f"Не удалось прочитать листы: {e}")

    def _on_objects_project_changed(self, index):
        if index < 0:
            self.selected_objects_project_id = None
            self.objects_project_info.setText("Id: -   |   Code: -")
            return

        source_index = self.objects_project_combo.itemData(index)
        if source_index is None:
            self.selected_objects_project_id = None
            self.objects_project_info.setText("Id: -   |   Code: -")
            return

        try:
            source_index = int(source_index)
        except Exception:
            self.selected_objects_project_id = None
            self.objects_project_info.setText("Id: -   |   Code: -")
            return

        if source_index < 0 or source_index >= len(self.objects_projects):
            self.selected_objects_project_id = None
            self.objects_project_info.setText("Id: -   |   Code: -")
            return

        project = self.objects_projects[source_index]
        self.selected_objects_project_id = normalize_object_lookup_value(project.get("Id"))
        self.objects_project_info.setText(
            f"Id: {normalize_object_lookup_value(project.get('Id'))}   |   "
            f"Code: {normalize_object_lookup_value(project.get('Code'))}"
        )
        self.objects_items = []
        self.objects_indexes = {"code_to_id": {}, "title_to_objects": {}}
        self._populate_objects_parent_combo(self.objects_parent_search_edit.text())
        self._start_objects_load()

    def _get_objects_parent_mode(self):
        if hasattr(self, "objects_parent_mode_combo"):
            mode = self.objects_parent_mode_combo.currentData()
            if mode:
                return mode
        return getattr(self, "objects_parent_mode", "excel")

    def _on_objects_parent_mode_changed(self, index):
        mode = self.objects_parent_mode_combo.currentData() if hasattr(self, "objects_parent_mode_combo") else "excel"
        if not mode:
            mode = "excel"
        self.objects_parent_mode = mode

        excel_mode = mode == "excel"
        if hasattr(self, "objects_parent_group"):
            self.objects_parent_group.setVisible(not excel_mode)

        if hasattr(self, "objects_parent_mode_hint"):
            if excel_mode:
                self.objects_parent_mode_hint.setText("ParentCode ищется среди объектов проекта и строк этого Excel. Ручной родитель скрыт.")
            else:
                self.objects_parent_mode_hint.setText("Все строки привязываются к одному базовому родителю. Ручной выбор обязателен.")

        if hasattr(self, "objects_create_btn"):
            if excel_mode:
                self.objects_create_btn.setToolTip("Родители будут найдены по ParentCode из Excel")
            else:
                self.objects_create_btn.setToolTip("Используется один базовый родитель для всех строк")

    def _on_objects_parent_changed(self, index):
        if index < 0:
            self.selected_object_parent_id = None
            self.objects_parent_info.setText("Id: -   |   Code: -")
            return

        source_index = self.objects_parent_combo.itemData(index)
        if source_index is None:
            self.selected_object_parent_id = None
            self.objects_parent_info.setText("Id: -   |   Code: -")
            return

        try:
            source_index = int(source_index)
        except Exception:
            self.selected_object_parent_id = None
            self.objects_parent_info.setText("Id: -   |   Code: -")
            return

        if source_index < 0 or source_index >= len(self.objects_items):
            self.selected_object_parent_id = None
            self.objects_parent_info.setText("Id: -   |   Code: -")
            return

        obj = self.objects_items[source_index]
        self.selected_object_parent_id = normalize_object_lookup_value(obj.get("Id"))
        self.objects_parent_info.setText(
            f"Id: {normalize_object_lookup_value(obj.get('Id'))}   |   "
            f"Code: {normalize_object_lookup_value(obj.get('Code'))}"
        )

    def _object_project_display(self, project):
        code = normalize_object_lookup_value(project.get("Code"))
        title = normalize_object_lookup_value(project.get("Title"))
        if code and title:
            return f"{code}. {title}"
        if title:
            return title
        if code:
            return code
        return normalize_object_lookup_value(project.get("Id"))

    def _object_parent_display(self, obj):
        combined = normalize_object_lookup_value(obj.get("Combined"))
        code = normalize_object_lookup_value(obj.get("Code"))
        title = normalize_object_lookup_value(obj.get("Title"))
        if combined:
            return combined
        if code and title:
            return f"{code}. {title}"
        if title:
            return title
        if code:
            return code
        return normalize_object_lookup_value(obj.get("Id"))

    def _populate_objects_project_combo(self, filter_text=""):
        needle = normalize_object_lookup_value(filter_text).lower()

        self.objects_project_combo.blockSignals(True)
        self.objects_project_combo.clear()

        for source_index, project in enumerate(self.objects_projects):
            searchable = " ".join(
                [
                    normalize_object_lookup_value(project.get("Id")),
                    normalize_object_lookup_value(project.get("Code")),
                    normalize_object_lookup_value(project.get("Title")),
                    normalize_object_lookup_value(project.get("FullTitle")),
                ]
            ).lower()
            if needle and needle not in searchable:
                continue
            self.objects_project_combo.addItem(self._object_project_display(project), source_index)

        has_items = self.objects_project_combo.count() > 0
        self.objects_project_combo.setEnabled(has_items)
        self.objects_project_combo.blockSignals(False)

        if has_items:
            self.objects_project_combo.setCurrentIndex(0)
            self._on_objects_project_changed(0)
        else:
            self.selected_objects_project_id = None
            self.objects_project_info.setText("Id: -   |   Code: -")

    def _populate_objects_parent_combo(self, filter_text=""):
        needle = normalize_object_lookup_value(filter_text).lower()

        self.objects_parent_combo.blockSignals(True)
        self.objects_parent_combo.clear()

        for source_index, obj in enumerate(self.objects_items):
            searchable = " ".join(
                [
                    normalize_object_lookup_value(obj.get("Id")),
                    normalize_object_lookup_value(obj.get("Code")),
                    normalize_object_lookup_value(obj.get("Title")),
                    normalize_object_lookup_value(obj.get("Combined")),
                    normalize_object_lookup_value(obj.get("Path")),
                    normalize_object_lookup_value(obj.get("ParentCode")),
                    normalize_object_lookup_value(obj.get("ParentTitle")),
                    normalize_object_lookup_value(obj.get("ProjectTitle")),
                ]
            ).lower()
            if needle and needle not in searchable:
                continue
            self.objects_parent_combo.addItem(self._object_parent_display(obj), source_index)

        has_items = self.objects_parent_combo.count() > 0
        self.objects_parent_combo.setEnabled(has_items)
        self.objects_parent_combo.blockSignals(False)

        if has_items:
            self.objects_parent_combo.setCurrentIndex(0)
            self._on_objects_parent_changed(0)
        else:
            self.selected_object_parent_id = None
            self.objects_parent_info.setText("Id: -   |   Code: -")

    def _filter_objects_parent_combo(self, text):
        self._populate_objects_parent_combo(text)

    def _filter_objects_project_combo(self, text):
        self._populate_objects_project_combo(text)

    def _show_selected_project_details(self):
        index = self.objects_project_combo.currentIndex()
        if index < 0:
            QMessageBox.information(self, "Проект", "Проект не выбран.")
            return
        source_index = self.objects_project_combo.itemData(index)
        if source_index is None:
            QMessageBox.information(self, "Проект", "Проект не выбран.")
            return
        try:
            source_index = int(source_index)
        except Exception:
            QMessageBox.information(self, "Проект", "Проект не выбран.")
            return
        if source_index < 0 or source_index >= len(self.objects_projects):
            QMessageBox.information(self, "Проект", "Проект не выбран.")
            return
        project = self.objects_projects[source_index]
        details = "\n".join(
            [
                f"Id: {normalize_object_lookup_value(project.get('Id'))}",
                f"Code: {normalize_object_lookup_value(project.get('Code'))}",
                f"Title: {normalize_object_lookup_value(project.get('Title'))}",
                f"FullTitle: {normalize_object_lookup_value(project.get('FullTitle'))}",
            ]
        )
        QMessageBox.information(self, "Проект — Подробнее", details)

    def _show_selected_parent_details(self):
        index = self.objects_parent_combo.currentIndex()
        if index < 0:
            QMessageBox.information(self, "Родительский объект", "Родительский объект не выбран.")
            return
        source_index = self.objects_parent_combo.itemData(index)
        if source_index is None:
            QMessageBox.information(self, "Родительский объект", "Родительский объект не выбран.")
            return
        try:
            source_index = int(source_index)
        except Exception:
            QMessageBox.information(self, "Родительский объект", "Родительский объект не выбран.")
            return
        if source_index < 0 or source_index >= len(self.objects_items):
            QMessageBox.information(self, "Родительский объект", "Родительский объект не выбран.")
            return
        obj = self.objects_items[source_index]
        details = "\n".join(
            [
                f"Id: {normalize_object_lookup_value(obj.get('Id'))}",
                f"Code: {normalize_object_lookup_value(obj.get('Code'))}",
                f"Title: {normalize_object_lookup_value(obj.get('Title'))}",
                f"Combined: {normalize_object_lookup_value(obj.get('Combined'))}",
                f"ParentId: {normalize_object_lookup_value(obj.get('ParentId'))}",
                f"ParentCode: {normalize_object_lookup_value(obj.get('ParentCode'))}",
                f"ParentTitle: {normalize_object_lookup_value(obj.get('ParentTitle'))}",
                f"Path: {normalize_object_lookup_value(obj.get('Path'))}",
                f"ProjectId: {normalize_object_lookup_value(obj.get('ProjectId'))}",
                f"ProjectTitle: {normalize_object_lookup_value(obj.get('ProjectTitle'))}",
            ]
        )
        QMessageBox.information(self, "Родительский объект — Подробнее", details)

    def _invalidate_objects_preview(self, *_args):
        if hasattr(self, "objects_preview_table"):
            clear_inline_preview(self.objects_preview_table)
        if hasattr(self, "objects_preview_details_btn"):
            self.objects_preview_details_btn.setEnabled(False)
        if hasattr(self, "objects_file_edit") and self.objects_file_edit.text():
            self.objects_summary.setText("Лист изменён · запустите предпросмотр заново")

    def _objects_preview_specs(self):
        sheet = self.objects_sheet_combo.currentText() if self.objects_sheet_combo.isEnabled() else ""
        return [(sheet, ("Code", "Title"))]

    def _preview_objects_excel(self):
        sheets = run_inline_excel_preview(
            self, self.objects_preview_table, self.objects_file_edit.text(),
            self._objects_preview_specs(), self.objects_summary,
        )
        self.objects_preview_details_btn.setEnabled(bool(sheets))

    def _show_objects_preview_details(self):
        show_excel_preview(
            self, "Подробный предпросмотр — объекты строительства", self.objects_file_edit.text(),
            self._objects_preview_specs(),
        )

    def _log_objects(self, text):
        self.objects_log_signal.log.emit(text)

    def _show_objects_log(self):
        if self._objects_log_dialog is None or not self._objects_log_dialog.isVisible():
            self._objects_log_dialog = LogDialog(self)
            self._objects_log_dialog.setWindowTitle("Log — Объекты строительства")
            for chunk in self._objects_log_buffer:
                self._objects_log_dialog.append(chunk)
            self._objects_log_dialog.show()
        else:
            self._objects_log_dialog.raise_()
            self._objects_log_dialog.activateWindow()

    def _append_objects_log(self, text):
        if text == "__SHOW_OBJECTS_LOG__":
            self._show_objects_log()
            return
        self._objects_log_buffer.append(text)
        try:
            message = str(text).strip().splitlines()[-1]
            if message:
                self.statusBar().showMessage(message[:180], 8000)
        except Exception:
            pass
        if self._objects_log_dialog and self._objects_log_dialog.isVisible():
            self._objects_log_dialog.append(text)

    def _clear_objects_log(self):
        self._objects_log_buffer.clear()
        if self._objects_log_dialog:
            self._objects_log_dialog.clear_log()

    def _start_objects_load(self):
        if self._objects_parents_loading:
            return
        if not self.selected_objects_project_id:
            return
        auth_args = self._get_worker_auth_args()
        if not auth_args:
            return
        self._objects_parents_loading = True
        self.objects_load_btn.setEnabled(False)
        self._clear_objects_log()
        threading.Thread(
            target=self._run_objects_load_worker,
            args=(self.selected_objects_project_id, *auth_args),
            daemon=True,
        ).start()

    def _start_objects_projects_load(self):
        if self._objects_projects_loading:
            return
        auth_args = self._get_worker_auth_args()
        if not auth_args:
            return
        self._objects_projects_loading = True
        self.objects_projects_load_btn.setEnabled(False)
        self._clear_objects_log()
        threading.Thread(
            target=self._run_objects_projects_load_worker,
            args=auth_args,
            daemon=True,
        ).start()

    def _start_objects_create(self):
        file_path = self.objects_file_edit.text()
        if not file_path or not os.path.exists(file_path):
            QMessageBox.warning(self, "Ошибка", "Выберите существующий Excel файл")
            return
        sheet_name = self.objects_sheet_combo.currentText() if self.objects_sheet_combo.isEnabled() else 0
        if not sheet_name and sheet_name != 0:
            QMessageBox.warning(self, "Ошибка", "Выберите лист Excel")
            return

        if not self.selected_objects_project_id:
            QMessageBox.warning(self, "Ошибка", "Выберите проект")
            return

        parent_mode = self._get_objects_parent_mode()
        if parent_mode == "base" and not self.selected_object_parent_id:
            QMessageBox.warning(self, "Ошибка", "Выберите базовый родительский объект")
            return

        auth_args = self._get_worker_auth_args()
        if not auth_args:
            QMessageBox.warning(self, "Ошибка", "Заполните BASE_URL, USERNAME и PASSWORD")
            return

        self.objects_create_btn.setEnabled(False)
        self._clear_objects_log()
        threading.Thread(
            target=self._run_objects_create_worker,
            args=(file_path, sheet_name, parent_mode, self.selected_objects_project_id, self.selected_object_parent_id, *auth_args),
            daemon=True,
        ).start()

    def _apply_loaded_objects(self, items):
        self.objects_items = items or []
        self.objects_indexes = build_object_indexes(self.objects_items)
        self._populate_objects_parent_combo(self.objects_parent_search_edit.text())

    def _apply_loaded_projects(self, projects):
        self.objects_projects = projects or []
        self._populate_objects_project_combo(self.objects_project_search_edit.text())

    def _finish_objects_projects_load(self):
        self._objects_projects_loading = False
        self.objects_projects_load_btn.setEnabled(True)

    def _finish_objects_load(self):
        self._objects_parents_loading = False
        self.objects_load_btn.setEnabled(True)

    def _run_objects_projects_load_worker(
        self,
        auth_mode,
        client_id,
        base_url,
        username,
        password,
        sso_url,
        realm,
        broker_alias,
        adfs_url,
    ):
        try:
            self._log_objects("=" * 60 + "\n")
            self._log_objects("Загрузка списка проектов\n")
            self._log_objects(f"Base URL: {base_url}\n")
            self._log_objects(f"Auth mode: {auth_mode}\n")
            self._log_objects("=" * 60 + "\n\n")

            self._log_objects("[1/3] Авторизация...\n")
            access_token = get_access_token(
                session,
                auth_mode,
                client_id,
                base_url,
                username,
                password,
                sso_base_url=sso_url,
                realm=realm,
                broker_alias=broker_alias,
                adfs_base_url=adfs_url,
            )
            self._log_objects("  access_token получен\n\n")

            self._log_objects("[2/3] Загрузка проектов...\n")
            projects = get_projects_for_select(session, access_token, base_url)

            prepared_projects = []
            for item in projects:
                project = {
                    "Id": normalize_object_lookup_value(item.get("Id")),
                    "Code": normalize_object_lookup_value(item.get("Code")),
                    "Title": normalize_object_lookup_value(item.get("Title")),
                    "FullTitle": normalize_object_lookup_value(item.get("FullTitle")),
                }
                prepared_projects.append(project)

            self._log_objects(f"  Загружено проектов: {len(prepared_projects)}\n\n")
            self._log_objects("[3/3] Передача данных в UI поток...\n")
            self.objects_load_signal.projects_loaded.emit(prepared_projects)
            self._log_objects("  Список проектов обновлен\n")
        except Exception as e:
            self._log_objects(f"\nФАТАЛЬНАЯ ОШИБКА: {e}\n")
            self._log_objects(traceback.format_exc())
            self.objects_log_signal.log.emit("__SHOW_OBJECTS_LOG__")
        finally:
            self.objects_load_signal.projects_finished.emit()

    def _run_objects_load_worker(
        self,
        selected_project_id,
        auth_mode,
        client_id,
        base_url,
        username,
        password,
        sso_url,
        realm,
        broker_alias,
        adfs_url,
    ):
        try:
            self._log_objects("=" * 60 + "\n")
            self._log_objects("Загрузка структуры объектов строительства\n")
            self._log_objects(f"Base URL: {base_url}\n")
            self._log_objects(f"Auth mode: {auth_mode}\n")
            self._log_objects("=" * 60 + "\n\n")

            self._log_objects("[1/3] Авторизация...\n")
            access_token = get_access_token(
                session,
                auth_mode,
                client_id,
                base_url,
                username,
                password,
                sso_base_url=sso_url,
                realm=realm,
                broker_alias=broker_alias,
                adfs_base_url=adfs_url,
            )
            self._log_objects("  access_token получен\n\n")

            self._log_objects("[2/3] Загрузка объектов структуры...\n")
            safe_project_id = normalize_object_lookup_value(selected_project_id)
            if not safe_project_id:
                raise Exception("Не выбран ProjectId для загрузки структуры объектов")
            self._log_objects(f"  ProjectId={safe_project_id}\n")
            items = get_object_structures_odata(session, access_token, base_url, safe_project_id)
            self._log_objects("  OData загрузка выполнена успешно\n")

            prepared_items = []
            for item in items:
                obj = {
                    "Id": normalize_object_lookup_value(item.get("Id")),
                    "Code": normalize_object_lookup_value(item.get("Code")),
                    "Title": normalize_object_lookup_value(item.get("Title")),
                    "Combined": normalize_object_lookup_value(item.get("Combined")),
                    "ParentId": normalize_object_lookup_value(item.get("ParentId")),
                    "ParentTitle": normalize_object_lookup_value(item.get("ParentTitle")),
                    "ParentCode": normalize_object_lookup_value(item.get("ParentCode")),
                    "ParentCombined": normalize_object_lookup_value(item.get("ParentCombined")),
                    "Path": normalize_object_lookup_value(item.get("Path")),
                    "ProjectId": normalize_object_lookup_value(item.get("ProjectId")),
                    "ProjectTitle": normalize_object_lookup_value(item.get("ProjectTitle")),
                }
                prepared_items.append(obj)

            self._log_objects(f"  Загружено объектов: {len(prepared_items)}\n\n")
            self._log_objects("[3/3] Передача данных в UI поток...\n")
            self.objects_load_signal.loaded.emit(prepared_items)
            self._log_objects("  Список родителей обновлен\n")
        except Exception as e:
            self._log_objects(f"\nФАТАЛЬНАЯ ОШИБКА: {e}\n")
            self._log_objects(traceback.format_exc())
            self.objects_log_signal.log.emit("__SHOW_OBJECTS_LOG__")
        finally:
            self.objects_load_signal.finished.emit()

    def _run_objects_create_worker(
        self,
        excel_file,
        sheet_name,
        parent_mode,
        selected_project_id,
        selected_parent_id,
        auth_mode,
        client_id,
        base_url,
        username,
        password,
        sso_url,
        realm,
        broker_alias,
        adfs_url,
    ):
        try:
            self._log_objects("=" * 60 + "\n")
            self._log_objects("Запуск создания объектов строительства из Excel\n")
            self._log_objects(f"Base URL: {base_url}\n")
            self._log_objects(f"Auth mode: {auth_mode}\n")
            self._log_objects(f"ProjectId: {normalize_object_lookup_value(selected_project_id)}\n")
            self._log_objects(f"Режим родителя: {'Из Excel по ParentCode' if parent_mode == 'excel' else 'Один базовый родитель для всех строк'}\n")
            self._log_objects(f"Excel: {excel_file}\n")
            self._log_objects(f"Лист: {sheet_name}\n")
            self._log_objects("=" * 60 + "\n\n")

            self._log_objects("[1/5] Чтение Excel...\n")
            records = read_objects_import_excel(excel_file, sheet_name)
            self._log_objects(f"  Загружено строк: {len(records)}\n\n")

            self._log_objects("[2/5] Авторизация...\n")
            access_token = get_access_token(
                session,
                auth_mode,
                client_id,
                base_url,
                username,
                password,
                sso_base_url=sso_url,
                realm=realm,
                broker_alias=broker_alias,
                adfs_base_url=adfs_url,
            )
            self._log_objects("  access_token получен\n\n")

            self._log_objects("[3/5] Загрузка текущей структуры объектов проекта...\n")
            safe_project_id = normalize_object_lookup_value(selected_project_id)
            if not safe_project_id:
                raise Exception("Не выбран ProjectId для загрузки структуры объектов")
            self._log_objects(f"  ProjectId={safe_project_id}\n")
            fresh_items = get_object_structures_odata(session, access_token, base_url, safe_project_id)
            self._log_objects(f"  Загружено объектов проекта: {len(fresh_items)}\n\n")

            indexes = build_object_indexes(fresh_items)

            self._log_objects("[4/5] Предварительная валидация...\n")
            self._log_objects(
                f"  Исходные индексы: codes={len(indexes['code_to_id'])}, titles={len(indexes['title_to_objects'])}\n"
            )
            records, validation_errors = validate_and_order_objects_import_records(records, indexes, parent_mode, selected_parent_id)
            if validation_errors:
                self._log_objects("  ОШИБКИ ВАЛИДАЦИИ:\n")
                for error_text in validation_errors:
                    self._log_objects(f"    {error_text}\n")
                self._log_objects("\nРабота остановлена. Исправьте ошибки и повторите.\n")
                self.objects_log_signal.log.emit("__SHOW_OBJECTS_LOG__")
                return
            self._log_objects("  Валидация пройдена\n\n")

            self._log_objects("[5/5] Создание объектов...\n")
            self._log_objects("-" * 60 + "\n")
            stats = {"created": 0, "skipped": 0, "errors": 0}
            created_objects = []

            for i, row in enumerate(records, 1):
                code = normalize_object_lookup_value(row.get("Code"))
                title = normalize_object_lookup_value(row.get("Title"))
                parent_code = normalize_object_lookup_value(row.get("ParentCode"))
                parent_title = normalize_object_lookup_value(row.get("ParentTitle"))
                row_num = row.get("RowNum")

                self._log_objects(
                    f"\n  [{i}] row={row_num}, Code='{code}', Title='{title}', ParentCode='{parent_code}', ParentTitle='{parent_title}'\n"
                )

                if not code or not title:
                    self._log_objects("      ОШИБКА: обязательные поля Code и Title должны быть заполнены\n")
                    stats["errors"] += 1
                    continue

                parent_id, parent_error = resolve_object_parent_id(row, indexes, parent_mode, selected_parent_id)
                if parent_error or not parent_id:
                    self._log_objects(f"      ОШИБКА: {parent_error or 'не удалось определить ParentId'}\n")
                    stats["errors"] += 1
                    continue

                payload = {
                    "Code": code,
                    "Title": title,
                    "ParentId": parent_id,
                    "IsActive": True,
                    "Order": 0,
                    "Url": "",
                }

                ok, status_code, resp_data = create_object_structure(session, access_token, base_url, payload)
                if not ok:
                    self._log_objects(f"      ОШИБКА: HTTP {status_code}\n")
                    self._log_objects(f"      Ответ: {str(resp_data)[:400]}\n")
                    stats["errors"] += 1
                    continue

                new_id = ""
                if isinstance(resp_data, dict):
                    new_id = normalize_object_lookup_value(resp_data.get("Id"))
                    if not new_id:
                        data_payload = resp_data.get("Data")
                        if isinstance(data_payload, dict):
                            new_id = normalize_object_lookup_value(data_payload.get("Id"))
                elif isinstance(resp_data, str):
                    new_id = normalize_object_lookup_value(resp_data)

                if not new_id:
                    self._log_objects("      ОШИБКА: объект создан, но Id в ответе не найден\n")
                    stats["errors"] += 1
                    continue

                created_obj = {
                    "Id": new_id,
                    "Code": code,
                    "Title": title,
                    "Combined": "",
                    "ParentId": normalize_object_lookup_value(parent_id),
                    "ParentTitle": "",
                    "ParentCode": parent_code,
                    "ParentCombined": "",
                    "Path": "",
                    "ProjectId": "",
                    "ProjectTitle": "",
                }
                add_object_to_indexes(indexes, created_obj)
                created_objects.append(created_obj)
                self._log_objects(f"      СОЗДАНО: Id={new_id}\n")
                stats["created"] += 1

            if created_objects:
                self.objects_items = list(fresh_items)
                self.objects_items.extend(created_objects)
                self.objects_indexes = indexes

            self._log_objects("\n" + "=" * 60 + "\n")
            self._log_objects("ИТОГИ\n")
            self._log_objects(f"  Создано: {stats['created']}\n")
            self._log_objects(f"  Пропущено: {stats['skipped']}\n")
            self._log_objects(f"  Ошибки: {stats['errors']}\n")
            self._log_objects("=" * 60 + "\n")
            self.objects_log_signal.log.emit("__SHOW_OBJECTS_LOG__")
        except Exception as e:
            self._log_objects(f"\nФАТАЛЬНАЯ ОШИБКА: {e}\n")
            self._log_objects(traceback.format_exc())
            self.objects_log_signal.log.emit("__SHOW_OBJECTS_LOG__")
        finally:
            self.objects_log_signal.finished.emit()
