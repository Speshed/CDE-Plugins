from __future__ import annotations

import json
import os
import re
import threading
import traceback

import pandas as pd
from PySide6.QtWidgets import (
    QComboBox, QFileDialog, QFrame, QGridLayout, QHBoxLayout, QLabel, QLineEdit,
    QMessageBox, QPushButton, QVBoxLayout, QWidget,
)

from ...auth import *
from ...config import *
from ...api import *
from ...excel import *
from ..dialogs import LogDialog
from ..components import (
    clear_inline_preview, make_file_row, make_info_banner, make_preview_table, make_preview_header, mark_file_button,
    run_inline_excel_preview, show_excel_preview,
)


class RoutesMixin:
    def _build_routes_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(2, 8, 2, 2)
        layout.setSpacing(8)

        file_card = QFrame()
        file_card.setObjectName("card")
        file_layout = QVBoxLayout(file_card)
        file_layout.setContentsMargins(14, 8, 14, 9)
        file_layout.setSpacing(6)
        self.routes_file_edit = QLineEdit()
        self.routes_file_edit.setReadOnly(True)
        self.routes_file_edit.setPlaceholderText("Excel-файл не выбран")
        self.routes_template_btn = QPushButton("Скачать шаблон")
        self.routes_template_btn.clicked.connect(self._start_routes_template)
        self.routes_file_btn = QPushButton("Загрузить файл")
        self.routes_file_btn.clicked.connect(self._choose_routes_file)
        file_layout.addWidget(make_file_row(
            self.asset_dir, self.is_dark_theme,
            "Excel-файл маршрутов согласований",
            "Шаблон содержит маршруты, условия и продолжительность этапов",
            self.routes_template_btn, self.routes_file_btn,
        ))

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(6)
        labels = ["Маршруты", "Условия маршрута", "Продолжительность"]
        combos = []
        for row_index, text in enumerate(labels):
            label = QLabel(text)
            label.setObjectName("fieldLabel")
            combo = QComboBox()
            combo.setEnabled(False)
            grid.addWidget(label, row_index, 0)
            grid.addWidget(combo, row_index, 1)
            combos.append(combo)
        self.routes_sheet_combo, self.routes_conditions_sheet_combo, self.routes_duration_sheet_combo = combos
        for combo in combos:
            combo.currentTextChanged.connect(self._invalidate_routes_preview)
        grid.setColumnStretch(1, 1)
        file_layout.addLayout(grid)
        layout.addWidget(file_card)

        action_card = QFrame()
        action_card.setObjectName("card")
        action_layout = QVBoxLayout(action_card)
        action_layout.setContentsMargins(14, 10, 14, 11)
        action_layout.setSpacing(8)
        title = make_preview_header("Создание маршрутов согласований")
        action_layout.addWidget(title)
        action_layout.addWidget(make_info_banner(
            self.asset_dir, self.is_dark_theme,
            "Предпросмотр проверяет выбранные три листа одновременно. Никакие маршруты при предпросмотре не создаются.",
        ))
        self.routes_preview_table = make_preview_table()
        action_layout.addWidget(self.routes_preview_table)
        self.routes_summary = QLabel("Загрузите Excel-файл и проверьте соответствие листов")
        self.routes_summary.setObjectName("rowSubtitle")
        self.routes_summary.setWordWrap(True)
        footer = QHBoxLayout()
        footer.setSpacing(8)
        footer.addWidget(self.routes_summary, 1)
        routes_log_btn = QPushButton("Log")
        routes_log_btn.setObjectName("logButton")
        routes_log_btn.clicked.connect(self._show_route_log)
        footer.addWidget(routes_log_btn)
        self.routes_details_btn = QPushButton("Подробнее")
        self.routes_details_btn.setObjectName("modeSwitch")
        self.routes_details_btn.setEnabled(False)
        self.routes_details_btn.clicked.connect(self._show_routes_preview_details)
        footer.addWidget(self.routes_details_btn)
        self.routes_preview_btn = QPushButton("Предпросмотр")
        self.routes_preview_btn.setObjectName("loginButton")
        self.routes_preview_btn.setEnabled(False)
        self.routes_preview_btn.clicked.connect(self._preview_routes_excel)
        footer.addWidget(self.routes_preview_btn)
        self.routes_start_btn = QPushButton("Создать маршруты")
        self.routes_start_btn.setObjectName("greenAction")
        self.routes_start_btn.setMinimumWidth(190)
        self.routes_start_btn.clicked.connect(self._start_routes)
        footer.addWidget(self.routes_start_btn)
        action_layout.addLayout(footer)
        layout.addWidget(action_card)
        layout.addSpacing(2)
        self.tabs.addTab(tab, "Согласования")

    def _choose_routes_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл", "", "Excel (*.xlsx *.xls)"
        )
        if path:
            self.routes_file_edit.setText(path)
            mark_file_button(self.routes_file_btn, True)
            self._refresh_action_icons()
            try:
                xls = pd.ExcelFile(path)
                sheets = xls.sheet_names

                # Fill routes sheet combo
                self.routes_sheet_combo.blockSignals(True)
                self.routes_sheet_combo.clear()
                self.routes_sheet_combo.addItems(sheets)
                # Try to find default: sheet containing "Маршруты"
                default_routes_idx = 0
                for i, name in enumerate(sheets):
                    if "Маршруты" in name:
                        default_routes_idx = i
                        break
                self.routes_sheet_combo.setCurrentIndex(default_routes_idx)
                self.routes_sheet_combo.setEnabled(True)
                self.routes_sheet_combo.blockSignals(False)

                # Fill conditions sheet combo
                self.routes_conditions_sheet_combo.blockSignals(True)
                self.routes_conditions_sheet_combo.clear()
                self.routes_conditions_sheet_combo.addItems(sheets)
                # Try to find default with priority:
                # 1. sheet containing "Настройка"
                # 2. sheet containing "Условия"
                # 3. sheet containing "5.1"
                default_cond_idx = 0
                for i, name in enumerate(sheets):
                    if "Настройка" in name:
                        default_cond_idx = i
                        break
                else:
                    for i, name in enumerate(sheets):
                        if "Условия" in name:
                            default_cond_idx = i
                            break
                    else:
                        for i, name in enumerate(sheets):
                            if "5.1" in name:
                                default_cond_idx = i
                                break
                self.routes_conditions_sheet_combo.setCurrentIndex(default_cond_idx)
                self.routes_conditions_sheet_combo.setEnabled(True)
                self.routes_conditions_sheet_combo.blockSignals(False)

                # Fill duration sheet combo (5.2)
                self.routes_duration_sheet_combo.blockSignals(True)
                self.routes_duration_sheet_combo.clear()
                self.routes_duration_sheet_combo.addItems(sheets)
                # Try to find default with priority:
                # 1. sheet containing "5.2"
                # 2. sheet containing "Продолжительность"
                # 3. sheet containing "Настройка"
                default_duration_idx = 0
                for i, name in enumerate(sheets):
                    if "5.2" in name:
                        default_duration_idx = i
                        break
                else:
                    for i, name in enumerate(sheets):
                        if "Продолжительность" in name:
                            default_duration_idx = i
                            break
                    else:
                        for i, name in enumerate(sheets):
                            if "Настройка" in name:
                                default_duration_idx = i
                                break
                self.routes_duration_sheet_combo.setCurrentIndex(default_duration_idx)
                self.routes_duration_sheet_combo.setEnabled(True)
                self.routes_duration_sheet_combo.blockSignals(False)
                self.routes_preview_btn.setEnabled(bool(sheets))
                self.routes_details_btn.setEnabled(False)
                clear_inline_preview(self.routes_preview_table)
                self.routes_summary.setText(f"Файл выбран · листов: {len(sheets)} · запустите предпросмотр")

            except Exception as e:
                self.routes_sheet_combo.clear()
                self.routes_sheet_combo.setEnabled(False)
                self.routes_conditions_sheet_combo.clear()
                self.routes_conditions_sheet_combo.setEnabled(False)
                self.routes_duration_sheet_combo.clear()
                self.routes_duration_sheet_combo.setEnabled(False)
                self.routes_preview_btn.setEnabled(False)
                self.routes_details_btn.setEnabled(False)
                self.routes_summary.setText("Не удалось прочитать Excel-файл")
                QMessageBox.warning(self, "Ошибка", f"Не удалось прочитать листы: {e}")

    def _invalidate_routes_preview(self, *_args):
        if hasattr(self, "routes_preview_table"):
            clear_inline_preview(self.routes_preview_table)
        if hasattr(self, "routes_details_btn"):
            self.routes_details_btn.setEnabled(False)
        if hasattr(self, "routes_file_edit") and self.routes_file_edit.text():
            self.routes_summary.setText("Выбор листов изменён · запустите предпросмотр заново")

    def _routes_preview_specs(self):
        return [
            (self.routes_sheet_combo.currentText(), ("Маршрут", "Проект")),
            (self.routes_conditions_sheet_combo.currentText(), ("Маршрут", "Вид документа")),
            (self.routes_duration_sheet_combo.currentText(), ("Маршрут", "Этап 1, дней")),
        ]

    def _preview_routes_excel(self):
        sheets = run_inline_excel_preview(
            self, self.routes_preview_table, self.routes_file_edit.text(),
            self._routes_preview_specs(), self.routes_summary,
        )
        self.routes_details_btn.setEnabled(bool(sheets))

    def _show_routes_preview_details(self):
        show_excel_preview(
            self, "Подробный предпросмотр — маршруты согласований", self.routes_file_edit.text(),
            self._routes_preview_specs(),
        )

    def _log_r(self, text):
        self.route_log_signal.log.emit(text)

    def _show_route_log(self):
        if self._route_log_dialog is None or not self._route_log_dialog.isVisible():
            self._route_log_dialog = LogDialog(self)
            self._route_log_dialog.setWindowTitle("Log — Маршруты согласований")
            for chunk in self._route_log_buffer:
                self._route_log_dialog.append(chunk)
            self._route_log_dialog.show()
        else:
            self._route_log_dialog.raise_()
            self._route_log_dialog.activateWindow()

    def _append_route_log(self, text):
        if text == "__SHOW_ROUTE_LOG__":
            self._show_route_log()
            return
        self._route_log_buffer.append(text)
        try:
            message = str(text).strip().splitlines()[-1]
            if message:
                self.statusBar().showMessage(message[:180], 8000)
        except Exception:
            pass
        if self._route_log_dialog and self._route_log_dialog.isVisible():
            self._route_log_dialog.append(text)

    def _clear_routes_log(self):
        self._route_log_buffer.clear()
        if self._route_log_dialog:
            self._route_log_dialog.clear_log()

    def _start_routes_template(self):
        self._save_template_file(
            "Сохранить шаблон маршрутов согласований",
            "Project Point Согласования.xlsx",
            write_routes_import_template,
        )

    def _start_routes(self):
        file_path = self.routes_file_edit.text()
        if not file_path or not os.path.exists(file_path):
            QMessageBox.warning(self, "Ошибка", "Выберите существующий Excel файл")
            return

        # Validate sheet selections
        if not self.routes_sheet_combo.isEnabled() or self.routes_sheet_combo.currentIndex() < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите лист 'Маршруты'")
            return
        if not self.routes_conditions_sheet_combo.isEnabled() or self.routes_conditions_sheet_combo.currentIndex() < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите лист 'Условия маршрута'")
            return
        if not self.routes_duration_sheet_combo.isEnabled() or self.routes_duration_sheet_combo.currentIndex() < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите лист 'Продолжительность'")
            return

        routes_sheet = self.routes_sheet_combo.currentText()
        conditions_sheet = self.routes_conditions_sheet_combo.currentText()
        duration_sheet = self.routes_duration_sheet_combo.currentText()

        auth_args = self._get_worker_auth_args()
        if not auth_args:
            QMessageBox.warning(self, "Ошибка", "Заполните BASE_URL, USERNAME и PASSWORD")
            return

        self.routes_start_btn.setEnabled(False)
        self._clear_routes_log()
        threading.Thread(
            target=self._run_routes_worker,
            args=(file_path, routes_sheet, conditions_sheet, duration_sheet, *auth_args),
            daemon=True,
        ).start()

    def _run_routes_worker(
        self,
        excel_file,
        routes_sheet_name,
        conditions_sheet_name,
        duration_sheet_name,
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
            self._log_r("=" * 60 + "\n")
            self._log_r("Запуск процесса создания маршрутов согласований\n")
            self._log_r(f"Base URL: {base_url}\n")
            self._log_r(f"Auth mode: {auth_mode}\n")
            self._log_r("=" * 60 + "\n\n")

            # Step 1: Read Excel files
            self._log_r("[1/9] Чтение Excel файлов...\n")

            self._log_r(f"  Лист маршрутов: '{routes_sheet_name}'\n")
            self._log_r(f"  Лист условий: '{conditions_sheet_name}'\n")
            self._log_r(f"  Лист продолжительности: '{duration_sheet_name}'\n")

            def normalize_string(val):
                if pd.isna(val):
                    return ""
                return str(val).strip()

            def collapse_spaces(s):
                """Collapse multiple whitespace characters into single space."""
                return re.sub(r'\s+', ' ', s).strip()

            def parse_excel_checkbox(val):
                """Parse Excel checkbox values: ✓, ✔ -> True, ✗ or empty -> False."""
                if pd.isna(val):
                    return False
                s = str(val).strip()
                # Check for checkmark symbols (U+2713, U+2714) and common truthy strings
                truthy_symbols = {"✓", "✔"}
                if s in truthy_symbols:
                    return True
                # Also support text-based truthy values (case-insensitive)
                s_lower = s.lower()
                truthy_strings = {"true", "1", "да", "yes"}
                return s_lower in truthy_strings

            def parse_stage_duration(val):
                """Parse stage duration: convert to int if valid, else None."""
                if pd.isna(val):
                    return None
                try:
                    return int(float(val))
                except (ValueError, TypeError):
                    return None

            def read_sheet_52_stage_settings(excel_file_path, sheet_name):
                """Read and parse sheet 5.2 stage settings.

                Args:
                    excel_file_path: path to Excel file
                    sheet_name: name of the duration/stage settings sheet

                Returns:
                    dict: {route_name: {stage_num: settings}}
                    or empty dict if sheet not found
                """
                if not sheet_name:
                    return {}

                try:
                    # Read sheet 5.2 - data starts from row 6 (index 5 in pandas)
                    df52_raw = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=None)
                    df52 = df52_raw.iloc[5:].reset_index(drop=True)  # Skip first 5 rows
                except Exception as e:
                    return {}

                stage_settings = {}

                # Column mapping for sheet 5.2 (updated structure with Утверждающий column):
                # col 0: route_name
                # col 1-6: WorkingDaysDuration for stages 1-6
                # col 7: ApproveWorkingDaysDuration (Утверждающий)
                # col 11-14: Stage 1 flags (AutoComplete, CompleteOnSingleTaskCompleted, EndProcessIfRejected, DigitalSignatureRequired)
                # col 15-18: Stage 2 flags
                # col 19-22: Stage 3 flags
                # col 23-26: Stage 4 flags
                # col 27-30: Stage 5 flags
                # col 31-34: Stage 6 flags
                # col 35: ApproveDigitalSignatureRequired (if present)

                STAGE_FLAG_COLS = {
                    1: {"auto": 11, "single": 12, "reject": 13, "sign": 14},
                    2: {"auto": 15, "single": 16, "reject": 17, "sign": 18},
                    3: {"auto": 19, "single": 20, "reject": 21, "sign": 22},
                    4: {"auto": 23, "single": 24, "reject": 25, "sign": 26},
                    5: {"auto": 27, "single": 28, "reject": 29, "sign": 30},
                    6: {"auto": 31, "single": 32, "reject": 33, "sign": 34},
                }

                # Column for approve duration (Утверждающий)
                APPROVE_DURATION_COL = 7
                # Column for approve digital signature (optional, col 35)
                APPROVE_SIGN_COL = 35

                for idx, row in df52.iterrows():
                    route_name = normalize_string(row.iloc[0] if len(row) > 0 else "")
                    if not route_name:
                        continue

                    route_stages = {}

                    for stage_num in range(1, 7):
                        # Parse duration from columns 1-6
                        duration_col = stage_num  # col 1 for stage 1, col 2 for stage 2, etc.
                        duration = parse_stage_duration(row.iloc[duration_col] if len(row) > duration_col else None)

                        # Parse flags from flag columns
                        flag_cols = STAGE_FLAG_COLS.get(stage_num, {})
                        auto_complete = parse_excel_checkbox(row.iloc[flag_cols.get("auto", 0)] if len(row) > flag_cols.get("auto", 0) else None)
                        single_complete = parse_excel_checkbox(row.iloc[flag_cols.get("single", 0)] if len(row) > flag_cols.get("single", 0) else None)
                        end_on_reject = parse_excel_checkbox(row.iloc[flag_cols.get("reject", 0)] if len(row) > flag_cols.get("reject", 0) else None)
                        digital_sign = parse_excel_checkbox(row.iloc[flag_cols.get("sign", 0)] if len(row) > flag_cols.get("sign", 0) else None)

                        route_stages[stage_num] = {
                            "WorkingDaysDuration": duration,
                            "AutoComplete": auto_complete,
                            "CompleteOnSingleTaskCompleted": single_complete,
                            "EndProcessIfRejected": end_on_reject,
                            "DigitalSignatureRequired": digital_sign,
                        }

                    # Parse approve duration (Утверждающий)
                    approve_duration = parse_stage_duration(row.iloc[APPROVE_DURATION_COL] if len(row) > APPROVE_DURATION_COL else None)

                    # Parse approve digital signature (optional, col 35)
                    approve_digital_sign = parse_excel_checkbox(row.iloc[APPROVE_SIGN_COL] if len(row) > APPROVE_SIGN_COL else None)

                    stage_settings[route_name] = {
                        "stages": route_stages,
                        "approve": {
                            "ApproveWorkingDaysDuration": approve_duration,
                            "ApproveDigitalSignatureRequired": approve_digital_sign,
                        }
                    }

                return stage_settings

            # Read sheet 5.2 using the user-selected sheet name
            stage_settings_raw = read_sheet_52_stage_settings(excel_file, duration_sheet_name)

            if stage_settings_raw:
                self._log_r(f"  Лист продолжительности '{duration_sheet_name}': {len(stage_settings_raw)} строк настроек этапов\n")
                # Debug: log parsed values for first few routes
                for i, (route_name, settings) in enumerate(stage_settings_raw.items()):
                    if i >= 3:  # Log only first 3 routes to avoid spam
                        break
                    approve_cfg = settings.get("approve", {})
                    approve_duration = approve_cfg.get("ApproveWorkingDaysDuration")
                    approve_sign = approve_cfg.get("ApproveDigitalSignatureRequired")
                    self._log_r(f"    [{route_name}] ApproveWorkingDaysDuration={approve_duration}, ApproveDigitalSignatureRequired={approve_sign}\n")
                    # Log stage 3 flags as example
                    stages = settings.get("stages", {})
                    if 3 in stages:
                        s3 = stages[3]
                        self._log_r(f"    [{route_name}] Этап 3: AutoComplete={s3.get('AutoComplete')}, CompleteOnSingleTaskCompleted={s3.get('CompleteOnSingleTaskCompleted')}, EndProcessIfRejected={s3.get('EndProcessIfRejected')}, DigitalSignatureRequired={s3.get('DigitalSignatureRequired')}\n")
            else:
                self._log_r(f"  ВНИМАНИЕ: Не удалось прочитать лист '{duration_sheet_name}', применяются значения этапов по умолчанию\n")
                stage_settings_raw = {}
            
            # Read routes sheet - skip header rows
            df5_raw = pd.read_excel(excel_file, sheet_name=routes_sheet_name, header=None)
            df5 = df5_raw.iloc[4:].reset_index(drop=True)  # Skip first 4 rows
            
            # Parse routes sheet with proper stage grouping
            # Columns structure based on actual Excel:
            # 0: маршрут, 1: проект, 2: администратор, 3: тип_согласования
            # Stage columns groups: 1=4-13 (10 cols), 2=14-16 (3 cols), 3=17-18 (2 cols), 
            #                       4=19-20 (2 cols), 5=21-22 (2 cols), 6=23-24 (2 cols)
            STAGE_COL_GROUPS = {
                1: list(range(4, 14)),   # cols 4-13
                2: list(range(14, 17)),  # cols 14-16
                3: list(range(17, 19)),  # cols 17-18
                4: list(range(19, 21)),  # cols 19-20
                5: list(range(21, 23)),  # cols 21-22
                6: list(range(23, 25)),  # cols 23-24
            }
            
            routes_data = []
            for idx, row in df5.iterrows():
                route_name = normalize_string(row.iloc[0])
                if not route_name:
                    continue
                
                # Parse stages by column groups
                stages = []
                for stage_num, col_indices in STAGE_COL_GROUPS.items():
                    stage_users = []
                    for col_idx in col_indices:
                        if col_idx < len(row):
                            user_val = normalize_string(row.iloc[col_idx])
                            if user_val:
                                stage_users.append(user_val)
                    if stage_users:
                        stages.append({
                            'stage_num': stage_num,
                            'users': stage_users
                        })
                
                routes_data.append({
                    'row_idx': idx,
                    'route_name': route_name,
                    'project': normalize_string(row.iloc[1] if len(row) > 1 else ""),
                    'admin': normalize_string(row.iloc[2] if len(row) > 2 else ""),
                    'approval_type': normalize_string(row.iloc[3] if len(row) > 3 else ""),
                    'stages': stages,
                })
            
            # Read conditions sheet (route configuration)
            df51_raw = pd.read_excel(excel_file, sheet_name=conditions_sheet_name, header=None)
            df51 = df51_raw.iloc[5:].reset_index(drop=True)  # header is row 5; data starts at row 6
            
            config_data = []
            for idx, row in df51.iterrows():
                route_name = normalize_string(row.iloc[0] if len(row) > 0 else "")
                if not route_name:
                    continue
                
                config_data.append({
                    'row_idx': idx,
                    'route_name': route_name,
                    'вид_документа': normalize_string(row.iloc[1] if len(row) > 1 else ""),
                    'тип_документа': normalize_string(row.iloc[2] if len(row) > 2 else ""),
                    'проект': normalize_string(row.iloc[3] if len(row) > 3 else ""),
                    'объект_строительства': normalize_string(row.iloc[4] if len(row) > 4 else ""),
                    'заказчик': normalize_string(row.iloc[5] if len(row) > 5 else ""),
                    'разработчик': normalize_string(row.iloc[6] if len(row) > 6 else ""),
                    'дисциплина': normalize_string(row.iloc[7] if len(row) > 7 else ""),
                    'категория': normalize_string(row.iloc[8] if len(row) > 8 else ""),
                    'цель_выпуска': normalize_string(row.iloc[9] if len(row) > 9 else ""),
                })
            
            self._log_r(f"  Лист маршрутов '{routes_sheet_name}': {len(routes_data)} строк\n")
            self._log_r(f"  Лист условий '{conditions_sheet_name}': {len(config_data)} строк\n")
            
            # Step 2: Build occurrence index for merge (to handle duplicate route names like АХОВ_ПЗУ)
            self._log_r("\n[2/9] Построение occurrence index для merge...\n")
            
            def add_occurrence_index(data_list):
                """Add occurrence index for duplicate handling."""
                name_counts = {}
                for item in data_list:
                    name = item['route_name']
                    if name not in name_counts:
                        name_counts[name] = 0
                    name_counts[name] += 1
                    item['occurrence'] = name_counts[name]
                return data_list
            
            routes_data = add_occurrence_index(routes_data)
            config_data = add_occurrence_index(config_data)

            # Build stage_settings with occurrence index for 5.2
            # stage_settings_raw is keyed by route_name only, need to add occurrence
            stage_settings = {}
            name_counts_52 = {}
            for route_name, stages_cfg in stage_settings_raw.items():
                if route_name not in name_counts_52:
                    name_counts_52[route_name] = 0
                name_counts_52[route_name] += 1
                occurrence = name_counts_52[route_name]
                stage_settings[(route_name, occurrence)] = stages_cfg

            # Merge by route_name + occurrence
            merged_data = []
            routes_by_key = {(r['route_name'], r['occurrence']): r for r in routes_data}
            
            for config in config_data:
                key = (config['route_name'], config['occurrence'])
                if key in routes_by_key:
                    route = routes_by_key[key]
                    merged_data.append({
                        'index': len(merged_data) + 1,
                        'route_name': route['route_name'],
                        'occurrence': route['occurrence'],
                        'project_from_route': route['project'],
                        'admin': route['admin'],
                        'approval_type': route['approval_type'],
                        'stages': route['stages'],
                        'вид_документа': config['вид_документа'],
                        'тип_документа': config['тип_документа'],
                        'project_from_config': config['проект'],
                        'object_structure': config['объект_строительства'],
                        'заказчик': config['заказчик'],
                        'разработчик': config['разработчик'],
                        'дисциплина': config['дисциплина'],
                        'категория': config['категория'],
                        'цель_выпуска': config['цель_выпуска'],
                    })
                else:
                    self._log_r(f"  ВНИМАНИЕ: Не найден маршрут для config: {key}\n")
            
            self._log_r(f"  Объединено записей: {len(merged_data)}\n\n")

            # ============================================================================
            # UNIFIED REFERENCE RESOLUTION SYSTEM
            # ============================================================================

            def normalize_lookup_string(value):
                """Normalize string for lookup comparison.

                Performs:
                - trim
                - collapse multiple spaces
                - preserve case for exact match, use lower() for case-insensitive
                """
                if value is None:
                    return ""
                if pd.isna(value):
                    return ""
                s = str(value).strip()
                return collapse_spaces(s)

            def parse_code_title_value(raw_value):
                """Universal parser for 'Code + Title' format.

                Supports:
                - numeric codes: '11.001', '002'
                - text codes: 'Тест.1', 'ABC-01', 'КПП'
                - mixed codes: 'A-123', 'Test_01'

                Returns dict:
                - full_normalized: normalized full string
                - code: extracted code or None
                - title: extracted title or None
                """
                if not raw_value:
                    return {"full_normalized": "", "code": None, "title": None}

                full_normalized = normalize_lookup_string(raw_value)
                if not full_normalized:
                    return {"full_normalized": "", "code": None, "title": None}

                # Pattern: alphanumeric code (letters, digits, dots, hyphens, underscores)
                # followed by separator (dot or space) and title
                match = re.match(r'^([A-Za-zА-Яа-я0-9_.-]+)[\.\s]+(.+)$', full_normalized)
                if match:
                    extracted_code = match.group(1).strip()
                    extracted_title = normalize_lookup_string(match.group(2))
                    if extracted_code and extracted_title:
                        return {
                            "full_normalized": full_normalized,
                            "code": extracted_code,
                            "title": extracted_title,
                        }

                # Pattern: just code without title
                if re.match(r'^[A-Za-zА-Яа-я0-9_.-]+$', full_normalized):
                    return {
                        "full_normalized": full_normalized,
                        "code": full_normalized,
                        "title": None,
                    }

                # Fallback: treat entire string as title
                return {
                    "full_normalized": full_normalized,
                    "code": None,
                    "title": full_normalized,
                }

            def build_generic_lookup(items, code_field=None, title_field=None, path_field=None, name_field=None):
                """Build lookup structure for reference resolution.

                Args:
                    items: list of dicts from API
                    code_field: field name for code (e.g., 'Code')
                    title_field: field name for title (e.g., 'Title', 'FullName')
                    path_field: field name for path (e.g., 'Path')
                    name_field: alternative name field (e.g., 'Name')

                Returns dict with lookup indices.
                """
                lookup = {
                    "by_id": {},
                    "by_code": {},
                    "by_code_lower": {},
                    "by_title": {},
                    "by_title_lower": {},
                    "by_path": {},
                    "by_path_lower": {},
                    "all": items,
                }

                for item in items:
                    item_id = item.get('Id')
                    if item_id:
                        lookup["by_id"][item_id] = item

                    # Code field
                    if code_field:
                        code = normalize_lookup_string(item.get(code_field))
                        if code:
                            lookup["by_code"][code] = item
                            lookup["by_code_lower"][code.lower()] = item

                    # Title field (can be 'Title', 'FullName', etc.)
                    title_fields = [f for f in [title_field, name_field] if f]
                    for tf in title_fields:
                        title = normalize_lookup_string(item.get(tf))
                        if title:
                            lookup["by_title"][title] = item
                            lookup["by_title_lower"][title.lower()] = item

                    # Path field
                    if path_field:
                        path = normalize_lookup_string(item.get(path_field))
                        if path:
                            lookup["by_path"][path] = item
                            lookup["by_path_lower"][path.lower()] = item

                return lookup

            def resolve_reference_value(parsed_value, lookup, item_type="item", min_similarity=0.8):
                """Unified reference resolution with safe matching.

                Args:
                    parsed_value: result from parse_code_title_value()
                    lookup: result from build_generic_lookup()
                    item_type: type name for logging (e.g., 'project', 'object')
                    min_similarity: minimum similarity ratio for cautious fallback

                Returns:
                    (item_id, match_info) or (None, error_message)
                """
                full = parsed_value["full_normalized"]
                code = parsed_value["code"]
                title = parsed_value["title"]

                attempted = []

                # 1. Exact match by Path (most reliable for hierarchical data)
                if full in lookup["by_path"]:
                    item = lookup["by_path"][full]
                    return item['Id'], f"Path exact: '{full}'"
                attempted.append("Path exact")

                if full.lower() in lookup["by_path_lower"]:
                    item = lookup["by_path_lower"][full.lower()]
                    return item['Id'], f"Path case-insensitive: '{full}'"
                attempted.append("Path case-insensitive")

                # 2. Exact match by Title/Name
                if full in lookup["by_title"]:
                    item = lookup["by_title"][full]
                    return item['Id'], f"Title exact: '{full}'"
                attempted.append("Title exact")

                if full.lower() in lookup["by_title_lower"]:
                    item = lookup["by_title_lower"][full.lower()]
                    return item['Id'], f"Title case-insensitive: '{full}'"
                attempted.append("Title case-insensitive")

                # 3. Exact match by Code
                if full in lookup["by_code"]:
                    item = lookup["by_code"][full]
                    return item['Id'], f"Code exact: '{full}'"
                attempted.append("Code exact")

                if full.lower() in lookup["by_code_lower"]:
                    item = lookup["by_code_lower"][full.lower()]
                    return item['Id'], f"Code case-insensitive: '{full}'"
                attempted.append("Code case-insensitive")

                # 4. If we have extracted code + title
                if code and title:
                    # Match by code first (exact only)
                    if code in lookup["by_code"]:
                        item = lookup["by_code"][code]
                        return item['Id'], f"Code extracted: '{code}' -> '{item.get('Title', item.get('FullName', ''))}'"

                    if code.lower() in lookup["by_code_lower"]:
                        item = lookup["by_code_lower"][code.lower()]
                        return item['Id'], f"Code extracted case-insensitive: '{code}'"
                    attempted.append(f"Code='{code}'")

                    # Then match by title (exact only)
                    if title in lookup["by_title"]:
                        item = lookup["by_title"][title]
                        return item['Id'], f"Title extracted: '{title}' -> Code='{item.get('Code', 'N/A')}'"

                    if title.lower() in lookup["by_title_lower"]:
                        item = lookup["by_title_lower"][title.lower()]
                        return item['Id'], f"Title extracted case-insensitive: '{title}'"
                    attempted.append(f"Title='{title}'")

                # 5. Cautious fallback: only for Title/Path with similarity threshold
                # NO substring fallback for Code to avoid false matches
                for item in lookup["all"]:
                    # Check title similarity
                    item_title = normalize_lookup_string(item.get('Title') or item.get('FullName') or '')
                    if item_title and full and len(full) >= 3 and len(item_title) >= 3:
                        # Check if strings are similar enough
                        if full.lower() == item_title.lower():
                            return item['Id'], f"Title normalized match: '{full}'"

                        # Cautious: full containment with length check
                        if (full in item_title and len(full) >= len(item_title) * min_similarity) or \
                           (item_title in full and len(item_title) >= len(full) * min_similarity):
                            return item['Id'], f"Title cautious match: '{full}' ~ '{item_title}'"

                    # Check path similarity
                    item_path = normalize_lookup_string(item.get('Path', ''))
                    if item_path and full and len(full) >= 3 and len(item_path) >= 3:
                        if full.lower() == item_path.lower():
                            return item['Id'], f"Path normalized match: '{full}'"

                        if (full in item_path and len(full) >= len(item_path) * min_similarity) or \
                           (item_path in full and len(item_path) >= len(full) * min_similarity):
                            return item['Id'], f"Path cautious match: '{full}' ~ '{item_path}'"

                attempted.append("cautious fallback")

                # No reliable match found
                error_msg = f"{item_type} not found after attempts: {', '.join(attempted)}"
                return None, error_msg

            # Legacy wrapper for object structure lookup building
            def build_object_structure_lookup(obj_structs):
                """Build lookup for object structures using generic system."""
                return build_generic_lookup(
                    obj_structs,
                    code_field='Code',
                    title_field='Title',
                    path_field='Path'
                )

            # Legacy wrapper for object structure parsing
            def parse_object_structure_excel_value(raw_value):
                """Parse object structure value - wrapper for universal parser."""
                return parse_code_title_value(raw_value)

            def resolve_object_structure_id(parsed_obj, lookup, obj_structs_count):
                """Resolve object structure ID using unified reference resolution.

                Args:
                    parsed_obj: result of parse_code_title_value()
                    lookup: result from build_object_structure_lookup()
                    obj_structs_count: count for logging purposes

                Returns:
                    (object_id, match_info) or (None, error_message)
                """
                object_id, match_info = resolve_reference_value(
                    parsed_obj, lookup, item_type="Объект строительства", min_similarity=0.8
                )

                if object_id:
                    return object_id, match_info

                # Build detailed error message
                error_msg = f"{match_info}; в проекте {obj_structs_count} объектов"
                return None, error_msg

            # Step 3: Authenticate
            self._log_r("[3/9] Авторизация...\n")
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
            self._log_r("  access_token получен\n\n")

            # Log stage settings source info
            if duration_sheet_name and stage_settings_raw:
                self._log_r(f"INFO: Настройки этапов загружаются из листа '{duration_sheet_name}'\n")
            else:
                self._log_r("INFO: Лист продолжительности не выбран или пуст, применяются значения по умолчанию для этапов\n")
            self._log_r("      (AutoComplete=True, WorkingDaysDuration=2 если не указано в 5.2)\n\n")
            
            # Step 4: Load reference data
            self._log_r("[4/9] Загрузка справочников...\n")
            
            # Cache for lookups
            user_cache = {}
            approve_member_cache = {}
            object_structures_cache = {}  # project_id -> list of structures
            
            # Load projects - build unified lookup
            all_projects = get_all_projects(session, access_token, base_url)
            projects_lookup = build_generic_lookup(
                all_projects,
                code_field='Code',
                title_field='Title'
            )
            self._log_r(f"  Проектов: {len(all_projects)}\n")

            def resolve_project_id(excel_project_name):
                """Resolve project ID using unified reference resolution system.

                Supports formats:
                - '11.001. Жилой комплекс 11.001' (Code. Title)
                - '11.001 Жилой комплекс 11.001' (Code Title)
                - 'Жилой комплекс 11.001' (Title only)
                """
                if not excel_project_name:
                    return None, None, None

                parsed = parse_code_title_value(excel_project_name)
                project_id, match_info = resolve_reference_value(
                    parsed, projects_lookup, item_type="Проект"
                )

                if project_id:
                    project = projects_lookup["by_id"].get(project_id, {})
                    return project_id, project.get('Code'), project.get('Title')

                return None, None, None

            # Load content types - uses unified lookup system
            all_content_types = get_all_content_types(session, access_token, base_url)
            ct_lookup = build_generic_lookup(all_content_types, title_field='Title')
            self._log_r(f"  Типов документов: {len(all_content_types)}\n")
            
            # Load disciplines - uses unified lookup system
            all_disciplines = get_all_disciplines(session, access_token, base_url)
            disc_lookup = build_generic_lookup(all_disciplines, title_field='Title')
            self._log_r(f"  Дисциплин: {len(all_disciplines)}\n")
            
            # Load departments - uses unified lookup system (FullName as primary)
            all_departments = get_all_departments(session, access_token, base_url)
            dept_lookup = build_generic_lookup(all_departments, title_field='FullName')
            self._log_r(f"  Орг. единиц: {len(all_departments)}\n")
            
            # Load project stages - uses 'Title' field
            all_stages = get_active_project_stages(session, access_token, base_url)
            stage_by_name = {}
            for s in all_stages:
                name = normalize_string(s.get('Title'))
                if name:
                    stage_by_name[name] = s
                    stage_by_name[name.lower()] = s
            self._log_r(f"  Этапов проектов: {len(all_stages)}\n\n")
            
            # Step 5: Load existing routes for deduplication
            self._log_r("[5/9] Загрузка существующих маршрутов...\n")
            existing_routes = get_all_approval_routes(session, access_token, base_url)
            
            # Build composite key lookup for existing routes using IDs
            existing_by_key = {}
            for er in existing_routes:
                # Build key from IDs (more reliable than names)
                key = (
                    normalize_string(er.get('Name')),
                    er.get('ProjectId'),  # Use ID, not name
                    er.get('ContentTypeId'),
                    er.get('ObjectStructureId'),
                    er.get('DisciplineId'),
                    er.get('CustomerId'),
                    er.get('DeveloperId'),
                )
                existing_by_key[key] = er
            
            self._log_r(f"  Загружено существующих маршрутов: {len(existing_routes)}\n\n")
            
            # Helper function for exact user matching with ambiguity detection
            def resolve_user_exact(name, search_func, cache_dict, user_type="user"):
                """Resolve user by exact match with ambiguity detection."""
                if not name:
                    return None, "empty"
                
                if name in cache_dict:
                    return cache_dict[name], None
                
                results = search_func(name)
                if not results:
                    return None, f"{user_type} not found: '{name}'"
                
                # Normalize search name
                search_normalized = collapse_spaces(name.lower())
                
                # Look for exact match
                exact_matches = []
                for r in results:
                    display_name = r.get('DisplayName') or r.get('FullName') or r.get('Name', '')
                    if collapse_spaces(display_name.lower()) == search_normalized:
                        exact_matches.append(r)
                
                if len(exact_matches) == 1:
                    cache_dict[name] = exact_matches[0]['Id']
                    return exact_matches[0]['Id'], None
                elif len(exact_matches) > 1:
                    return None, f"{user_type} ambiguity for '{name}': {len(exact_matches)} exact matches"
                
                # No exact match, check if single result
                if len(results) == 1:
                    cache_dict[name] = results[0]['Id']
                    return results[0]['Id'], None
                else:
                    return None, f"{user_type} ambiguity for '{name}': {len(results)} results, no exact match"
            
            # Step 6: Validate all routes (Phase 1 - Pre-validation)
            self._log_r("[6/9] Валидация всех маршрутов...\n")
            self._log_r("-" * 60 + "\n")

            validated_routes = []
            validation_errors = []

            for item in merged_data:
                idx = item['index']
                route_name = item['route_name']
                route_errors = []

                self._log_r(f"\n  [{idx}] Валидация маршрута: '{route_name}' (occurrence: {item['occurrence']})\n")

                try:
                    # Resolve project
                    project_name = item['project_from_config'] or item['project_from_route']
                    project_id = None
                    all_projects_flag = False

                    if project_name and project_name.lower() not in ['все', '']:
                        project_id, matched_code, matched_title = resolve_project_id(project_name)
                        if not project_id:
                            route_errors.append(f"Проект не найден: '{project_name}'")
                        else:
                            self._log_r(f"      OK: Проект сопоставлен: '{project_name}' -> '{matched_title}'\n")
                    else:
                        all_projects_flag = True

                    # Resolve content type using unified system
                    ct_name = item['тип_документа'] or item['вид_документа']
                    ct_id = None
                    all_ct_flag = True

                    if ct_name and ct_name.lower() not in ['все', '']:
                        all_ct_flag = False
                        parsed_ct = parse_code_title_value(ct_name)
                        ct_id, ct_match_info = resolve_reference_value(
                            parsed_ct, ct_lookup, item_type="Тип документа"
                        )
                        if not ct_id:
                            route_errors.append(f"Тип документа не найден: '{ct_name}' ({ct_match_info})")
                        else:
                            self._log_r(f"      OK: Тип документа сопоставлен: '{ct_name}' -> {ct_match_info}\n")

                    # Resolve object structure
                    obj_struct_name = item['object_structure']
                    obj_struct_id = None
                    all_obj_struct_flag = True

                    if obj_struct_name and obj_struct_name.lower() not in ['все', '']:
                        all_obj_struct_flag = False

                        if not project_id:
                            route_errors.append("Невозможно найти объект строительства без проекта")
                        else:
                            if project_id not in object_structures_cache:
                                object_structures_cache[project_id] = get_object_structures_for_project(
                                    session, access_token, base_url, project_id
                                )

                            obj_structs = object_structures_cache[project_id]
                            obj_lookup = build_object_structure_lookup(obj_structs)
                            parsed_obj = parse_object_structure_excel_value(obj_struct_name)
                            obj_struct_id, match_result = resolve_object_structure_id(
                                parsed_obj, obj_lookup, len(obj_structs)
                            )

                            if not obj_struct_id:
                                route_errors.append(f"Объект строительства не найден: '{obj_struct_name}'")
                            else:
                                self._log_r(f"      OK: Объект сопоставлен: '{obj_struct_name}'\n")

                    # Resolve discipline using unified system
                    disc_name = item['дисциплина']
                    disc_id = None
                    all_disc_flag = True

                    if disc_name and disc_name.lower() not in ['все', '']:
                        all_disc_flag = False
                        parsed_disc = parse_code_title_value(disc_name)
                        disc_id, disc_match_info = resolve_reference_value(
                            parsed_disc, disc_lookup, item_type="Дисциплина"
                        )
                        if not disc_id:
                            route_errors.append(f"Дисциплина не найдена: '{disc_name}' ({disc_match_info})")
                        else:
                            self._log_r(f"      OK: Дисциплина сопоставлена: '{disc_name}' -> {disc_match_info}\n")

                    # Resolve developer using unified system
                    dev_name = item['разработчик']
                    dev_id = None
                    all_dev_flag = True

                    if dev_name and dev_name.lower() not in ['все', '']:
                        all_dev_flag = False
                        parsed_dev = parse_code_title_value(dev_name)
                        dev_id, dev_match_info = resolve_reference_value(
                            parsed_dev, dept_lookup, item_type="Разработчик"
                        )
                        if not dev_id:
                            route_errors.append(f"Разработчик не найден: '{dev_name}' ({dev_match_info})")
                        else:
                            self._log_r(f"      OK: Разработчик сопоставлен: '{dev_name}' -> {dev_match_info}\n")

                    # Resolve customer using unified system
                    cust_name = item['заказчик']
                    cust_id = None
                    all_cust_flag = True

                    if cust_name and cust_name.lower() not in ['все', '']:
                        all_cust_flag = False
                        parsed_cust = parse_code_title_value(cust_name)
                        cust_id, cust_match_info = resolve_reference_value(
                            parsed_cust, dept_lookup, item_type="Заказчик"
                        )
                        if not cust_id:
                            route_errors.append(f"Заказчик не найден: '{cust_name}' ({cust_match_info})")
                        else:
                            self._log_r(f"      OK: Заказчик сопоставлен: '{cust_name}' -> {cust_match_info}\n")

                    # Resolve admin user
                    admin_name = item['admin']
                    admin_id, admin_error = resolve_user_exact(
                        admin_name,
                        lambda n: search_users_by_name(session, access_token, base_url, n),
                        user_cache,
                        "admin"
                    )

                    if admin_error:
                        route_errors.append(f"Администратор: {admin_error}")
                    else:
                        self._log_r(f"      OK: Администратор сопоставлен: '{admin_name}'\n")

                    # Resolve stage users and validate stage settings
                    stages_payload = []
                    stage_validation_errors = []

                    route_key = (route_name, item['occurrence'])
                    route_stage_settings = stage_settings.get(route_key, {})
                    route_stages_cfg = route_stage_settings.get("stages", {})
                    route_approve_cfg = route_stage_settings.get("approve", {})

                    # Log approve duration if available
                    approve_duration = route_approve_cfg.get("ApproveWorkingDaysDuration")
                    if approve_duration:
                        self._log_r(f"      OK: Длительность утверждающего для маршрута '{route_name}' = {approve_duration}\n")
                    else:
                        self._log_r(f"      ВНИМАНИЕ: Для маршрута '{route_name}' длительность утверждающего не указана, используется fallback ApproveWorkingDaysDuration=4\n")

                    for stage_info in item['stages']:
                        stage_num = stage_info['stage_num']
                        stage_users = []

                        for user_name in stage_info['users']:
                            user_id, user_error = resolve_user_exact(
                                user_name,
                                lambda n: search_approve_members_by_name(session, access_token, base_url, n),
                                approve_member_cache,
                                "approve_member"
                            )

                            if user_error:
                                stage_validation_errors.append(f"Этап {stage_num}: {user_error}")
                            else:
                                stage_users.append(user_id)

                        if not stage_users:
                            stage_validation_errors.append(f"Этап {stage_num} не имеет валидных участников")
                            continue

                        # Get stage settings and validate conflicts
                        stage_cfg = route_stages_cfg.get(stage_num, {})
                        auto_complete = stage_cfg.get("AutoComplete", True)
                        digital_sign = stage_cfg.get("DigitalSignatureRequired", False)

                        # Check for conflicting flags: AutoComplete + DigitalSignatureRequired
                        if auto_complete and digital_sign:
                            stage_validation_errors.append(
                                f"ОШИБКА ВАЛИДАЦИИ: этап {stage_num} — конфликт флагов: "
                                f"'Автоматически завершить при истечении срока' (AutoComplete=True) и "
                                f"'Подписание ЭП' (DigitalSignatureRequired=True) одновременно"
                            )

                        stages_payload.append({
                            "stage_num": stage_num,
                            "users": stage_users,
                            "stage_cfg": stage_cfg,
                        })

                    if stage_validation_errors:
                        route_errors.extend(stage_validation_errors)

                    if not stages_payload:
                        route_errors.append("Нет валидных этапов")

                    # Collect all errors for this route
                    if route_errors:
                        for err in route_errors:
                            self._log_r(f"      ОШИБКА ВАЛИДАЦИИ: {err}\n")
                            validation_errors.append(f"[{route_name}] {err}")
                    else:
                        self._log_r(f"      OK: Маршрут прошел валидацию\n")
                        # Store validated route data for execution phase
                        validated_routes.append({
                            'item': item,
                            'project_id': project_id,
                            'all_projects_flag': all_projects_flag,
                            'ct_id': ct_id,
                            'all_ct_flag': all_ct_flag,
                            'obj_struct_id': obj_struct_id,
                            'all_obj_struct_flag': all_obj_struct_flag,
                            'disc_id': disc_id,
                            'all_disc_flag': all_disc_flag,
                            'dev_id': dev_id,
                            'all_dev_flag': all_dev_flag,
                            'cust_id': cust_id,
                            'all_cust_flag': all_cust_flag,
                            'admin_id': admin_id,
                            'stages_payload': stages_payload,
                            'approve_cfg': route_approve_cfg,
                            'project_name': project_name,
                            'obj_struct_name': obj_struct_name,
                        })

                except Exception as e:
                    self._log_r(f"      ОШИБКА ВАЛИДАЦИИ: {e}\n")
                    validation_errors.append(f"[{route_name}] {str(e)}")

            # Check if any validation errors occurred
            if validation_errors:
                self._log_r("\n" + "=" * 60 + "\n")
                self._log_r("ОШИБКИ ВАЛИДАЦИИ\n")
                self._log_r("=" * 60 + "\n")
                self._log_r(f"Найдено ошибок: {len(validation_errors)}\n\n")
                for i, err in enumerate(validation_errors, 1):
                    self._log_r(f"  {i}. {err}\n")
                self._log_r("\n" + "=" * 60 + "\n")
                self._log_r("ВЫПОЛНЕНИЕ ОСТАНОВЛЕНО\n")
                self._log_r("Создание/обновление маршрутов не выполнялось из-за ошибок валидации.\n")
                self._log_r("Исправьте ошибки и повторите попытку.\n")
                self._log_r("=" * 60 + "\n")
                self.route_log_signal.log.emit("__SHOW_ROUTE_LOG__")
                return

            self._log_r("\n" + "=" * 60 + "\n")
            self._log_r(f"Валидация пройдена: {len(validated_routes)} маршрутов готовы к импорту\n")
            self._log_r("=" * 60 + "\n")

            # Step 7: Execute - Create/Update routes (Phase 2 - Execution)
            self._log_r("\n[7/9] Создание маршрутов...\n")
            self._log_r("-" * 60 + "\n")

            stats = {"created": 0, "skipped": 0, "errors": 0}

            for route_data in validated_routes:
                item = route_data['item']
                route_name = item['route_name']
                idx = item['index']

                self._log_r(f"\n  [{idx}] Создание маршрута: '{route_name}'\n")

                try:
                    # Build deduplication key
                    existing_key = (
                        route_name,
                        route_data['project_id'] if not route_data['all_projects_flag'] else None,
                        route_data['ct_id'] if not route_data['all_ct_flag'] else None,
                        route_data['obj_struct_id'] if not route_data['all_obj_struct_flag'] else None,
                        route_data['disc_id'] if not route_data['all_disc_flag'] else None,
                        route_data['cust_id'] if not route_data['all_cust_flag'] else None,
                        route_data['dev_id'] if not route_data['all_dev_flag'] else None,
                    )

                    if existing_key in existing_by_key:
                        self._log_r(f"      ПРОПУСК: Маршрут уже существует\n")
                        stats["skipped"] += 1
                        continue

                    # Build stage payload for API
                    api_stages = []
                    for stage in route_data['stages_payload']:
                        stage_num = stage['stage_num']
                        stage_cfg = stage['stage_cfg']

                        api_stages.append({
                            "Name": f"Этап {stage_num}",
                            "AutoComplete": stage_cfg.get("AutoComplete", True),
                            "CompleteOnSingleTaskCompleted": stage_cfg.get("CompleteOnSingleTaskCompleted", False),
                            "EndProcessIfRejected": stage_cfg.get("EndProcessIfRejected", False),
                            "SortOrder": stage_num - 1,
                            "Users": [{"UserId": uid} for uid in stage['users']],
                            "WorkingDaysDuration": stage_cfg.get("WorkingDaysDuration") or 2,
                            "DigitalSignatureRequired": stage_cfg.get("DigitalSignatureRequired", False),
                        })

                    # Determine approver
                    last_stage = api_stages[-1]
                    approver_id = last_stage["Users"][-1]["UserId"] if last_stage["Users"] else None
                    
                    # Get approve duration from 5.2 or use fallback
                    approve_cfg = route_data.get('approve_cfg', {})
                    approve_duration = approve_cfg.get('ApproveWorkingDaysDuration') or 4
                    
                    # Build payload
                    approve_stage_name = item['approval_type'] or "Утверждение"

                    payload = {
                        "Id": None,
                        "ApproveStageName": approve_stage_name,
                        "IsActive": True,
                        "Name": route_name,
                        "AdminId": route_data['admin_id'],
                        "Type": 0,
                        "CoverLetter": "Уважаемые коллеги!\nВам направлен сопроводительный документ. Просьба ознакомиться.",
                        "CopyToUsers": [{"UserId": route_data['admin_id']}],
                        "Stages": api_stages,
                        "ApproverId": approver_id,
                        "ApproveAutoComplete": False,
                        "ApproveWorkingDaysDuration": approve_duration,
                        "ApproveDigitalSignatureRequired": route_data.get('approve_cfg', {}).get('ApproveDigitalSignatureRequired', False),
                        "AllContentTypes": route_data['all_ct_flag'],
                        "ContentTypeId": route_data['ct_id'] if not route_data['all_ct_flag'] else None,
                        "ProjectId": route_data['project_id'] if not route_data['all_projects_flag'] else None,
                        "AllProjects": route_data['all_projects_flag'],
                        "ObjectStructureId": route_data['obj_struct_id'] if not route_data['all_obj_struct_flag'] else None,
                        "AllObjectStructures": route_data['all_obj_struct_flag'],
                        "BlueprintMarkId": None,
                        "AllBlueprintMarks": True,
                        "DepartmentId": None,
                        "AllDepartments": True,
                        "DisciplineId": route_data['disc_id'] if not route_data['all_disc_flag'] else None,
                        "AllDisciplines": route_data['all_disc_flag'],
                        "ReleaseTargetId": None,
                        "AllReleaseTargets": True,
                        "DeveloperId": route_data['dev_id'] if not route_data['all_dev_flag'] else None,
                        "AllDevelopers": route_data['all_dev_flag'],
                        "CustomerId": route_data['cust_id'] if not route_data['all_cust_flag'] else None,
                        "AllCustomers": route_data['all_cust_flag'],
                        "AllProjectStages": True,
                        "ProjectStageId": None,
                        "StampPlacementEnabled": False,
                        "StampApprovalStatus": None,
                        "StampApproveResultIds": None,
                        "StampAnyResult": False,
                        "StampTemplateId": None,
                        "StampFirstPageOnly": False,
                        "TransmittalCreationEnabled": False,
                        "TransmittalApprovalStatus": None,
                        "TransmittalCategoryId": None,
                        "TransmittalSenderId": None,
                        "TransmittalReceiverId": None,
                        "ApproveCompletionActionsSpecified": True,
                    }

                    # Create route
                    success, status_code, resp_data = create_approval_route(
                        session, access_token, base_url, payload
                    )

                    if success:
                        route_id = resp_data if isinstance(resp_data, str) else resp_data.get('Id', '?')
                        self._log_r(f"      СОЗДАН  ID={route_id}\n")
                        stats["created"] += 1
                        existing_by_key[existing_key] = {"Id": route_id}
                    else:
                        self._log_r(f"      ОШИБКА: HTTP {status_code}\n")
                        resp_str = (
                            json.dumps(resp_data, ensure_ascii=False)
                            if isinstance(resp_data, (dict, list))
                            else str(resp_data)
                        )
                        self._log_r(f"      Ответ: {resp_str[:500]}\n")
                        stats["errors"] += 1

                except Exception as e:
                    self._log_r(f"      ОШИБКА: {e}\n")
                    self._log_r(f"      {traceback.format_exc()}\n")
                    stats["errors"] += 1

            self._log_r("\n" + "=" * 60 + "\n")
            self._log_r("[8/9] ИТОГИ\n")
            self._log_r(f"  Создано:   {stats['created']}\n")
            self._log_r(f"  Пропущено: {stats['skipped']}\n")
            self._log_r(f"  Ошибки:    {stats['errors']}\n")
            self._log_r("=" * 60 + "\n")
            self.route_log_signal.log.emit("__SHOW_ROUTE_LOG__")
        
        except Exception as e:
            self._log_r(f"\nФАТАЛЬНАЯ ОШИБКА: {e}\n")
            self._log_r(traceback.format_exc())
            self.route_log_signal.log.emit("__SHOW_ROUTE_LOG__")
        finally:
            self.route_log_signal.finished.emit()
