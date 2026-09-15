from __future__ import annotations

import json
import os
import threading
import traceback

import pandas as pd
from PySide6.QtWidgets import (
    QComboBox, QFileDialog, QFrame, QHBoxLayout, QLabel, QLineEdit,
    QMessageBox, QPushButton, QVBoxLayout, QWidget,
)

from ...auth import *
from ...config import *
from ...api import *
from ...excel import *
from ..dialogs import LogDialog
from ..components import (
    clear_inline_preview, make_file_row, make_info_banner, make_preview_table, mark_file_button,
    run_inline_excel_preview, show_excel_preview,
)


class ParametersMixin:
    def _build_params_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(2, 8, 2, 2)
        layout.setSpacing(8)

        file_card = QFrame()
        file_card.setObjectName("card")
        file_layout = QVBoxLayout(file_card)
        file_layout.setContentsMargins(14, 8, 14, 9)
        file_layout.setSpacing(6)

        self.params_file_edit = QLineEdit()
        self.params_file_edit.setReadOnly(True)
        self.params_file_edit.setPlaceholderText("Excel-файл не выбран")

        self.params_template_btn = QPushButton("Скачать шаблон")
        self.params_template_btn.clicked.connect(self._start_params_template)
        self.params_file_btn = QPushButton("Загрузить файл")
        self.params_file_btn.clicked.connect(self._choose_params_file)
        file_layout.addWidget(make_file_row(
            self.asset_dir, self.is_dark_theme,
            "Excel-файл параметров",
            "Скачайте шаблон, заполните его и загрузите обратно",
            self.params_template_btn, self.params_file_btn,
        ))

        sheet_row = QHBoxLayout()
        sheet_row.setSpacing(8)
        sheet_lbl = QLabel("Лист Excel")
        sheet_lbl.setObjectName("fieldLabel")
        sheet_row.addWidget(sheet_lbl)
        self.params_sheet_combo = QComboBox()
        self.params_sheet_combo.setEnabled(False)
        self.params_sheet_combo.currentTextChanged.connect(self._invalidate_params_preview)
        sheet_row.addWidget(self.params_sheet_combo, 1)
        file_layout.addLayout(sheet_row)
        layout.addWidget(file_card)

        action_card = QFrame()
        action_card.setObjectName("card")
        action_layout = QVBoxLayout(action_card)
        action_layout.setContentsMargins(14, 10, 14, 11)
        action_layout.setSpacing(8)
        title = QLabel("Создание параметров")
        title.setObjectName("cardTitle")
        action_layout.addWidget(title)
        action_layout.addWidget(make_info_banner(
            self.asset_dir, self.is_dark_theme,
            "Предпросмотр читает Excel локально и ничего не отправляет на сервер. "
            "Перед запуском проверьте названия колонок и значения параметров.",
        ))
        self.params_preview_table = make_preview_table()
        action_layout.addWidget(self.params_preview_table)

        self.params_summary = QLabel("Загрузите Excel-файл для предпросмотра")
        self.params_summary.setObjectName("rowSubtitle")
        self.params_summary.setWordWrap(True)

        footer = QHBoxLayout()
        footer.setSpacing(8)
        footer.addWidget(self.params_summary, 1)
        params_log_btn = QPushButton("Log")
        params_log_btn.setObjectName("logButton")
        params_log_btn.clicked.connect(self._show_param_log)
        footer.addWidget(params_log_btn)
        self.params_details_btn = QPushButton("Подробнее")
        self.params_details_btn.setObjectName("modeSwitch")
        self.params_details_btn.setEnabled(False)
        self.params_details_btn.clicked.connect(self._show_params_preview_details)
        footer.addWidget(self.params_details_btn)
        self.params_preview_btn = QPushButton("Предпросмотр")
        self.params_preview_btn.setObjectName("loginButton")
        self.params_preview_btn.setEnabled(False)
        self.params_preview_btn.clicked.connect(self._preview_params_excel)
        footer.addWidget(self.params_preview_btn)
        self.params_start_btn = QPushButton("Запустить создание параметров")
        self.params_start_btn.setObjectName("greenAction")
        self.params_start_btn.setMinimumWidth(250)
        self.params_start_btn.clicked.connect(self._start_params)
        footer.addWidget(self.params_start_btn)
        action_layout.addLayout(footer)
        layout.addWidget(action_card)
        layout.addSpacing(2)

        self.tabs.addTab(tab, "Создание параметров")

    def _choose_params_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл", "", "Excel (*.xlsx *.xls)"
        )
        if path:
            self.params_file_edit.setText(path)
            mark_file_button(self.params_file_btn, True)
            self._refresh_action_icons()
            try:
                xls = pd.ExcelFile(path)
                sheets = xls.sheet_names
                self.params_sheet_combo.blockSignals(True)
                self.params_sheet_combo.clear()
                self.params_sheet_combo.addItems(sheets)
                self.params_sheet_combo.setEnabled(True)
                self.params_sheet_combo.blockSignals(False)
                self.params_preview_btn.setEnabled(bool(sheets))
                self.params_details_btn.setEnabled(False)
                clear_inline_preview(self.params_preview_table)
                self.params_summary.setText(f"Файл выбран · листов: {len(sheets)} · запустите предпросмотр")
            except Exception as e:
                self.params_sheet_combo.clear()
                self.params_sheet_combo.setEnabled(False)
                self.params_preview_btn.setEnabled(False)
                self.params_details_btn.setEnabled(False)
                self.params_summary.setText("Не удалось прочитать Excel-файл")
                QMessageBox.warning(self, "Ошибка", f"Не удалось прочитать листы: {e}")

    def _invalidate_params_preview(self, *_args):
        if hasattr(self, "params_preview_table"):
            clear_inline_preview(self.params_preview_table)
        if hasattr(self, "params_details_btn"):
            self.params_details_btn.setEnabled(False)
        if hasattr(self, "params_file_edit") and self.params_file_edit.text():
            self.params_summary.setText("Лист изменён · запустите предпросмотр заново")

    def _params_preview_specs(self):
        sheet = self.params_sheet_combo.currentText() if self.params_sheet_combo.isEnabled() else ""
        return [(sheet, ("Наименование атрибута", "Внутреннее имя атрибута", "Тип данных"))]

    def _preview_params_excel(self):
        sheets = run_inline_excel_preview(
            self, self.params_preview_table, self.params_file_edit.text(),
            self._params_preview_specs(), self.params_summary,
        )
        self.params_details_btn.setEnabled(bool(sheets))

    def _show_params_preview_details(self):
        show_excel_preview(
            self, "Подробный предпросмотр — параметры", self.params_file_edit.text(),
            self._params_preview_specs(),
        )

    def _log_p(self, text):
        self.param_log_signal.log.emit(text)

    def _show_param_log(self):
        if self._param_log_dialog is None or not self._param_log_dialog.isVisible():
            self._param_log_dialog = LogDialog(self)
            self._param_log_dialog.setWindowTitle("Log — Создание параметров")
            for chunk in self._param_log_buffer:
                self._param_log_dialog.append(chunk)
            self._param_log_dialog.show()
        else:
            self._param_log_dialog.raise_()
            self._param_log_dialog.activateWindow()

    def _append_param_log(self, text):
        if text == "__SHOW_LOG__":
            self._show_param_log()
            return
        self._param_log_buffer.append(text)
        try:
            message = str(text).strip().splitlines()[-1]
            if message:
                self.statusBar().showMessage(message[:180], 8000)
        except Exception:
            pass
        if self._param_log_dialog and self._param_log_dialog.isVisible():
            self._param_log_dialog.append(text)

    def _clear_params_log(self):
        self._param_log_buffer.clear()
        if self._param_log_dialog:
            self._param_log_dialog.clear_log()

    def _start_params_template(self):
        self._save_template_file(
            "Сохранить шаблон параметров",
            "projectpoint_params_template.xlsx",
            write_params_import_template,
        )

    def _start_params(self):
        file_path = self.params_file_edit.text()
        if not file_path or not os.path.exists(file_path):
            QMessageBox.warning(self, "Ошибка", "Выберите существующий Excel файл")
            return
        sheet_name = self.params_sheet_combo.currentText() if self.params_sheet_combo.isEnabled() else 0
        auth_args = self._get_worker_auth_args()
        if not auth_args:
            QMessageBox.warning(self, "Ошибка", "Заполните BASE_URL, USERNAME и PASSWORD")
            return
        self.params_start_btn.setEnabled(False)
        self._clear_params_log()
        threading.Thread(
            target=self._run_params_worker,
            args=(file_path, sheet_name, *auth_args),
            daemon=True,
        ).start()

    def _run_params_worker(
        self,
        excel_file,
        sheet_name,
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
            self._log_p("=" * 60 + "\n")
            self._log_p("Запуск процесса создания атрибутов\n")
            self._log_p(f"Base URL: {base_url}\n")
            self._log_p(f"Auth mode: {auth_mode}\n")
            self._log_p("=" * 60 + "\n\n")

            self._log_p(f"[1/5] Чтение и валидация Excel (лист: {sheet_name})...\n")
            records = read_excel(excel_file, sheet_name=sheet_name)
            self._log_p(f"  Загружено {len(records)} записей\n")

            all_errors = []
            for i, row in enumerate(records, 1):
                all_errors.extend(validate_row(row, i))
            if all_errors:
                msg = "\n".join(all_errors)
                self._log_p(f"\nОШИБКИ ВАЛИДАЦИИ:\n{msg}\n")
                self._log_p("Работа остановлена. Исправьте ошибки и повторите.\n")
                self.param_log_signal.log.emit("__SHOW_LOG__")
                return
            self._log_p("  Все строки прошли валидацию\n\n")

            self._log_p("[2/5] Авторизация...\n")
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
            self._log_p("  access_token получен\n\n")

            self._log_p("[3/5] Загрузка существующих атрибутов...\n")
            existing_fields = get_all_custom_fields(session, access_token, base_url)
            existing_by_name = {}
            existing_by_system = {}
            if isinstance(existing_fields, list):
                for f in existing_fields:
                    n = str(f.get("Name", "")).strip()
                    s = str(f.get("SystemName", "")).strip()
                    if n:
                        existing_by_name[n] = f
                    if s:
                        existing_by_system[s] = f
                self._log_p(f"  Загружено {len(existing_fields)} существующих атрибутов\n\n")
            else:
                self._log_p(f"  Ответ: {existing_fields}\n\n")

            self._log_p("[4/5] Обработка записей...\n")
            self._log_p("-" * 60 + "\n")

            stats = {"created": 0, "skipped": 0, "errors": 0}

            for i, row in enumerate(records, 1):
                name_raw = row.get("Name", "")
                sys_raw = row.get("SystemName", "")
                display_name = str(name_raw).strip() if pd.notna(name_raw) else f"Строка {i}"
                display_sys = str(sys_raw).strip() if pd.notna(sys_raw) else ""

                self._log_p(f"\n  [{i}] Name='{display_name}', SystemName='{display_sys}'\n")

                payload, errors = map_excel_row_to_payload(row)
                if errors:
                    for err in errors:
                        self._log_p(f"      ОШИБКА: {err}\n")
                    stats["errors"] += 1
                    continue

                payload_name = payload["Name"]
                payload_sys = payload["SystemName"]

                if payload_sys in existing_by_system:
                    self._log_p(f"      ПРОПУСК (SystemName '{payload_sys}' существует)\n")
                    stats["skipped"] += 1
                    continue
                if payload_name in existing_by_name:
                    self._log_p(f"      ПРОПУСК (Name '{payload_name}' существует)\n")
                    stats["skipped"] += 1
                    continue

                warning = payload.pop("_warning", None)
                if warning:
                    self._log_p(f"      ВНИМАНИЕ: {warning}\n")

                payload_json = json.dumps(payload, ensure_ascii=False, indent=2)
                self._log_p(f"      Payload:\n{payload_json}\n")

                success, status_code, resp_data = create_custom_field(
                    session, access_token, base_url, payload
                )

                if success:
                    field_id = resp_data.get("Id", "?") if isinstance(resp_data, dict) else "?"
                    self._log_p(f"      СОЗДАН  ID={field_id}\n")
                    existing_by_name[payload_name] = payload
                    existing_by_system[payload_sys] = payload
                    stats["created"] += 1
                else:
                    self._log_p(f"      ОШИБКА: HTTP {status_code}\n")
                    resp_str = (
                        json.dumps(resp_data, ensure_ascii=False)
                        if isinstance(resp_data, (dict, list))
                        else str(resp_data)
                    )
                    self._log_p(f"      Ответ: {resp_str}\n")
                    stats["errors"] += 1

            self._log_p("\n" + "=" * 60 + "\n")
            self._log_p("[5/5] ИТОГИ\n")
            self._log_p(f"  Создано:               {stats['created']}\n")
            self._log_p(f"  Пропущено (дубликаты): {stats['skipped']}\n")
            self._log_p(f"  Ошибки:                {stats['errors']}\n")
            self._log_p("=" * 60 + "\n")
            self.param_log_signal.log.emit("__SHOW_LOG__")

        except Exception as e:
            self._log_p(f"\nФАТАЛЬНАЯ ОШИБКА: {e}\n")
            self._log_p(traceback.format_exc())
            self.param_log_signal.log.emit("__SHOW_LOG__")
        finally:
            self.param_log_signal.finished.emit()
