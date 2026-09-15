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

from ...auth import get_access_token, session
from ...api.project_stages import create_project_stage, get_all_project_stages, update_project_stage
from ...excel.project_stages import read_project_stages_excel, project_stage_payload_equal
from ...excel.templates import write_project_stages_import_template
from ..dialogs import LogDialog
from ..components import (
    clear_inline_preview, make_file_row, make_info_banner, make_preview_table, make_status_chip, mark_file_button,
    run_inline_excel_preview, show_excel_preview,
)


class ProjectStagesMixin:
    def _build_project_stages_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(2, 8, 2, 2)
        layout.setSpacing(8)

        file_card = QFrame()
        file_card.setObjectName("card")
        file_layout = QVBoxLayout(file_card)
        file_layout.setContentsMargins(14, 8, 14, 9)
        file_layout.setSpacing(6)
        self.stages_file_edit = QLineEdit()
        self.stages_file_edit.setReadOnly(True)
        self.stages_file_edit.setPlaceholderText("Excel-файл не выбран")
        self.stages_template_btn = QPushButton("Скачать шаблон")
        self.stages_template_btn.clicked.connect(self._start_project_stages_template)
        self.stages_file_btn = QPushButton("Загрузить файл")
        self.stages_file_btn.clicked.connect(self._choose_project_stages_file)
        file_layout.addWidget(make_file_row(
            self.asset_dir, self.is_dark_theme,
            "Excel-файл видов документов",
            "Code используется как ключ при создании и обновлении",
            self.stages_template_btn, self.stages_file_btn,
        ))
        sheet_row = QHBoxLayout()
        sheet_row.setSpacing(8)
        sheet_label = QLabel("Лист Excel")
        sheet_label.setObjectName("fieldLabel")
        sheet_row.addWidget(sheet_label)
        self.stages_sheet_combo = QComboBox()
        self.stages_sheet_combo.setEnabled(False)
        self.stages_sheet_combo.currentTextChanged.connect(self._invalidate_project_stages_preview)
        sheet_row.addWidget(self.stages_sheet_combo, 1)
        file_layout.addLayout(sheet_row)
        layout.addWidget(file_card)

        action_card = QFrame()
        action_card.setObjectName("card")
        action_layout = QVBoxLayout(action_card)
        action_layout.setContentsMargins(14, 10, 14, 11)
        action_layout.setSpacing(8)
        title = QLabel("Создание и обновление видов документов")
        title.setObjectName("cardTitle")
        action_layout.addWidget(title)
        action_layout.addWidget(make_info_banner(
            self.asset_dir, self.is_dark_theme,
            "Сопоставление выполняется по Code: новый код создаётся, существующий обновляется, "
            "полностью совпадающая запись пропускается.",
        ))
        status_row = QHBoxLayout()
        status_row.setSpacing(7)
        status_lbl = QLabel("Результат обработки:")
        status_lbl.setObjectName("fieldLabel")
        status_row.addWidget(status_lbl)
        status_row.addWidget(make_status_chip("Создать", "success"))
        status_row.addWidget(make_status_chip("Обновить", "pending"))
        status_row.addWidget(make_status_chip("Пропустить", "neutral"))
        status_row.addStretch(1)
        action_layout.addLayout(status_row)
        self.stages_preview_table = make_preview_table()
        action_layout.addWidget(self.stages_preview_table)

        self.stages_summary = QLabel("Загрузите Excel-файл для предпросмотра")
        self.stages_summary.setObjectName("rowSubtitle")
        self.stages_summary.setWordWrap(True)
        footer = QHBoxLayout()
        footer.setSpacing(8)
        footer.addWidget(self.stages_summary, 1)
        log_btn = QPushButton("Log")
        log_btn.setObjectName("logButton")
        log_btn.clicked.connect(self._show_project_stages_log)
        footer.addWidget(log_btn)
        self.stages_details_btn = QPushButton("Подробнее")
        self.stages_details_btn.setObjectName("modeSwitch")
        self.stages_details_btn.setEnabled(False)
        self.stages_details_btn.clicked.connect(self._show_project_stages_preview_details)
        footer.addWidget(self.stages_details_btn)
        self.stages_preview_btn = QPushButton("Предпросмотр")
        self.stages_preview_btn.setObjectName("loginButton")
        self.stages_preview_btn.setEnabled(False)
        self.stages_preview_btn.clicked.connect(self._preview_project_stages_excel)
        footer.addWidget(self.stages_preview_btn)
        self.stages_start_btn = QPushButton("Создать / обновить")
        self.stages_start_btn.setObjectName("greenAction")
        self.stages_start_btn.setMinimumWidth(190)
        self.stages_start_btn.clicked.connect(self._start_project_stages)
        footer.addWidget(self.stages_start_btn)
        action_layout.addLayout(footer)
        layout.addWidget(action_card)
        layout.addSpacing(2)
        self.tabs.addTab(tab, "Виды документов")

    def _choose_project_stages_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите Excel файл", "", "Excel (*.xlsx *.xls)")
        if not path:
            return
        self.stages_file_edit.setText(path)
        mark_file_button(self.stages_file_btn, True)
        self._refresh_action_icons()
        try:
            xls = pd.ExcelFile(path)
            self.stages_sheet_combo.clear()
            self.stages_sheet_combo.addItems(xls.sheet_names)
            self.stages_sheet_combo.setEnabled(bool(xls.sheet_names))
            self.stages_preview_btn.setEnabled(bool(xls.sheet_names))
            self.stages_details_btn.setEnabled(False)
            clear_inline_preview(self.stages_preview_table)
            self.stages_summary.setText(f"Файл выбран · листов: {len(xls.sheet_names)} · запустите предпросмотр")
        except Exception as exc:
            self.stages_sheet_combo.clear()
            self.stages_sheet_combo.setEnabled(False)
            self.stages_preview_btn.setEnabled(False)
            self.stages_details_btn.setEnabled(False)
            self.stages_summary.setText("Не удалось прочитать Excel-файл")
            QMessageBox.warning(self, "Ошибка", f"Не удалось прочитать листы: {exc}")

    def _invalidate_project_stages_preview(self, *_args):
        if hasattr(self, "stages_preview_table"):
            clear_inline_preview(self.stages_preview_table)
        if hasattr(self, "stages_details_btn"):
            self.stages_details_btn.setEnabled(False)
        if hasattr(self, "stages_file_edit") and self.stages_file_edit.text():
            self.stages_summary.setText("Лист изменён · запустите предпросмотр заново")

    def _project_stages_preview_specs(self):
        sheet = self.stages_sheet_combo.currentText() if self.stages_sheet_combo.isEnabled() else ""
        return [(sheet, ("Code", "Title"))]

    def _preview_project_stages_excel(self):
        sheets = run_inline_excel_preview(
            self, self.stages_preview_table, self.stages_file_edit.text(),
            self._project_stages_preview_specs(), self.stages_summary,
        )
        self.stages_details_btn.setEnabled(bool(sheets))

    def _show_project_stages_preview_details(self):
        show_excel_preview(
            self, "Подробный предпросмотр — виды документов", self.stages_file_edit.text(),
            self._project_stages_preview_specs(),
        )

    def _log_stage(self, text):
        self.stage_log_signal.log.emit(text)

    def _append_project_stages_log(self, text):
        if text == "__SHOW_STAGE_LOG__":
            self._show_project_stages_log()
            return
        self._stage_log_buffer.append(text)
        try:
            message = str(text).strip().splitlines()[-1]
            if message:
                self.statusBar().showMessage(message[:180], 8000)
        except Exception:
            pass
        if self._stage_log_dialog and self._stage_log_dialog.isVisible():
            self._stage_log_dialog.append(text)

    def _show_project_stages_log(self):
        if self._stage_log_dialog is None or not self._stage_log_dialog.isVisible():
            self._stage_log_dialog = LogDialog(self)
            self._stage_log_dialog.setWindowTitle("Log — Виды документов")
            for chunk in self._stage_log_buffer:
                self._stage_log_dialog.append(chunk)
            self._stage_log_dialog.show()
        else:
            self._stage_log_dialog.raise_()
            self._stage_log_dialog.activateWindow()

    def _clear_project_stages_log(self):
        self._stage_log_buffer.clear()
        if self._stage_log_dialog:
            self._stage_log_dialog.clear_log()

    def _start_project_stages_template(self):
        self._save_template_file(
            "Сохранить шаблон видов документов",
            "projectpoint_document_views_template.xlsx",
            write_project_stages_import_template,
        )

    def _start_project_stages(self):
        file_path = self.stages_file_edit.text()
        if not file_path or not os.path.exists(file_path):
            QMessageBox.warning(self, "Ошибка", "Выберите существующий Excel файл")
            return
        if not self.stages_sheet_combo.isEnabled() or not self.stages_sheet_combo.currentText():
            QMessageBox.warning(self, "Ошибка", "Выберите лист Excel")
            return
        auth_args = self._get_worker_auth_args()
        if not auth_args:
            QMessageBox.warning(self, "Ошибка", "Заполните параметры подключения")
            return

        self.stages_start_btn.setEnabled(False)
        self._clear_project_stages_log()
        threading.Thread(
            target=self._run_project_stages_worker,
            args=(file_path, self.stages_sheet_combo.currentText(), *auth_args),
            daemon=True,
        ).start()

    def _run_project_stages_worker(
        self, excel_file, sheet_name, auth_mode, client_id, base_url, username, password,
        sso_url, realm, broker_alias, adfs_url,
    ):
        try:
            self._log_stage("=" * 60 + "\n")
            self._log_stage("Создание / обновление видов документов\n")
            self._log_stage("=" * 60 + "\n\n")

            self._log_stage(f"[1/4] Чтение Excel (лист: {sheet_name})...\n")
            desired = read_project_stages_excel(excel_file, sheet_name=sheet_name)
            if not desired:
                raise ValueError("В выбранном листе нет записей")
            self._log_stage(f"  Корректных записей: {len(desired)}\n\n")

            self._log_stage("[2/4] Авторизация...\n")
            token = get_access_token(
                session, auth_mode, client_id, base_url, username, password,
                sso_base_url=sso_url, realm=realm, broker_alias=broker_alias, adfs_base_url=adfs_url,
            )
            self._log_stage("  access_token получен\n\n")

            self._log_stage("[3/4] Загрузка существующих видов документов...\n")
            existing = get_all_project_stages(session, token, base_url)
            existing_by_code = {
                str(item.get("Code", "")).strip().casefold(): item
                for item in existing
                if str(item.get("Code", "")).strip()
            }
            self._log_stage(f"  На сервере: {len(existing_by_code)}\n\n")

            stats = {"created": 0, "updated": 0, "skipped": 0, "errors": 0}
            self._log_stage("[4/4] Обработка...\n")
            for item in desired:
                code = item["Code"]
                current = existing_by_code.get(code.casefold())
                if current and project_stage_payload_equal(current, item):
                    self._log_stage(f"  {code}: без изменений — пропуск\n")
                    stats["skipped"] += 1
                    continue

                if current:
                    payload = {"Id": current.get("Id"), **item}
                    success, status, response = update_project_stage(session, token, base_url, payload)
                    action = "обновлён"
                    stat_key = "updated"
                else:
                    payload = dict(item)
                    success, status, response = create_project_stage(session, token, base_url, payload)
                    action = "создан"
                    stat_key = "created"

                if success:
                    self._log_stage(f"  {code}: {action}\n")
                    stats[stat_key] += 1
                    if current is None:
                        new_id = response.get("Id") if isinstance(response, dict) else None
                        existing_by_code[code.casefold()] = {"Id": new_id, **item}
                else:
                    details = json.dumps(response, ensure_ascii=False) if isinstance(response, (dict, list)) else str(response)
                    self._log_stage(f"  {code}: ОШИБКА HTTP {status}: {details[:500]}\n")
                    stats["errors"] += 1

            self._log_stage("\n" + "=" * 60 + "\n")
            self._log_stage(
                f"Создано: {stats['created']} | Обновлено: {stats['updated']} | "
                f"Без изменений: {stats['skipped']} | Ошибок: {stats['errors']}\n"
            )
            self._log_stage("=" * 60 + "\n")
            self.stage_log_signal.log.emit("__SHOW_STAGE_LOG__")
        except Exception as exc:
            self._log_stage(f"\nФАТАЛЬНАЯ ОШИБКА: {exc}\n")
            self._log_stage(traceback.format_exc())
            self.stage_log_signal.log.emit("__SHOW_STAGE_LOG__")
        finally:
            # Same pattern as the existing feature workers; kept localized for compatibility.
            self.stage_log_signal.finished.emit()
