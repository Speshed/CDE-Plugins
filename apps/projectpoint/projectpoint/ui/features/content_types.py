from __future__ import annotations

import json
import os
import re
import threading
import traceback

import numpy as np
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
    clear_inline_preview, make_file_row, make_info_banner, make_preview_table, make_status_chip, mark_file_button,
    run_inline_excel_preview, show_excel_preview,
)


class ContentTypesMixin:
    def _build_types_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(2, 8, 2, 2)
        layout.setSpacing(8)

        file_card = QFrame()
        file_card.setObjectName("card")
        file_layout = QVBoxLayout(file_card)
        file_layout.setContentsMargins(14, 8, 14, 9)
        file_layout.setSpacing(6)
        self.types_file_edit = QLineEdit()
        self.types_file_edit.setReadOnly(True)
        self.types_file_edit.setPlaceholderText("Excel-файл не выбран")
        self.types_template_btn = QPushButton("Скачать шаблон")
        self.types_template_btn.clicked.connect(self._start_types_template)
        self.types_file_btn = QPushButton("Загрузить файл")
        self.types_file_btn.clicked.connect(self._choose_types_file)
        file_layout.addWidget(make_file_row(
            self.asset_dir, self.is_dark_theme,
            "Excel-файл типов документов",
            "Загрузите шаблон и проверьте состав типов перед запуском",
            self.types_template_btn, self.types_file_btn,
        ))
        sheet_row = QHBoxLayout()
        sheet_row.setSpacing(8)
        sheet_lbl = QLabel("Лист Excel")
        sheet_lbl.setObjectName("fieldLabel")
        sheet_row.addWidget(sheet_lbl)
        self.types_sheet_combo = QComboBox()
        self.types_sheet_combo.setEnabled(False)
        self.types_sheet_combo.currentTextChanged.connect(self._invalidate_types_preview)
        sheet_row.addWidget(self.types_sheet_combo, 1)
        file_layout.addLayout(sheet_row)
        layout.addWidget(file_card)

        action_card = QFrame()
        action_card.setObjectName("card")
        action_layout = QVBoxLayout(action_card)
        action_layout.setContentsMargins(14, 10, 14, 11)
        action_layout.setSpacing(8)
        title = QLabel("Создание и обновление типов документов")
        title.setObjectName("cardTitle")
        action_layout.addWidget(title)
        action_layout.addWidget(make_info_banner(
            self.asset_dir, self.is_dark_theme,
            "Предпросмотр показывает данные выбранного листа до отправки на сервер. "
            "При запуске существующие типы сопоставляются по Code.",
        ))
        status_row = QHBoxLayout()
        status_row.setSpacing(7)
        status_lbl = QLabel("Операции:")
        status_lbl.setObjectName("fieldLabel")
        status_row.addWidget(status_lbl)
        status_row.addWidget(make_status_chip("Создать", "success"))
        status_row.addWidget(make_status_chip("Обновить", "pending"))
        status_row.addStretch(1)
        action_layout.addLayout(status_row)
        self.types_preview_table = make_preview_table()
        action_layout.addWidget(self.types_preview_table)
        self.types_summary = QLabel("Загрузите Excel-файл для предпросмотра")
        self.types_summary.setObjectName("rowSubtitle")
        self.types_summary.setWordWrap(True)
        footer = QHBoxLayout()
        footer.setSpacing(8)
        footer.addWidget(self.types_summary, 1)
        types_log_btn = QPushButton("Log")
        types_log_btn.setObjectName("logButton")
        types_log_btn.clicked.connect(self._show_type_log)
        footer.addWidget(types_log_btn)
        self.types_details_btn = QPushButton("Подробнее")
        self.types_details_btn.setObjectName("modeSwitch")
        self.types_details_btn.setEnabled(False)
        self.types_details_btn.clicked.connect(self._show_types_preview_details)
        footer.addWidget(self.types_details_btn)
        self.types_preview_btn = QPushButton("Предпросмотр")
        self.types_preview_btn.setObjectName("loginButton")
        self.types_preview_btn.setEnabled(False)
        self.types_preview_btn.clicked.connect(self._preview_types_excel)
        footer.addWidget(self.types_preview_btn)
        self.types_start_btn = QPushButton("Создать / обновить типы")
        self.types_start_btn.setObjectName("greenAction")
        self.types_start_btn.setMinimumWidth(220)
        self.types_start_btn.clicked.connect(self._start_types)
        footer.addWidget(self.types_start_btn)
        action_layout.addLayout(footer)
        layout.addWidget(action_card)
        layout.addSpacing(2)
        self.tabs.addTab(tab, "Типы документов")

    def _choose_types_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл", "", "Excel (*.xlsx *.xls)"
        )
        if path:
            self.types_file_edit.setText(path)
            mark_file_button(self.types_file_btn, True)
            self._refresh_action_icons()
            try:
                xls = pd.ExcelFile(path)
                sheets = xls.sheet_names
                self.types_sheet_combo.blockSignals(True)
                self.types_sheet_combo.clear()
                self.types_sheet_combo.addItems(sheets)
                self.types_sheet_combo.setEnabled(True)
                self.types_sheet_combo.blockSignals(False)
                self.types_preview_btn.setEnabled(bool(sheets))
                self.types_details_btn.setEnabled(False)
                clear_inline_preview(self.types_preview_table)
                self.types_summary.setText(f"Файл выбран · листов: {len(sheets)} · запустите предпросмотр")
            except Exception as e:
                self.types_sheet_combo.clear()
                self.types_sheet_combo.setEnabled(False)
                self.types_preview_btn.setEnabled(False)
                self.types_details_btn.setEnabled(False)
                self.types_summary.setText("Не удалось прочитать Excel-файл")
                QMessageBox.warning(self, "Ошибка", f"Не удалось прочитать листы: {e}")

    def _invalidate_types_preview(self, *_args):
        if hasattr(self, "types_preview_table"):
            clear_inline_preview(self.types_preview_table)
        if hasattr(self, "types_details_btn"):
            self.types_details_btn.setEnabled(False)
        if hasattr(self, "types_file_edit") and self.types_file_edit.text():
            self.types_summary.setText("Лист изменён · запустите предпросмотр заново")

    def _types_preview_specs(self):
        sheet = self.types_sheet_combo.currentText() if self.types_sheet_combo.isEnabled() else ""
        return [(sheet, ("Code", "Title"))]

    def _preview_types_excel(self):
        sheets = run_inline_excel_preview(
            self, self.types_preview_table, self.types_file_edit.text(),
            self._types_preview_specs(), self.types_summary,
        )
        self.types_details_btn.setEnabled(bool(sheets))

    def _show_types_preview_details(self):
        show_excel_preview(
            self, "Подробный предпросмотр — типы документов", self.types_file_edit.text(),
            self._types_preview_specs(),
        )

    def _log_t(self, text):
        self.type_log_signal.log.emit(text)

    def _show_type_log(self):
        if self._type_log_dialog is None or not self._type_log_dialog.isVisible():
            self._type_log_dialog = LogDialog(self)
            self._type_log_dialog.setWindowTitle("Log — Типы документов")
            for chunk in self._type_log_buffer:
                self._type_log_dialog.append(chunk)
            self._type_log_dialog.show()
        else:
            self._type_log_dialog.raise_()
            self._type_log_dialog.activateWindow()

    def _append_type_log(self, text):
        if text == "__SHOW_TYPE_LOG__":
            self._show_type_log()
            return
        self._type_log_buffer.append(text)
        try:
            message = str(text).strip().splitlines()[-1]
            if message:
                self.statusBar().showMessage(message[:180], 8000)
        except Exception:
            pass
        if self._type_log_dialog and self._type_log_dialog.isVisible():
            self._type_log_dialog.append(text)

    def _clear_types_log(self):
        self._type_log_buffer.clear()
        if self._type_log_dialog:
            self._type_log_dialog.clear_log()

    def _start_types_template(self):
        self._save_template_file(
            "Сохранить шаблон типов документов",
            "projectpoint_content_types_template.xlsx",
            write_content_types_import_template,
        )

    def _start_types(self):
        file_path = self.types_file_edit.text()
        if not file_path or not os.path.exists(file_path):
            QMessageBox.warning(self, "Ошибка", "Выберите существующий Excel файл")
            return
        sheet_name = self.types_sheet_combo.currentText() if self.types_sheet_combo.isEnabled() else 0
        auth_args = self._get_worker_auth_args()
        if not auth_args:
            QMessageBox.warning(self, "Ошибка", "Заполните BASE_URL, USERNAME и PASSWORD")
            return
        self.types_start_btn.setEnabled(False)
        self._clear_types_log()
        threading.Thread(
            target=self._run_types_worker,
            args=(file_path, sheet_name, *auth_args),
            daemon=True,
        ).start()

    def _run_types_worker(
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
            self._log_t("=" * 60 + "\n")
            self._log_t("Запуск процесса создания/обновления типов документов\n")
            self._log_t(f"Base URL: {base_url}\n")
            self._log_t(f"Auth mode: {auth_mode}\n")
            self._log_t("=" * 60 + "\n\n")

            self._log_t("[1/6] Авторизация...\n")
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
            self._log_t("  access_token получен\n\n")

            self._log_t("[2/6] Загрузка справочников...\n")

            self._log_t("  Загружаем ProjectStages...\n")
            all_stages = get_all_project_stages(session, access_token, base_url)
            stage_lookup_to_id = {}
            for stage in all_stages:
                stage_id = stage.get("Id")
                if not stage_id:
                    continue
                for key in (stage.get("Title"), stage.get("Code")):
                    key = str(key or "").strip()
                    if key:
                        stage_lookup_to_id[key.casefold()] = stage_id
            self._log_t(f"  Загружено {len(all_stages)} видов документов\n")

            self._log_t("  Загружаем ContentTypes...\n")
            all_ct = get_all_content_types(session, access_token, base_url)
            existing_by_code = {
                str(ct.get("Code", "")).strip(): ct
                for ct in all_ct
                if ct.get("Code")
            }
            self._log_t(f"  Загружено {len(existing_by_code)} типов\n")

            self._log_t("  Загружаем CustomFields...\n")
            all_cf = get_all_custom_fields(session, access_token, base_url)
            cf_name_to_id = {}
            for cf in all_cf:
                name = str(cf.get("Name", "")).strip()
                cf_id = cf.get("Id")
                if name and cf_id:
                    cf_name_to_id[name] = cf_id
            self._log_t(f"  Загружено {len(cf_name_to_id)} пользовательских полей\n\n")

            self._log_t(f"[3/6] Чтение Excel (лист: {sheet_name})...\n")
            df = read_excel_table(excel_file, sheet_name=sheet_name, required_headers=("Code", "Title"))
            documents = df.to_dict(orient="records")
            self._log_t(f"  Загружено {len(documents)} записей\n\n")

            self._log_t("[4/6] Подготовка записей...\n")
            prepared_docs = []
            for doc in documents:
                project_stages_ids = []
                if "ProjectStages" in doc:
                    raw_value = doc["ProjectStages"]
                    if pd.notna(raw_value) and str(raw_value).strip():
                        names = [n.strip() for n in re.split(r"[;,]", str(raw_value)) if n.strip()]
                        for name in names:
                            stage_id = stage_lookup_to_id.get(name.casefold())
                            if stage_id:
                                if stage_id not in project_stages_ids:
                                    project_stages_ids.append(stage_id)
                            else:
                                self._log_t(f"  ВНИМАНИЕ: Вид документа не найден: '{name}'\n")
                    del doc["ProjectStages"]
                doc["ProjectStagesIds"] = project_stages_ids
                doc["UseProjectStage"] = bool(project_stages_ids)

                if "DisplayParameters" in doc:
                    val = doc["DisplayParameters"]
                    if pd.isna(val):
                        doc["DisplayParameters"] = 0
                    else:
                        try:
                            doc["DisplayParameters"] = int(float(val))
                        except (ValueError, TypeError):
                            self._log_t(f"  ВНИМАНИЕ: Некорректное DisplayParameters: '{val}' -> 0\n")
                            doc["DisplayParameters"] = 0

                if "CustomFields" in doc:
                    custom_fields_payload = []
                    raw_cf = doc["CustomFields"]
                    if pd.notna(raw_cf) and str(raw_cf).strip():
                        cf_names = [n.strip() for n in str(raw_cf).split(";") if n.strip()]
                        for cf_name in cf_names:
                            if cf_name in cf_name_to_id:
                                custom_fields_payload.append({
                                    "ContentTypeId": None,
                                    "IsRequired": False,
                                    "CopyOnCreateDocumentInDocumentGroup": True,
                                    "CustomField": None,
                                    "CustomFieldId": cf_name_to_id[cf_name],
                                })
                            else:
                                self._log_t(f"  ОШИБКА: Пользовательское поле не найдено: '{cf_name}'\n")
                                raise Exception(
                                    f"Пользовательское поле '{cf_name}' не найдено в системе. "
                                    f"Проверьте имя в Excel."
                                )
                    del doc["CustomFields"]
                    doc["_CustomFieldsPayload"] = custom_fields_payload

                for key in list(doc.keys()):
                    val = doc[key]
                    is_scalar_na = False
                    try:
                        is_scalar_na = pd.isna(val) and not isinstance(val, (list, dict, np.ndarray))
                    except (ValueError, TypeError):
                        pass
                    if is_scalar_na:
                        if key == "Description":
                            doc[key] = ""
                        elif key in BOOL_FIELDS:
                            doc[key] = False
                        elif key in NUMERIC_FIELDS:
                            doc[key] = 0
                        else:
                            doc[key] = ""
                    elif isinstance(val, str):
                        if key in BOOL_FIELDS:
                            clean = val.strip().lower()
                            if clean in ("true", "истина", "1", "да"):
                                doc[key] = True
                            elif clean in ("false", "ложь", "0", "нет"):
                                doc[key] = False
                            else:
                                doc[key] = False
                        elif key in NUMERIC_FIELDS:
                            try:
                                doc[key] = int(float(val))
                            except (ValueError, TypeError):
                                doc[key] = 0
                    elif isinstance(val, (int, float)):
                        if key in NUMERIC_FIELDS:
                            doc[key] = int(val)

                prepared_docs.append(doc)

            self._log_t(f"  Подготовлено {len(prepared_docs)} записей\n\n")

            self._log_t("[5/6] Обработка записей...\n")
            self._log_t("-" * 60 + "\n")

            stats = {"created": 0, "updated": 0, "skipped": 0, "errors": 0}

            for i, doc_data in enumerate(prepared_docs, 1):
                code = str(doc_data.get("Code", "")).strip()
                if not code:
                    self._log_t(f"\n  [{i}] ПРОПУСК: нет Code\n")
                    stats["skipped"] += 1
                    continue

                self._log_t(f"\n  [{i}] Code='{code}'\n")

                if code in existing_by_code:
                    existing = existing_by_code[code]
                    payload = {
                        "Id": existing["Id"],
                        "Code": doc_data.get("Code", existing.get("Code")),
                        "Title": doc_data.get("Title", existing.get("Title")),
                        "Description": doc_data.get("Description", existing.get("Description", "")),
                        "ProjectStagesIds": doc_data["ProjectStagesIds"],
                    }
                    for field in BOOL_FIELDS:
                        payload[field] = doc_data.get(field, existing.get(field, False))
                    for field in NUMERIC_FIELDS:
                        payload[field] = doc_data.get(field, existing.get(field, 0))

                    cf_payload = doc_data.get("_CustomFieldsPayload")
                    if cf_payload is not None:
                        payload["CustomFields"] = cf_payload
                    else:
                        payload["CustomFields"] = existing.get("CustomFields", [])

                    self._log_t(f"      ОБНОВЛЕНИЕ (ID={existing['Id']})\n")
                    payload_json = json.dumps(payload, ensure_ascii=False, indent=2)
                    self._log_t(f"      Payload:\n{payload_json}\n")

                    success, msg, detail = check_and_update_content_type(
                        session, access_token, base_url, payload
                    )
                    if success:
                        self._log_t(f"      ОБНОВЛЁН УСПЕШНО\n")
                        stats["updated"] += 1
                    else:
                        self._log_t(f"      ОШИБКА ОБНОВЛЕНИЯ: {msg}\n")
                        if detail:
                            self._log_t(f"      Детали: {detail}\n")
                        stats["errors"] += 1
                else:
                    cf_payload = doc_data.pop("_CustomFieldsPayload", [])
                    doc_data["CustomFields"] = cf_payload

                    self._log_t(f"      СОЗДАНИЕ (новый тип)\n")
                    payload_json = json.dumps(doc_data, ensure_ascii=False, indent=2)
                    self._log_t(f"      Payload:\n{payload_json}\n")

                    success, status_code, resp_data = create_content_type(
                        session, access_token, base_url, doc_data
                    )
                    if success:
                        ct_id = resp_data.get("Id", "?") if isinstance(resp_data, dict) else "?"
                        self._log_t(f"      СОЗДАН  ID={ct_id}\n")
                        existing_by_code[code] = doc_data
                        stats["created"] += 1
                    else:
                        self._log_t(f"      ОШИБКА: HTTP {status_code}\n")
                        resp_str = (
                            json.dumps(resp_data, ensure_ascii=False)
                            if isinstance(resp_data, (dict, list))
                            else str(resp_data)
                        )
                        self._log_t(f"      Ответ: {resp_str}\n")
                        stats["errors"] += 1

            self._log_t("\n" + "=" * 60 + "\n")
            self._log_t("[6/6] ИТОГИ\n")
            self._log_t(f"  Создано:    {stats['created']}\n")
            self._log_t(f"  Обновлено:  {stats['updated']}\n")
            self._log_t(f"  Пропущено:  {stats['skipped']}\n")
            self._log_t(f"  Ошибки:     {stats['errors']}\n")
            self._log_t("=" * 60 + "\n")
            self.type_log_signal.log.emit("__SHOW_TYPE_LOG__")

        except Exception as e:
            self._log_t(f"\nФАТАЛЬНАЯ ОШИБКА: {e}\n")
            self._log_t(traceback.format_exc())
            self.type_log_signal.log.emit("__SHOW_TYPE_LOG__")
        finally:
            self.type_log_signal.finished.emit()
