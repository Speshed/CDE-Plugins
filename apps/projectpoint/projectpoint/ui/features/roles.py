from __future__ import annotations

import json
import os
import threading
import traceback

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox, QFileDialog, QFrame, QGridLayout, QHBoxLayout, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMessageBox, QPushButton, QVBoxLayout, QWidget,
)

from ...auth import *
from ...config import *
from ...api import *
from ...excel import *
from ...api.roles import _normalize_lookup_text
from ...excel.objects import _normalize_excel_col
from ...excel.roles import _find_roles_permission_header_row
from ..dialogs import LogDialog
from ..components import (
    clear_inline_preview, make_file_row, make_info_banner, make_preview_table, make_status_chip, mark_file_button,
    run_inline_excel_preview, show_excel_preview,
)


class RolesMixin:
    def _build_roles_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(2, 8, 2, 2)
        layout.setSpacing(8)

        file_card = QFrame()
        file_card.setObjectName("card")
        file_layout = QVBoxLayout(file_card)
        file_layout.setContentsMargins(14, 8, 14, 9)
        file_layout.setSpacing(6)
        self.roles_file_edit = QLineEdit()
        self.roles_file_edit.setReadOnly(True)
        self.roles_file_edit.setPlaceholderText("Excel-файл не выбран")
        self.roles_template_btn = QPushButton("Скачать шаблон")
        self.roles_template_btn.clicked.connect(self._start_roles_template)
        self.roles_file_btn = QPushButton("Загрузить файл")
        self.roles_file_btn.clicked.connect(self._choose_roles_file)
        file_layout.addWidget(make_file_row(
            self.asset_dir, self.is_dark_theme,
            "Excel-файл ролей и привилегий",
            "Выберите листы-модули, которые нужно проверить и импортировать",
            self.roles_template_btn, self.roles_file_btn,
        ))

        settings = QGridLayout()
        settings.setHorizontalSpacing(10)
        settings.setVerticalSpacing(6)
        sheets_lbl = QLabel("Листы для импорта")
        sheets_lbl.setObjectName("fieldLabel")
        settings.addWidget(sheets_lbl, 0, 0)
        self.roles_sheets_list = QListWidget()
        self.roles_sheets_list.setMinimumHeight(46)
        self.roles_sheets_list.setMaximumHeight(112)
        self.roles_sheets_list.setUniformItemSizes(True)
        placeholder = QListWidgetItem("Загрузите Excel-файл — здесь появятся листы")
        placeholder.setFlags(Qt.NoItemFlags)
        self.roles_sheets_list.addItem(placeholder)
        self.roles_sheets_list.itemChanged.connect(self._roles_sheet_selection_changed)
        self.roles_sheets_list.setFixedHeight(46)
        settings.addWidget(self.roles_sheets_list, 1, 0, 1, 2)
        existing_lbl = QLabel("Если роль уже существует")
        existing_lbl.setObjectName("fieldLabel")
        self.roles_existing_combo = QComboBox()
        self.roles_existing_combo.addItem("Пропустить существующую роль", "skip")
        self.roles_existing_combo.addItem("Обновить существующую роль", "update")
        self.roles_existing_combo.addItem("Остановить импорт с ошибкой", "error")
        settings.addWidget(existing_lbl, 2, 0)
        settings.addWidget(self.roles_existing_combo, 2, 1)
        template_module_lbl = QLabel("Модули в шаблоне")
        template_module_lbl.setObjectName("fieldLabel")
        self.roles_template_module_combo = QComboBox()
        self.roles_template_module_combo.addItem("Все видимые модули — отдельными листами", None)
        self.roles_template_module_combo.addItem("Только: Документы", "Документы")
        self.roles_template_module_combo.addItem("Только: Контроль Качества", "Контроль Качества")
        self.roles_template_module_combo.addItem("Только: Настройки", "Настройки")
        self.roles_template_module_combo.addItem("Только: Отчеты", "Отчеты")
        self.roles_template_module_combo.addItem("Только: Приёмка Работ", "Приёмка Работ")
        settings.addWidget(template_module_lbl, 3, 0)
        settings.addWidget(self.roles_template_module_combo, 3, 1)
        settings.setColumnMinimumWidth(0, 170)
        settings.setColumnStretch(1, 1)
        file_layout.addLayout(settings)
        layout.addWidget(file_card)

        action_card = QFrame()
        action_card.setObjectName("card")
        action_layout = QVBoxLayout(action_card)
        action_layout.setContentsMargins(14, 10, 14, 11)
        action_layout.setSpacing(8)
        title = QLabel("Импорт ролей")
        title.setObjectName("cardTitle")
        action_layout.addWidget(title)
        action_layout.addWidget(make_info_banner(
            self.asset_dir, self.is_dark_theme,
            "Предпросмотр показывает только выбранные листы Excel. Перед импортом сервер повторно проверяет доступные права и модули.",
        ))
        status_row = QHBoxLayout()
        status_row.setSpacing(7)
        lbl = QLabel("Существующая роль:")
        lbl.setObjectName("fieldLabel")
        status_row.addWidget(lbl)
        status_row.addWidget(make_status_chip("Пропустить", "neutral"))
        status_row.addWidget(make_status_chip("Обновить", "pending"))
        status_row.addWidget(make_status_chip("Ошибка", "danger"))
        status_row.addStretch(1)
        action_layout.addLayout(status_row)
        self.roles_preview_table = make_preview_table()
        action_layout.addWidget(self.roles_preview_table)
        self.roles_summary = QLabel("Загрузите Excel-файл для предпросмотра")
        self.roles_summary.setObjectName("rowSubtitle")
        self.roles_summary.setWordWrap(True)
        footer = QHBoxLayout()
        footer.setSpacing(8)
        footer.addWidget(self.roles_summary, 1)
        roles_log_btn = QPushButton("Log")
        roles_log_btn.setObjectName("logButton")
        roles_log_btn.clicked.connect(self._show_roles_log)
        footer.addWidget(roles_log_btn)
        self.roles_details_btn = QPushButton("Подробнее")
        self.roles_details_btn.setObjectName("modeSwitch")
        self.roles_details_btn.setEnabled(False)
        self.roles_details_btn.clicked.connect(self._show_roles_preview_details)
        footer.addWidget(self.roles_details_btn)
        self.roles_preview_btn = QPushButton("Предпросмотр")
        self.roles_preview_btn.setObjectName("loginButton")
        self.roles_preview_btn.setEnabled(False)
        self.roles_preview_btn.clicked.connect(self._preview_roles_excel)
        footer.addWidget(self.roles_preview_btn)
        self.roles_import_btn = QPushButton("Импортировать")
        self.roles_import_btn.setObjectName("greenAction")
        self.roles_import_btn.setMinimumWidth(160)
        self.roles_import_btn.clicked.connect(self._start_roles_import)
        footer.addWidget(self.roles_import_btn)
        action_layout.addLayout(footer)
        layout.addWidget(action_card)
        layout.addSpacing(2)
        self.tabs.addTab(tab, "Роли и привилегии")

    def _choose_roles_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл импорта ролей", "", "Excel (*.xlsx *.xls)"
        )
        if path:
            self.roles_file_edit.setText(path)
            mark_file_button(self.roles_file_btn, True)
            self._refresh_action_icons()
            self._populate_roles_sheet_checklist(path)
            selected = self._get_selected_roles_sheets() or []
            self.roles_preview_btn.setEnabled(bool(selected))
            self.roles_details_btn.setEnabled(False)
            clear_inline_preview(self.roles_preview_table)
            self.roles_summary.setText(f"Файл выбран · листов для импорта: {len(selected)} · запустите предпросмотр")

    def _populate_roles_sheet_checklist(self, path):
        if not hasattr(self, "roles_sheets_list"):
            return
        self.roles_sheets_list.clear()
        try:
            xls = pd.ExcelFile(path)
            for sheet_name in xls.sheet_names:
                normalized = _normalize_excel_col(sheet_name)
                if normalized in {"инструкция", "readme", "настройка роли"}:
                    continue
                try:
                    df = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=80)
                    if _find_roles_permission_header_row(df) is None:
                        continue
                except Exception:
                    continue
                item = QListWidgetItem(sheet_name)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked)
                self.roles_sheets_list.addItem(item)
            if self.roles_sheets_list.count() == 0:
                item = QListWidgetItem("Не найдены листы с колонками прав")
                item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
                item.setCheckState(Qt.Unchecked)
                self.roles_sheets_list.addItem(item)
            self._update_roles_sheet_list_height()
        except Exception as e:
            placeholder = QListWidgetItem("Не удалось прочитать список листов")
            placeholder.setFlags(Qt.NoItemFlags)
            self.roles_sheets_list.addItem(placeholder)
            self._update_roles_sheet_list_height()
            QMessageBox.warning(self, "Ошибка", f"Не удалось прочитать листы Excel: {e}")

    def _update_roles_sheet_list_height(self):
        if not hasattr(self, "roles_sheets_list"):
            return
        count = max(1, self.roles_sheets_list.count())
        # About 28 px per checklist row, capped so the settings card never grows excessively.
        self.roles_sheets_list.setFixedHeight(max(46, min(28 * count + 10, 112)))

    def _get_selected_roles_sheets(self):
        if not hasattr(self, "roles_sheets_list") or self.roles_sheets_list.count() == 0:
            return None
        selected = []
        for i in range(self.roles_sheets_list.count()):
            item = self.roles_sheets_list.item(i)
            if item.flags() & Qt.ItemIsEnabled and item.checkState() == Qt.Checked:
                selected.append(item.text())
        return selected

    def _roles_sheet_selection_changed(self, _item=None):
        self._update_roles_sheet_list_height()
        selected = self._get_selected_roles_sheets() or []
        self.roles_preview_btn.setEnabled(bool(self.roles_file_edit.text() and selected))
        self.roles_details_btn.setEnabled(False)
        clear_inline_preview(self.roles_preview_table)
        if self.roles_file_edit.text():
            self.roles_summary.setText(f"Выбрано листов для импорта: {len(selected)} · предпросмотр нужно обновить")

    def _roles_preview_specs(self):
        selected = self._get_selected_roles_sheets() or []
        return [(sheet, ()) for sheet in selected]

    def _preview_roles_excel(self):
        sheets = run_inline_excel_preview(
            self, self.roles_preview_table, self.roles_file_edit.text(),
            self._roles_preview_specs(), self.roles_summary,
        )
        self.roles_details_btn.setEnabled(bool(sheets))

    def _show_roles_preview_details(self):
        show_excel_preview(
            self, "Подробный предпросмотр — роли и привилегии", self.roles_file_edit.text(),
            self._roles_preview_specs(),
        )

    def _log_roles(self, text):
        self.roles_log_signal.log.emit(text)

    def _show_roles_log(self):
        if self._roles_log_dialog is None or not self._roles_log_dialog.isVisible():
            self._roles_log_dialog = LogDialog(self)
            self._roles_log_dialog.setWindowTitle("Log — Роли и привилегии")
            for chunk in self._roles_log_buffer:
                self._roles_log_dialog.append(chunk)
            self._roles_log_dialog.show()
        else:
            self._roles_log_dialog.raise_()
            self._roles_log_dialog.activateWindow()

    def _append_roles_log(self, text):
        if text == "__SHOW_ROLES_LOG__":
            self._show_roles_log()
            return
        self._roles_log_buffer.append(text)
        try:
            message = str(text).strip().splitlines()[-1]
            if message:
                self.statusBar().showMessage(message[:180], 8000)
        except Exception:
            pass
        if self._roles_log_dialog and self._roles_log_dialog.isVisible():
            self._roles_log_dialog.append(text)

    def _clear_roles_log(self):
        self._roles_log_buffer.clear()
        if self._roles_log_dialog:
            self._roles_log_dialog.clear_log()

    def _start_roles_template(self):
        default_name = "projectpoint_roles_import_template.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить шаблон импорта ролей",
            default_name,
            "Excel (*.xlsx)"
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        selected_module = self.roles_template_module_combo.currentData() if hasattr(self, "roles_template_module_combo") else None

        self.roles_template_btn.setEnabled(False)
        self.roles_import_btn.setEnabled(False)
        self._clear_roles_log()
        threading.Thread(
            target=self._run_roles_template_worker,
            args=(file_path, selected_module),
            daemon=True,
        ).start()

    def _start_roles_import(self):
        file_path = self.roles_file_edit.text()
        if not file_path or not os.path.exists(file_path):
            QMessageBox.warning(self, "Ошибка", "Выберите существующий Excel файл")
            return

        selected_sheets = self._get_selected_roles_sheets()
        if selected_sheets is not None and not selected_sheets:
            QMessageBox.warning(self, "Ошибка", "Выберите хотя бы один лист модуля для импорта")
            return

        auth_args = self._get_worker_auth_args()
        if not auth_args:
            QMessageBox.warning(self, "Ошибка", "Заполните BASE_URL, USERNAME и PASSWORD")
            return

        existing_mode = self.roles_existing_combo.currentData() or "skip"
        self.roles_template_btn.setEnabled(False)
        self.roles_import_btn.setEnabled(False)
        self._clear_roles_log()
        threading.Thread(
            target=self._run_roles_import_worker,
            args=(file_path, selected_sheets, existing_mode, *auth_args),
            daemon=True,
        ).start()

    def _run_roles_template_worker(
        self,
        output_file,
        selected_module,
    ):
        try:
            self._log_roles("=" * 60 + "\n")
            self._log_roles("Формирование шаблона импорта ролей без подключения\n")
            self._log_roles(f"Файл шаблона: {output_file}\n")
            self._log_roles(f"Фильтр модуля: {selected_module or 'все видимые модули'}\n")
            self._log_roles(f"Источник справочника: {BUILTIN_ROLE_PERMISSIONS_SOURCE}\n")
            self._log_roles("=" * 60 + "\n\n")

            self._log_roles("[1/2] Загрузка встроенного справочника прав...\n")
            all_permissions = get_builtin_role_permissions()
            active_modules = None
            visible_module_ids = get_visible_permission_module_ids(active_modules)
            self._log_roles(f"  Прав во встроенном справочнике: {len(all_permissions)}\n")
            self._log_roles(f"  Модули прав, видимые как в интерфейсе: {sorted(visible_module_ids)}\n")
            self._log_roles("  Примечание: при импорте будет выполнена повторная проверка по актуальному серверу.\n\n")

            self._log_roles("[2/2] Формирование Excel-шаблона...\n")
            selected_modules = [selected_module] if selected_module else None
            result = write_roles_import_template(
                all_permissions,
                output_file,
                active_modules=active_modules,
                selected_modules=selected_modules,
            )
            self._log_roles(f"  Шаблон сохранен: {result['filename']}\n")
            self._log_roles(f"  Модулей в шаблоне: {result.get('modules_count', 0)}\n")
            if result.get('sheets'):
                self._log_roles(f"  Листы: {', '.join(result['sheets'])}\n")
            self._log_roles(f"  Прав в шаблоне: {result['permissions_count']}\n")
            self._log_roles("=" * 60 + "\n")
            self.roles_log_signal.log.emit("__SHOW_ROLES_LOG__")

        except Exception as e:
            self._log_roles(f"\nФАТАЛЬНАЯ ОШИБКА: {e}\n")
            self._log_roles(traceback.format_exc())
            self.roles_log_signal.log.emit("__SHOW_ROLES_LOG__")
        finally:
            self.roles_log_signal.finished.emit()

    def _run_roles_import_worker(
        self,
        excel_file,
        selected_sheets,
        existing_mode,
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
            self._log_roles("=" * 60 + "\n")
            self._log_roles("Запуск импорта ролей и привилегий из Excel\n")
            self._log_roles(f"Base URL: {base_url}\n")
            self._log_roles(f"Auth mode: {auth_mode}\n")
            self._log_roles(f"Excel: {excel_file}\n")
            self._log_roles(f"Листы для импорта: {', '.join(selected_sheets) if selected_sheets else 'все найденные листы'}\n")
            self._log_roles(f"Режим существующих ролей: {existing_mode}\n")
            self._log_roles("=" * 60 + "\n\n")

            self._log_roles("[1/6] Чтение Excel...\n")
            roles_meta, permission_rows = read_roles_import_excel(excel_file, include_sheets=selected_sheets)
            self._log_roles(f"  Ролей в Excel: {len(roles_meta)}\n")
            self._log_roles(f"  Строк прав в Excel: {len(permission_rows)}\n\n")

            self._log_roles("[2/6] Авторизация...\n")
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
            self._log_roles("  access_token получен\n\n")

            self._log_roles("[3/6] Загрузка справочников Project Point...\n")
            all_permissions = get_all_permissions(session, access_token, base_url)
            active_modules = get_active_modules(session, access_token, base_url)
            existing_roles = get_all_roles(session, access_token, base_url)
            existing_names = {
                _normalize_lookup_text(role.get("Name")).lower(): role
                for role in existing_roles
                if _normalize_lookup_text(role.get("Name"))
            }
            self._log_roles(f"  Прав в справочнике: {len(all_permissions)}\n")
            self._log_roles(f"  Активных модулей: {len(active_modules)}\n")
            self._log_roles(f"  Существующих ролей: {len(existing_roles)}\n\n")

            self._log_roles("[4/6] Валидация и подготовка payload...\n")
            payloads, changed_counts, duplicate_count, imported_permission_ids_by_role = build_role_payloads_from_import(
                roles_meta,
                permission_rows,
                all_permissions,
                active_modules=active_modules,
            )
            if duplicate_count:
                self._log_roles(f"  ВНИМАНИЕ: найдено повторных строк прав: {duplicate_count}. Использовано последнее значение.\n")
            self._log_roles(f"  Payload подготовлен для ролей: {len(payloads)}\n\n")

            if existing_mode == "error":
                duplicates = [p["Name"] for p in payloads if p["Name"].lower() in existing_names]
                if duplicates:
                    raise Exception("Роли уже существуют в пространстве: " + ", ".join(duplicates[:50]))

            self._log_roles("[5/6] Создание / обновление ролей через RoleService...\n")
            self._log_roles("-" * 60 + "\n")

            stats = {"created": 0, "updated": 0, "skipped": 0, "errors": 0}
            for idx, payload in enumerate(payloads, 1):
                role_name = payload["Name"]
                role_key = role_name.lower()
                existing_role = existing_names.get(role_key)
                self._log_roles(f"\n  [{idx}] Роль: '{role_name}'\n")

                true_permissions = changed_counts.get(role_name, 0)
                imported_permission_ids = imported_permission_ids_by_role.get(role_name, set())
                self._log_roles(f"      Прав с выданным доступом: {true_permissions}\n")
                self._log_roles(f"      Прав из выбранных листов: {len(imported_permission_ids)}\n")

                if existing_role and existing_mode == "skip":
                    self._log_roles("      ПРОПУСК: роль уже существует\n")
                    stats["skipped"] += 1
                    continue

                if existing_role and existing_mode == "update":
                    update_payload = build_update_role_payload(existing_role, payload, imported_permission_ids)
                    self._log_roles(f"      ОБНОВЛЕНИЕ существующей роли ID={update_payload.get('Id')}\n")
                    self._log_roles(f"      Всего прав в update payload: {len(update_payload.get('Permissions', []))}\n")
                    success, status_code, resp_data = update_role(session, access_token, base_url, update_payload)
                    if success:
                        self._log_roles("      ОБНОВЛЕНА\n")
                        existing_names[role_key] = {**existing_role, **update_payload}
                        stats["updated"] += 1
                    else:
                        self._log_roles(f"      ОШИБКА ОБНОВЛЕНИЯ: HTTP {status_code}\n")
                        resp_str = (
                            json.dumps(resp_data, ensure_ascii=False)
                            if isinstance(resp_data, (dict, list))
                            else str(resp_data)
                        )
                        self._log_roles(f"      Ответ: {resp_str[:1000]}\n")
                        stats["errors"] += 1
                    continue

                self._log_roles(f"      Всего прав в create payload: {len(payload.get('Permissions', []))}\n")
                success, status_code, resp_data = create_role(session, access_token, base_url, payload)
                if success:
                    role_id = resp_data.get("Id", "?") if isinstance(resp_data, dict) else str(resp_data)
                    self._log_roles(f"      СОЗДАНА  ID={role_id}\n")
                    existing_names[role_key] = {"Id": role_id, "Name": role_name, "Permissions": payload.get("Permissions", [])}
                    stats["created"] += 1
                else:
                    self._log_roles(f"      ОШИБКА: HTTP {status_code}\n")
                    resp_str = (
                        json.dumps(resp_data, ensure_ascii=False)
                        if isinstance(resp_data, (dict, list))
                        else str(resp_data)
                    )
                    self._log_roles(f"      Ответ: {resp_str[:1000]}\n")
                    stats["errors"] += 1

            self._log_roles("\n" + "=" * 60 + "\n")
            self._log_roles("[6/6] ИТОГИ\n")
            self._log_roles(f"  Создано:    {stats['created']}\n")
            self._log_roles(f"  Обновлено:  {stats['updated']}\n")
            self._log_roles(f"  Пропущено:  {stats['skipped']}\n")
            self._log_roles(f"  Ошибки:     {stats['errors']}\n")
            self._log_roles("=" * 60 + "\n")
            self.roles_log_signal.log.emit("__SHOW_ROLES_LOG__")

        except Exception as e:
            self._log_roles(f"\nФАТАЛЬНАЯ ОШИБКА: {e}\n")
            self._log_roles(traceback.format_exc())
            self.roles_log_signal.log.emit("__SHOW_ROLES_LOG__")
        finally:
            self.roles_log_signal.finished.emit()
