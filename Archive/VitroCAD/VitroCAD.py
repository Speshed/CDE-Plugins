# -*- coding: utf-8 -*-
"""VitroCAD — единая оболочка утилит в стиле Larix Platform Plugin."""
from __future__ import annotations

import sys
import traceback
from pathlib import Path
from typing import Any, Dict, Optional

from PySide6 import QtCore, QtGui, QtWidgets

from modules import folder_creator, permissions, schedule_sync
from ui_theme import (
    ThemeToggle, build_vitrocad_qss, themed_icon, themed_pixmap, shared_asset_dir,
    initial_dark_theme, persist_dark_theme, apply_windows_titlebar_theme,
)
from shared.ui_components import (
    configure_preview_table,
    make_details_button,
    show_table_details,
    add_standard_header_controls,
    install_status_bar,
    mark_destructive_buttons,
)
from shared.theme_core import shared_style_overrides, install_window_state


APP_TITLE = "VitroCAD"
APP_VERSION = "1.2.0"
DEFAULT_SERVER = folder_creator.DEFAULT_SERVER
ALT_SERVER = schedule_sync.DEFAULT_BASE_URL


class NoWheelComboBox(QtWidgets.QComboBox):
    def wheelEvent(self, event: QtGui.QWheelEvent) -> None:
        view = self.view()
        if view is not None and view.isVisible():
            super().wheelEvent(event)
            return
        event.ignore()


class NoWheelTabBar(QtWidgets.QTabBar):
    def wheelEvent(self, event: QtGui.QWheelEvent) -> None:
        event.accept()


class PasswordLineEdit(QtWidgets.QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._eye_btn = QtWidgets.QToolButton(self)
        self._eye_btn.setObjectName("passwordEye")
        self._eye_btn.setCursor(QtCore.Qt.PointingHandCursor)
        self._eye_btn.setFixedSize(30, 28)
        self._eye_btn.setIconSize(QtCore.QSize(18, 18))
        self.setTextMargins(0, 0, 38, 0)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        x = self.width() - self._eye_btn.width() - 8
        y = (self.height() - self._eye_btn.height()) // 2
        self._eye_btn.move(x, y)

    def set_eye_clicked(self, callback) -> None:
        self._eye_btn.clicked.connect(callback)


class SharedAuthWorker(QtCore.QThread):
    """Одна авторизация для всех трёх инструментов.

    Вход в /api/security/login выполняется один раз. Полученный токен затем
    передаётся всем модулям. Чтение служебных метаданных каждого раздела
    выполняется здесь же, чтобы не блокировать интерфейс.
    """

    succeeded = QtCore.Signal(object)
    failed = QtCore.Signal(str)
    progress = QtCore.Signal(str)

    def __init__(self, server: str, login: str, password: str):
        super().__init__()
        self.server = server.rstrip("/")
        self.login_value = login
        self.password_value = password

    def run(self) -> None:
        try:
            self.progress.emit("Авторизация...")
            base_client = folder_creator.VitroClient(self.server)
            token = base_client.login(self.login_value, self.password_value)
            result: Dict[str, Any] = {
                "server": self.server,
                "login": self.login_value,
                "token": token,
                "folder_content_types": [],
                "permissions_content_types": [],
                "schedule_api": None,
                "schedule_content_type": None,
                "warnings": [],
            }

            try:
                self.progress.emit("Загрузка типов для структуры папок...")
                result["folder_content_types"] = base_client.get_content_types(folder_creator.LIST_ID)
            except Exception as exc:
                result["warnings"].append(f"Структура папок: {exc}")

            try:
                self.progress.emit("Загрузка типов для прав доступа...")
                perm_client = permissions.VitroClient(self.server, token)
                result["permissions_content_types"] = perm_client.get_content_types(
                    permissions.DEFAULT_FILE_LIST_ID
                )
            except Exception as exc:
                result["warnings"].append(f"Права доступа: {exc}")

            try:
                self.progress.emit("Загрузка метаданных план-графика...")
                api = schedule_sync.LarixApi(
                    self.server,
                    schedule_sync.DEFAULT_LIST_ID,
                    schedule_sync.DEFAULT_CONTENT_TYPE_ID,
                    schedule_sync.VERIFY_SSL,
                )
                api.authorization_token = token
                api.current_user_name = self.login_value
                api._apply_auth_variant(0)
                content_type = api.load_content_type()
                api.default_instance = api.create_new_instance(api.list_id)
                result["schedule_api"] = api
                result["schedule_content_type"] = content_type
            except Exception as exc:
                result["warnings"].append(f"План-график: {exc}")

            self.succeeded.emit(result)
        except Exception as exc:
            self.failed.emit(f"{exc}\n\n{traceback.format_exc()}")


class ToolShell(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Larix CDE — Vitro")
        self.resize(1200, 700)
        self.setMinimumSize(1020, 620)
        self.asset_dir = shared_asset_dir(Path(__file__).resolve().parent)
        self.is_dark_theme = initial_dark_theme(False)
        self._password_visible = False
        self._authorized = False
        self._auth_worker: Optional[SharedAuthWorker] = None

        icon_path = self.asset_dir / "icon.ico"
        if icon_path.exists():
            self.setWindowIcon(QtGui.QIcon(str(icon_path)))

        # Контроллеры сохраняются как QMainWindow: вся существующая бизнес-логика,
        # воркеры и диалоги продолжают работать без переписывания API-кода.
        self.folder_page = folder_creator.MainWindow()
        self.permissions_page = permissions.MainWindow()
        self.schedule_page = schedule_sync.MainWindow()
        self.controllers = [self.folder_page, self.permissions_page, self.schedule_page]
        self._legacy_centrals = []

        for controller in self.controllers:
            try:
                controller.statusBar().messageChanged.connect(
                    lambda message, shell=self: shell.statusBar().showMessage(message)
                )
            except Exception:
                pass

        self._rebuild_tool_pages()
        self._setup_ui()
        install_status_bar(self)
        mark_destructive_buttons(self)
        self._apply_theme()
        install_window_state(self, "cde_tool")

    # ------------------------------------------------------------------
    # Общие UI helpers — повторяют композицию референсного Larix plugin
    # ------------------------------------------------------------------
    def _make_badge(self, asset_name: str, tone: str = "orange", size: int = 32) -> QtWidgets.QLabel:
        badge = QtWidgets.QLabel()
        badge.setObjectName("iconBadge")
        badge.setProperty("tone", tone)
        badge.setAlignment(QtCore.Qt.AlignCenter)
        badge.setFixedSize(size, size)
        badge.setProperty("iconAsset", asset_name)
        badge.setProperty("iconSize", 18)
        badge.setPixmap(themed_pixmap(self.asset_dir, asset_name, self.is_dark_theme, 18))
        return badge

    def _add_card_header(
        self,
        layout: QtWidgets.QVBoxLayout,
        title: str,
        asset_name: str,
        tone: str = "orange",
        right_widget: Optional[QtWidgets.QWidget] = None,
    ) -> None:
        row = QtWidgets.QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)
        row.addWidget(self._make_badge(asset_name, tone, 32))
        label = QtWidgets.QLabel(title)
        label.setObjectName("cardTitle")
        row.addWidget(label)
        row.addStretch(1)
        if right_widget is not None:
            row.addWidget(right_widget)
        layout.addLayout(row)

    @staticmethod
    def _field_label(text: str) -> QtWidgets.QLabel:
        label = QtWidgets.QLabel(text)
        label.setObjectName("fieldLabel")
        return label

    @staticmethod
    def _row_title(text: str) -> QtWidgets.QLabel:
        label = QtWidgets.QLabel(text)
        label.setObjectName("rowTitle")
        return label

    @staticmethod
    def _row_subtitle(text: str) -> QtWidgets.QLabel:
        label = QtWidgets.QLabel(text)
        label.setObjectName("rowSubtitle")
        label.setWordWrap(True)
        return label

    def _make_card(self, title: str, asset: str, tone: str = "orange"):
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(14, 11, 14, 13)
        layout.setSpacing(10)
        self._add_card_header(layout, title, asset, tone)
        return card, layout

    def _make_file_card(
        self,
        title: str,
        subtitle: str,
        primary_button: QtWidgets.QPushButton,
        secondary_button: Optional[QtWidgets.QPushButton] = None,
        tone: str = "green",
    ) -> tuple[QtWidgets.QFrame, QtWidgets.QLabel]:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(14, 11, 14, 8)
        layout.setSpacing(0)

        row_widget = QtWidgets.QWidget()
        row_widget.setObjectName("fileRowLast")
        row = QtWidgets.QHBoxLayout(row_widget)
        row.setContentsMargins(0, 10, 0, 10)
        row.setSpacing(12)
        row.addWidget(self._make_badge("Excel.png", tone, 32))

        text_col = QtWidgets.QVBoxLayout()
        text_col.setContentsMargins(0, 0, 0, 0)
        text_col.setSpacing(2)
        text_col.addWidget(self._row_title(title))
        subtitle_label = self._row_subtitle(subtitle)
        text_col.addWidget(subtitle_label)
        row.addLayout(text_col, 1)

        if secondary_button is not None:
            secondary_button.setObjectName("downloadButton")
            secondary_button.setFixedHeight(34)
            secondary_button.setMinimumWidth(170)
            row.addWidget(secondary_button)
        primary_button.setObjectName("uploadButton")
        primary_button.setFixedHeight(34)
        primary_button.setMinimumWidth(170)
        row.addWidget(primary_button)
        layout.addWidget(row_widget)
        return card, subtitle_label

    def _make_info_banner(self, text: str) -> QtWidgets.QFrame:
        frame = QtWidgets.QFrame()
        frame.setObjectName("infoBanner")
        row = QtWidgets.QHBoxLayout(frame)
        row.setContentsMargins(10, 8, 10, 8)
        row.setSpacing(8)
        icon = QtWidgets.QLabel()
        icon.setObjectName("inlineIcon")
        icon.setFixedSize(20, 20)
        icon.setProperty("iconAsset", "information.png")
        icon.setProperty("iconSize", 16)
        icon.setPixmap(themed_pixmap(self.asset_dir, "information.png", self.is_dark_theme, 16))
        label = QtWidgets.QLabel(text)
        label.setObjectName("infoText")
        label.setWordWrap(True)
        row.addWidget(icon, 0, QtCore.Qt.AlignTop)
        row.addWidget(label, 1)
        return frame

    # ------------------------------------------------------------------
    # Полная перекомпоновка трёх старых окон в карточки Larix
    # ------------------------------------------------------------------
    def _replace_central(self, page: QtWidgets.QMainWindow, central: QtWidgets.QWidget) -> None:
        old = page.takeCentralWidget()
        page.setCentralWidget(central)
        page.statusBar().hide()
        if old is not None:
            # Часть служебных полей старых модулей остаётся источником настроек
            # для бизнес-логики, хотя больше не показывается пользователю.
            # Сохраняем старый central скрытым, вместо deleteLater.
            old.hide()
            old.setParent(None)
            self._legacy_centrals.append(old)

    def _rebuild_folder_page(self) -> None:
        p = self.folder_page
        root = QtWidgets.QWidget()
        root.setObjectName("tabPage")
        layout = QtWidgets.QVBoxLayout(root)
        layout.setContentsMargins(2, 12, 2, 2)
        layout.setSpacing(12)

        p.btn_pick_excel.setText("Загрузить файл")
        p.btn_types.setText("Типы с сервера")
        file_card, self.folder_file_subtitle = self._make_file_card(
            "Excel-файл структуры",
            "Выберите заполненный Excel-файл с иерархией папок проекта.",
            p.btn_pick_excel,
            p.btn_types,
        )
        layout.addWidget(file_card)
        p.ed_excel.hide()
        p.ed_excel.textChanged.connect(self._update_folder_file_caption)

        content = QtWidgets.QHBoxLayout()
        content.setSpacing(12)

        tree_card, tree_layout = self._make_card(
            "Целевая папка / проект", "folder_icon_variant_1.png", "orange"
        )
        tree_layout.addWidget(
            self._make_info_banner("Выберите в дереве папку или проект, внутри которого нужно создать структуру.")
        )
        p.tree_folders.setObjectName("matrixTable")
        tree_layout.addWidget(p.tree_folders, 1)
        content.addWidget(tree_card, 1)

        plan_card, plan_layout = self._make_card("Предпросмотр плана", "comparison.png", "green")
        p.tbl_plan.setObjectName("matrixTable")
        configure_preview_table(p.tbl_plan, min_height=135, max_height=220)
        plan_layout.addWidget(p.tbl_plan, 1)
        action_row = QtWidgets.QHBoxLayout()
        action_row.setSpacing(8)
        action_row.addStretch(1)
        self.folder_details_btn = make_details_button()
        self.folder_details_btn.clicked.connect(
            lambda: show_table_details(
                self, p.tbl_plan, "Структура папок — Подробнее",
                "Полный план создания структуры папок VitroCAD."
            )
        )
        action_row.addWidget(self.folder_details_btn)
        p.btn_preview.setText("Построить предпросмотр")
        p.btn_preview.setObjectName("secondaryButton")
        p.btn_preview.setFixedHeight(38)
        p.btn_create.setText("Создать структуру")
        p.btn_create.setObjectName("orangeAction")
        p.btn_create.setFixedHeight(38)
        p.btn_create.setMinimumWidth(180)
        action_row.addWidget(p.btn_preview)
        action_row.addWidget(p.btn_create)
        plan_layout.addLayout(action_row)
        content.addWidget(plan_card, 2)
        layout.addLayout(content, 1)

        p.btn_log.hide()
        p.lbl_auth.hide()
        p.ed_server.hide(); p.ed_login.hide(); p.ed_password.hide(); p.btn_login.hide(); p.btn_toggle_password.hide()
        self._replace_central(p, root)

    def _rebuild_permissions_page(self) -> None:
        p = self.permissions_page
        root = QtWidgets.QWidget()
        root.setObjectName("tabPage")
        layout = QtWidgets.QVBoxLayout(root)
        layout.setContentsMargins(2, 12, 2, 2)
        layout.setSpacing(12)

        p.btn_excel.setText("Загрузить файл")
        file_card, self.permissions_file_subtitle = self._make_file_card(
            "Excel-файл матрицы прав",
            "Таблица путей к папкам и прав пользователей/групп.",
            p.btn_excel,
        )
        layout.addWidget(file_card)
        p.ed_excel.hide()
        p.ed_excel.textChanged.connect(self._update_permissions_file_caption)

        plan_card, plan_layout = self._make_card("Проверка и применение прав", "access_icon_variant_1.png", "orange")
        plan_layout.addWidget(
            self._make_info_banner(
                "Папки ищутся по полным путям из Excel. «Проверить» строит безопасный предпросмотр; «Применить права» сначала проверяет данные и затем запрашивает подтверждение."
            )
        )
        p.tbl_plan.setObjectName("matrixTable")
        configure_preview_table(p.tbl_plan, min_height=135, max_height=220)
        plan_layout.addWidget(p.tbl_plan, 1)
        action_row = QtWidgets.QHBoxLayout()
        action_row.setSpacing(8)
        action_row.addStretch(1)
        self.permissions_details_btn = make_details_button()
        self.permissions_details_btn.clicked.connect(
            lambda: show_table_details(
                self, p.tbl_plan, "Права доступа — Подробнее",
                "Полный результат проверки матрицы прав VitroCAD."
            )
        )
        action_row.addWidget(self.permissions_details_btn)
        p.btn_preview.setText("Проверить")
        p.btn_preview.setObjectName("secondaryButton")
        p.btn_preview.setFixedHeight(38)
        p.btn_apply.setText("Применить права")
        p.btn_apply.setObjectName("orangeAction")
        p.btn_apply.setFixedHeight(38)
        p.btn_apply.setMinimumWidth(170)
        action_row.addWidget(p.btn_preview)
        action_row.addWidget(p.btn_apply)
        plan_layout.addLayout(action_row)
        layout.addWidget(plan_card, 1)

        p.btn_log.hide()
        p.lbl_auth.hide()
        p.cmb_server.hide(); p.ed_login.hide(); p.ed_password.hide(); p.btn_login.hide(); p.btn_toggle_password.hide()
        p.tree.hide()
        self._replace_central(p, root)

    def _rebuild_schedule_page(self) -> None:
        p = self.schedule_page
        root = QtWidgets.QWidget()
        root.setObjectName("tabPage")
        layout = QtWidgets.QVBoxLayout(root)
        layout.setContentsMargins(2, 12, 2, 2)
        layout.setSpacing(12)

        p.btn_excel.setText("Загрузить файл")
        file_card, self.schedule_file_subtitle = self._make_file_card(
            "Excel-файл план-графика",
            "Загрузите план-график и выберите лист для сопоставления с VitroCAD.",
            p.btn_excel,
        )
        file_inner = file_card.layout()
        sheet_row = QtWidgets.QHBoxLayout()
        sheet_row.setSpacing(10)
        sheet_row.addWidget(self._field_label("Лист Excel"))
        p.cb_sheet.setMinimumWidth(260)
        sheet_row.addWidget(p.cb_sheet, 1)
        file_inner.addLayout(sheet_row)
        layout.addWidget(file_card)

        p.lbl_excel.hide()
        # lbl_excel получает имя файла из существующей логики, поэтому используем
        # его как источник текста для новой подписи.
        p.btn_excel.clicked.connect(
            lambda: QtCore.QTimer.singleShot(0, lambda: self._update_schedule_file_caption(p.lbl_excel.text()))
        )

        middle = QtWidgets.QHBoxLayout()
        middle.setSpacing(12)

        required_card, required_layout = self._make_card("Проверка параметров", "comparison.png", "orange")
        p.lbl_required.setObjectName("rowSubtitle")
        required_layout.addWidget(p.lbl_required)
        p.required_table.setObjectName("matrixTable")
        required_layout.addWidget(p.required_table, 1)
        p.btn_required.setText("Проверить обязательные поля")
        p.btn_required.setObjectName("secondaryButton")
        p.btn_required.setFixedHeight(36)
        required_layout.addWidget(p.btn_required, 0, QtCore.Qt.AlignRight)
        middle.addWidget(required_card, 1)

        preview_card, preview_layout = self._make_card("Предпросмотр данных", "structure.png", "green")
        p.preview_table.setObjectName("matrixTable")
        configure_preview_table(p.preview_table, min_height=135, max_height=220)
        preview_layout.addWidget(p.preview_table, 1)
        preview_footer = QtWidgets.QHBoxLayout()
        preview_footer.addStretch(1)
        self.schedule_details_btn = make_details_button()
        self.schedule_details_btn.clicked.connect(
            lambda: show_table_details(
                self, p.preview_table, "План-график — Подробнее",
                "Полный локальный предпросмотр данных план-графика перед синхронизацией."
            )
        )
        preview_footer.addWidget(self.schedule_details_btn)
        preview_layout.addLayout(preview_footer)
        middle.addWidget(preview_card, 2)
        layout.addLayout(middle, 1)

        run_card, run_layout = self._make_card("Синхронизация план-графика", "sync.png", "orange")
        run_layout.addWidget(
            self._make_info_banner(
                "Программа сверяет полный путь из столбца «Название»: найденные элементы обновляются, отсутствующие создаются. Ничего не удаляется."
            )
        )
        run_layout.addWidget(p.progress)
        run_row = QtWidgets.QHBoxLayout()
        run_row.addStretch(1)
        p.btn_import.setText("Синхронизировать")
        p.btn_import.setObjectName("orangeAction")
        p.btn_import.setFixedHeight(40)
        p.btn_import.setMinimumWidth(190)
        run_row.addWidget(p.btn_import)
        run_layout.addLayout(run_row)
        layout.addWidget(run_card)

        p.btn_log.hide(); p.btn_theme.hide(); p.lbl_auth.hide()
        p.ed_server.hide(); p.ed_login.hide(); p.ed_password.hide(); p.btn_login.hide()
        self._replace_central(p, root)

    def _rebuild_tool_pages(self) -> None:
        self._rebuild_folder_page()
        self._rebuild_permissions_page()
        self._rebuild_schedule_page()

        # Иконки действий — тот же набор, что в референсном проекте.
        icon_map = [
            (self.folder_page.btn_pick_excel, "upload.png"),
            (self.folder_page.btn_types, "structure.png"),
            (self.folder_page.btn_preview, "comparison.png"),
            (self.folder_page.btn_create, "krug_galka.png"),
            (self.permissions_page.btn_excel, "upload.png"),
            (self.permissions_page.btn_preview, "comparison.png"),
            (self.permissions_page.btn_apply, "krug_galka.png"),
            (self.schedule_page.btn_excel, "upload.png"),
            (self.schedule_page.btn_required, "comparison.png"),
            (self.schedule_page.btn_import, "sync.png"),
        ]
        for button, asset in icon_map:
            button.setProperty("vitroIconAsset", asset)
            button.setIconSize(QtCore.QSize(16, 16))

    # ------------------------------------------------------------------
    # Общая оболочка
    # ------------------------------------------------------------------
    def _setup_ui(self) -> None:
        central = QtWidgets.QWidget()
        central.setObjectName("centralRoot")
        self.setCentralWidget(central)
        central_layout = QtWidgets.QVBoxLayout(central)
        central_layout.setContentsMargins(0, 0, 0, 0)
        central_layout.setSpacing(0)

        scroll = QtWidgets.QScrollArea()
        scroll.setObjectName("mainScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        central_layout.addWidget(scroll)

        page = QtWidgets.QWidget()
        page.setObjectName("page")
        scroll.setWidget(page)
        layout = QtWidgets.QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 18)
        layout.setSpacing(12)

        header = QtWidgets.QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 4)
        header.setSpacing(12)
        title_col = QtWidgets.QVBoxLayout()
        title_col.setSpacing(2)
        title = QtWidgets.QLabel("Vitro")
        title.setObjectName("pageTitle")
        subtitle = QtWidgets.QLabel("Структура проектов, права доступа и синхронизация план-графика.")
        subtitle.setObjectName("pageSubtitle")
        title_col.addWidget(title)
        title_col.addWidget(subtitle)
        header.addLayout(title_col, 1)
        self.theme_toggle = ThemeToggle(self.asset_dir, self)
        self.theme_toggle.setToolTip("Светлая / тёмная тема")
        self.theme_toggle.toggled.connect(self._toggle_theme)
        self.back_to_manager_button = add_standard_header_controls(header, self, self.theme_toggle)
        layout.addLayout(header)

        # ОДНА общая авторизация на всё приложение.
        auth_card = QtWidgets.QFrame()
        auth_card.setObjectName("card")
        auth_layout = QtWidgets.QVBoxLayout(auth_card)
        auth_layout.setContentsMargins(14, 11, 14, 13)
        auth_layout.setSpacing(8)
        self.auth_text = QtWidgets.QLabel("Статус: Не авторизован")
        self.auth_text.setObjectName("authStatus")
        self._add_card_header(
            auth_layout,
            "Подключение к Larix",
            "free-icon-login-2623062.png",
            "purple",
            self.auth_text,
        )

        grid = QtWidgets.QGridLayout()
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(6)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        grid.setColumnMinimumWidth(2, 112)

        self.cb_server = NoWheelComboBox()
        self.cb_server.setEditable(True)
        self.cb_server.addItems(list(dict.fromkeys([DEFAULT_SERVER, ALT_SERVER])))
        self.cb_server.setCurrentText(DEFAULT_SERVER)
        self.ed_login = QtWidgets.QLineEdit()
        self.ed_login.setPlaceholderText("Логин")
        self.ed_password = PasswordLineEdit()
        self.ed_password.setPlaceholderText("Пароль")
        self.ed_password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.ed_password.set_eye_clicked(self._toggle_password_visibility)
        self.btn_login = QtWidgets.QPushButton("Войти")
        self.btn_login.setObjectName("loginButton")
        self.btn_login.setFixedHeight(36)
        self.btn_login.setMinimumWidth(112)
        self.btn_login.setCursor(QtCore.Qt.PointingHandCursor)

        grid.addWidget(self._field_label("Сервер"), 0, 0, 1, 3)
        grid.addWidget(self.cb_server, 1, 0, 1, 3)
        grid.addWidget(self._field_label("Логин"), 2, 0)
        grid.addWidget(self._field_label("Пароль"), 2, 1)
        grid.addWidget(self.ed_login, 3, 0)
        grid.addWidget(self.ed_password, 3, 1)
        grid.addWidget(self.btn_login, 3, 2)
        auth_layout.addLayout(grid)
        layout.addWidget(auth_card)

        self.main_tabs = QtWidgets.QTabWidget()
        self.main_tabs.setTabBar(NoWheelTabBar())
        self.main_tabs.setObjectName("workspaceTabs")
        self.main_tabs.setDocumentMode(True)
        self.main_tabs.setMovable(False)
        self.main_tabs.tabBar().setDrawBase(False)
        self.main_tabs.tabBar().setExpanding(True)
        self.main_tabs.tabBar().setUsesScrollButtons(False)
        self.main_tabs.tabBar().setElideMode(QtCore.Qt.ElideNone)

        for controller, text in (
            (self.folder_page, "Структура папок"),
            (self.permissions_page, "Права доступа"),
            (self.schedule_page, "План-график"),
        ):
            controller.setParent(self.main_tabs)
            controller.setWindowFlags(QtCore.Qt.WindowType.Widget)
            controller.setMinimumSize(0, 0)
            controller.setObjectName("toolPage")
            self.main_tabs.addTab(controller, text)
        layout.addWidget(self.main_tabs, 1)

        # Один общий лог, как в референсе.
        log_bar = QtWidgets.QFrame()
        log_bar.setObjectName("logBar")
        log_layout = QtWidgets.QHBoxLayout(log_bar)
        log_layout.setContentsMargins(10, 0, 8, 0)
        log_layout.setSpacing(6)
        self.btn_log = QtWidgets.QPushButton("Открыть лог")
        self.btn_log.setObjectName("logButton")
        self.btn_log.setCursor(QtCore.Qt.PointingHandCursor)
        log_layout.addWidget(self.btn_log)
        log_layout.addStretch(1)
        chevron = QtWidgets.QLabel("›")
        chevron.setObjectName("chevron")
        log_layout.addWidget(chevron)
        layout.addWidget(log_bar)

        self.btn_login.clicked.connect(self._do_login)
        self.ed_login.returnPressed.connect(self._do_login)
        self.ed_password.returnPressed.connect(self._do_login)
        self.cb_server.currentTextChanged.connect(self._server_changed)
        self.btn_log.clicked.connect(self._show_current_log)
        self.main_tabs.currentChanged.connect(lambda _i: self._update_log_tooltip())
        self._update_password_eye()
        self._update_log_tooltip()

    # ------------------------------------------------------------------
    # Файловые подписи
    # ------------------------------------------------------------------
    def _update_folder_file_caption(self, path: str) -> None:
        self.folder_file_subtitle.setText(
            Path(path).name if path.strip() else "Выберите заполненный Excel-файл с иерархией папок проекта."
        )

    def _update_permissions_file_caption(self, path: str) -> None:
        self.permissions_file_subtitle.setText(
            Path(path).name if path.strip() else "Таблица путей к папкам и прав пользователей/групп."
        )

    def _update_schedule_file_caption(self, text: str) -> None:
        value = text.strip()
        self.schedule_file_subtitle.setText(
            value if value and value != "Файл не выбран" else "Загрузите план-график и выберите лист для сопоставления с VitroCAD."
        )

    # ------------------------------------------------------------------
    # Авторизация
    # ------------------------------------------------------------------
    def _toggle_password_visibility(self) -> None:
        self._password_visible = not self._password_visible
        self.ed_password.setEchoMode(
            QtWidgets.QLineEdit.Normal if self._password_visible else QtWidgets.QLineEdit.Password
        )
        self._update_password_eye()

    def _update_password_eye(self) -> None:
        name = "free-icon-hide-11238328.png" if self._password_visible else "free-icon-eye-2455724.png"
        self.ed_password._eye_btn.setIcon(themed_icon(self.asset_dir, name, self.is_dark_theme))
        self.ed_password._eye_btn.setToolTip("Скрыть пароль" if self._password_visible else "Показать пароль")

    def _set_auth_status(self, state: str, details: str = "") -> None:
        # state: off / busy / ok / error
        if state == "ok":
            text = "Статус: Авторизован"
            color = "#4CAF73"
        elif state == "busy":
            text = "Статус: Подключение..."
            color = "#F7921E"
        elif state == "error":
            text = "Статус: Ошибка подключения"
            color = "#D94A4A"
        else:
            text = "Статус: Не авторизован"
            color = "#777777"
        self.auth_text.setText(text)
        self.auth_text.setStyleSheet(
            f"color: {color}; font-weight: 600; background: transparent;"
        )
        self.auth_text.setToolTip(details)

    def _set_auth_busy(self, busy: bool) -> None:
        self.cb_server.setEnabled(not busy)
        self.ed_login.setEnabled(not busy)
        self.ed_password.setEnabled(not busy)
        self.btn_login.setEnabled(not busy)
        self.btn_login.setText("Подключение..." if busy else ("Переподключиться" if self._authorized else "Войти"))

    def _server_changed(self, _text: str) -> None:
        if self._authorized:
            self._clear_authorization()

    def _clear_authorization(self) -> None:
        self._authorized = False
        self.folder_page.auth_token = None
        self.permissions_page.auth_token = ""
        self.schedule_page.api = None
        try:
            self.schedule_page._set_authorized(False)
            self.schedule_page._refresh_actions()
        except Exception:
            pass
        self._set_auth_status("off")
        self.btn_login.setText("Войти")

    def _do_login(self) -> None:
        if self._auth_worker is not None and self._auth_worker.isRunning():
            return
        server = self.cb_server.currentText().strip().rstrip("/")
        login = self.ed_login.text().strip()
        password = self.ed_password.text()
        if not server or not login or not password:
            QtWidgets.QMessageBox.warning(self, "Не хватает данных", "Заполните сервер, логин и пароль.")
            return

        self._clear_authorization()
        self._set_auth_busy(True)
        self._set_auth_status("busy")
        self._auth_worker = SharedAuthWorker(server, login, password)
        self._auth_worker.progress.connect(lambda text: self.auth_text.setToolTip(text))
        self._auth_worker.succeeded.connect(self._auth_ok)
        self._auth_worker.failed.connect(self._auth_failed)
        self._auth_worker.finished.connect(lambda: self._set_auth_busy(False))
        self._auth_worker.start()

    def _auth_ok(self, result: Dict[str, Any]) -> None:
        server = result["server"]
        token = result["token"]
        warnings = list(result.get("warnings") or [])

        # Структура папок
        fp = self.folder_page
        fp.base_url = server
        fp.auth_token = token
        fp.lbl_auth.setText("Авторизован")
        try:
            cts = result.get("folder_content_types") or []
            if cts:
                folder_creator.register_server_content_types(cts)
            fp._load_root_folders()
            fp._log("Авторизация получена из общей панели VitroCAD.")
        except Exception as exc:
            warnings.append(f"Структура папок: {exc}")

        # Права доступа
        pp = self.permissions_page
        pp.base_url = server
        pp.auth_token = token
        pp.file_list_id = permissions.clean_text(pp.ed_file_list_id.text()) or permissions.DEFAULT_FILE_LIST_ID
        pp.scope_list_id = permissions.clean_text(pp.ed_scope_list_id.text()) or permissions.DEFAULT_SCOPE_LIST_ID
        pp.permission_content_type_id = (
            permissions.clean_text(pp.ed_permission_ct_id.text())
            or permissions.DEFAULT_USER_PERMISSION_CONTENT_TYPE_ID
        )
        pp.lbl_auth.setText("Авторизован")
        try:
            cts = result.get("permissions_content_types") or []
            pp.content_types = {
                permissions.clean_text(ct.get("id")): ct
                for ct in cts
                if permissions.clean_text(ct.get("id"))
            }
            pp._load_root_tree()
            pp._log("Авторизация получена из общей панели VitroCAD.\n")
        except Exception as exc:
            warnings.append(f"Права доступа: {exc}")

        # План-график
        sp = self.schedule_page
        api = result.get("schedule_api")
        if api is not None:
            sp.api = api
            sp._set_authorized(True)
            sp._log("✅ Авторизация получена из общей панели VitroCAD.")
            try:
                sp._update_required_table()
            except Exception as exc:
                warnings.append(f"План-график: {exc}")
        else:
            sp.api = None
            sp._set_authorized(False)
            sp._log("⚠ Авторизация выполнена, но метаданные план-графика на выбранном сервере не загружены.")
        sp._refresh_actions()

        self._authorized = True
        details = "\n".join(warnings)
        self._set_auth_status("ok", details)
        self.btn_login.setText("Переподключиться")
        if warnings:
            # Не показываем модальное предупреждение: общий вход успешен, а причины
            # недоступности отдельных разделов доступны по наведению на статус.
            self.auth_text.setText("Статус: Авторизован · есть предупреждения")

    def _auth_failed(self, error: str) -> None:
        self._authorized = False
        self._set_auth_status("error", error)
        self.btn_login.setText("Войти")
        box = QtWidgets.QMessageBox(self)
        box.setWindowTitle("Ошибка подключения")
        box.setText(error.split("\n\n", 1)[0])
        box.setStyleSheet(self.styleSheet())
        box.exec()

    # ------------------------------------------------------------------
    # Лог
    # ------------------------------------------------------------------
    def _show_current_log(self) -> None:
        index = self.main_tabs.currentIndex()
        if index == 0:
            self.folder_page._show_log_dialog()
        elif index == 1:
            self.permissions_page._show_log()
        else:
            self.schedule_page._show_log()

    def _update_log_tooltip(self) -> None:
        names = ["структуры папок", "прав доступа", "план-графика"]
        idx = max(0, min(self.main_tabs.currentIndex(), len(names) - 1))
        self.btn_log.setToolTip(f"Открыть журнал раздела «{names[idx]}»")

    # ------------------------------------------------------------------
    # Тема
    # ------------------------------------------------------------------
    def _toggle_theme(self, dark: bool) -> None:
        self.is_dark_theme = bool(dark)
        persist_dark_theme(self.is_dark_theme)
        self._apply_theme()

    def _refresh_icons(self) -> None:
        # Верхние вкладки в референсе намеренно без иконок.
        for i in range(self.main_tabs.count()):
            self.main_tabs.setTabIcon(i, QtGui.QIcon())

        for widget in self.findChildren(QtWidgets.QLabel):
            asset = widget.property("iconAsset")
            if asset:
                size = int(widget.property("iconSize") or 18)
                widget.setPixmap(themed_pixmap(self.asset_dir, str(asset), self.is_dark_theme, size))

        for page in self.controllers:
            for button in page.findChildren(QtWidgets.QAbstractButton):
                asset = button.property("vitroIconAsset")
                if asset:
                    button.setIcon(themed_icon(self.asset_dir, str(asset), self.is_dark_theme))

        self.btn_login.setIcon(themed_icon(self.asset_dir, "free-icon-login-2623062.png", self.is_dark_theme))
        self.btn_login.setIconSize(QtCore.QSize(16, 16))
        self.btn_log.setIcon(themed_icon(self.asset_dir, "information.png", self.is_dark_theme))
        self.btn_log.setIconSize(QtCore.QSize(15, 15))
        self._update_password_eye()

    def _apply_theme(self) -> None:
        qss = build_vitrocad_qss(self.is_dark_theme, self.asset_dir) + shared_style_overrides(self.is_dark_theme)
        self.setStyleSheet(qss)
        app = QtWidgets.QApplication.instance()
        if app is not None:
            app.setProperty("larix_dark", bool(self.is_dark_theme))

        for controller in self.controllers:
            controller.setStyleSheet(qss)
            if hasattr(controller, "dark_theme"):
                controller.dark_theme = self.is_dark_theme

        self.theme_toggle.blockSignals(True)
        self.theme_toggle.setChecked(self.is_dark_theme, animate=False)
        self.theme_toggle.blockSignals(False)
        self._refresh_icons()
        self._set_native_titlebar_theme(self.is_dark_theme)
        QtCore.QTimer.singleShot(0, lambda: self._set_native_titlebar_theme(self.is_dark_theme))
        QtCore.QTimer.singleShot(120, lambda: self._set_native_titlebar_theme(self.is_dark_theme))

    def _set_native_titlebar_theme(self, dark: bool) -> None:
        apply_windows_titlebar_theme(self, bool(dark))


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setStyle("Fusion")
    window = ToolShell()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
