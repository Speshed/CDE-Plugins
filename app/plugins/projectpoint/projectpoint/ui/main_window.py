from __future__ import annotations

import sys

from PySide6 import QtCore, QtGui
from PySide6.QtWidgets import (
    QComboBox, QFileDialog, QFrame, QHBoxLayout, QLabel, QLineEdit, QMainWindow,
    QMessageBox, QPushButton, QScrollArea, QTabWidget, QVBoxLayout, QWidget,
)

from ..config import *
from ..styles import (
    ThemeToggle, UI_ASSET_DIR, build_qss, themed_icon,
    initial_dark_theme, persist_dark_theme, apply_windows_titlebar_theme,
)
from shared.ui_components import (
    NoWheelTabBar, add_standard_header_controls, install_status_bar, mark_destructive_buttons,
)
from shared.theme_core import shared_style_overrides, install_window_state
from .dialogs import ConnectionDialog, LogDialog, LogSignal, ObjectsLoadSignal
from .features import ProjectStagesMixin, ParametersMixin, ContentTypesMixin, RoutesMixin, RolesMixin, ObjectsMixin, ConnectionMixin

__all__ = ["MainWindow"]


class MainWindow(ParametersMixin, ProjectStagesMixin, ContentTypesMixin, RoutesMixin, RolesMixin, ObjectsMixin, ConnectionMixin, QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Larix CDE — Project Point")
        self.resize(1200, 700)
        self.setMinimumSize(1020, 620)
        self.asset_dir = UI_ASSET_DIR
        self.is_dark_theme = initial_dark_theme(False)
        app_icon = self.asset_dir / "icon.ico"
        if app_icon.exists():
            self.setWindowIcon(QtGui.QIcon(str(app_icon)))

        self.param_log_signal = LogSignal()
        self.param_log_signal.log.connect(self._append_param_log)
        self.param_log_signal.finished.connect(self._finish_param_worker)
        self._param_log_buffer = []
        self._param_log_dialog = None

        self.stage_log_signal = LogSignal()
        self.stage_log_signal.log.connect(self._append_project_stages_log)
        self.stage_log_signal.finished.connect(self._finish_project_stages_worker)
        self._stage_log_buffer = []
        self._stage_log_dialog = None

        self.type_log_signal = LogSignal()
        self.type_log_signal.log.connect(self._append_type_log)
        self.type_log_signal.finished.connect(self._finish_type_worker)
        self._type_log_buffer = []
        self._type_log_dialog = None

        self.route_log_signal = LogSignal()
        self.route_log_signal.log.connect(self._append_route_log)
        self.route_log_signal.finished.connect(self._finish_route_worker)
        self._route_log_buffer = []
        self._route_log_dialog = None

        self.roles_log_signal = LogSignal()
        self.roles_log_signal.log.connect(self._append_roles_log)
        self.roles_log_signal.finished.connect(self._finish_roles_worker)
        self._roles_log_buffer = []
        self._roles_log_dialog = None

        self.objects_log_signal = LogSignal()
        self.objects_log_signal.log.connect(self._append_objects_log)
        self.objects_log_signal.finished.connect(self._finish_objects_create_worker)
        self._objects_log_buffer = []
        self._objects_log_dialog = None

        self.objects_load_signal = ObjectsLoadSignal()
        self.objects_load_signal.projects_loaded.connect(self._apply_loaded_projects)
        self.objects_load_signal.projects_finished.connect(self._finish_objects_projects_load)
        self.objects_load_signal.loaded.connect(self._apply_loaded_objects)
        self.objects_load_signal.finished.connect(self._finish_objects_load)

        self.objects_projects = []
        self.selected_objects_project_id = None
        self.objects_items = []
        self.objects_indexes = {"code_to_id": {}, "title_to_objects": {}}
        self.selected_object_parent_id = None
        self.objects_parent_mode = "excel"
        self._objects_tab_initialized = False
        self._objects_projects_loading = False
        self._objects_parents_loading = False

        self._build_ui()
        install_status_bar(self)
        mark_destructive_buttons(self)
        self._apply_styles()
        install_window_state(self, "cde_tool")

    def _build_ui(self):
        central = QWidget()
        central.setObjectName("centralRoot")
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(18, 14, 18, 14)
        root.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 2)
        header.setSpacing(12)

        title_col = QVBoxLayout()
        title_col.setSpacing(2)
        title = QLabel("Project Point")
        title.setObjectName("pageTitle")
        subtitle = QLabel("Параметры, виды документов, типы, маршруты, роли и объекты.")
        subtitle.setObjectName("pageSubtitle")
        title_col.addWidget(title)
        title_col.addWidget(subtitle)
        header.addLayout(title_col, 1)

        self.theme_toggle = ThemeToggle(self.asset_dir, self)
        self.theme_toggle.setToolTip("Светлая / тёмная тема")
        self.theme_toggle.toggled.connect(self._toggle_theme)
        self.back_to_manager_button = add_standard_header_controls(header, self, self.theme_toggle)
        root.addLayout(header)

        # Compact connection card. Detailed credentials are edited in a separate dialog.
        auth_card = QFrame()
        auth_card.setObjectName("card")
        auth_layout = QVBoxLayout(auth_card)
        auth_layout.setContentsMargins(14, 9, 14, 10)
        auth_layout.setSpacing(6)

        auth_header = QHBoxLayout()
        auth_title = QLabel("Подключение к Project Point")
        auth_title.setObjectName("cardTitle")
        auth_header.addWidget(auth_title)
        auth_header.addStretch(1)
        self.auth_status = QLabel("Не подключено")
        self.auth_status.setObjectName("authStatus")
        self.auth_status.setProperty("tone", "pending")
        auth_header.addWidget(self.auth_status)
        auth_layout.addLayout(auth_header)

        self.fields = {}

        # Hidden connection widgets/state used by existing auth and worker methods.
        self.profile_combo = QComboBox(self)
        self.profile_combo.addItem("projectpoint.areal.ru", "Production")
        self.profile_combo.addItem("ibim-test.cloud.projectpoint.ru", "Test")
        self.profile_combo.addItem("eyurevich.cloud.projectpoint.ru", "Eyurevich")
        self.profile_combo.addItem("Custom", "Custom")
        self.profile_combo.currentIndexChanged.connect(self._on_profile_changed)
        self.profile_combo.setVisible(False)

        hidden_defaults = {
            "BASE_URL": CONNECTION_PRESETS["Production"]["base_url"],
            "USERNAME": "",
            "PASSWORD": "",
            "CLIENT_ID": DEFAULT_CLIENT_ID,
            "SSO_BASE_URL": DEFAULT_SSO_BASE_URL,
            "REALM": DEFAULT_REALM,
            "BROKER_ALIAS": DEFAULT_BROKER_ALIAS,
            "ADFS_URL": DEFAULT_ADFS_BASE_URL,
        }
        for key, value in hidden_defaults.items():
            edit = QLineEdit(value, self)
            if key == "PASSWORD":
                edit.setEchoMode(QLineEdit.Password)
            edit.setVisible(False)
            self.fields[key] = edit

        # Hidden compatibility controls referenced by older profile/toggle logic.
        self.connection_details_group = QWidget(self)
        self.connection_details_group.setVisible(False)
        self.connection_toggle_btn = QPushButton("Показать настройки", self)
        self.connection_toggle_btn.setVisible(False)
        self.advanced_group = QWidget(self)
        self.advanced_group.setVisible(False)
        self.advanced_toggle_btn = QPushButton("▼ Расширенные настройки", self)
        self.advanced_toggle_btn.setVisible(False)

        compact_row = QHBoxLayout()
        compact_row.setSpacing(10)
        self.connection_summary_label = QLabel("")
        self.connection_summary_label.setObjectName("rowSubtitle")
        compact_row.addWidget(self.connection_summary_label, 1)

        self.auth_btn = QPushButton("Вход")
        self.auth_btn.setObjectName("loginButton")
        self.auth_btn.setMinimumHeight(36)
        self.auth_btn.setMinimumWidth(120)
        self.auth_btn.setCursor(QtCore.Qt.PointingHandCursor)
        self.auth_btn.clicked.connect(self._show_connection_dialog)
        compact_row.addWidget(self.auth_btn)
        auth_layout.addLayout(compact_row)

        self.auth_config_source = QLabel("")
        self.auth_config_source.setObjectName("rowSubtitle")
        auth_layout.addWidget(self.auth_config_source)

        self._on_profile_changed(self.profile_combo.currentIndex())
        self._update_connection_summary()
        root.addWidget(auth_card)

        self.tabs = QTabWidget()
        self.tabs.setTabBar(NoWheelTabBar())
        self.tabs.setObjectName("workspaceTabs")
        self.tabs.setDocumentMode(True)
        self.tabs.setMovable(False)
        self.tabs.tabBar().setDrawBase(False)
        # The six workspace tabs act as the main navigation. Keep them evenly
        # distributed across the full available width instead of packing them
        # against the left edge with title-dependent widths.
        self.tabs.tabBar().setExpanding(True)
        self.tabs.tabBar().setUsesScrollButtons(False)
        self.tabs.tabBar().setElideMode(QtCore.Qt.ElideRight)
        root.addWidget(self.tabs, 1)

        self._build_params_tab()
        self._build_project_stages_tab()
        self._build_types_tab()
        self._build_routes_tab()
        self._build_roles_tab()
        self._build_objects_tab()
        for index in range(self.tabs.count()):
            self.tabs.widget(index).setObjectName("tabPage")
        self._wrap_workspace_tabs_in_scroll_areas()
        self.tabs.currentChanged.connect(self._on_tab_changed)

    def _wrap_workspace_tabs_in_scroll_areas(self):
        """Make every workspace page scroll vertically instead of clipping its bottom edge."""
        pages = []
        for index in range(self.tabs.count()):
            pages.append((
                self.tabs.widget(index),
                self.tabs.tabIcon(index),
                self.tabs.tabText(index),
                self.tabs.tabToolTip(index),
            ))

        while self.tabs.count():
            self.tabs.removeTab(0)

        for page, icon, title, tooltip in pages:
            scroll = QScrollArea()
            scroll.setObjectName("workspaceScrollArea")
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
            scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
            scroll.setAlignment(QtCore.Qt.AlignTop)
            scroll.setWidget(page)
            index = self.tabs.addTab(scroll, icon, title)
            if tooltip:
                self.tabs.setTabToolTip(index, tooltip)

    def _toggle_theme(self, dark=None):
        self.is_dark_theme = (not self.is_dark_theme) if dark is None else bool(dark)
        persist_dark_theme(self.is_dark_theme)
        self._apply_styles()

    def _apply_styles(self):
        self.setStyleSheet(build_qss(self.is_dark_theme, self.asset_dir) + shared_style_overrides(self.is_dark_theme))
        app = QtCore.QCoreApplication.instance()
        if app is not None:
            app.setProperty("projectpoint_dark", bool(self.is_dark_theme))
        self.theme_toggle.blockSignals(True)
        self.theme_toggle.setChecked(self.is_dark_theme, animate=False)
        self.theme_toggle.blockSignals(False)
        self._refresh_action_icons()
        self._set_native_titlebar_theme(self.is_dark_theme)
        QtCore.QTimer.singleShot(0, lambda: self._set_native_titlebar_theme(self.is_dark_theme))

    def _set_button_icon(self, button, asset_name, size=16):
        if button is None:
            return
        button.setProperty("iconAsset", asset_name)
        button.setIcon(themed_icon(self.asset_dir, asset_name, self.is_dark_theme))
        button.setIconSize(QtCore.QSize(size, size))

    def _refresh_action_icons(self):
        for button in self.findChildren(QPushButton):
            text = button.text().strip().lower()
            asset = button.property("iconAsset")
            if bool(button.property("fileLoaded")):
                asset = "krug_galka.png"
            elif not asset:
                if text == "вход":
                    asset = "free-icon-login-2623062.png"
                elif text in {"обзор", "выбрать файл", "загрузить файл"}:
                    asset = "upload.png"
                elif "шаблон" in text:
                    asset = "free-icon-download-126488.png"
                elif text in {"log", "подробнее", "детали"}:
                    asset = "information.png"
                elif "предпросмотр" in text or "сверить" in text:
                    asset = "preview.png"
                elif text == "обновить":
                    asset = "free-icon-refresh-5234214.png"
                elif any(word in text for word in ("запустить", "создать", "импортировать", "применить")):
                    asset = "import.png"
            if asset:
                self._set_button_icon(button, str(asset))
        self._refresh_static_icons()

    def _refresh_static_icons(self):
        for label in self.findChildren(QLabel):
            asset = label.property("iconAsset")
            if not asset:
                continue
            try:
                size = int(label.property("iconSize") or 18)
            except Exception:
                size = 18
            label.setPixmap(themed_icon(self.asset_dir, str(asset), self.is_dark_theme).pixmap(QtCore.QSize(size, size)))

    def _set_native_titlebar_theme(self, dark: bool):
        apply_windows_titlebar_theme(self, bool(dark))

    def _finish_param_worker(self):
        self.params_start_btn.setEnabled(True)

    def _finish_project_stages_worker(self):
        self.stages_start_btn.setEnabled(True)

    def _finish_type_worker(self):
        self.types_start_btn.setEnabled(True)

    def _finish_route_worker(self):
        self.routes_start_btn.setEnabled(True)

    def _finish_roles_worker(self):
        self.roles_template_btn.setEnabled(True)
        self.roles_import_btn.setEnabled(True)

    def _finish_objects_create_worker(self):
        self.objects_create_btn.setEnabled(True)

    def _on_tab_changed(self, index):
        if index != getattr(self, "objects_tab_index", -1):
            return
        if self._objects_tab_initialized:
            return
        self._objects_tab_initialized = True
        auth_args = self._get_worker_auth_args()
        if auth_args:
            self._start_objects_projects_load()

    def _save_template_file(self, title, default_name, writer_func):
        file_path, _ = QFileDialog.getSaveFileName(self, title, default_name, "Excel (*.xlsx)")
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"
        try:
            result = writer_func(file_path)
            QMessageBox.information(self, "Готово", f"Шаблон сохранен:\n{result['filename']}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сформировать шаблон:\n{e}")
