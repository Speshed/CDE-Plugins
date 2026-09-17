from __future__ import annotations

from PySide6.QtCore import Signal, QObject
from PySide6.QtGui import QFont, QTextCursor
from PySide6.QtWidgets import QComboBox, QDialog, QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTextEdit, QVBoxLayout

from ..config import CONNECTION_PRESETS

__all__ = ["LogSignal", "ObjectsLoadSignal", "LogDialog", "ConnectionDialog"]


class LogSignal(QObject):
    log = Signal(str)
    finished = Signal()


class ObjectsLoadSignal(QObject):
    projects_loaded = Signal(list)
    projects_finished = Signal()
    loaded = Signal(list)
    finished = Signal()


class LogDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Log")
        self.resize(800, 500)
        self.setMinimumSize(500, 300)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        self.text_edit = QTextEdit()
        self.text_edit.setObjectName("logArea")
        self.text_edit.setReadOnly(True)
        self.text_edit.setFont(QFont("Cascadia Code", 10))
        layout.addWidget(self.text_edit)

    def append(self, text):
        self.text_edit.moveCursor(QTextCursor.End)
        self.text_edit.insertPlainText(text)
        self.text_edit.ensureCursorVisible()

    def clear_log(self):
        self.text_edit.clear()


class ConnectionDialog(QDialog):
    """Separate connection dialog so the main window stays compact."""

    def __init__(self, parent):
        super().__init__(parent)
        self.parent_window = parent
        self.setWindowTitle("Вход / подключение")
        self.resize(720, 420)
        self.setMinimumSize(620, 360)

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(10)

        title = QLabel("Подключение к Project Point")
        title.setFont(QFont("Segoe UI", 14, QFont.Bold))
        title.setObjectName("pageTitle")
        root.addWidget(title)

        form_group = QGroupBox("Параметры входа")
        form_layout = QGridLayout(form_group)
        form_layout.setHorizontalSpacing(12)
        form_layout.setVerticalSpacing(8)
        form_layout.setColumnStretch(1, 1)

        self.profile_combo = QComboBox()
        for idx in range(parent.profile_combo.count()):
            self.profile_combo.addItem(parent.profile_combo.itemText(idx), parent.profile_combo.itemData(idx))
        self.profile_combo.setCurrentIndex(parent.profile_combo.currentIndex())
        self.profile_combo.currentIndexChanged.connect(self._on_profile_changed)
        form_layout.addWidget(self._label("Профиль"), 0, 0)
        form_layout.addWidget(self.profile_combo, 0, 1)

        self.edits = {}
        rows = [
            ("BASE_URL", "BASE_URL", False),
            ("USERNAME", "USERNAME", False),
            ("PASSWORD", "PASSWORD", True),
            ("CLIENT_ID", "CLIENT_ID", False),
            ("SSO_BASE_URL", "SSO_BASE_URL", False),
            ("REALM", "REALM", False),
            ("BROKER_ALIAS", "BROKER_ALIAS", False),
            ("ADFS_URL", "ADFS_URL", False),
        ]
        for row_idx, (key, label, is_password) in enumerate(rows, start=1):
            edit = QLineEdit(parent.fields[key].text())
            edit.setMinimumHeight(34)
            if is_password:
                edit.setEchoMode(QLineEdit.Password)
            self.edits[key] = edit
            form_layout.addWidget(self._label(label), row_idx, 0)
            form_layout.addWidget(edit, row_idx, 1)

        root.addWidget(form_group)

        hint = QLabel(
            "Для стандартных профилей параметры авторизации заполнены автоматически. "
            "Для Custom можно вручную изменить BASE_URL, CLIENT_ID и ADFS/SSO-настройки."
        )
        hint.setWordWrap(True)
        hint.setObjectName("rowSubtitle")
        root.addWidget(hint)

        buttons = QHBoxLayout()
        buttons.addStretch()
        cancel_btn = QPushButton("Отмена")
        cancel_btn.setObjectName("secondaryBtn")
        cancel_btn.clicked.connect(self.reject)
        buttons.addWidget(cancel_btn)

        save_btn = QPushButton("Сохранить")
        save_btn.setObjectName("secondaryBtn")
        save_btn.clicked.connect(self._save_and_close)
        buttons.addWidget(save_btn)

        connect_btn = QPushButton("Подключиться")
        connect_btn.clicked.connect(self._save_and_connect)
        buttons.addWidget(connect_btn)
        root.addLayout(buttons)

        self._apply_profile_editability()

    def _label(self, text):
        label = QLabel(text)
        label.setObjectName("fieldLabel")
        return label

    def _apply_profile_editability(self):
        profile_key = self.profile_combo.currentData()
        is_custom = profile_key == "Custom"
        for key in ["BASE_URL", "CLIENT_ID", "SSO_BASE_URL", "REALM", "BROKER_ALIAS", "ADFS_URL"]:
            self.edits[key].setReadOnly(not is_custom)

    def _on_profile_changed(self, index):
        profile_key = self.profile_combo.currentData()
        if profile_key not in CONNECTION_PRESETS:
            return

        preset = CONNECTION_PRESETS[profile_key]
        self.edits["BASE_URL"].setText(preset["base_url"])
        self.edits["CLIENT_ID"].setText(preset["client_id"])
        self.edits["SSO_BASE_URL"].setText(preset["sso_base_url"])
        self.edits["REALM"].setText(preset["realm"])
        self.edits["BROKER_ALIAS"].setText(preset["broker_alias"])
        self.edits["ADFS_URL"].setText(preset["adfs_base_url"])
        self._apply_profile_editability()

    def _apply_to_parent(self):
        parent = self.parent_window
        selected_key = self.profile_combo.currentData()
        for idx in range(parent.profile_combo.count()):
            if parent.profile_combo.itemData(idx) == selected_key:
                parent.profile_combo.setCurrentIndex(idx)
                break

        for key, edit in self.edits.items():
            parent.fields[key].setText(edit.text().strip())

        parent._update_connection_summary()
        parent._set_auth_status("Не подключено", "pending")
        parent.auth_config_source.setText("")

    def _save_and_close(self):
        self._apply_to_parent()
        self.accept()

    def _save_and_connect(self):
        self._apply_to_parent()
        self.accept()
        self.parent_window._test_auth()
