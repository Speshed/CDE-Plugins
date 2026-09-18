from __future__ import annotations

from PySide6.QtCore import Signal, QObject
from PySide6.QtGui import QFont, QTextCursor
from PySide6.QtWidgets import (
    QDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
)

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
    """Universal Project Point connection dialog.

    There are no environment presets anymore.  The user supplies only the
    Project Point address and, optionally, login/password for the automatic
    HTTP attempt.  If the site requires MFA/CAPTCHA/passkey/another SSO step,
    the main window opens the real login page in the embedded browser.
    """

    def __init__(self, parent):
        super().__init__(parent)
        self.parent_window = parent
        self.setWindowTitle("Вход / подключение")
        self.resize(680, 320)
        self.setMinimumSize(580, 300)

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(10)

        title = QLabel("Подключение к Project Point")
        title.setFont(QFont("Segoe UI", 14, QFont.Bold))
        title.setObjectName("pageTitle")
        root.addWidget(title)

        form_group = QGroupBox("Авторизация")
        form_layout = QGridLayout(form_group)
        form_layout.setHorizontalSpacing(12)
        form_layout.setVerticalSpacing(8)
        form_layout.setColumnStretch(1, 1)

        self.edits = {}
        rows = [
            ("BASE_URL", "Адрес Project Point", False),
            ("USERNAME", "Логин", False),
            ("PASSWORD", "Пароль", True),
        ]
        for row_idx, (key, label, is_password) in enumerate(rows):
            edit = QLineEdit(parent.fields[key].text())
            edit.setMinimumHeight(34)
            if key == "BASE_URL":
                edit.setPlaceholderText("https://projectpoint.company.ru")
            if is_password:
                edit.setEchoMode(QLineEdit.Password)
            self.edits[key] = edit
            form_layout.addWidget(self._label(label), row_idx, 0)
            form_layout.addWidget(edit, row_idx, 1)

        root.addWidget(form_group)

        hint = QLabel(
            "Программа сама определяет ADFS / Keycloak / OIDC из текущего сайта. "
            "Если обычный вход по логину и паролю не подходит или включены MFA, CAPTCHA, "
            "passkey, подтверждение на телефоне либо другой SSO-сценарий, автоматически "
            "откроется встроенный браузер и защита проходится штатно. "
            "Логин и пароль можно оставить пустыми, чтобы сразу открыть страницу входа."
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

    def _label(self, text):
        label = QLabel(text)
        label.setObjectName("fieldLabel")
        return label

    def _apply_to_parent(self):
        parent = self.parent_window
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
