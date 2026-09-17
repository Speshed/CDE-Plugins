# -*- coding: utf-8 -*-
"""Минимальное подключение к Project Point через встроенный Chromium.

Зависимости:
    pip install "PySide6>=6.5" requests

Сценарий:
    1) пользователь вводит адрес Project Point;
    2) программа открывает сайт во встроенном QWebEngineView;
    3) сам сайт проходит свой обычный SSO / MFA / redirect flow;
    4) программа перехватывает Bearer access_token в API-запросе Project Point;
    5) токен проверяется реальным GetMyProfile;
    6) дополнительно загружается список доступных проектов;
    7) requests.Session остаётся готовой для дальнейшей работы с API;
    8) web-сессия хранится только в памяти и сбрасывается после закрытия приложения.
"""

from __future__ import annotations

import hashlib
import logging
from pathlib import Path
import sys
import threading
from datetime import datetime
from urllib.parse import parse_qsl, urlsplit, urlunsplit

import requests
from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtWebEngineCore import (
    QWebEnginePage,
    QWebEngineProfile,
    QWebEngineScript,
    QWebEngineUrlRequestInterceptor,
)
from PySide6.QtWebEngineWidgets import QWebEngineView


APP_TITLE = "Project Point"
REQUEST_TIMEOUT = 30

# Внутренняя утилита: намеренно принимаем любые серверные TLS-сертификаты.
TRUST_ALL_CERTIFICATES = True

requests.packages.urllib3.disable_warnings(  # type: ignore[attr-defined]
    requests.packages.urllib3.exceptions.InsecureRequestWarning  # type: ignore[attr-defined]
)


class InternalProjectPointSession(requests.Session):
    """HTTP-сессия для внутренних Project Point.

    verify=False передаётся именно в каждый request, а не только хранится
    в Session.verify. Это не даёт REQUESTS_CA_BUNDLE/CURL_CA_BUNDLE из
    окружения Windows снова включить проверку TLS.
    """

    def request(self, method, url, **kwargs):
        if TRUST_ALL_CERTIFICATES:
            kwargs["verify"] = False
        return super().request(method, url, **kwargs)


def new_api_session() -> requests.Session:
    return InternalProjectPointSession()


def _log_directory() -> Path:
    """Логи сохраняются прямо рядом с файлом ProjectPoint.pyw."""
    return Path(__file__).resolve().parent


LOG_DIR = _log_directory()
LOG_FILE = LOG_DIR / f"ProjectPoint_{datetime.now():%Y%m%d_%H%M%S}.log"


def _build_logger() -> logging.Logger:
    logger = logging.getLogger("ProjectPoint")
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    logger.handlers.clear()

    handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    handler.setLevel(logging.DEBUG)
    handler.setFormatter(
        logging.Formatter(
            "%(asctime)s.%(msecs)03d | %(levelname)-7s | %(threadName)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )
    logger.addHandler(handler)
    return logger


LOGGER = _build_logger()


def safe_url_for_log(url: str) -> str:
    """URL без OAuth code/token/state и других значений query/fragment."""
    try:
        parsed = urlsplit(url)
        query_keys = [key for key, _ in parse_qsl(parsed.query, keep_blank_values=True)]
        query = ""
        if query_keys:
            query = "?" + "&".join(f"{key}=…" for key in dict.fromkeys(query_keys))
        fragment = "#…" if parsed.fragment else ""
        return urlunsplit((parsed.scheme, parsed.netloc, parsed.path, query, fragment))
    except Exception:
        return "<некорректный URL>"


def token_description(token: str) -> str:
    """Безопасный идентификатор кандидата без записи самого токена."""
    digest = hashlib.sha256(token.encode("utf-8", errors="ignore")).hexdigest()[:10]
    return f"len={len(token)}, sha256[:10]={digest}"


def normalize_url(value: str) -> str:
    value = value.strip()
    if not value:
        raise ValueError("Укажите адрес Project Point.")
    if "://" not in value:
        value = "https://" + value
    parsed = urlsplit(value)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("Некорректный адрес сайта.")
    return urlunsplit((parsed.scheme, parsed.netloc, parsed.path or "/", "", ""))


def api_base_from_request(request_url: str) -> str:
    """Возвращает всё до /api + /api из реального запроса браузера."""
    parsed = urlsplit(request_url)
    path = parsed.path
    index = path.lower().find("/api/")
    if index < 0:
        return ""
    api_path = path[:index].rstrip("/") + "/api"
    return urlunsplit((parsed.scheme, parsed.netloc, api_path, "", ""))


def api_base_candidates(project_url: str, request_url: str) -> list[str]:
    """Строит короткий список вероятных API-баз Project Point.

    Приоритет у реального /api-запроса. Если Bearer встретился раньше, чем
    браузер сделал такой запрос, пробуем локализованные /en/api, /ru/api и /api.
    """
    result: list[str] = []

    def add(value: str) -> None:
        value = value.rstrip("/")
        if value and value not in result:
            result.append(value)

    exact = api_base_from_request(request_url)
    if exact:
        add(exact)

    parsed_project = urlsplit(project_url)
    parsed_request = urlsplit(request_url)

    roots: list[tuple[str, str, str]] = []
    for parsed in (parsed_project, parsed_request):
        if parsed.scheme in {"http", "https"} and parsed.netloc:
            root = (parsed.scheme, parsed.netloc, "")
            if root not in roots:
                roots.append(root)

    locales: list[str] = []
    for parsed in (parsed_project, parsed_request):
        first = next((part for part in parsed.path.split("/") if part), "").lower()
        if first in {"ru", "en"} and first not in locales:
            locales.append(first)
    for fallback in ("ru", "en"):
        if fallback not in locales:
            locales.append(fallback)

    for scheme, netloc, _ in roots:
        for locale in locales:
            add(urlunsplit((scheme, netloc, f"/{locale}/api", "", "")))
        add(urlunsplit((scheme, netloc, "/api", "", "")))

    return result


def unpack_list(payload):
    """Нормализует типовые ответы Project Point в список."""
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        for key in ("Items", "items", "value", "Value"):
            value = payload.get(key)
            if isinstance(value, list):
                return value
    return None


def display_user(profile) -> str:
    if not isinstance(profile, dict):
        return ""
    for key in (
        "FullName",
        "DisplayName",
        "Name",
        "UserName",
        "Username",
        "Login",
        "Email",
        "Title",
    ):
        value = profile.get(key)
        if value:
            return str(value).strip()
    return ""


def display_project(project, index: int) -> str:
    if not isinstance(project, dict):
        return str(project)

    title = ""
    for key in ("Title", "FullTitle", "Name", "ProjectName"):
        value = project.get(key)
        if value:
            title = str(value).strip()
            break

    code = ""
    for key in ("Code", "ProjectCode", "ShortName"):
        value = project.get(key)
        if value:
            code = str(value).strip()
            break

    if code and title and code.casefold() not in title.casefold():
        return f"{code} — {title}"
    if title:
        return title
    if code:
        return code

    project_id = project.get("Id") or project.get("ID") or project.get("id")
    return f"Проект {project_id or index}"


TOKEN_CAPTURE_SCRIPT = r"""
(() => {
    if (window.__ppTokenCaptureInstalled) return;
    window.__ppTokenCaptureInstalled = true;
    window.__ppCapturedAccessToken = "";

    const saveToken = (value) => {
        try {
            if (value == null) return;
            let text = String(value).trim();
            if (/^Bearer\s+/i.test(text)) text = text.replace(/^Bearer\s+/i, "").trim();
            if (text.length >= 20) window.__ppCapturedAccessToken = text;
        } catch (_) {}
    };

    const inspectHeaders = (headers) => {
        try {
            if (!headers) return;
            if (typeof Headers !== "undefined" && headers instanceof Headers) {
                saveToken(headers.get("authorization"));
                return;
            }
            if (Array.isArray(headers)) {
                for (const pair of headers) {
                    if (Array.isArray(pair) && String(pair[0]).toLowerCase() === "authorization") {
                        saveToken(pair[1]);
                    }
                }
                return;
            }
            if (typeof headers === "object") {
                for (const [key, value] of Object.entries(headers)) {
                    if (String(key).toLowerCase() === "authorization") saveToken(value);
                }
            }
        } catch (_) {}
    };

    try {
        const originalFetch = window.fetch;
        if (typeof originalFetch === "function") {
            window.fetch = function(input, init) {
                try {
                    if (init) inspectHeaders(init.headers);
                    if (typeof Request !== "undefined" && input instanceof Request) {
                        inspectHeaders(input.headers);
                    }
                } catch (_) {}
                return originalFetch.apply(this, arguments);
            };
        }
    } catch (_) {}

    try {
        const originalSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
        XMLHttpRequest.prototype.setRequestHeader = function(name, value) {
            try {
                if (String(name).toLowerCase() === "authorization") saveToken(value);
            } catch (_) {}
            return originalSetRequestHeader.apply(this, arguments);
        };
    } catch (_) {}
})();
"""


TOKEN_READ_SCRIPT = r"""
(() => {
    const direct = window.__ppCapturedAccessToken;
    if (direct) return direct;

    const preferredKeys = new Set([
        "access_token", "accesstoken", "access-token", "bearer", "bearertoken"
    ]);

    const normalizeKey = (value) => String(value || "").toLowerCase().replace(/[^a-z0-9_-]/g, "");
    const looksUseful = (value) => {
        if (typeof value !== "string") return false;
        const text = value.trim().replace(/^Bearer\s+/i, "");
        return text.length >= 40;
    };

    const inspectObject = (value, depth = 0) => {
        if (!value || depth > 4) return "";
        if (typeof value === "string") {
            try { return inspectObject(JSON.parse(value), depth + 1); } catch (_) { return ""; }
        }
        if (typeof value !== "object") return "";

        for (const [key, item] of Object.entries(value)) {
            const normalized = normalizeKey(key);
            if (preferredKeys.has(normalized) && looksUseful(item)) {
                return String(item).trim().replace(/^Bearer\s+/i, "");
            }
        }
        for (const item of Object.values(value)) {
            const nested = inspectObject(item, depth + 1);
            if (nested) return nested;
        }
        return "";
    };

    const inspectStorage = (storage) => {
        try {
            for (let i = 0; i < storage.length; i++) {
                const key = storage.key(i);
                const raw = storage.getItem(key);
                const normalized = normalizeKey(key);
                if (preferredKeys.has(normalized) && looksUseful(raw)) {
                    return String(raw).trim().replace(/^Bearer\s+/i, "");
                }
                const nested = inspectObject(raw);
                if (nested) return nested;
            }
        } catch (_) {}
        return "";
    };

    return inspectStorage(localStorage) || inspectStorage(sessionStorage) || "";
})();
"""


def install_token_capture(profile: QWebEngineProfile) -> None:
    """Перехватывает Bearer на JS-уровне до запуска frontend Project Point."""
    LOGGER.debug("Устанавливаю JS-перехватчик Bearer на DocumentCreation/MainWorld")
    script = QWebEngineScript()
    script.setName("ProjectPointTokenCapture")
    script.setInjectionPoint(QWebEngineScript.InjectionPoint.DocumentCreation)
    script.setWorldId(QWebEngineScript.ScriptWorldId.MainWorld)
    script.setRunsOnSubFrames(True)
    script.setSourceCode(TOKEN_CAPTURE_SCRIPT)
    profile.scripts().insert(script)


class TokenInterceptor(QWebEngineUrlRequestInterceptor):
    tokenFound = QtCore.Signal(str, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.project_host = ""
        self.last_token = ""

    def watch(self, project_url: str) -> None:
        self.project_host = (urlsplit(project_url).hostname or "").casefold()
        self.last_token = ""
        LOGGER.info("Network interceptor: наблюдаю host=%s", self.project_host or "<пусто>")

    def interceptRequest(self, info) -> None:  # noqa: N802 - Qt API
        request_url = info.requestUrl().toString()
        parsed = urlsplit(request_url)

        # Bearer может появиться ещё до запроса с /api/. Ограничиваем перехват
        # самим Project Point или запросами, инициированными его страницей, чтобы
        # не забирать токен со страницы внешнего SSO-провайдера.
        request_host = (parsed.hostname or "").casefold()
        first_party_host = info.firstPartyUrl().host().casefold()
        if request_host != self.project_host and first_party_host != self.project_host:
            return

        try:
            headers = info.httpHeaders()  # Qt 6.5+
        except AttributeError:
            return

        for name, value in headers.items():
            if bytes(name).lower() != b"authorization":
                continue
            authorization = bytes(value).decode("utf-8", errors="ignore").strip()
            if not authorization.lower().startswith("bearer "):
                continue
            token = authorization[7:].strip()
            if token and token != self.last_token:
                self.last_token = token
                LOGGER.info(
                    "Network interceptor: найден Bearer (%s), request=%s",
                    token_description(token),
                    safe_url_for_log(request_url),
                )
                self.tokenFound.emit(token, request_url)
            return


class BrowserView(QWebEngineView):
    """Открывает target=_blank / popup в том же встроенном окне."""

    def createWindow(self, _window_type):  # noqa: N802 - Qt API
        return self


class LoginDialog(QtWidgets.QDialog):
    tokenFound = QtCore.Signal(str, str)

    def __init__(self, profile: QWebEngineProfile, url: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Вход в Project Point")
        self.resize(1080, 760)
        self.setModal(False)
        self.project_host = (urlsplit(url).hostname or "").casefold()
        self._last_js_token = ""
        LOGGER.info("Создаю окно WebView: %s", safe_url_for_log(url))

        self.view = BrowserView(self)
        page = QWebEnginePage(profile, self.view)
        page.certificateError.connect(self._accept_certificate_error)
        self.view.setPage(page)

        self.status = QtWidgets.QLabel("Открываю страницу входа…")
        self.status.setObjectName("browserStatus")

        back = QtWidgets.QPushButton("←")
        back.setFixedWidth(40)
        back.setToolTip("Назад")
        back.clicked.connect(self.view.back)

        reload_button = QtWidgets.QPushButton("Обновить")
        reload_button.clicked.connect(self.view.reload)

        close_button = QtWidgets.QPushButton("Закрыть")
        close_button.clicked.connect(self.reject)

        bar = QtWidgets.QHBoxLayout()
        bar.addWidget(back)
        bar.addWidget(self.status, 1)
        bar.addWidget(reload_button)
        bar.addWidget(close_button)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.addLayout(bar)
        layout.addWidget(self.view, 1)

        self.view.urlChanged.connect(self._url_changed)
        self.view.loadStarted.connect(self._load_started)
        self.view.loadFinished.connect(self._load_finished)
        self.finished.connect(lambda code: LOGGER.info("LoginDialog finished: code=%s", code))
        self.destroyed.connect(lambda *_: LOGGER.info("LoginDialog destroyed"))

        self.token_timer = QtCore.QTimer(self)
        self.token_timer.setInterval(600)
        self.token_timer.timeout.connect(self._poll_token)
        self.token_timer.start()

        self.view.setUrl(QtCore.QUrl(url))

    def _accept_certificate_error(self, error) -> None:
        """Для внутренних адресов принимаем любые ошибки TLS-сертификата."""
        LOGGER.warning("WebView: ошибка TLS-сертификата автоматически принята")
        error.acceptCertificate()

    def _url_changed(self, qurl: QtCore.QUrl) -> None:
        self.status.setText(qurl.host() or "Загрузка…")
        LOGGER.info("WebView URL -> %s", safe_url_for_log(qurl.toString()))

    def _load_started(self) -> None:
        LOGGER.debug("WebView loadStarted: %s", safe_url_for_log(self.view.url().toString()))

    def _load_finished(self, ok: bool) -> None:
        LOGGER.info(
            "WebView loadFinished: ok=%s, url=%s",
            ok,
            safe_url_for_log(self.view.url().toString()),
        )
        if not ok:
            self.status.setText("Не удалось загрузить страницу")

    def _poll_token(self) -> None:
        current_url = self.view.url().toString()
        current_host = (urlsplit(current_url).hostname or "").casefold()
        if current_host != self.project_host:
            return
        self.view.page().runJavaScript(TOKEN_READ_SCRIPT, self._token_from_javascript)

    def _token_from_javascript(self, value) -> None:
        if not isinstance(value, str):
            return
        token = value.strip()
        if not token or token == self._last_js_token:
            return
        self._last_js_token = token
        LOGGER.info(
            "JavaScript capture: найден token-кандидат (%s), page=%s",
            token_description(token),
            safe_url_for_log(self.view.url().toString()),
        )
        self.tokenFound.emit(token, self.view.url().toString())


class VerifySignals(QtCore.QObject):
    finished = QtCore.Signal(object)
    failed = QtCore.Signal(str)


class ProjectPointWindow(QtWidgets.QWidget):
    """Небольшое окно подключения. После проверки self.session готова к API."""

    def __init__(self):
        super().__init__()
        LOGGER.info("Главное окно создаётся")
        self.setWindowTitle(APP_TITLE)
        self.setFixedWidth(620)

        self.access_token = ""
        self.api_base_url = ""
        self.session = new_api_session()
        self.login_dialog: LoginDialog | None = None
        self.verifying = False
        self.current_project_url = ""
        self.attempted_tokens: set[str] = set()

        # Временный off-the-record профиль: cookies, cache и SSO-сессия
        # существуют только до закрытия приложения и не сохраняются на диск.
        self.profile = QWebEngineProfile(self)
        install_token_capture(self.profile)
        self.interceptor = TokenInterceptor(self.profile)
        self.profile.setUrlRequestInterceptor(self.interceptor)
        self.interceptor.tokenFound.connect(self._token_found)

        self.verify_signals = VerifySignals(self)
        self.verify_signals.finished.connect(self._verification_finished)
        self.verify_signals.failed.connect(self._verification_failed)

        self.url_edit = QtWidgets.QLineEdit()
        self.url_edit.setPlaceholderText("https://projectpoint.company.ru")
        self.url_edit.returnPressed.connect(self.connect_to_projectpoint)

        self.connect_button = QtWidgets.QPushButton("Подключиться")
        self.connect_button.setObjectName("primaryButton")
        self.connect_button.clicked.connect(self.connect_to_projectpoint)

        self.log_button = QtWidgets.QPushButton("Лог")
        self.log_button.setToolTip(str(LOG_FILE))
        self.log_button.clicked.connect(self.open_log)

        self.status = QtWidgets.QLabel("Не подключено")
        self.status.setObjectName("statusLabel")

        self.details = QtWidgets.QLabel(
            "Откроется встроенное окно входа. Логин, пароль и MFA вводятся прямо на странице Project Point."
        )
        self.details.setWordWrap(True)
        self.details.setObjectName("hintLabel")

        self.projects_title = QtWidgets.QLabel("")
        self.projects_title.setObjectName("projectsTitle")
        self.projects_title.hide()

        self.projects_list = QtWidgets.QListWidget()
        self.projects_list.setMinimumHeight(120)
        self.projects_list.setMaximumHeight(220)
        self.projects_list.hide()

        title = QtWidgets.QLabel("Подключение к Project Point")
        title.setObjectName("titleLabel")

        form = QtWidgets.QHBoxLayout()
        form.addWidget(self.url_edit, 1)
        form.addWidget(self.connect_button)
        form.addWidget(self.log_button)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(12)
        layout.addWidget(title)
        layout.addLayout(form)
        layout.addWidget(self.status)
        layout.addWidget(self.details)
        layout.addWidget(self.projects_title)
        layout.addWidget(self.projects_list)

        self.setStyleSheet(
            """
            QWidget { font-family: 'Segoe UI'; font-size: 13px; background: #f6f7f9; color: #202124; }
            QLineEdit { background: white; border: 1px solid #d5d8de; border-radius: 7px; padding: 9px 10px; }
            QPushButton { border: 1px solid #d5d8de; border-radius: 7px; padding: 8px 13px; background: white; }
            QPushButton:hover { background: #eef0f3; }
            QPushButton#primaryButton { background: #202124; color: white; border: none; padding: 9px 16px; }
            QPushButton#primaryButton:hover { background: #34363a; }
            QLabel#titleLabel { font-size: 20px; font-weight: 600; }
            QLabel#statusLabel, QLabel#projectsTitle { font-weight: 600; }
            QLabel#hintLabel, QLabel#browserStatus { color: #676b73; }
            QListWidget { background: white; border: 1px solid #d5d8de; border-radius: 7px; padding: 4px; }
            QListWidget::item { padding: 5px 7px; }
            """
        )

    def open_log(self) -> None:
        LOGGER.debug("Пользователь открыл файл лога")
        if not QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(str(LOG_FILE))):
            QtWidgets.QMessageBox.information(
                self,
                APP_TITLE,
                f"Лог сохранён здесь:\n{LOG_FILE}",
            )

    def _clear_projects(self) -> None:
        self.projects_list.clear()
        self.projects_list.hide()
        self.projects_title.clear()
        self.projects_title.hide()

    def connect_to_projectpoint(self) -> None:
        LOGGER.info("Нажата кнопка «Подключиться»")
        try:
            project_url = normalize_url(self.url_edit.text())
        except ValueError as exc:
            LOGGER.warning("Некорректный URL: %s", exc)
            QtWidgets.QMessageBox.warning(self, APP_TITLE, str(exc))
            return

        LOGGER.info("Старт подключения: %s", safe_url_for_log(project_url))
        self.url_edit.setText(project_url)
        self.current_project_url = project_url
        self.access_token = ""
        self.api_base_url = ""
        self.session = new_api_session()
        self.verifying = False
        self.attempted_tokens.clear()
        self.interceptor.watch(project_url)
        self._clear_projects()

        self.status.setText("Ожидание входа…")
        self.details.setText(
            "Завершите вход в открывшемся окне. После получения токена программа сама проверит API Project Point."
        )

        if self.login_dialog is not None:
            LOGGER.info("Закрываю предыдущее окно входа перед новым подключением")
            self.login_dialog.close()
        self.login_dialog = LoginDialog(self.profile, project_url, self)
        self.login_dialog.tokenFound.connect(self._token_found)
        self.login_dialog.rejected.connect(self._login_closed)
        self.login_dialog.show()
        self.login_dialog.raise_()
        self.login_dialog.activateWindow()
        LOGGER.info("Окно входа показано: visible=%s", self.login_dialog.isVisible())

    @QtCore.Slot(str, str)
    def _token_found(self, token: str, request_url: str) -> None:
        token = token.strip()
        if not token:
            LOGGER.debug("Получен пустой token-кандидат — игнорирую")
            return
        if self.verifying:
            LOGGER.debug("Token-кандидат (%s) пришёл во время проверки — игнорирую", token_description(token))
            return
        if self.access_token:
            LOGGER.debug("Token-кандидат пришёл после успешного подключения — игнорирую")
            return
        if token in self.attempted_tokens:
            LOGGER.debug("Повторный token-кандидат (%s) — игнорирую", token_description(token))
            return
        self.attempted_tokens.add(token)

        LOGGER.info(
            "Новый token-кандидат: %s, source_url=%s",
            token_description(token),
            safe_url_for_log(request_url),
        )
        candidates = api_base_candidates(self.current_project_url, request_url)
        LOGGER.info("API-кандидаты: %s", " | ".join(candidates) if candidates else "<нет>")
        if not candidates:
            LOGGER.warning("Не удалось построить ни одного API-кандидата")
            return

        self.verifying = True
        self.status.setText("Токен получен — проверяю API…")
        self.details.setText(
            "Bearer найден. Подбираю API Project Point и подтверждаю токен через GetMyProfile…"
        )

        user_agent = self.profile.httpUserAgent()
        threading.Thread(
            target=self._verify_worker,
            args=(token, candidates, user_agent),
            daemon=True,
        ).start()

    def _verify_worker(self, token: str, api_candidates: list[str], user_agent: str) -> None:
        LOGGER.info("Verify worker: старт (%s), кандидатов API=%d", token_description(token), len(api_candidates))
        session = new_api_session()
        session.headers.update(
            {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
                "User-Agent": user_agent,
            }
        )

        try:
            api_base_url = ""
            profile_data = None
            failures: list[str] = []

            for candidate in api_candidates:
                profile_url = candidate.rstrip("/") + "/Core/UsersService/GetMyProfile"
                LOGGER.info("GetMyProfile -> %s", safe_url_for_log(profile_url))
                try:
                    response = session.get(profile_url, timeout=REQUEST_TIMEOUT)
                except requests.exceptions.SSLError as exc:
                    LOGGER.warning("GetMyProfile SSL error на %s: %s", candidate, exc)
                    failures.append(f"{candidate}: SSL error: {exc}")
                    continue
                except requests.RequestException as exc:
                    LOGGER.warning("GetMyProfile network error на %s: %s", candidate, exc)
                    failures.append(f"{candidate}: {exc}")
                    continue

                LOGGER.info("GetMyProfile <- HTTP %s для %s", response.status_code, candidate)
                if response.status_code != 200:
                    failures.append(f"{candidate}: HTTP {response.status_code}")
                    continue
                try:
                    candidate_profile = response.json()
                except ValueError:
                    LOGGER.warning("GetMyProfile на %s вернул HTTP 200, но не JSON", candidate)
                    failures.append(f"{candidate}: HTTP 200, но ответ не JSON")
                    continue

                api_base_url = candidate
                profile_data = candidate_profile
                LOGGER.info("GetMyProfile подтвердил токен. API=%s", api_base_url)
                break

            if not api_base_url:
                details = "; ".join(failures[-6:])
                raise RuntimeError(
                    "GetMyProfile не подтвердил токен ни на одном предполагаемом API."
                    + (f" Проверено: {details}" if details else "")
                )

            # Эти варианты взяты из API-модуля присланного проекта.
            project_paths = (
                "/Core/ProjectService/GetAll",
                "/Core/ProjectService/GetAllForSelect",
                "/Core/ProjectServiceOData/",
            )
            projects = None
            projects_error = ""
            for path in project_paths:
                url = api_base_url.rstrip("/") + path
                LOGGER.info("ProjectService -> %s", safe_url_for_log(url))
                try:
                    project_response = session.get(url, timeout=REQUEST_TIMEOUT)
                except requests.RequestException as exc:
                    LOGGER.warning("ProjectService network error %s: %s", path, exc)
                    projects_error = f"{path}: {exc}"
                    continue

                LOGGER.info("ProjectService <- HTTP %s для %s", project_response.status_code, path)
                if project_response.status_code != 200:
                    projects_error = f"{path}: HTTP {project_response.status_code}"
                    continue
                try:
                    payload = project_response.json()
                except ValueError:
                    projects_error = f"{path}: ответ не JSON"
                    continue

                candidate = unpack_list(payload)
                if candidate is not None:
                    projects = candidate
                    projects_error = ""
                    LOGGER.info("Список проектов получен: count=%d через %s", len(projects), path)
                    break
                projects_error = f"{path}: неизвестный формат ответа"

            LOGGER.info("Verify worker: успешная проверка завершена, передаю результат в UI")
            self.verify_signals.finished.emit(
                {
                    "token": token,
                    "api_base_url": api_base_url,
                    "session": session,
                    "profile": profile_data,
                    "projects": projects,
                    "projects_error": projects_error,
                }
            )
        except Exception as exc:
            LOGGER.exception("Verify worker: проверка завершилась ошибкой: %s", exc)
            self.verify_signals.failed.emit(str(exc))

    @QtCore.Slot(object)
    def _verification_finished(self, result) -> None:
        LOGGER.info("UI: получен успешный результат проверки")
        self.verifying = False
        self.access_token = result["token"]
        self.api_base_url = result["api_base_url"]
        self.session = result["session"]

        user = display_user(result.get("profile"))
        projects = result.get("projects")
        projects_error = result.get("projects_error") or ""

        self.status.setText("Подключено — API подтвердил авторизацию")

        lines = [
            f"API: {self.api_base_url}",
            "GetMyProfile: HTTP 200 — токен принят сервером.",
        ]
        if user:
            lines.append(f"Пользователь: {user}")

        if projects is not None:
            lines.append(f"Доступно проектов: {len(projects)}")
            self.projects_title.setText(f"Доступные проекты ({len(projects)})")
            self.projects_title.show()
            self.projects_list.clear()
            for index, project in enumerate(projects, 1):
                self.projects_list.addItem(display_project(project, index))
            if projects:
                self.projects_list.show()
            else:
                self.projects_list.addItem("Список пуст")
                self.projects_list.show()
        else:
            lines.append("Список проектов получить не удалось, но GetMyProfile успешно подтвердил подключение.")
            if projects_error:
                lines.append(f"ProjectService: {projects_error}")

        self.details.setText("\n".join(lines))

        # QWebEngineView иногда продолжает держать окно QDialog видимым даже
        # после accept(). Закрываем окно входа принудительно после того, как
        # GetMyProfile реально подтвердил токен.
        dialog = self.login_dialog
        self.login_dialog = None
        if dialog is not None:
            LOGGER.info(
                "Закрываю LoginDialog: visible_before=%s, active=%s",
                dialog.isVisible(),
                dialog.isActiveWindow(),
            )
            dialog.token_timer.stop()
            LOGGER.debug("LoginDialog: token_timer.stop()")
            dialog.view.stop()
            LOGGER.debug("LoginDialog: WebView.stop()")
            dialog.hide()
            LOGGER.info("LoginDialog: hide(), visible=%s", dialog.isVisible())
            close_result = dialog.close()
            LOGGER.info("LoginDialog: close()=%s, visible=%s", close_result, dialog.isVisible())
            dialog.deleteLater()
            LOGGER.info("LoginDialog: deleteLater() поставлен в очередь")
        else:
            LOGGER.warning("Успешная авторизация подтверждена, но self.login_dialog уже None")

        self.show()
        self.raise_()
        self.activateWindow()
        LOGGER.info("Главное окно активировано после подключения")

    @QtCore.Slot(str)
    def _verification_failed(self, message: str) -> None:
        LOGGER.warning("UI: token-кандидат не подтверждён: %s", message)
        self.verifying = False
        self.access_token = ""
        self.api_base_url = ""
        self._clear_projects()
        self.status.setText("Найден токен-кандидат — продолжаю ждать подтверждение")
        self.details.setText(
            "Один из найденных токенов GetMyProfile не принял. Окно входа остаётся открытым; "
            "программа продолжает искать рабочий access token.\n"
            f"Последняя проверка: {message}"
        )

    def _login_closed(self) -> None:
        LOGGER.info(
            "Сигнал LoginDialog.rejected: access_token=%s, verifying=%s",
            bool(self.access_token),
            self.verifying,
        )
        if not self.access_token and not self.verifying:
            self.status.setText("Вход не завершён")
            self.details.setText("Нажмите «Подключиться», чтобы открыть окно входа ещё раз.")
        self.login_dialog = None

    def api_url(self, path: str) -> str:
        """Удобный helper для следующего кода приложения."""
        if not self.api_base_url or not self.access_token:
            raise RuntimeError("Сначала подключитесь к Project Point.")
        return self.api_base_url.rstrip("/") + "/" + path.lstrip("/")

    def api_get(self, path: str, **kwargs):
        return self.session.get(self.api_url(path), timeout=REQUEST_TIMEOUT, **kwargs)


def main() -> int:
    LOGGER.info("=" * 72)
    LOGGER.info("Запуск Project Point connector")
    LOGGER.info("Python=%s", sys.version.replace("\n", " "))
    LOGGER.info("Qt=%s", QtCore.qVersion())
    LOGGER.warning(
        "TLS certificate verification disabled=%s (forced per request, internal Project Point mode)",
        TRUST_ALL_CERTIFICATES,
    )
    LOGGER.info("Log file=%s", LOG_FILE)

    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.aboutToQuit.connect(lambda: LOGGER.info("QApplication aboutToQuit"))

    window = ProjectPointWindow()
    window.show()
    LOGGER.info("Главное окно показано")
    exit_code = app.exec()
    LOGGER.info("QApplication завершён: exit_code=%s", exit_code)
    logging.shutdown()
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
