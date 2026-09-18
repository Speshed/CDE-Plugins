from __future__ import annotations

import hashlib
import threading
from urllib.parse import urlsplit, urlunsplit

import requests
from PySide6 import QtCore, QtWidgets
from PySide6.QtWebEngineCore import (
    QWebEnginePage,
    QWebEngineProfile,
    QWebEngineScript,
    QWebEngineUrlRequestInterceptor,
)
from PySide6.QtWebEngineWidgets import QWebEngineView

REQUEST_TIMEOUT = 30


def normalize_project_url(value: str) -> str:
    value = (value or "").strip()
    if not value:
        raise ValueError("Укажите адрес Project Point.")
    if "://" not in value:
        value = "https://" + value
    parsed = urlsplit(value)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("Некорректный адрес Project Point.")
    path = parsed.path or "/"
    return urlunsplit((parsed.scheme, parsed.netloc, path, "", ""))


def project_root_url(value: str) -> str:
    normalized = normalize_project_url(value)
    parsed = urlsplit(normalized)
    parts = [part for part in parsed.path.split("/") if part]
    if parts and parts[-1].casefold() in {"ru", "en"}:
        parts.pop()
    path = "/" + "/".join(parts) if parts else ""
    return urlunsplit((parsed.scheme, parsed.netloc, path.rstrip("/"), "", ""))


def api_base_from_request(request_url: str) -> str:
    parsed = urlsplit(request_url)
    index = parsed.path.casefold().find("/api/")
    if index < 0:
        return ""
    api_path = parsed.path[:index].rstrip("/") + "/api"
    return urlunsplit((parsed.scheme, parsed.netloc, api_path, "", ""))


def api_base_candidates(project_url: str, request_url: str = "") -> list[str]:
    result: list[str] = []

    def add(value: str) -> None:
        value = value.rstrip("/")
        if value and value not in result:
            result.append(value)

    if request_url:
        add(api_base_from_request(request_url))

    root = project_root_url(project_url)
    parsed = urlsplit(root)
    root_path = parsed.path.rstrip("/")
    for locale in ("ru", "en"):
        add(urlunsplit((parsed.scheme, parsed.netloc, f"{root_path}/{locale}/api", "", "")))
    add(urlunsplit((parsed.scheme, parsed.netloc, f"{root_path}/api", "", "")))
    return result


def token_description(token: str) -> str:
    digest = hashlib.sha256(token.encode("utf-8", errors="ignore")).hexdigest()[:10]
    return f"len={len(token)}, sha256[:10]={digest}"


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
                    if (Array.isArray(pair) && String(pair[0]).toLowerCase() === "authorization") saveToken(pair[1]);
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
                    if (typeof Request !== "undefined" && input instanceof Request) inspectHeaders(input.headers);
                } catch (_) {}
                return originalFetch.apply(this, arguments);
            };
        }
    } catch (_) {}
    try {
        const originalSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
        XMLHttpRequest.prototype.setRequestHeader = function(name, value) {
            try { if (String(name).toLowerCase() === "authorization") saveToken(value); } catch (_) {}
            return originalSetRequestHeader.apply(this, arguments);
        };
    } catch (_) {}
})();
"""


TOKEN_READ_SCRIPT = r"""
(() => {
    const direct = window.__ppCapturedAccessToken;
    if (direct) return direct;
    const keys = new Set(["access_token", "accesstoken", "access-token", "bearer", "bearertoken"]);
    const norm = (v) => String(v || "").toLowerCase().replace(/[^a-z0-9_-]/g, "");
    const useful = (v) => typeof v === "string" && v.trim().replace(/^Bearer\s+/i, "").length >= 40;
    const inspectObject = (value, depth = 0) => {
        if (!value || depth > 4) return "";
        if (typeof value === "string") {
            try { return inspectObject(JSON.parse(value), depth + 1); } catch (_) { return ""; }
        }
        if (typeof value !== "object") return "";
        for (const [key, item] of Object.entries(value)) {
            if (keys.has(norm(key)) && useful(item)) return String(item).trim().replace(/^Bearer\s+/i, "");
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
                if (keys.has(norm(key)) && useful(raw)) return String(raw).trim().replace(/^Bearer\s+/i, "");
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

    def interceptRequest(self, info) -> None:  # noqa: N802
        request_url = info.requestUrl().toString()
        parsed = urlsplit(request_url)
        request_host = (parsed.hostname or "").casefold()
        first_party_host = info.firstPartyUrl().host().casefold()
        if request_host != self.project_host and first_party_host != self.project_host:
            return
        try:
            headers = info.httpHeaders()
        except AttributeError:
            return
        for name, value in headers.items():
            if bytes(name).lower() != b"authorization":
                continue
            authorization = bytes(value).decode("utf-8", errors="ignore").strip()
            if not authorization.casefold().startswith("bearer "):
                continue
            token = authorization[7:].strip()
            if token and token != self.last_token:
                self.last_token = token
                self.tokenFound.emit(token, request_url)
            return


class BrowserView(QWebEngineView):
    def createWindow(self, _window_type):  # noqa: N802
        return self


class _VerifySignals(QtCore.QObject):
    success = QtCore.Signal(str, str)
    failure = QtCore.Signal(str)


class UniversalBrowserAuthDialog(QtWidgets.QDialog):
    authenticated = QtCore.Signal(str, str)

    def __init__(self, project_url: str, parent=None):
        super().__init__(parent)
        self.project_url = normalize_project_url(project_url)
        self.project_host = (urlsplit(self.project_url).hostname or "").casefold()
        self.setWindowTitle("Вход в Project Point")
        self.resize(1080, 760)
        self.setModal(False)

        self._attempted_tokens: set[str] = set()
        self._verifying = False
        self._last_js_token = ""

        # Off-the-record profile: login cookies/session exist only until the app closes.
        self.profile = QWebEngineProfile(self)
        install_token_capture(self.profile)
        self.interceptor = TokenInterceptor(self.profile)
        self.profile.setUrlRequestInterceptor(self.interceptor)
        self.interceptor.watch(self.project_url)
        self.interceptor.tokenFound.connect(self._token_found)

        self.verify_signals = _VerifySignals(self)
        self.verify_signals.success.connect(self._verification_success)
        self.verify_signals.failure.connect(self._verification_failure)

        self.view = BrowserView(self)
        page = QWebEnginePage(self.profile, self.view)
        page.certificateError.connect(self._accept_certificate_error)
        self.view.setPage(page)

        self.status = QtWidgets.QLabel(
            "Войдите на странице Project Point. MFA, CAPTCHA, passkey и другие проверки проходятся здесь штатно."
        )
        self.status.setWordWrap(True)

        back_btn = QtWidgets.QPushButton("←")
        back_btn.setFixedWidth(40)
        back_btn.clicked.connect(self.view.back)
        reload_btn = QtWidgets.QPushButton("Обновить")
        reload_btn.clicked.connect(self.view.reload)
        close_btn = QtWidgets.QPushButton("Закрыть")
        close_btn.clicked.connect(self.reject)

        bar = QtWidgets.QHBoxLayout()
        bar.addWidget(back_btn)
        bar.addWidget(self.status, 1)
        bar.addWidget(reload_btn)
        bar.addWidget(close_btn)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.addLayout(bar)
        layout.addWidget(self.view, 1)

        self.view.urlChanged.connect(self._url_changed)
        self.token_timer = QtCore.QTimer(self)
        self.token_timer.setInterval(600)
        self.token_timer.timeout.connect(self._poll_token)
        self.token_timer.start()
        self.finished.connect(lambda _code: self.token_timer.stop())

        self.view.setUrl(QtCore.QUrl(self.project_url))

    def _accept_certificate_error(self, error) -> None:
        error.acceptCertificate()

    def _url_changed(self, qurl: QtCore.QUrl) -> None:
        host = qurl.host() or "Загрузка…"
        if not self._verifying:
            self.status.setText(f"Страница входа: {host}")

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
        self._token_found(token, self.view.url().toString())

    @QtCore.Slot(str, str)
    def _token_found(self, token: str, request_url: str) -> None:
        if not token or token in self._attempted_tokens or self._verifying:
            return
        self._attempted_tokens.add(token)
        candidates = api_base_candidates(self.project_url, request_url)
        if not candidates:
            return
        self._verifying = True
        self.status.setText("Токен найден. Проверяю подключение через GetMyProfile…")
        threading.Thread(
            target=self._verify_worker,
            args=(token, candidates),
            daemon=True,
        ).start()

    def _verify_worker(self, token: str, candidates: list[str]) -> None:
        failures: list[str] = []
        session = requests.Session()
        session.verify = False
        try:
            requests.packages.urllib3.disable_warnings()  # type: ignore[attr-defined]
        except Exception:
            pass
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "User-Agent": self.profile.httpUserAgent(),
        }
        for candidate in candidates:
            url = candidate.rstrip("/") + "/Core/UsersService/GetMyProfile"
            try:
                response = session.get(url, headers=headers, timeout=REQUEST_TIMEOUT, verify=False)
            except requests.RequestException as exc:
                failures.append(f"{candidate}: {exc}")
                continue
            if response.status_code == 200:
                try:
                    response.json()
                except ValueError:
                    failures.append(f"{candidate}: HTTP 200, но ответ не JSON")
                    continue
                self.verify_signals.success.emit(token, candidate)
                return
            failures.append(f"{candidate}: HTTP {response.status_code}")
        self.verify_signals.failure.emit("; ".join(failures[-5:]) or "GetMyProfile не подтвердил токен")

    @QtCore.Slot(str, str)
    def _verification_success(self, token: str, api_base: str) -> None:
        self._verifying = False
        self.status.setText(f"Подключено. API: {api_base}")
        self.authenticated.emit(token, api_base)
        self.token_timer.stop()
        self.view.stop()
        self.accept()

    @QtCore.Slot(str)
    def _verification_failure(self, message: str) -> None:
        self._verifying = False
        self.status.setText(
            "Найденный токен не принят API. Продолжаю ждать рабочую сессию. " + message
        )
