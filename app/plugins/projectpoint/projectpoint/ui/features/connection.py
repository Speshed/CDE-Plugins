from __future__ import annotations

import threading

from ...auth import (
    clear_cached_access_token,
    get_cached_access_token,
    normalize_project_base_url,
    set_cached_access_token,
    session,
)
from ...browser_auth import UniversalBrowserAuthDialog, normalize_project_url, project_root_url
from ...universal_http_auth import authenticate_projectpoint_http
from ..dialogs import ConnectionDialog


class ConnectionMixin:
    def _update_connection_summary(self):
        base_widget = self.fields.get("BASE_URL") if hasattr(self, "fields") else None
        user_widget = self.fields.get("USERNAME") if hasattr(self, "fields") else None
        base_url = base_widget.text().strip() if base_widget else ""
        username = user_widget.text().strip() if user_widget else ""
        username_text = username if username else "ввод на странице входа"
        self.connection_summary_label.setText(
            f"Адрес: {base_url or 'не указан'}   |   Пользователь: {username_text}"
        )

    def _show_connection_dialog(self):
        dialog = ConnectionDialog(self)
        dialog.exec()

    def _val(self, key):
        widget = self.fields.get(key)
        return widget.text().strip() if widget is not None else ""

    def _normalized_base_url(self):
        raw = self._val("BASE_URL")
        if not raw:
            return ""
        try:
            return project_root_url(normalize_project_url(raw))
        except Exception:
            return normalize_project_base_url(raw)

    def _get_worker_auth_args(self):
        """Return the legacy tuple shape expected by feature workers.

        Authentication itself is no longer profile-based.  The GUI performs one
        universal login, verifies the token, and stores it in memory.  Workers
        simply reuse that verified session token through ``get_access_token``.
        """
        base_url = self._normalized_base_url()
        if not base_url or not get_cached_access_token(base_url):
            return None
        return (
            "browser_session",
            "",
            base_url,
            "",
            "",
            None,
            None,
            None,
            None,
        )

    def _auth_args_with_discovery(self):
        """Compatibility helper retained for older call sites."""
        base_url = self._normalized_base_url()
        if not base_url:
            return None
        return {
            "base_url": base_url,
            "username": self._val("USERNAME"),
            "password": self._val("PASSWORD"),
            "config": {"auth_mode": "universal"},
        }

    def _auto_connect(self):
        if self._val("BASE_URL"):
            self._test_auth()

    def _test_auth(self):
        raw_url = self._val("BASE_URL")
        if not raw_url:
            self._set_auth_status("Укажите адрес Project Point", "danger")
            return

        try:
            normalized = project_root_url(normalize_project_url(raw_url))
        except Exception as exc:
            self._set_auth_status(str(exc), "danger")
            return

        self.fields["BASE_URL"].setText(normalized)
        clear_cached_access_token(normalized)
        self.auth_btn.setEnabled(False)
        self._set_auth_status("Авторизация…", "pending")
        self.auth_config_source.setText("Определяю способ входа автоматически")

        username = self._val("USERNAME")
        password = self._val("PASSWORD")

        # If credentials are present, first try the HTTP/OIDC discovery supplied
        # by the user.  It is fast and requires no extra window for ordinary
        # Forms Authentication.  Any failure falls back to the real browser.
        if username and password:
            threading.Thread(
                target=self._http_auth_worker,
                args=(normalized, username, password),
                daemon=True,
            ).start()
            return

        self._open_browser_auth(
            "Логин или пароль не заполнены — открываю штатную страницу входа."
        )

    def _http_auth_worker(self, base_url: str, username: str, password: str):
        try:
            result = authenticate_projectpoint_http(
                base_url,
                username,
                password,
                verify=False,
            )
            payload = {"ok": True, "result": result}
        except Exception as exc:
            payload = {"ok": False, "error": str(exc)}
        self.auth_http_finished.emit(payload)

    def _handle_http_auth_result(self, payload):
        if payload.get("ok"):
            result = payload["result"]
            token = result["access_token"]
            base_url = result.get("base_url") or self._normalized_base_url()
            base_url = normalize_project_base_url(base_url)
            self.fields["BASE_URL"].setText(base_url)
            set_cached_access_token(base_url, token)
            self._set_auth_status("Подключено", "success")
            provider = result.get("provider") or "OIDC/OAuth"
            self.auth_config_source.setText(
                f"Автоматически определено: {provider}. Токен проверен через GetMyProfile."
            )
            self.auth_btn.setEnabled(True)
            return

        reason = payload.get("error") or "Автоматический вход не завершён."
        self._open_browser_auth(reason)

    def _open_browser_auth(self, reason: str = ""):
        self.auth_config_source.setText(
            "Встроенный браузер: используется штатная авторизация сайта"
            + (f". Причина перехода: {reason}" if reason else "")
        )
        self._set_auth_status("Ожидаю вход в браузере…", "pending")

        base_url = self._normalized_base_url()
        if not base_url:
            self.auth_btn.setEnabled(True)
            self._set_auth_status("Укажите адрес Project Point", "danger")
            return

        dialog = UniversalBrowserAuthDialog(base_url, self)
        self._browser_auth_dialog = dialog
        dialog.authenticated.connect(self._on_browser_authenticated)
        dialog.finished.connect(self._on_browser_auth_finished)
        dialog.open()

    def _on_browser_authenticated(self, token: str, api_base: str):
        base_url = self._normalized_base_url()
        set_cached_access_token(base_url, token)
        self._set_auth_status("Подключено", "success")
        self.auth_config_source.setText(
            f"Штатная браузерная сессия. API подтверждён: {api_base}"
        )
        self.auth_btn.setEnabled(True)

    def _on_browser_auth_finished(self, _result: int):
        dialog = getattr(self, "_browser_auth_dialog", None)
        self._browser_auth_dialog = None
        self.auth_btn.setEnabled(True)
        if dialog is None:
            return
        if not get_cached_access_token(self._normalized_base_url()):
            self._set_auth_status("Вход не завершён", "pending")

    def _set_auth_status(self, text, tone="neutral"):
        self.auth_status.setText(text)
        self.auth_status.setProperty("tone", tone)
        style = self.auth_status.style()
        style.unpolish(self.auth_status)
        style.polish(self.auth_status)
        self.auth_status.update()
        if hasattr(self, "connection_summary_label"):
            self._update_connection_summary()
