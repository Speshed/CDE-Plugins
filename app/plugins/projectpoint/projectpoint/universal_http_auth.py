#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Project Point: универсальная HTTP-авторизация без браузера.

Поддерживаемые сценарии:
- ADFS OAuth/OIDC с HTML Forms Authentication;
- Keycloak OpenID Connect Authorization Code + PKCE;
- обычный OIDC Authorization Code, если конфигурация обнаруживается во frontend.

Пользователь вводит только:
- адрес экземпляра Project Point;
- логин;
- пароль.

Программа:
1. Находит рабочую локализованную страницу Project Point (/ru/, /en/ и т. п.).
2. Загружает main.*.js и другие скрипты страницы.
3. Собирает все возможные конфигурации ADFS/Keycloak/OIDC.
4. Проверяет кандидатов и выбирает тот, который реально открывает форму входа.
5. Проходит HTML-формы и HTTP-переходы через requests.Session.
6. Получает access_token и проверяет его запросом GetMyProfile.

Ограничения:
- CAPTCHA, WebAuthn/passkey и подтверждение на телефоне автоматически не обходятся;
- MFA/OTP распознаётся, но код необходимо обрабатывать отдельно;
- Windows Integrated Authentication (Negotiate/NTLM) требует отдельной поддержки;
- конфигурация, сформированная исключительно исполняемым JavaScript без строковых
  endpoint/client_id, может не обнаружиться без браузера.

Зависимость:
    pip install requests
"""

from __future__ import annotations

import base64
import getpass
import hashlib
import html as html_module
import json
import re
import secrets
import sys
import traceback
from dataclasses import dataclass, field, replace
from html.parser import HTMLParser
from typing import Iterable, Optional
from urllib.parse import (
    parse_qs,
    parse_qsl,
    urlencode,
    urljoin,
    urlsplit,
    urlunsplit,
)

try:
    import requests
    from requests import Response, Session
    from requests.exceptions import RequestException, SSLError
except ImportError:
    print("Не найден модуль requests.")
    print("Установите его командой: pip install requests")
    input("\nНажмите Enter, чтобы закрыть программу...")
    raise SystemExit(1)


TIMEOUT = 30
MAX_AUTH_STEPS = 30
MAX_REDIRECTS_DURING_DISCOVERY = 10
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/150.0.0.0 Safari/537.36"
)

AUTH_BLOCK_NAMES = [
    "adfs",
    "keycloak",
    "oidc",
    "openid",
    "openId",
    "openIdConnect",
    "identityServer",
    "authentication",
    "authorization",
]

CLIENT_ID_NAMES = ["clientId", "clientID", "client_id"]
RESPONSE_TYPE_NAMES = ["responseType", "response_type"]
RESPONSE_MODE_NAMES = ["responseMode", "response_mode"]
TOKEN_URL_NAMES = ["tokenUrl", "tokenURL", "tokenEndpoint", "token_endpoint"]
AUTHORIZE_URL_NAMES = [
    "authorizeUrl",
    "authorizationUrl",
    "authorizationEndpoint",
    "authorization_endpoint",
]
AUTHORITY_NAMES = [
    "authority",
    "issuer",
    "url",
    "serverUrl",
    "authServerUrl",
    "keycloakUrl",
    "baseUrl",
    "adfsUrl",
]

USERNAME_NAMES = {
    "username",
    "user",
    "userid",
    "user_name",
    "login",
    "email",
    "mail",
    "upn",
    "account",
    "identifier",
}
PASSWORD_NAMES = {
    "password",
    "passwd",
    "pass",
    "userpassword",
    "credential",
}
OTP_MARKERS = {
    "otp",
    "totp",
    "one-time",
    "onetime",
    "verificationcode",
    "verification_code",
    "authcode",
    "auth_code",
    "smscode",
    "sms_code",
    "mfa",
}
CAPTCHA_MARKERS = {"captcha", "recaptcha", "hcaptcha"}


@dataclass
class AuthConfig:
    provider: str
    authority_url: str
    authorize_url: str
    token_url: Optional[str]
    client_id: str
    response_type: str
    response_mode: Optional[str]
    scope: str
    source: str
    score: int = 0

    def identity(self) -> tuple[str, str, str, str]:
        return (
            self.provider.lower(),
            self.authorize_url.rstrip("/").lower(),
            self.client_id,
            self.response_type,
        )


@dataclass
class OAuthResult:
    access_token: Optional[str] = None
    id_token: Optional[str] = None
    code: Optional[str] = None


@dataclass
class AuthTransaction:
    config: AuthConfig
    redirect_uri: str
    state: str
    nonce: str
    code_verifier: str = ""
    authorize_request_url: str = ""


@dataclass
class InputField:
    name: str
    input_type: str = "text"
    value: str = ""
    element_id: str = ""
    autocomplete: str = ""
    placeholder: str = ""
    checked: bool = False
    disabled: bool = False


@dataclass
class HtmlForm:
    action: str = ""
    method: str = "post"
    form_id: str = ""
    enctype: str = "application/x-www-form-urlencoded"
    inputs: list[InputField] = field(default_factory=list)

    def named_inputs(self) -> list[InputField]:
        return [item for item in self.inputs if item.name and not item.disabled]


class StepCounter:
    def __init__(self) -> None:
        self._value = 1

    def next(self) -> int:
        value = self._value
        self._value += 1
        return value


class PageParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.scripts: list[str] = []
        self.links: list[str] = []
        self.forms: list[HtmlForm] = []
        self._current_form: Optional[HtmlForm] = None
        self._current_button: Optional[InputField] = None
        self._button_text: list[str] = []

    def handle_starttag(
        self,
        tag: str,
        attrs: list[tuple[str, Optional[str]]],
    ) -> None:
        attributes = {key.lower(): (value or "") for key, value in attrs}
        tag_name = tag.lower()

        if tag_name == "script":
            src = attributes.get("src", "").strip()
            if src:
                self.scripts.append(src)
            return

        if tag_name == "link":
            href = attributes.get("href", "").strip()
            if href:
                self.links.append(href)
            return

        if tag_name == "form":
            form = HtmlForm(
                action=attributes.get("action", ""),
                method=(attributes.get("method") or "get").lower(),
                form_id=attributes.get("id", ""),
                enctype=(
                    attributes.get("enctype")
                    or "application/x-www-form-urlencoded"
                ).lower(),
            )
            self.forms.append(form)
            self._current_form = form
            return

        if self._current_form is None:
            return

        if tag_name == "input":
            name = attributes.get("name", "")
            if not name:
                return
            self._current_form.inputs.append(
                InputField(
                    name=name,
                    input_type=(attributes.get("type") or "text").lower(),
                    value=attributes.get("value", ""),
                    element_id=attributes.get("id", ""),
                    autocomplete=attributes.get("autocomplete", ""),
                    placeholder=attributes.get("placeholder", ""),
                    checked="checked" in attributes,
                    disabled="disabled" in attributes,
                )
            )
            return

        if tag_name == "button":
            name = attributes.get("name", "")
            if not name:
                return
            self._current_button = InputField(
                name=name,
                input_type=(attributes.get("type") or "submit").lower(),
                value=attributes.get("value", ""),
                element_id=attributes.get("id", ""),
                disabled="disabled" in attributes,
            )
            self._button_text = []

    def handle_data(self, data: str) -> None:
        if self._current_button is not None:
            self._button_text.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag_name = tag.lower()
        if tag_name == "button" and self._current_button is not None:
            if not self._current_button.value:
                self._current_button.value = " ".join(self._button_text).strip()
            if self._current_form is not None:
                self._current_form.inputs.append(self._current_button)
            self._current_button = None
            self._button_text = []
        elif tag_name == "form":
            self._current_form = None


# ---------------------------------------------------------------------------
# Консоль и HTTP
# ---------------------------------------------------------------------------


def configure_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass


def pause() -> None:
    try:
        input("\nНажмите Enter, чтобы закрыть программу...")
    except (EOFError, KeyboardInterrupt):
        pass


def safe_url_for_log(url: str) -> str:
    parsed = urlsplit(url)
    query_keys = [key for key, _ in parse_qsl(parsed.query, keep_blank_values=True)]
    query = ""
    if query_keys:
        query = "?" + "&".join(f"{key}=…" for key in dict.fromkeys(query_keys))
    return urlunsplit((parsed.scheme, parsed.netloc, parsed.path, query, ""))


def status_line(step: int, method: str, url: str, response: Response) -> None:
    print(f"[{step}] {method:<4} {safe_url_for_log(url)}")
    print(f"    -> HTTP {response.status_code} {response.reason}")


def request(
    session: Session,
    method: str,
    url: str,
    *,
    verify: bool | str,
    step: int,
    show_status: bool = True,
    **kwargs,
) -> Response:
    try:
        response = session.request(
            method,
            url,
            timeout=TIMEOUT,
            verify=verify,
            **kwargs,
        )
    except SSLError as exc:
        raise RuntimeError(
            "Ошибка проверки SSL-сертификата. Укажите путь к корпоративному CA "
            "либо временно отключите проверку SSL только для теста."
        ) from exc
    except RequestException as exc:
        raise RuntimeError(f"Сетевая ошибка: {exc}") from exc

    if show_status:
        status_line(step, method.upper(), url, response)
    return response


# ---------------------------------------------------------------------------
# Поиск страницы и frontend-скриптов
# ---------------------------------------------------------------------------


def normalize_input_url(raw_url: str) -> str:
    value = raw_url.strip()
    if not value:
        raise ValueError("Адрес сайта не указан.")
    if "://" not in value:
        value = "https://" + value

    parsed = urlsplit(value)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("Некорректный HTTP/HTTPS-адрес сайта.")

    path = parsed.path or "/"
    return urlunsplit((parsed.scheme, parsed.netloc, path, "", ""))


def candidate_page_urls(input_url: str) -> list[str]:
    parsed = urlsplit(input_url)
    path = parsed.path or "/"
    result: list[str] = []

    def add(candidate_path: str) -> None:
        if not candidate_path.startswith("/"):
            candidate_path = "/" + candidate_path
        if not candidate_path.endswith("/"):
            candidate_path += "/"
        candidate = urlunsplit(
            (parsed.scheme, parsed.netloc, candidate_path, "", "")
        )
        if candidate not in result:
            result.append(candidate)

    add(path)

    if not re.search(r"/(?:ru|en)(?:/|$)", path, flags=re.IGNORECASE):
        base_path = path.rstrip("/")
        add(base_path + "/ru/")
        add(base_path + "/en/")

    return result


def parse_page(html: str) -> PageParser:
    parser = PageParser()
    parser.feed(html)
    return parser


def looks_like_projectpoint_page(response: Response) -> bool:
    if response.status_code >= 400:
        return False
    parser = parse_page(response.text)
    script_names = [urlsplit(src).path.rsplit("/", 1)[-1] for src in parser.scripts]
    return any(
        re.fullmatch(r"main(?:\.[A-Za-z0-9_-]+)?\.js", name, re.IGNORECASE)
        for name in script_names
    )


def find_project_page(
    session: Session,
    raw_url: str,
    *,
    verify: bool | str,
    steps: StepCounter,
) -> tuple[str, str]:
    input_url = normalize_input_url(raw_url)
    errors: list[str] = []

    for candidate in candidate_page_urls(input_url):
        try:
            response = request(
                session,
                "GET",
                candidate,
                verify=verify,
                step=steps.next(),
                allow_redirects=True,
            )
        except RuntimeError as exc:
            errors.append(f"{candidate}: {exc}")
            continue

        if looks_like_projectpoint_page(response):
            final_url = response.url
            if not final_url.endswith("/"):
                final_url += "/"
            return final_url, response.text

        errors.append(
            f"{candidate}: HTTP {response.status_code}, main.*.js не найден"
        )

    details = "\n".join(f"  - {item}" for item in errors)
    raise RuntimeError(
        "Не удалось найти рабочую страницу Project Point.\n" + details
    )


def find_script_urls(html: str, page_url: str) -> list[str]:
    parser = parse_page(html)
    result: list[str] = []
    for src in parser.scripts:
        absolute = urljoin(page_url, src)
        filename = urlsplit(absolute).path.rsplit("/", 1)[-1]
        if filename.lower().endswith(".js") and absolute not in result:
            result.append(absolute)

    result.sort(
        key=lambda value: (
            0
            if re.fullmatch(
                r"main(?:\.[A-Za-z0-9_-]+)?\.js",
                urlsplit(value).path.rsplit("/", 1)[-1],
                re.IGNORECASE,
            )
            else 1
        )
    )
    return result


def download_frontend_scripts(
    session: Session,
    html: str,
    page_url: str,
    *,
    verify: bool | str,
    steps: StepCounter,
) -> list[tuple[str, str]]:
    scripts: list[tuple[str, str]] = []
    urls = find_script_urls(html, page_url)
    if not urls:
        raise RuntimeError("В HTML страницы не найдены JavaScript-файлы.")

    for url in urls:
        try:
            response = request(
                session,
                "GET",
                url,
                verify=verify,
                step=steps.next(),
                allow_redirects=True,
            )
            if response.status_code == 200 and response.text:
                scripts.append((response.url, response.text))
        except RuntimeError as exc:
            print(f"    Пропускаю скрипт: {exc}")

    if not scripts:
        raise RuntimeError("Не удалось загрузить ни один frontend-скрипт.")
    return scripts


# ---------------------------------------------------------------------------
# Извлечение конфигурации из JavaScript
# ---------------------------------------------------------------------------


def extract_balanced_object(text: str, opening_brace: int) -> Optional[str]:
    depth = 0
    quote: Optional[str] = None
    escaped = False

    for index in range(opening_brace, len(text)):
        char = text[index]
        if quote is not None:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == quote:
                quote = None
            continue

        if char in {"'", '"', "`"}:
            quote = char
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[opening_brace : index + 1]
    return None


def decode_js_string(value: str) -> str:
    value = html_module.unescape(value)
    value = value.replace(r"\/", "/")
    value = re.sub(
        r"\\u([0-9a-fA-F]{4})",
        lambda match: chr(int(match.group(1), 16)),
        value,
    )
    return value.strip()


def extract_property(text: str, names: Iterable[str]) -> Optional[str]:
    alternatives = "|".join(re.escape(name) for name in names)
    pattern = re.compile(
        rf"(?:[\"']?(?:{alternatives})[\"']?)\s*:\s*([\"'])(.*?)\1",
        flags=re.IGNORECASE | re.DOTALL,
    )
    match = pattern.search(text)
    return decode_js_string(match.group(2)) if match else None


def iter_property_matches(text: str, names: Iterable[str]) -> Iterable[re.Match[str]]:
    alternatives = "|".join(re.escape(name) for name in names)
    pattern = re.compile(
        rf"(?:[\"']?(?:{alternatives})[\"']?)\s*:\s*([\"'])(.*?)\1",
        flags=re.IGNORECASE | re.DOTALL,
    )
    return pattern.finditer(text)


def find_named_blocks(js: str, names: Iterable[str]) -> list[tuple[str, str]]:
    result: list[tuple[str, str]] = []
    for name in names:
        pattern = re.compile(rf"[\"']?{re.escape(name)}[\"']?\s*:\s*\{{", re.I)
        for match in pattern.finditer(js):
            opening = js.find("{", match.start())
            block = extract_balanced_object(js, opening)
            if block:
                result.append((name, block))
    return result


def extract_http_urls(text: str) -> list[str]:
    normalized = text.replace(r"\/", "/")
    raw_urls = re.findall(r"https?://[^\s\"'`<>\\]+", normalized, flags=re.I)
    result: list[str] = []
    for value in raw_urls:
        cleaned = html_module.unescape(value).rstrip("),;]}>.")
        if cleaned not in result:
            result.append(cleaned)
    return result


def strip_endpoint_suffix(url: str, suffixes: Iterable[str]) -> str:
    original = url.rstrip("/")
    lowered = original.lower()
    for suffix in suffixes:
        suffix_value = suffix.rstrip("/")
        if lowered.endswith(suffix_value.lower()):
            return original[: -len(suffix_value)].rstrip("/")
    return original


def infer_provider(context: str, urls: list[str], hint: str = "") -> str:
    combined = (hint + " " + context[:2000] + " " + " ".join(urls)).lower()
    if "/realms/" in combined and "openid-connect" in combined:
        return "keycloak"
    if "keycloak" in combined or ("realm" in combined and "authserver" in combined):
        return "keycloak"
    if "login.microsoftonline.com" in combined:
        return "entra-id"
    if "/adfs/" in combined or "adfs" in combined:
        return "adfs"
    if "/connect/authorize" in combined:
        return "identityserver"
    return "oidc"


def keycloak_issuer(raw_url: str, realm: Optional[str]) -> Optional[str]:
    value = raw_url.rstrip("/")
    lowered = value.lower()
    marker = "/protocol/openid-connect/"
    if marker in lowered:
        return value[: lowered.index(marker)].rstrip("/")
    if re.search(r"/realms/[^/?#]+$", value, flags=re.I):
        return value
    if realm:
        return f"{value}/realms/{realm.strip('/')}"
    return None


def normalize_adfs_authority(raw_url: str) -> str:
    value = strip_endpoint_suffix(
        raw_url,
        [
            "/oauth2/authorize",
            "/oauth2/authorize/",
            "/oauth2/token",
            "/oauth2/token/",
            "/.well-known/openid-configuration",
        ],
    )
    parsed = urlsplit(value)
    path = parsed.path.rstrip("/")
    if not re.search(r"/adfs$", path, flags=re.I):
        if "/adfs/" in path.lower():
            index = path.lower().index("/adfs/")
            path = path[: index + len("/adfs")]
        elif "adfs" in parsed.netloc.lower():
            path = path + "/adfs"
    return urlunsplit((parsed.scheme, parsed.netloc, path, "", "")).rstrip("/")


def choose_semantic_url(urls: list[str], provider: str) -> Optional[str]:
    if provider == "keycloak":
        for url in urls:
            low = url.lower()
            if "/realms/" in low and "/protocol/openid-connect/auth" in low:
                return url
        for url in urls:
            if "/realms/" in url.lower():
                return url
    elif provider == "adfs":
        for url in urls:
            if "/adfs/oauth2/authorize" in url.lower():
                return url
        for url in urls:
            if "/adfs" in url.lower():
                return url
    else:
        for url in urls:
            low = url.lower()
            if any(
                marker in low
                for marker in [
                    "/oauth2/authorize",
                    "/oauth2/v2.0/authorize",
                    "/connect/authorize",
                    "/protocol/openid-connect/auth",
                ]
            ):
                return url
    return None


def build_candidate_from_context(
    context: str,
    source: str,
    *,
    hint: str = "",
    extra_score: int = 0,
) -> Optional[AuthConfig]:
    client_id = extract_property(context, CLIENT_ID_NAMES)
    if not client_id:
        return None

    urls = extract_http_urls(context)
    explicit_authorize = extract_property(context, AUTHORIZE_URL_NAMES)
    raw_authority = extract_property(context, AUTHORITY_NAMES)
    token_url = extract_property(context, TOKEN_URL_NAMES)
    realm = extract_property(context, ["realm"])

    provider = infer_provider(context, urls, hint)
    semantic_url = choose_semantic_url(urls, provider)
    base_value = explicit_authorize or semantic_url or raw_authority
    if not base_value:
        return None

    response_type = extract_property(context, RESPONSE_TYPE_NAMES)
    response_mode = extract_property(context, RESPONSE_MODE_NAMES)
    scope = extract_property(context, ["scope"]) or "openid"

    score = extra_score
    if explicit_authorize:
        score += 30
    if semantic_url:
        score += 20
    if raw_authority:
        score += 10
    if response_type:
        score += 5

    if provider == "keycloak":
        issuer = keycloak_issuer(base_value, realm)
        if not issuer:
            return None
        authorize_url = (
            explicit_authorize
            if explicit_authorize and "/protocol/openid-connect/auth" in explicit_authorize.lower()
            else f"{issuer}/protocol/openid-connect/auth"
        )
        token_endpoint = token_url or f"{issuer}/protocol/openid-connect/token"
        return AuthConfig(
            provider="keycloak",
            authority_url=issuer,
            authorize_url=authorize_url,
            token_url=token_endpoint,
            client_id=client_id,
            response_type=response_type or "code",
            response_mode=response_mode or "query",
            scope=scope,
            source=source,
            score=score + 20,
        )

    if provider == "adfs":
        authority = normalize_adfs_authority(base_value)
        authorize_url = (
            explicit_authorize
            if explicit_authorize and "/oauth2/authorize" in explicit_authorize.lower()
            else f"{authority}/oauth2/authorize/"
        )
        return AuthConfig(
            provider="adfs",
            authority_url=authority,
            authorize_url=authorize_url,
            token_url=token_url or f"{authority}/oauth2/token/",
            client_id=client_id,
            response_type=response_type or "id_token token",
            response_mode=response_mode,
            scope=scope,
            source=source,
            score=score + 20,
        )

    authorize_url = explicit_authorize or semantic_url
    if not authorize_url:
        return None
    authority = strip_endpoint_suffix(
        authorize_url,
        [
            "/oauth2/v2.0/authorize",
            "/oauth2/authorize",
            "/connect/authorize",
        ],
    )
    return AuthConfig(
        provider=provider,
        authority_url=authority,
        authorize_url=authorize_url,
        token_url=token_url,
        client_id=client_id,
        response_type=response_type or "code",
        response_mode=response_mode or "query",
        scope=scope,
        source=source,
        score=score,
    )


def collect_auth_candidates(
    scripts: list[tuple[str, str]],
) -> list[AuthConfig]:
    candidates: list[AuthConfig] = []

    for script_url, js in scripts:
        type_auth = extract_property(js, ["typeAuth", "type_auth"]) or ""

        for block_name, block in find_named_blocks(js, AUTH_BLOCK_NAMES):
            bonus = 40
            if type_auth and block_name.lower() in type_auth.lower():
                bonus += 30
            candidate = build_candidate_from_context(
                block,
                f"{script_url} → блок {block_name}",
                hint=block_name,
                extra_score=bonus,
            )
            if candidate:
                candidates.append(candidate)

        # Контексты вокруг каждого clientId помогают при сильно минифицированном JS.
        for match in iter_property_matches(js, CLIENT_ID_NAMES):
            start = max(0, match.start() - 5000)
            end = min(len(js), match.end() + 5000)
            context = js[start:end]
            candidate = build_candidate_from_context(
                context,
                f"{script_url} → рядом с clientId",
                hint=type_auth,
                extra_score=10,
            )
            if candidate:
                candidates.append(candidate)

        # Резерв: контексты вокруг endpoint URL.
        normalized_js = js.replace(r"\/", "/")
        endpoint_pattern = re.compile(
            r"https?://[^\s\"'`<>\\]+(?:"
            r"/adfs(?:/oauth2/authorize/?)?"
            r"|/realms/[^\s\"'`<>\\]+"
            r"|/connect/authorize"
            r"|/oauth2(?:/v2\.0)?/authorize"
            r")",
            re.I,
        )
        for match in endpoint_pattern.finditer(normalized_js):
            start = max(0, match.start() - 5000)
            end = min(len(normalized_js), match.end() + 5000)
            context = normalized_js[start:end]
            candidate = build_candidate_from_context(
                context,
                f"{script_url} → рядом с endpoint",
                hint=type_auth,
                extra_score=15,
            )
            if candidate:
                candidates.append(candidate)

    deduplicated: dict[tuple[str, str, str, str], AuthConfig] = {}
    for candidate in candidates:
        key = candidate.identity()
        previous = deduplicated.get(key)
        if previous is None or candidate.score > previous.score:
            deduplicated[key] = candidate

    result = sorted(deduplicated.values(), key=lambda item: item.score, reverse=True)
    if not result:
        raise RuntimeError(
            "Во frontend-скриптах не удалось найти client_id и параметры ADFS/Keycloak/OIDC."
        )
    return result


# ---------------------------------------------------------------------------
# OIDC Discovery и построение authorize-запроса
# ---------------------------------------------------------------------------


def discovery_url(config: AuthConfig) -> str:
    return config.authority_url.rstrip("/") + "/.well-known/openid-configuration"


def enrich_from_discovery(
    session: Session,
    config: AuthConfig,
    *,
    verify: bool | str,
    steps: StepCounter,
) -> AuthConfig:
    url = discovery_url(config)
    try:
        response = request(
            session,
            "GET",
            url,
            verify=verify,
            step=steps.next(),
            allow_redirects=True,
        )
    except RuntimeError:
        return config

    if response.status_code != 200:
        return config
    try:
        payload = response.json()
    except ValueError:
        return config

    authorize_url = payload.get("authorization_endpoint") or config.authorize_url
    token_url = payload.get("token_endpoint") or config.token_url
    issuer = payload.get("issuer") or config.authority_url
    return replace(
        config,
        authority_url=str(issuer).rstrip("/"),
        authorize_url=str(authorize_url),
        token_url=str(token_url) if token_url else None,
        score=config.score + 15,
    )


def generate_pkce_pair() -> tuple[str, str]:
    verifier = secrets.token_urlsafe(64)
    digest = hashlib.sha256(verifier.encode("ascii")).digest()
    challenge = base64.urlsafe_b64encode(digest).decode("ascii").rstrip("=")
    return verifier, challenge


def uses_authorization_code(response_type: str) -> bool:
    return "code" in response_type.lower().split()


def create_transaction(config: AuthConfig, redirect_uri: str) -> AuthTransaction:
    state = secrets.token_urlsafe(36)
    nonce = secrets.token_urlsafe(36)
    verifier = ""

    params: dict[str, str] = {
        "response_type": config.response_type,
        "client_id": config.client_id,
        "redirect_uri": redirect_uri,
        "scope": config.scope,
        "state": state,
        "nonce": nonce,
    }
    if config.response_mode:
        params["response_mode"] = config.response_mode

    if uses_authorization_code(config.response_type):
        verifier, challenge = generate_pkce_pair()
        params["code_challenge"] = challenge
        params["code_challenge_method"] = "S256"

    parsed = urlsplit(config.authorize_url)
    existing = dict(parse_qsl(parsed.query, keep_blank_values=True))
    existing.update(params)
    authorize_request_url = urlunsplit(
        (
            parsed.scheme,
            parsed.netloc,
            parsed.path,
            urlencode(existing),
            "",
        )
    )

    return AuthTransaction(
        config=config,
        redirect_uri=redirect_uri,
        state=state,
        nonce=nonce,
        code_verifier=verifier,
        authorize_request_url=authorize_request_url,
    )


# ---------------------------------------------------------------------------
# Формы и OAuth-результат
# ---------------------------------------------------------------------------


def normalize_field_token(value: str) -> str:
    return re.sub(r"[^a-z0-9_-]+", "", value.lower())


def find_username_field(form: HtmlForm) -> Optional[InputField]:
    fields = form.named_inputs()

    for item in fields:
        if item.autocomplete.lower() == "username":
            return item

    for item in fields:
        name = normalize_field_token(item.name)
        element_id = normalize_field_token(item.element_id)
        if name in USERNAME_NAMES or element_id in USERNAME_NAMES:
            return item

    for item in fields:
        if item.input_type.lower() == "email":
            return item

    for item in fields:
        if item.input_type.lower() in {"text", "email", "tel"}:
            return item
    return None


def find_password_field(form: HtmlForm) -> Optional[InputField]:
    fields = form.named_inputs()
    for item in fields:
        if item.input_type.lower() == "password":
            return item
    for item in fields:
        name = normalize_field_token(item.name)
        element_id = normalize_field_token(item.element_id)
        if name in PASSWORD_NAMES or element_id in PASSWORD_NAMES:
            return item
    return None


def form_marker_set(form: HtmlForm) -> set[str]:
    result: set[str] = set()
    for item in form.named_inputs():
        result.add(normalize_field_token(item.name))
        result.add(normalize_field_token(item.element_id))
        result.add(normalize_field_token(item.placeholder))
    return {item for item in result if item}


def is_mfa_form(form: HtmlForm) -> bool:
    markers = form_marker_set(form)
    return any(any(token in marker for token in OTP_MARKERS) for marker in markers)


def is_captcha_form(form: HtmlForm, html: str) -> bool:
    markers = form_marker_set(form)
    combined = " ".join(markers) + " " + html[:10000].lower()
    return any(marker in combined for marker in CAPTCHA_MARKERS)


def is_auto_submit_form(form: HtmlForm) -> bool:
    fields = form.named_inputs()
    if not fields:
        return False
    meaningful = [
        item
        for item in fields
        if item.input_type.lower()
        not in {"hidden", "submit", "button", "image"}
    ]
    return not meaningful


def choose_interactive_form(forms: list[HtmlForm]) -> Optional[HtmlForm]:
    password_forms = [form for form in forms if find_password_field(form)]
    if password_forms:
        return password_forms[0]

    username_forms = [form for form in forms if find_username_field(form)]
    if username_forms:
        return username_forms[0]
    return None


def form_payload(
    form: HtmlForm,
    *,
    username: Optional[str] = None,
    password: Optional[str] = None,
) -> dict[str, str]:
    payload: dict[str, str] = {}
    username_field = find_username_field(form)
    password_field = find_password_field(form)

    for item in form.named_inputs():
        input_type = item.input_type.lower()
        if input_type in {"file", "image", "reset", "button"}:
            continue
        if input_type in {"checkbox", "radio"} and not item.checked:
            continue
        if input_type == "submit":
            # Отправляем первую именованную submit-кнопку, как это делает браузер.
            if item.name not in payload:
                payload[item.name] = item.value
            continue
        payload[item.name] = item.value

    if username is not None and username_field is not None:
        payload[username_field.name] = username
    if password is not None and password_field is not None:
        payload[password_field.name] = password

    # Типовая форма ADFS.
    for key in list(payload):
        if key.lower() == "authmethod":
            payload[key] = payload[key] or "FormsAuthentication"

    return payload


def submit_form(
    session: Session,
    response: Response,
    form: HtmlForm,
    payload: dict[str, str],
    *,
    verify: bool | str,
    steps: StepCounter,
) -> Response:
    action = urljoin(response.url, form.action or response.url)
    method = (form.method or "get").upper()
    parsed_source = urlsplit(response.url)
    origin = urlunsplit((parsed_source.scheme, parsed_source.netloc, "", "", ""))
    headers = {
        "Referer": response.url,
        "Origin": origin,
        "Accept": (
            "text/html,application/xhtml+xml,application/xml;q=0.9,"
            "image/avif,image/webp,*/*;q=0.8"
        ),
    }

    kwargs: dict = {
        "headers": headers,
        "allow_redirects": False,
    }
    if method == "GET":
        kwargs["params"] = payload
    else:
        kwargs["data"] = payload
        headers["Content-Type"] = "application/x-www-form-urlencoded"

    return request(
        session,
        method,
        action,
        verify=verify,
        step=steps.next(),
        **kwargs,
    )


def parse_oauth_values(url: str) -> dict[str, list[str]]:
    parsed = urlsplit(url)
    result: dict[str, list[str]] = {}
    for source in (parsed.query, parsed.fragment):
        for key, values in parse_qs(source, keep_blank_values=True).items():
            result.setdefault(key, []).extend(values)
    return result


def oauth_result_from_values(
    values: dict[str, list[str]],
    expected_state: str,
) -> Optional[OAuthResult]:
    interesting = {"access_token", "id_token", "code", "error"}
    if not interesting.intersection(values):
        return None

    returned_state = values.get("state", [None])[0]
    if returned_state != expected_state:
        raise RuntimeError(
            "Проверка OAuth state не пройдена: результат отклонён."
        )

    error = values.get("error", [None])[0]
    if error:
        description = values.get("error_description", [""])[0]
        raise RuntimeError(
            f"Сервер авторизации вернул ошибку: {error}. {description}".strip()
        )

    return OAuthResult(
        access_token=values.get("access_token", [None])[0],
        id_token=values.get("id_token", [None])[0],
        code=values.get("code", [None])[0],
    )


def extract_oauth_result(url: str, expected_state: str) -> Optional[OAuthResult]:
    return oauth_result_from_values(parse_oauth_values(url), expected_state)


def extract_oauth_result_from_form(
    form: HtmlForm,
    expected_state: str,
) -> Optional[OAuthResult]:
    values: dict[str, list[str]] = {}
    for item in form.named_inputs():
        values.setdefault(item.name, []).append(item.value)
    return oauth_result_from_values(values, expected_state)


def html_text(html: str) -> str:
    cleaned = re.sub(r"<script\b.*?</script>", " ", html, flags=re.I | re.S)
    cleaned = re.sub(r"<style\b.*?</style>", " ", cleaned, flags=re.I | re.S)
    cleaned = re.sub(r"<[^>]+>", " ", cleaned)
    cleaned = html_module.unescape(cleaned)
    return re.sub(r"\s+", " ", cleaned).strip()


def error_hint(html: str) -> str:
    text = html_text(html)
    markers = [
        "incorrect",
        "invalid",
        "authentication failed",
        "невер",
        "ошиб",
        "парол",
        "учетн",
        "учётн",
        "locked",
        "заблок",
        "expired",
        "истек",
    ]
    lowered = text.lower()
    if any(marker in lowered for marker in markers):
        return text[:700]
    return ""


# ---------------------------------------------------------------------------
# Проверка кандидатов и выполнение авторизации
# ---------------------------------------------------------------------------


def prepare_candidate(
    session: Session,
    transaction: AuthTransaction,
    *,
    verify: bool | str,
    steps: StepCounter,
) -> tuple[Response, Optional[OAuthResult]]:
    response = request(
        session,
        "GET",
        transaction.authorize_request_url,
        verify=verify,
        step=steps.next(),
        allow_redirects=False,
    )

    result = extract_oauth_result(response.url, transaction.state)
    if result:
        return response, result

    for _ in range(MAX_REDIRECTS_DURING_DISCOVERY):
        if 300 <= response.status_code < 400:
            location = response.headers.get("Location")
            if not location:
                break
            next_url = urljoin(response.url, location)
            result = extract_oauth_result(next_url, transaction.state)
            if result:
                return response, result
            response = request(
                session,
                "GET",
                next_url,
                verify=verify,
                step=steps.next(),
                allow_redirects=False,
            )
            continue
        break

    return response, None


def response_has_supported_interaction(response: Response) -> bool:
    if response.status_code != 200:
        return False
    parser = parse_page(response.text)
    if choose_interactive_form(parser.forms):
        return True
    if any(is_auto_submit_form(form) for form in parser.forms):
        return True
    return False


def select_working_configuration(
    session: Session,
    candidates: list[AuthConfig],
    redirect_uri: str,
    *,
    verify: bool | str,
    steps: StepCounter,
) -> tuple[AuthTransaction, Response, Optional[OAuthResult]]:
    failures: list[str] = []

    print(f"\nНайдено кандидатов авторизации: {len(candidates)}")
    for index, original in enumerate(candidates, 1):
        print(
            f"\nПроверяю кандидат {index}/{len(candidates)}: "
            f"{original.provider}, client_id={original.client_id}"
        )
        print(f"    authorize: {safe_url_for_log(original.authorize_url)}")

        config = enrich_from_discovery(
            session,
            original,
            verify=verify,
            steps=steps,
        )
        transaction = create_transaction(config, redirect_uri)

        try:
            response, immediate = prepare_candidate(
                session,
                transaction,
                verify=verify,
                steps=steps,
            )
        except RuntimeError as exc:
            failures.append(f"{config.authorize_url}: {exc}")
            continue

        if immediate or response_has_supported_interaction(response):
            return transaction, response, immediate

        authenticate = response.headers.get("WWW-Authenticate", "")
        if response.status_code == 401 and authenticate:
            failures.append(
                f"{config.authorize_url}: требуется {authenticate}"
            )
        else:
            failures.append(
                f"{config.authorize_url}: HTTP {response.status_code}, "
                "поддерживаемая форма не найдена"
            )

    details = "\n".join(f"  - {item}" for item in failures[-10:])
    raise RuntimeError(
        "Ни один найденный OAuth/OIDC-кандидат не открыл поддерживаемую форму.\n"
        + details
    )


def run_form_state_machine(
    session: Session,
    transaction: AuthTransaction,
    initial_response: Response,
    username: str,
    password: str,
    *,
    verify: bool | str,
    steps: StepCounter,
) -> OAuthResult:
    response = initial_response
    username_sent = False
    password_sent = False

    for _ in range(MAX_AUTH_STEPS):
        result = extract_oauth_result(response.url, transaction.state)
        if result:
            return result

        if 300 <= response.status_code < 400:
            location = response.headers.get("Location")
            if not location:
                raise RuntimeError("HTTP redirect не содержит заголовок Location.")
            next_url = urljoin(response.url, location)
            result = extract_oauth_result(next_url, transaction.state)
            if result:
                return result
            response = request(
                session,
                "GET",
                next_url,
                verify=verify,
                step=steps.next(),
                allow_redirects=False,
            )
            continue

        if response.status_code == 401:
            authenticate = response.headers.get("WWW-Authenticate", "")
            if "negotiate" in authenticate.lower() or "ntlm" in authenticate.lower():
                raise RuntimeError(
                    "Сервер требует Windows Integrated Authentication "
                    f"({authenticate}). Обычный requests без SSPI это не пройдёт."
                )
            raise RuntimeError(f"Сервер авторизации вернул HTTP 401: {authenticate}")

        if response.status_code >= 400:
            raise RuntimeError(
                f"Сервер авторизации вернул HTTP {response.status_code}.\n"
                f"Ответ: {html_text(response.text)[:700]}"
            )

        parser = parse_page(response.text)

        # response_mode=form_post или SAML/WS-Fed auto-submit.
        auto_forms = [form for form in parser.forms if is_auto_submit_form(form)]
        for form in auto_forms:
            result = extract_oauth_result_from_form(form, transaction.state)
            if result:
                return result

        interactive = choose_interactive_form(parser.forms)
        if interactive is not None:
            if is_captcha_form(interactive, response.text):
                raise RuntimeError(
                    "Обнаружена CAPTCHA. Без интерактивного браузера продолжить нельзя."
                )
            if is_mfa_form(interactive):
                raise RuntimeError(
                    "Обнаружен этап MFA/OTP. В текущей версии требуется отдельный ввод кода."
                )

            username_field = find_username_field(interactive)
            password_field = find_password_field(interactive)

            # Если после уже отправленного пароля снова показана форма пароля,
            # это обычно неверные данные или дополнительное действие.
            if password_sent and password_field is not None:
                hint = error_hint(response.text)
                message = (
                    "После отправки логина и пароля сервер снова показал форму входа. "
                    "Вероятны неверные учётные данные либо дополнительный этап."
                )
                if hint:
                    message += f"\nОтвет страницы: {hint}"
                raise RuntimeError(message)

            send_username = username if username_field is not None else None
            send_password = password if password_field is not None else None
            payload = form_payload(
                interactive,
                username=send_username,
                password=send_password,
            )
            response = submit_form(
                session,
                response,
                interactive,
                payload,
                verify=verify,
                steps=steps,
            )
            username_sent = username_sent or send_username is not None
            password_sent = password_sent or send_password is not None
            continue

        if auto_forms:
            form = auto_forms[0]
            response = submit_form(
                session,
                response,
                form,
                form_payload(form),
                verify=verify,
                steps=steps,
            )
            continue

        hint = error_hint(response.text)
        message = (
            "Получена HTML-страница без поддерживаемой формы и без OAuth-результата. "
            "Возможны JavaScript-вход, выбор провайдера, MFA, CAPTCHA или SAML-сценарий."
        )
        if hint:
            message += f"\nОтвет страницы: {hint}"
        raise RuntimeError(message)

    raise RuntimeError(
        f"Превышено допустимое число шагов авторизации ({MAX_AUTH_STEPS})."
    )


def exchange_code_for_token(
    session: Session,
    transaction: AuthTransaction,
    code: str,
    *,
    verify: bool | str,
    steps: StepCounter,
) -> str:
    config = transaction.config
    if not config.token_url:
        raise RuntimeError(
            "Получен authorization code, но token endpoint определить не удалось."
        )

    data = {
        "grant_type": "authorization_code",
        "client_id": config.client_id,
        "code": code,
        "redirect_uri": transaction.redirect_uri,
    }
    if transaction.code_verifier:
        data["code_verifier"] = transaction.code_verifier

    response = request(
        session,
        "POST",
        config.token_url,
        verify=verify,
        step=steps.next(),
        data=data,
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
        },
        allow_redirects=False,
    )

    try:
        payload = response.json()
    except ValueError as exc:
        raise RuntimeError(
            "Token endpoint вернул не JSON. "
            f"HTTP {response.status_code}: {response.text[:700]}"
        ) from exc

    if response.status_code >= 400:
        error = payload.get("error", "unknown_error")
        description = payload.get("error_description", "")
        raise RuntimeError(
            f"Не удалось обменять code на token: {error}. {description}".strip()
        )

    access_token = payload.get("access_token")
    if not access_token:
        raise RuntimeError(
            "Token endpoint ответил успешно, но access_token отсутствует."
        )
    return str(access_token)


def finish_token(
    session: Session,
    transaction: AuthTransaction,
    result: OAuthResult,
    *,
    verify: bool | str,
    steps: StepCounter,
) -> str:
    if result.access_token:
        return result.access_token
    if result.code:
        return exchange_code_for_token(
            session,
            transaction,
            result.code,
            verify=verify,
            steps=steps,
        )
    raise RuntimeError(
        "Авторизация завершилась без access_token и без authorization code."
    )


# ---------------------------------------------------------------------------
# Project Point API
# ---------------------------------------------------------------------------


def project_base_url(page_url: str) -> str:
    parsed = urlsplit(page_url)
    path = parsed.path.rstrip("/")
    locale_match = re.search(r"/(?:ru|en)$", path, flags=re.I)
    if locale_match:
        path = path[: locale_match.start()]
    return urlunsplit((parsed.scheme, parsed.netloc, path.rstrip("/"), "", ""))


def profile_url(page_url: str) -> str:
    return urljoin(page_url, "api/Core/UsersService/GetMyProfile")


def check_profile(
    session: Session,
    page_url: str,
    access_token: str,
    *,
    verify: bool | str,
    steps: StepCounter,
) -> bool:
    url = profile_url(page_url)
    response = request(
        session,
        "GET",
        url,
        verify=verify,
        step=steps.next(),
        headers={
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
        },
        allow_redirects=False,
    )

    print("\nПРОВЕРКА ПОДКЛЮЧЕНИЯ")
    if response.status_code == 200:
        print("УСПЕХ: API Project Point принял access token.")
        try:
            payload = response.json()
            formatted = json.dumps(payload, ensure_ascii=False, indent=2)
            print("\nОтвет GetMyProfile:")
            print(formatted[:4000])
            if len(formatted) > 4000:
                print("\n...ответ сокращён...")
        except ValueError:
            print("Сервер вернул HTTP 200, но ответ не является JSON.")
            print(response.text[:1000])
        return True

    if response.status_code == 401:
        print("ОШИБКА: HTTP 401. Токен отсутствует, истёк или не принят API.")
    elif response.status_code == 403:
        print("ОШИБКА: HTTP 403. Токен принят, но у пользователя нет прав.")
    else:
        print(f"ОШИБКА: API вернул HTTP {response.status_code}.")
        print(response.text[:1000])
    return False


# ---------------------------------------------------------------------------
# Запуск
# ---------------------------------------------------------------------------


def choose_ssl_verification() -> bool | str:
    answer = input("Проверять SSL-сертификат? [Y/n/путь к CA]: ").strip()
    if not answer or answer.lower() in {"y", "yes", "д", "да"}:
        return True
    if answer.lower() in {"n", "no", "н", "нет"}:
        try:
            requests.packages.urllib3.disable_warnings()  # type: ignore[attr-defined]
        except Exception:
            pass
        print("ВНИМАНИЕ: SSL-проверка отключена только для этого запуска.")
        return False
    return answer


def print_candidates(candidates: list[AuthConfig]) -> None:
    print("\nОБНАРУЖЕННЫЕ КОНФИГУРАЦИИ")
    for index, item in enumerate(candidates, 1):
        print(f"  [{index}] provider:      {item.provider}")
        print(f"      client_id:     {item.client_id}")
        print(f"      authorize:     {safe_url_for_log(item.authorize_url)}")
        print(f"      response_type: {item.response_type}")
        print(f"      response_mode: {item.response_mode or '[по умолчанию]'}")
        print(f"      scope:         {item.scope}")
        print(f"      источник:      {item.source}")


def run() -> None:
    print("=" * 82)
    print("PROJECT POINT — УНИВЕРСАЛЬНАЯ HTTP-АВТОРИЗАЦИЯ ADFS / KEYCLOAK")
    print("=" * 82)
    print("Браузер не запускается. Адреса SSO и client_id ищутся автоматически.\n")

    raw_url = input("Адрес сайта Project Point: ").strip()
    verify = choose_ssl_verification()

    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": USER_AGENT,
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            "Connection": "keep-alive",
        }
    )
    steps = StepCounter()

    print("\nЭТАП 1. Поиск страницы Project Point\n")
    page_url, page_html = find_project_page(
        session,
        raw_url,
        verify=verify,
        steps=steps,
    )
    print(f"\nРабочая страница: {page_url}")

    print("\nЭТАП 2. Загрузка frontend-конфигурации\n")
    scripts = download_frontend_scripts(
        session,
        page_html,
        page_url,
        verify=verify,
        steps=steps,
    )
    candidates = collect_auth_candidates(scripts)
    print_candidates(candidates)

    print("\nЭТАП 3. Проверка найденных провайдеров\n")
    transaction, auth_page, immediate = select_working_configuration(
        session,
        candidates,
        page_url,
        verify=verify,
        steps=steps,
    )

    selected = transaction.config
    print("\nВЫБРАНА РАБОЧАЯ КОНФИГУРАЦИЯ")
    print(f"    provider:       {selected.provider}")
    print(f"    authority:      {selected.authority_url}")
    print(f"    authorize URL:  {selected.authorize_url}")
    print(f"    token URL:      {selected.token_url or '[не требуется/не найден]'}")
    print(f"    client_id:      {selected.client_id}")
    print(f"    response_type:  {selected.response_type}")
    print(f"    response_mode:  {selected.response_mode or '[по умолчанию]'}")
    print(f"    redirect_uri:   {transaction.redirect_uri}")

    if immediate is None:
        print("\nПубличная часть SSO проверена. Теперь можно вводить учётные данные.\n")
        username = input("Логин: ").strip()
        password = getpass.getpass("Пароль: ")
        if not username:
            raise ValueError("Логин не указан.")
        if not password:
            raise ValueError("Пароль не указан.")

        print("\nЭТАП 4. Авторизация\n")
        result = run_form_state_machine(
            session,
            transaction,
            auth_page,
            username,
            password,
            verify=verify,
            steps=steps,
        )
    else:
        print("\nSSO-сессия уже активна: сервер сразу вернул OAuth-результат.")
        result = immediate

    access_token = finish_token(
        session,
        transaction,
        result,
        verify=verify,
        steps=steps,
    )
    print("    access_token получен: " + access_token[:8] + "... [скрыт]")

    print("\nЭТАП 5. Проверка Project Point API\n")
    check_profile(
        session,
        page_url,
        access_token,
        verify=verify,
        steps=steps,
    )


def main() -> None:
    configure_console()
    try:
        run()
    except KeyboardInterrupt:
        print("\nОперация отменена пользователем.")
    except Exception as exc:
        print("\n" + "=" * 82)
        print("ПОДКЛЮЧЕНИЕ НЕ ЗАВЕРШЕНО")
        print("=" * 82)
        print(str(exc))

        try:
            debug = input("\nПоказать техническую трассировку? [y/N]: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            debug = ""
        if debug in {"y", "yes", "д", "да"}:
            traceback.print_exc()
    finally:
        pause()


if __name__ == "__main__":
    main()

# ---------------------------------------------------------------------------
# Reusable integration entry point for Larix CDE
# ---------------------------------------------------------------------------

class InteractiveAuthRequired(RuntimeError):
    """Raised when the HTTP-only flow cannot complete an interactive login."""


def authenticate_projectpoint_http(
    raw_url: str,
    username: str,
    password: str,
    *,
    verify: bool | str = False,
) -> dict:
    """Discover Project Point auth settings and return a verified access token.

    This is the programmatic counterpart of ``run()`` above.  It intentionally
    does not ask questions on stdin and is therefore safe to call from the GUI
    worker thread.  Any unsupported interactive step (MFA, CAPTCHA, WebAuthn,
    Windows Integrated Authentication, JS-only login, etc.) is surfaced as
    ``InteractiveAuthRequired`` so the caller can continue in the embedded
    browser instead of trying to bypass the protection.
    """
    if not raw_url or not raw_url.strip():
        raise ValueError("Адрес Project Point не указан.")
    if not username or not username.strip():
        raise ValueError("Логин не указан.")
    if not password:
        raise ValueError("Пароль не указан.")

    local_session = requests.Session()
    local_session.headers.update(
        {
            "User-Agent": USER_AGENT,
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            "Connection": "keep-alive",
        }
    )
    if verify is False:
        try:
            requests.packages.urllib3.disable_warnings()  # type: ignore[attr-defined]
        except Exception:
            pass

    steps = StepCounter()
    try:
        page_url, page_html = find_project_page(
            local_session,
            raw_url,
            verify=verify,
            steps=steps,
        )
        scripts = download_frontend_scripts(
            local_session,
            page_html,
            page_url,
            verify=verify,
            steps=steps,
        )
        candidates = collect_auth_candidates(scripts)
        transaction, auth_page, immediate = select_working_configuration(
            local_session,
            candidates,
            page_url,
            verify=verify,
            steps=steps,
        )

        if immediate is None:
            result = run_form_state_machine(
                local_session,
                transaction,
                auth_page,
                username.strip(),
                password,
                verify=verify,
                steps=steps,
            )
        else:
            result = immediate

        access_token = finish_token(
            local_session,
            transaction,
            result,
            verify=verify,
            steps=steps,
        )

        # Verify the exact token before returning it to the GUI.
        if not check_profile(
            local_session,
            page_url,
            access_token,
            verify=verify,
            steps=steps,
        ):
            raise RuntimeError("Project Point API не подтвердил полученный access token.")

        return {
            "access_token": access_token,
            "page_url": page_url,
            "base_url": project_base_url(page_url),
            "provider": transaction.config.provider,
            "client_id": transaction.config.client_id,
            "authority_url": transaction.config.authority_url,
        }
    except Exception as exc:
        message = str(exc)
        lowered = message.casefold()
        interactive_markers = (
            "captcha",
            "mfa",
            "otp",
            "webauthn",
            "passkey",
            "windows integrated",
            "negotiate",
            "ntlm",
            "javascript-вход",
            "поддерживаемая форма не найдена",
            "без интерактивного браузера",
            "дополнительный этап",
        )
        if any(marker in lowered for marker in interactive_markers):
            raise InteractiveAuthRequired(message) from exc
        raise
