from __future__ import annotations

import base64
import hashlib
import re
import secrets
from html import unescape
from urllib.parse import parse_qs, quote, urljoin, urlparse

import requests
import urllib3

from .config import *

session = requests.Session()
session.verify = False

__all__ = [
    "session", "force_disable_ssl_verification", "discover_auth_config", "merge_auth_config",
    "authenticate", "authenticate_adfs_direct", "get_access_token",
    "_build_code_verifier", "_build_code_challenge", "_extract_form", "_extract_broker_login_url",
    "_response_debug_snippet", "_extract_auth_code_from_url", "_extract_auth_code_from_response",
    "_safe_request_target", "_raise_auth_request_error", "_keycloak_login_diagnostics",
    "_is_sso_8444", "_auth_flow_debug", "_build_keycloak_config",
]


def force_disable_ssl_verification(sess=None):
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    target_session = sess or session
    try:
        target_session.verify = False
    except Exception:
        pass
    return target_session


def discover_auth_config(base_url, sess=None):
    """Auto-discover authentication configuration for a Project Point instance.
    
    Returns a dict with:
    - client_id
    - sso_base_url
    - realm
    - broker_alias
    - adfs_base_url
    - source: "auto" or None if discovery failed
    """
    if sess is None:
        sess = session
    
    config = {
        "client_id": None,
        "sso_base_url": None,
        "realm": None,
        "broker_alias": None,
        "adfs_base_url": None,
        "source": None,
    }
    
    try:
        # Step 1: Try to get the main page and follow redirects
        resp = sess.get(base_url.rstrip("/"), allow_redirects=True, timeout=30, verify=False)
        final_url = resp.url
        
        # Step 2: Look for Keycloak/OpenID patterns in URL
        # Pattern: /realms/{realm}/protocol/openid-connect/auth
        realm_match = re.search(r"/realms/([^/]+)/protocol/openid-connect", final_url)
        if realm_match:
            config["realm"] = realm_match.group(1)
            # Extract SSO base URL from the redirect
            sso_match = re.match(r"(https?://[^/]+)", final_url)
            if sso_match:
                config["sso_base_url"] = sso_match.group(1)
        
        # Step 3: Look for client_id in query parameters
        from urllib.parse import urlparse, parse_qs
        parsed = urlparse(final_url)
        query_params = parse_qs(parsed.query)
        if "client_id" in query_params:
            config["client_id"] = query_params["client_id"][0]
        
        # Step 4: Parse HTML for form action (ADFS detection)
        html = resp.text
        
        # Look for SAML/ADFS form
        saml_form_match = re.search(r'<form[^>]*action="([^"]*adfs[^"]*)"', html, re.IGNORECASE)
        if saml_form_match:
            adfs_url = saml_form_match.group(1)
            # Normalize ADFS URL
            if adfs_url.startswith("/"):
                parsed_base = urlparse(base_url)
                config["adfs_base_url"] = f"{parsed_base.scheme}://{parsed_base.netloc}{adfs_url}"
            elif adfs_url.startswith("http"):
                config["adfs_base_url"] = adfs_url
        
        # Step 5: Look for broker patterns in HTML/JS
        # Pattern: /broker/{alias}/login
        broker_match = re.search(r"/broker/([^/]+)/login", html)
        if broker_match:
            config["broker_alias"] = broker_match.group(1)
        
        # Step 6: Look for frontend config in HTML/JS
        # Common pattern: var config = {...} or window.__CONFIG__ = {...}
        config_js_match = re.search(r'(?:var|window\.)\s*config\s*=\s*({[^;]+})', html, re.IGNORECASE)
        if config_js_match:
            try:
                # Try to safely parse the config object
                config_str = config_js_match.group(1)
                # Extract key-value pairs
                client_match = re.search(r'["\']?clientId["\']?\s*:\s*["\']([^"\']+)["\']', config_str)
                if client_match and not config["client_id"]:
                    config["client_id"] = client_match.group(1)
                
                realm_match = re.search(r'["\']?realm["\']?\s*:\s*["\']([^"\']+)["\']', config_str)
                if realm_match and not config["realm"]:
                    config["realm"] = realm_match.group(1)
                
                sso_match = re.search(r'["\']?authServerUrl["\']?\s*:\s*["\']([^"\']+)["\']', config_str)
                if sso_match and not config["sso_base_url"]:
                    config["sso_base_url"] = sso_match.group(1)
            except Exception:
                pass
        
        # Step 7: If we have SSO base URL but no realm, try well-known endpoint
        if config["sso_base_url"] and not config["realm"]:
            try:
                well_known_url = f"{config['sso_base_url'].rstrip('/')}/realms"
                resp_realms = sess.get(well_known_url, timeout=10, verify=False)
                # Try to extract realm from response
                realm_list_match = re.search(r'realm"\s*:\s*"([^"]+)"', resp_realms.text)
                if realm_list_match:
                    config["realm"] = realm_list_match.group(1)
            except Exception:
                pass
        
        # Step 8: Try common defaults if still missing
        if not config["realm"]:
            config["realm"] = "projectpoint"
        if not config["client_id"]:
            config["client_id"] = "projectpoint-client"
        if not config["broker_alias"]:
            config["broker_alias"] = "adfs"
        
        # Check if we discovered anything meaningful
        if config["sso_base_url"] or (config["realm"] and config["realm"] != "projectpoint"):
            config["source"] = "auto"
        
    except Exception as e:
        print(f"Auto-discovery failed: {e}")
        pass
    
    return config


def merge_auth_config(cached, discovered, manual):
    """Merge auth config from multiple sources with priority:
    1. Manual (highest priority)
    2. Cached
    3. Discovered (lowest priority)
    """
    result = {
        "client_id": None,
        "sso_base_url": None,
        "realm": None,
        "broker_alias": None,
        "adfs_base_url": None,
        "source": "manual" if manual.get("client_id") else None,
    }
    
    # Priority order: manual > cached > discovered
    sources = [discovered, cached, manual]
    for source in sources:
        if not source:
            continue
        for key in result:
            if key != "source" and not result[key] and source.get(key):
                result[key] = source[key]
    
    # Determine final source
    if manual and manual.get("client_id"):
        result["source"] = "manual"
    elif cached and cached.get("client_id"):
        result["source"] = "cache"
    elif discovered and discovered.get("source") == "auto":
        result["source"] = "auto"
    
    return result


def _build_code_verifier():
    alphabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-._~"
    return "".join(secrets.choice(alphabet) for _ in range(96))


def _build_code_challenge(code_verifier):
    digest = hashlib.sha256(code_verifier.encode("ascii")).digest()
    return base64.urlsafe_b64encode(digest).rstrip(b"=").decode("ascii")


def _extract_form(html):
    form_start = html.lower().find("<form")
    if form_start < 0:
        raise Exception("HTML-форма не найдена")
    form_end = html.lower().find("</form>", form_start)
    if form_end < 0:
        raise Exception("Закрывающий тег формы не найден")
    form_html = html[form_start:form_end]

    action_marker = form_html.lower().find("action=")
    action = ""
    if action_marker >= 0:
        quote_char = form_html[action_marker + 7]
        if quote_char in "\"'":
            action_end = form_html.find(quote_char, action_marker + 8)
            action = unescape(form_html[action_marker + 8:action_end])

    fields = {}
    pos = 0
    lower = form_html.lower()
    while True:
        input_start = lower.find("<input", pos)
        if input_start < 0:
            break
        input_end = lower.find(">", input_start)
        if input_end < 0:
            break
        tag = form_html[input_start:input_end]

        def attr(name):
            marker = f'{name}="'
            idx = tag.lower().find(marker)
            if idx >= 0:
                value_end = tag.find('"', idx + len(marker))
                return unescape(tag[idx + len(marker):value_end])
            marker = f"{name}='"
            idx = tag.lower().find(marker)
            if idx >= 0:
                value_end = tag.find("'", idx + len(marker))
                return unescape(tag[idx + len(marker):value_end])
            return None

        name = attr("name")
        if name:
            fields[name] = attr("value") or ""
        pos = input_end + 1

    return action, fields


def _extract_broker_login_url(html, broker_alias):
    marker = f"/broker/{broker_alias}/login?"
    idx = html.find(marker)
    if idx < 0:
        raise Exception(f"Ссылка на брокер '{broker_alias}' не найдена")
    start = idx
    while start > 0 and html[start - 1] not in "\"'":
        start -= 1
    end = idx
    while end < len(html) and html[end] not in "\"'":
        end += 1
    return unescape(html[start:end])


def _response_debug_snippet(response):
    body = response.text[:250].replace("\n", " ").strip()
    return f"HTTP {response.status_code} {response.url} {body}".strip()


def _extract_auth_code_from_url(url):
    if not url:
        return None
    query = parse_qs(urlparse(url).query)
    return query.get("code", [None])[0]


def _extract_auth_code_from_response(response):
    code = _extract_auth_code_from_url(response.url)
    if code:
        return code
    for item in response.history:
        code = _extract_auth_code_from_url(item.headers.get("Location", ""))
        if code:
            return code
    return None


def _safe_request_target(url):
    if not url:
        return "URL не определен"
    parsed = urlparse(url)
    if parsed.scheme and parsed.netloc:
        return parsed._replace(query="", fragment="").geturl()
    return url.split("?", 1)[0]


def _raise_auth_request_error(stage, url, exc):
    target = _safe_request_target(url)
    if isinstance(exc, requests.exceptions.Timeout):
        raise Exception(
            f"Ошибка авторизации: таймаут на этапе {stage}. URL: {target}. "
            "Проверьте VPN, доступность SSO/ProjectPoint и сетевое соединение."
        ) from exc
    if isinstance(exc, requests.exceptions.ConnectionError):
        raise Exception(
            f"Ошибка авторизации: нет соединения на этапе {stage}. URL: {target}. "
            "Проверьте VPN, доступность SSO/ProjectPoint и сетевое соединение."
        ) from exc
    raise exc


def _keycloak_login_diagnostics(response):
    html = response.text or ""
    diagnostics = [f"URL: {_safe_request_target(response.url)}"]

    title_match = re.search(r"<title[^>]*>(.*?)</title>", html, re.IGNORECASE | re.DOTALL)
    if title_match:
        title = unescape(re.sub(r"\s+", " ", title_match.group(1))).strip()
        if title:
            diagnostics.append(f"title: {title[:200]}")

    error_snippets = []
    for match in re.finditer(r"<(?P<tag>[a-z0-9]+)[^>]*class=[\"'][^\"']*(?:error|alert)[^\"']*[\"'][^>]*>(?P<body>.*?)</(?P=tag)>", html, re.IGNORECASE | re.DOTALL):
        body = re.sub(r"<[^>]+>", " ", match.group("body"))
        body = unescape(re.sub(r"\s+", " ", body)).strip()
        if body:
            error_snippets.append(body[:200])
        if len(error_snippets) >= 3:
            break
    if error_snippets:
        diagnostics.append("error/alert: " + " | ".join(error_snippets))

    keywords = [name for name in ("otp", "totp", "code", "authenticator", "login-actions") if name in html.lower()]
    if keywords:
        diagnostics.append("keywords: " + ", ".join(keywords))

    return "\n".join(diagnostics)


def _is_sso_8444(sso_base_url):
    try:
        return urlparse(sso_base_url).port == 8444
    except Exception:
        return False


def _auth_flow_debug(flow_name, redirect_uri, target=None):
    parts = [f"flow={flow_name}", f"redirect_uri={_safe_request_target(redirect_uri)}"]
    if target:
        parts.append(f"target={_safe_request_target(target)}")
    return "; ".join(parts)


def _build_keycloak_config(base_url, sso_base_url, realm):
    base_url = base_url.rstrip("/")
    sso_base_url = sso_base_url.rstrip("/")
    issuer = f"{sso_base_url}/realms/{realm}"
    if _is_sso_8444(sso_base_url):
        redirect_uri = (
            f"{base_url}/ru/auth/realms/{realm}/protocol/openid-connect/logout"
            f"?iss={quote(issuer, safe='')}&iss={quote(issuer, safe='')}"
        )
    elif base_url == DEFAULT_BASE_URL:
        redirect_uri = f"{base_url}/en/"
    else:
        redirect_uri = (
            f"{base_url}/ru/auth/realms/{realm}/protocol/openid-connect/logout"
            f"?iss={quote(issuer, safe='')}&iss={quote(issuer, safe='')}"
        )
    return {
        "auth_url": f"{issuer}/protocol/openid-connect/auth",
        "token_url": f"{issuer}/protocol/openid-connect/token",
        "redirect_uri": redirect_uri,
    }


def authenticate(
    sess,
    client_id,
    base_url,
    username,
    password,
    sso_base_url=DEFAULT_SSO_BASE_URL,
    realm=DEFAULT_REALM,
    broker_alias=DEFAULT_BROKER_ALIAS,
    adfs_base_url=DEFAULT_ADFS_BASE_URL,
):
    sess = force_disable_ssl_verification(sess)
    cfg = _build_keycloak_config(base_url, sso_base_url, realm)
    use_direct_flow = not _is_sso_8444(sso_base_url)
    state = secrets.token_urlsafe(16)
    code_verifier = _build_code_verifier()
    code_challenge = _build_code_challenge(code_verifier)
    params = {
        "client_id": client_id,
        "redirect_uri": cfg["redirect_uri"],
        "state": state,
        "response_mode": "query",
        "response_type": "code",
        "scope": "openid",
        "code_challenge": code_challenge,
        "code_challenge_method": "S256",
    }

    try:
        auth_page = sess.get(
            cfg["auth_url"],
            params=params,
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=AUTH_TIMEOUT_SECONDS,
            verify=False,
        )
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
        _raise_auth_request_error("открытии страницы SSO", cfg["auth_url"], exc)
    auth_code = _extract_auth_code_from_response(auth_page)
    if auth_code:
        token_data = {
            "grant_type": "authorization_code",
            "code": auth_code,
            "redirect_uri": cfg["redirect_uri"],
            "client_id": client_id,
            "code_verifier": code_verifier,
        }
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        try:
            token_resp = sess.post(cfg["token_url"], data=token_data, headers=headers, timeout=AUTH_TIMEOUT_SECONDS, verify=False)
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
            _raise_auth_request_error("обмене кода на токен", cfg["token_url"], exc)
        if token_resp.status_code != 200:
            raise Exception(f"Ошибка получения токена: {_response_debug_snippet(token_resp)}")
        return token_resp.json()["access_token"]

    if auth_page.status_code != 200:
        raise Exception(f"Не удалось открыть страницу SSO: {_response_debug_snippet(auth_page)}")

    login_action, login_fields = _extract_form(auth_page.text)
    broker_url = None
    try:
        broker_url = _extract_broker_login_url(auth_page.text, broker_alias)
    except Exception:
        broker_url = None

    if use_direct_flow and login_action and "login-actions/authenticate" in login_action and not broker_url:
        login_target = urljoin(auth_page.url, login_action)
        login_fields.update(
            {
                "username": username,
                "password": password,
            }
        )
        try:
            login_resp = sess.post(
                login_target,
                data=login_fields,
                headers={"User-Agent": "Mozilla/5.0"},
                allow_redirects=False,
                timeout=AUTH_TIMEOUT_SECONDS,
                verify=False,
            )
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
            _raise_auth_request_error("отправке формы входа Keycloak", login_target, exc)

        if login_resp.status_code != 302:
            raise Exception(
                f"Keycloak не вернул redirect после входа: {_response_debug_snippet(login_resp)}\n"
                f"{_auth_flow_debug('direct_keycloak', cfg['redirect_uri'], login_target)}\n"
                f"{_keycloak_login_diagnostics(login_resp)}"
            )

        final_url = login_resp.headers.get("Location", "")
        if not final_url:
            raise Exception("Keycloak не вернул redirect с кодом авторизации")

        auth_code = _extract_auth_code_from_url(final_url)
        if not auth_code:
            raise Exception("Код авторизации не получен")

        token_data = {
            "grant_type": "authorization_code",
            "code": auth_code,
            "redirect_uri": cfg["redirect_uri"],
            "client_id": client_id,
            "code_verifier": code_verifier,
        }
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        try:
            token_resp = sess.post(cfg["token_url"], data=token_data, headers=headers, timeout=AUTH_TIMEOUT_SECONDS, verify=False)
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
            _raise_auth_request_error("обмене кода на токен", cfg["token_url"], exc)
        if token_resp.status_code != 200:
            raise Exception(f"Ошибка получения токена: {_response_debug_snippet(token_resp)}")
        return token_resp.json()["access_token"]

    if _is_sso_8444(sso_base_url) and not broker_url:
        raise Exception(
            "Production auth ожидает broker/ADFS ветку для SSO :8444, но broker link не найден. "
            f"{_auth_flow_debug('broker_adfs', cfg['redirect_uri'], auth_page.url)}\n"
            f"{_response_debug_snippet(auth_page)}"
        )

    if not broker_url:
        raise Exception(
            "Direct Keycloak flow не выбран: broker link отсутствует и direct flow недоступен для текущего SSO. "
            f"{_auth_flow_debug('broker_adfs', cfg['redirect_uri'], auth_page.url)}"
        )

    broker_target = urljoin(auth_page.url, broker_url)
    try:
        broker_resp = sess.get(broker_target, headers={"User-Agent": "Mozilla/5.0"}, timeout=AUTH_TIMEOUT_SECONDS, verify=False)
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
        _raise_auth_request_error("получении страницы брокера SAML", broker_target, exc)
    if broker_resp.status_code != 200:
        raise Exception(
            f"Брокер SAML вернул ошибку: {_response_debug_snippet(broker_resp)}\n"
            f"{_auth_flow_debug('broker_adfs', cfg['redirect_uri'], broker_target)}"
        )

    saml_action, saml_fields = _extract_form(broker_resp.text)
    if "SAMLRequest" not in saml_fields:
        raise Exception("В форме перехода к ADFS отсутствует SAMLRequest")
    adfs_entry_url = urljoin(adfs_base_url.rstrip("/") + "/", saml_action.lstrip("/"))
    try:
        adfs_entry = sess.post(
            adfs_entry_url,
            data=saml_fields,
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=AUTH_TIMEOUT_SECONDS,
            verify=False,
        )
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
        _raise_auth_request_error("получении формы входа ADFS", adfs_entry_url, exc)
    if adfs_entry.status_code != 200:
        raise Exception(f"ADFS не открыл форму входа: {_response_debug_snippet(adfs_entry)}")

    login_action, login_fields = _extract_form(adfs_entry.text)
    login_fields.update(
        {
            "UserName": username,
            "Password": password,
            "AuthMethod": login_fields.get("AuthMethod", "FormsAuthentication"),
        }
    )
    login_target = urljoin(adfs_entry.url, login_action)
    try:
        adfs_result = sess.post(
            login_target,
            data=login_fields,
            headers={"User-Agent": "Mozilla/5.0"},
            allow_redirects=True,
            timeout=AUTH_TIMEOUT_SECONDS,
            verify=False,
        )
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
        _raise_auth_request_error("отправке учетных данных в ADFS", login_target, exc)
    if adfs_result.status_code != 200:
        raise Exception(f"ADFS отклонил вход: {_response_debug_snippet(adfs_result)}")

    saml_response_action, saml_response_fields = _extract_form(adfs_result.text)
    if "SAMLResponse" not in saml_response_fields:
        raise Exception("ADFS не вернул SAMLResponse после входа")

    broker_finish_url = urljoin(adfs_result.url, saml_response_action)
    try:
        broker_finish = sess.post(
            broker_finish_url,
            data=saml_response_fields,
            headers={"User-Agent": "Mozilla/5.0"},
            allow_redirects=False,
            timeout=AUTH_TIMEOUT_SECONDS,
            verify=False,
        )
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
        _raise_auth_request_error("возврате к Keycloak", broker_finish_url, exc)
    if broker_finish.status_code != 302:
        raise Exception(f"Keycloak не выдал код авторизации: {_response_debug_snippet(broker_finish)}")

    final_url = broker_finish.headers.get("Location", "")
    if not final_url:
        raise Exception("Keycloak не вернул redirect с кодом авторизации")

    auth_code = _extract_auth_code_from_url(final_url)
    if not auth_code:
        raise Exception("Код авторизации не получен")

    token_data = {
        "grant_type": "authorization_code",
        "code": auth_code,
        "redirect_uri": cfg["redirect_uri"],
        "client_id": client_id,
        "code_verifier": code_verifier,
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    try:
        token_resp = sess.post(cfg["token_url"], data=token_data, headers=headers, timeout=AUTH_TIMEOUT_SECONDS, verify=False)
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
        _raise_auth_request_error("обмене кода на токен", cfg["token_url"], exc)
    if token_resp.status_code != 200:
        raise Exception(f"Ошибка получения токена: {_response_debug_snippet(token_resp)}")
    return token_resp.json()["access_token"]


def authenticate_adfs_direct(
    sess,
    client_id,
    base_url,
    username,
    password,
    adfs_base_url,
):
    sess = force_disable_ssl_verification(sess)
    """Direct ADFS OAuth authentication for test environment.
    
    Based on implementation from create_type_gui_2 (2).py
    Uses ADFS OAuth2 authorization code flow directly without Keycloak broker.
    """
    from urllib.parse import urlparse, parse_qs, urljoin
    
    # ADFS OAuth endpoints
    auth_url = f"{adfs_base_url.rstrip('/')}/adfs/oauth2/authorize/"
    token_url = f"{adfs_base_url.rstrip('/')}/adfs/oauth2/token/"
    redirect_uri = base_url.rstrip("/")
    
    # Generate state and nonce
    state = secrets.token_urlsafe(16)
    nonce = secrets.token_urlsafe(16)
    
    # Step 1: Get authorization code via direct POST with credentials
    auth_params = {
        "response_type": "code",
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "scope": "openid",
        "state": state,
        "nonce": nonce,
    }
    
    post_data = {
        "UserName": username,
        "Password": password,
        "AuthMethod": "FormsAuthentication",
    }
    
    # Submit credentials directly to ADFS
    try:
        response = sess.post(auth_url, params=auth_params, data=post_data, allow_redirects=False, timeout=AUTH_TIMEOUT_SECONDS, verify=False)
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
        _raise_auth_request_error("отправке запроса на ADFS", auth_url, exc)
    if response.status_code != 302:
        raise Exception(f"ADFS authentication failed: HTTP {response.status_code}")
    
    # Step 2: Follow redirect to get the code
    location = response.headers.get("Location")
    if not location:
        raise Exception("ADFS did not return redirect location")
    
    try:
        second_response = sess.get(location, allow_redirects=False, timeout=AUTH_TIMEOUT_SECONDS, verify=False)
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
        _raise_auth_request_error("получении redirect от ADFS", location, exc)
    if second_response.status_code != 302:
        raise Exception(f"ADFS redirect failed: HTTP {second_response.status_code}")
    
    # Step 3: Extract authorization code from final redirect URL
    final_url = second_response.headers.get("Location")
    if not final_url:
        raise Exception("ADFS did not return final redirect URL")
    
    parsed = urlparse(final_url)
    query = parse_qs(parsed.query)
    
    if "code" not in query:
        raise Exception("Authorization code not received from ADFS")
    
    auth_code = query["code"][0]
    
    # Step 4: Exchange code for access token
    token_data = {
        "grant_type": "authorization_code",
        "code": auth_code,
        "redirect_uri": redirect_uri,
        "client_id": client_id,
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    
    try:
        token_response = sess.post(token_url, data=token_data, headers=headers, timeout=AUTH_TIMEOUT_SECONDS, verify=False)
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
        _raise_auth_request_error("обмене кода на токен", token_url, exc)
    if token_response.status_code != 200:
        raise Exception(f"Token exchange failed: HTTP {token_response.status_code}")
    
    token_json = token_response.json()
    access_token = token_json.get("access_token")
    
    if not access_token:
        raise Exception("Access token not in response")
    
    return access_token


def get_access_token(
    sess,
    auth_mode,
    client_id,
    base_url,
    username,
    password,
    sso_base_url=None,
    realm=None,
    broker_alias=None,
    adfs_base_url=None,
):
    sess = force_disable_ssl_verification(sess)
    """Unified authentication dispatcher.
    
    Selects appropriate auth flow based on auth_mode:
    - keycloak_broker: Uses Keycloak with broker (production)
    - adfs_direct: Uses direct ADFS OAuth (test)
    """
    if auth_mode == "keycloak_broker":
        return authenticate(
            sess,
            client_id,
            base_url,
            username,
            password,
            sso_base_url=sso_base_url,
            realm=realm,
            broker_alias=broker_alias,
            adfs_base_url=adfs_base_url,
        )
    elif auth_mode == "adfs_direct":
        return authenticate_adfs_direct(
            sess,
            client_id,
            base_url,
            username,
            password,
            adfs_base_url,
        )
    else:
        raise Exception(f"Неизвестный auth_mode: {auth_mode}")
