from __future__ import annotations

from typing import Any

from ..auth import get_access_token, force_disable_ssl_verification

__all__ = ["api_headers", "_get_json_from_variants", "_post_json_to_variants"]


def api_headers(access_token, extra=None):
    h = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0",
    }
    if extra:
        h.update(extra)
    return h


def _get_json_from_variants(sess, access_token, urls, expected_type=None, timeout=60):
    sess = force_disable_ssl_verification(sess)
    """Try several API URL variants and return parsed JSON from the first valid response."""
    headers = api_headers(access_token)
    last_error = None

    for url in urls:
        try:
            resp = sess.get(url, headers=headers, timeout=timeout, verify=False)
        except Exception as e:
            last_error = f"{url}: {e}"
            continue

        if resp.status_code != 200:
            last_error = f"{url}: HTTP {resp.status_code}, {resp.text[:300]}"
            continue

        try:
            data = resp.json()
        except Exception:
            last_error = f"{url}: ответ не JSON, {resp.text[:300]}"
            continue

        if isinstance(data, dict) and "Items" in data:
            data = data["Items"]

        if expected_type and not isinstance(data, expected_type):
            raise Exception(f"{url}: ожидался тип {expected_type}, получено {type(data)}")

        return data

    raise Exception(f"API endpoint недоступен. Последняя ошибка: {last_error}")


def _post_json_to_variants(sess, access_token, urls, payload, timeout=60):
    sess = force_disable_ssl_verification(sess)
    """POST JSON to several API URL variants and return the first non-missing endpoint result."""
    headers = api_headers(access_token, {"Content-Type": "application/json"})
    last_status = None
    last_data = None

    for url in urls:
        resp = sess.post(url, headers=headers, json=payload, timeout=timeout, verify=False)
        last_status = resp.status_code
        try:
            last_data = resp.json()
        except Exception:
            last_data = resp.text

        if resp.status_code in (200, 201, 204):
            return True, resp.status_code, last_data

        # If endpoint exists but payload/business validation failed, preserve that
        # response instead of masking it with a fallback /api 404/405.
        if resp.status_code not in (404, 405):
            return False, resp.status_code, last_data

    return False, last_status, last_data
