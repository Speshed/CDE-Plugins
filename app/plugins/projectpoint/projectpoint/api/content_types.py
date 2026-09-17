from __future__ import annotations

from .common import api_headers

__all__ = ["get_all_content_types", "create_content_type", "check_and_update_content_type"]


def get_all_content_types(sess, access_token, base_url):
    url = f"{base_url.rstrip('/')}/ru/api/Core/ContentTypeService/GetAll"
    resp = sess.get(url, headers=api_headers(access_token))
    if resp.status_code != 200:
        raise Exception(f"ContentTypes error: HTTP {resp.status_code}")
    return resp.json()


def create_content_type(sess, access_token, base_url, data):
    url = f"{base_url.rstrip('/')}/ru/api/Core/ContentTypeService/Create"
    headers = api_headers(access_token, {"Content-Type": "application/json"})
    resp = sess.post(url, headers=headers, json=data)
    try:
        resp_data = resp.json()
    except Exception:
        resp_data = resp.text
    if resp.status_code in (200, 201):
        return True, resp.status_code, resp_data
    return False, resp.status_code, resp_data


def check_and_update_content_type(sess, access_token, base_url, data):
    headers = api_headers(access_token, {"Content-Type": "application/json"})

    check_url = f"{base_url.rstrip('/')}/ru/api/Archive/ContentTypeService/CheckUpdate"
    resp = sess.post(check_url, headers=headers, json=data)
    if resp.status_code != 200:
        return False, f"CheckUpdate HTTP {resp.status_code}", resp.text[:300]
    try:
        check_res = resp.json()
        if not check_res.get("IsValid", False):
            return False, "CheckUpdate: IsValid=False", str(check_res)
    except Exception as e:
        return False, f"CheckUpdate parse error: {e}", ""

    update_url = f"{base_url.rstrip('/')}/ru/api/Archive/ContentTypeService/Update"
    resp = sess.post(update_url, headers=headers, json=data)
    if resp.status_code in (200, 201):
        return True, "OK", ""
    return False, f"Update HTTP {resp.status_code}", resp.text[:300]
