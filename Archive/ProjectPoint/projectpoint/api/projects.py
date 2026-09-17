from __future__ import annotations

from .common import api_headers

__all__ = ["get_projects_for_select", "get_all_projects"]


def get_projects_for_select(sess, access_token, base_url):
    base = base_url.rstrip("/")
    urls = [
        f"{base}/ru/api/Core/ProjectService/GetAll",
        f"{base}/ru/api/Core/ProjectService/GetAllForSelect",
        f"{base}/ru/api/Core/ProjectServiceOData/",
    ]
    headers = api_headers(access_token)
    last_error = None
    for url in urls:
        try:
            resp = sess.get(url, headers=headers)
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

        if isinstance(data, dict):
            if isinstance(data.get("value"), list):
                return data["value"]
            if isinstance(data.get("Items"), list):
                return data["Items"]
        if isinstance(data, list):
            return data
        last_error = f"{url}: неожиданный формат ответа ({type(data)})"

    raise Exception(f"ProjectService endpoint недоступен. Последняя ошибка: {last_error}")


def get_all_projects(sess, access_token, base_url):
    """Compatibility name used by the supplied route worker."""
    return get_projects_for_select(sess, access_token, base_url)
