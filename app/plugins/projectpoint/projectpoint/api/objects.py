from __future__ import annotations

from .common import api_headers

__all__ = ["get_object_structures_for_select", "get_object_structures_odata", "get_object_structures_for_project", "create_object_structure"]


def get_object_structures_for_select(sess, access_token, base_url):
    url = f"{base_url.rstrip('/')}/ru/api/Core/ObjectStructureService/GetAllForSelect"
    resp = sess.get(url, headers=api_headers(access_token))
    if resp.status_code != 200:
        raise Exception(f"GetAllForSelect ObjectStructures error: HTTP {resp.status_code}, {resp.text[:300]}")
    try:
        data = resp.json()
    except Exception:
        raise Exception(f"GetAllForSelect ObjectStructures error: ответ не JSON, {resp.text[:300]}")

    if isinstance(data, dict):
        if isinstance(data.get("value"), list):
            return data["value"]
        if isinstance(data.get("Items"), list):
            return data["Items"]
    if isinstance(data, list):
        return data
    raise Exception(f"GetAllForSelect ObjectStructures error: неожиданный формат ответа ({type(data)})")


def get_object_structures_odata(sess, access_token, base_url, project_id):
    safe_project_id = str(project_id).strip()
    if not safe_project_id:
        raise Exception("ProjectId не задан для OData запроса")
    url = (
        f"{base_url.rstrip('/')}/ru/api/Core/ObjectStructureServiceOData/"
        f"?$filter=(ProjectId eq {safe_project_id})"
    )
    resp = sess.get(url, headers=api_headers(access_token))
    if resp.status_code != 200:
        raise Exception(f"ObjectStructureServiceOData error: HTTP {resp.status_code}, {resp.text[:300]}")
    try:
        data = resp.json()
    except Exception:
        raise Exception(f"ObjectStructureServiceOData error: ответ не JSON, {resp.text[:300]}")

    if isinstance(data, dict):
        if isinstance(data.get("value"), list):
            return data["value"]
        if isinstance(data.get("Items"), list):
            return data["Items"]
    if isinstance(data, list):
        return data
    raise Exception(f"ObjectStructureServiceOData error: неожиданный формат ответа ({type(data)})")


def create_object_structure(sess, access_token, base_url, payload):
    url = f"{base_url.rstrip('/')}/ru/api/Core/ObjectStructureService/Create"
    headers = api_headers(access_token, {"Content-Type": "application/json"})
    resp = sess.post(url, headers=headers, json=payload)
    try:
        resp_data = resp.json()
    except Exception:
        resp_data = resp.text
    return resp.status_code in (200, 201), resp.status_code, resp_data


def get_object_structures_for_project(sess, access_token, base_url, project_id):
    """Compatibility name backed by the existing ProjectId-filtered OData call."""
    return get_object_structures_odata(sess, access_token, base_url, project_id)
