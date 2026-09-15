from __future__ import annotations

from .common import api_headers, _get_json_from_variants

__all__ = ["get_all_approval_routes", "create_approval_route", "get_all_disciplines", "get_all_departments"]


def get_all_approval_routes(sess, access_token, base_url):
    url = f"{base_url.rstrip('/')}/ru/api/Archive/DocumentApproveRouteService/GetAll"
    resp = sess.get(url, headers=api_headers(access_token))
    if resp.status_code != 200:
        raise Exception(f"GetAll ApprovalRoutes error: HTTP {resp.status_code}, {resp.text[:300]}")
    return resp.json()


def create_approval_route(sess, access_token, base_url, payload):
    url = f"{base_url.rstrip('/')}/ru/api/Archive/DocumentApproveRouteService/Create"
    headers = api_headers(access_token, {"Content-Type": "application/json"})
    resp = sess.post(url, headers=headers, json=payload)
    try:
        resp_data = resp.json()
    except Exception:
        resp_data = resp.text
    return resp.status_code in (200, 201), resp.status_code, resp_data


def get_all_disciplines(sess, access_token, base_url):
    """Load disciplines.

    The supplied Project Point frontend bundle confirms the Core/DisciplineService/
    service root. GetAll follows the same read-only service contract used by the
    adjacent reference services; both locale and non-locale URL variants are tried.
    """
    root = base_url.rstrip("/")
    return _get_json_from_variants(
        sess,
        access_token,
        [
            f"{root}/ru/api/Core/DisciplineService/GetAll",
            f"{root}/api/Core/DisciplineService/GetAll",
        ],
        expected_type=list,
    )


def get_all_departments(sess, access_token, base_url):
    """Load organizational units from the frontend-confirmed DepartmentsService/GetAll."""
    root = base_url.rstrip("/")
    return _get_json_from_variants(
        sess,
        access_token,
        [
            f"{root}/ru/api/Core/DepartmentsService/GetAll",
            f"{root}/api/Core/DepartmentsService/GetAll",
        ],
        expected_type=list,
    )
