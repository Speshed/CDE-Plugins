from __future__ import annotations

from .common import _get_json_from_variants, _post_json_to_variants

__all__ = [
    "get_all_project_stages",
    "get_active_project_stages",
    "create_project_stage",
    "update_project_stage",
]


def _urls(base_url, action):
    root = base_url.rstrip("/")
    return [
        f"{root}/ru/api/Core/ProjectStageService/{action}",
        f"{root}/api/Core/ProjectStageService/{action}",
    ]


def get_all_project_stages(sess, access_token, base_url):
    return _get_json_from_variants(
        sess,
        access_token,
        _urls(base_url, "GetAll"),
        expected_type=list,
    )


def get_active_project_stages(sess, access_token, base_url):
    items = get_all_project_stages(sess, access_token, base_url)
    return [item for item in items if item.get("IsActive", True)]


def create_project_stage(sess, access_token, base_url, payload):
    return _post_json_to_variants(
        sess,
        access_token,
        _urls(base_url, "Create"),
        payload,
    )


def update_project_stage(sess, access_token, base_url, payload):
    if not payload.get("Id"):
        raise ValueError("Для обновления вида документа обязателен Id")
    return _post_json_to_variants(
        sess,
        access_token,
        _urls(base_url, "Update"),
        payload,
    )
