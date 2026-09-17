from __future__ import annotations

import json
import re

import pandas as pd
from pathlib import Path

from .common import api_headers, _get_json_from_variants, _post_json_to_variants
from ..auth import force_disable_ssl_verification

BUILTIN_ROLE_PERMISSIONS_SOURCE = "adfs.hgml.ru.har / Core/PermissionService/GetAll / 2026-04-29"
_RESOURCE = Path(__file__).resolve().parent.parent / "resources" / "builtin_role_permissions.json"


def get_builtin_role_permissions():
    with _RESOURCE.open("r", encoding="utf-8") as fh:
        return json.load(fh)

__all__ = [
    "BUILTIN_ROLE_PERMISSIONS_SOURCE", "get_builtin_role_permissions",
    "ROLE_PERMISSION_BLOCKED_MY_IDS", "ROLE_PERMISSION_BLOCKED_ORG_IDS", "ROLE_PERMISSION_EXCLUDED_NAMES",
    "ROLE_PERMISSION_COLUMNS", "PERMISSION_MODULE_FALLBACK_NAMES", "ACTIVE_MODULE_TO_PERMISSION_MODULE",
    "BASE_VISIBLE_PERMISSION_MODULES", "DEFAULT_VISIBLE_PERMISSION_MODULES", "ROLE_UI_MODULE_ORDER",
    "ROLE_UI_MODULE_ORDER_INDEX", "get_all_roles", "get_all_permissions", "get_active_modules",
    "get_visible_permission_module_ids", "is_permission_module_visible", "build_permission_module_title_map",
    "permission_module_display", "_normalize_lookup_text", "_is_cyrillic_char", "_is_latin_char",
    "_role_ui_text_sort_key", "_normalize_module_filter", "_role_ui_module_sort_key", "is_role_permission_visible",
    "get_permission_allowed_flags", "create_role", "update_role", "_existing_role_permission_info_id",
    "build_update_role_payload",
]


ROLE_PERMISSION_BLOCKED_MY_IDS = {8112}


ROLE_PERMISSION_BLOCKED_ORG_IDS = {
    8125, 8158, 8126, 8127, 8149, 8128, 8129, 8130, 8131,
    8102, 8133, 8101, 8113, 8147, 8103, 8132, 8134,
    8105, 8106, 8145, 8160, 8161, 8152,
    1106, 1110, 1105, 1104, 1108, 1102, 1111, 1101,
    1112, 1103, 9101, 1114,
}


ROLE_PERMISSION_EXCLUDED_NAMES = {
    "Удаление коммуникации (комментариев, файлов и фото)",
}


ROLE_PERMISSION_COLUMNS = {
    "My": "Свои элементы",
    "MyOrganization": "Моей орг. единицы",
    "OtherOrganization": "Другие орг. единицы",
}


PERMISSION_MODULE_FALLBACK_NAMES = {
    1: "Настройки",
    5: "Контроль Качества",
    6: "Приёмка Работ",
    8: "Документы",
    9: "Отчеты",
}


ACTIVE_MODULE_TO_PERMISSION_MODULE = {
    0: 1,  # Настройки
    4: 5,  # Контроль качества
    5: 6,  # Приёмка Работ / Планирование и контроль СМР
    7: 8,  # Документы
    9: 9,  # Отчеты
}


BASE_VISIBLE_PERMISSION_MODULES = {1, 9}


DEFAULT_VISIBLE_PERMISSION_MODULES = {1, 5, 6, 8, 9}


def get_all_roles(sess, access_token, base_url):
    """Load Project Point roles with permissions."""
    base = base_url.rstrip("/")
    return _get_json_from_variants(
        sess,
        access_token,
        [
            f"{base}/ru/api/Core/RoleService/GetAll",
            f"{base}/api/Core/RoleService/GetAll",
        ],
        expected_type=list,
    )


def get_all_permissions(sess, access_token, base_url):
    """Load Project Point permission dictionary used by RoleService/Create."""
    base = base_url.rstrip("/")
    return _get_json_from_variants(
        sess,
        access_token,
        [
            f"{base}/ru/api/Core/PermissionService/GetAll",
            f"{base}/api/Core/PermissionService/GetAll",
        ],
        expected_type=list,
    )


def get_active_modules(sess, access_token, base_url):
    base = base_url.rstrip("/")
    try:
        return _get_json_from_variants(
            sess,
            access_token,
            [
                f"{base}/ru/api/Core/ModuleService/GetActive",
                f"{base}/api/Core/ModuleService/GetActive",
            ],
            expected_type=list,
            timeout=30,
        )
    except Exception:
        return []


def get_visible_permission_module_ids(active_modules=None):
    """Return PermissionService.Module ids that are visible in the role UI.

    The frontend does not show every Module value returned by PermissionService/GetAll.
    It shows settings/report modules plus modules mapped from ModuleService/GetActive.
    """
    visible = set(BASE_VISIBLE_PERMISSION_MODULES)
    if active_modules:
        for module in active_modules:
            try:
                active_id = int(module.get("Id"))
            except Exception:
                continue
            permission_module = ACTIVE_MODULE_TO_PERMISSION_MODULE.get(active_id)
            if permission_module:
                visible.add(permission_module)
    else:
        visible.update(DEFAULT_VISIBLE_PERMISSION_MODULES)
    return visible


def is_permission_module_visible(permission, active_modules=None):
    try:
        module_id = int(permission.get("Module"))
    except Exception:
        return False
    return module_id in get_visible_permission_module_ids(active_modules)


def build_permission_module_title_map(active_modules=None):
    titles = dict(PERMISSION_MODULE_FALLBACK_NAMES)
    if active_modules:
        for module in active_modules:
            try:
                active_id = int(module.get("Id"))
            except Exception:
                continue
            permission_module = ACTIVE_MODULE_TO_PERMISSION_MODULE.get(active_id)
            title = _normalize_lookup_text(module.get("Title"))
            if permission_module and title:
                titles[permission_module] = title
    return titles


def permission_module_display(permission, module_title_map=None):
    try:
        module_id = int(permission.get("Module"))
    except Exception:
        return _normalize_lookup_text(permission.get("Module"))
    module_title_map = module_title_map or PERMISSION_MODULE_FALLBACK_NAMES
    return module_title_map.get(module_id, f"Модуль {module_id}")


def _normalize_lookup_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(value).strip())


def _is_cyrillic_char(ch):
    return "а" <= ch.lower() <= "я" or ch.lower() == "ё"


def _is_latin_char(ch):
    return "a" <= ch.lower() <= "z"


def _role_ui_text_sort_key(value):
    """Sort like the role UI in a Russian browser: Cyrillic names first, Latin codes like MDR after."""
    text = _normalize_lookup_text(value)
    first = ""
    for ch in text:
        if not ch.isspace():
            first = ch
            break
    if _is_cyrillic_char(first):
        alphabet_group = 0
    elif _is_latin_char(first):
        alphabet_group = 2
    else:
        alphabet_group = 1
    return (alphabet_group, text.replace("Ё", "Е").replace("ё", "е").casefold())


def _normalize_module_filter(value):
    return _normalize_lookup_text(value).lower().replace("ё", "е")


ROLE_UI_MODULE_ORDER = ["Документы", "Контроль Качества", "Настройки", "Отчеты", "Приёмка Работ"]


ROLE_UI_MODULE_ORDER_INDEX = {
    _normalize_module_filter(name): idx
    for idx, name in enumerate(ROLE_UI_MODULE_ORDER)
}


def _role_ui_module_sort_key(module_name):
    normalized = _normalize_module_filter(module_name)
    return (
        ROLE_UI_MODULE_ORDER_INDEX.get(normalized, len(ROLE_UI_MODULE_ORDER_INDEX)),
        _role_ui_text_sort_key(module_name),
    )


def is_role_permission_visible(permission):
    return _normalize_lookup_text(permission.get("Name")) not in ROLE_PERMISSION_EXCLUDED_NAMES


def get_permission_allowed_flags(permission_id):
    try:
        pid = int(permission_id)
    except Exception:
        pid = None
    return {
        "My": pid not in ROLE_PERMISSION_BLOCKED_MY_IDS,
        "MyOrganization": pid not in ROLE_PERMISSION_BLOCKED_ORG_IDS,
        "OtherOrganization": pid not in ROLE_PERMISSION_BLOCKED_ORG_IDS,
    }


def create_role(sess, access_token, base_url, payload):
    """Create Project Point role with RoleService/Create."""
    base = base_url.rstrip("/")
    return _post_json_to_variants(
        sess,
        access_token,
        [
            f"{base}/ru/api/Core/RoleService/Create",
            f"{base}/api/Core/RoleService/Create",
        ],
        payload,
        timeout=60,
    )


def update_role(sess, access_token, base_url, payload):
    sess = force_disable_ssl_verification(sess)
    """Update Project Point role through RoleService/Update.

    Verified against the supplied HAR: frontend sends POST
    /ru/api/Core/RoleService/Update with payload:
    {IsActive, Id, Name, Permissions[]}. Successful response is
    {"Status": 1, "ErrorText": null, "HistoryText": null}.
    """
    base = base_url.rstrip("/")
    urls = [
        f"{base}/ru/api/Core/RoleService/Update",
        f"{base}/api/Core/RoleService/Update",
    ]
    headers = api_headers(access_token, {"Content-Type": "application/json"})
    last_status = None
    last_data = None

    for url in urls:
        resp = sess.post(url, headers=headers, json=payload, timeout=60, verify=False)
        last_status = resp.status_code
        try:
            resp_data = resp.json()
        except Exception:
            resp_data = resp.text
        last_data = resp_data

        if resp.status_code in (200, 201, 204):
            if isinstance(resp_data, dict) and "Status" in resp_data:
                return resp_data.get("Status") == 1, resp.status_code, resp_data
            return True, resp.status_code, resp_data

        # If /ru endpoint exists but business validation failed, do not hide
        # the real server error behind the fallback endpoint.
        if resp.status_code not in (404, 405):
            return False, resp.status_code, resp_data

    return False, last_status, last_data


def _existing_role_permission_info_id(permission):
    raw_id = permission.get("PermissionsInfoId")
    if raw_id is None and isinstance(permission.get("PermissionsInfo"), dict):
        raw_id = permission["PermissionsInfo"].get("Id")
    if raw_id is None:
        return None
    try:
        return int(raw_id)
    except Exception:
        return None


def build_update_role_payload(existing_role, desired_payload, imported_permission_ids):
    """Build RoleService update payload that preserves permissions outside imported sheets."""
    existing_permissions = existing_role.get("Permissions") or []
    merged_permissions = []
    merged_by_info_id = {}

    for permission in existing_permissions:
        info_id = _existing_role_permission_info_id(permission)
        if info_id is None:
            continue
        item = {
            "Id": permission.get("Id"),
            "PermissionsInfoId": info_id,
            "My": bool(permission.get("My", False)),
            "MyOrganization": bool(permission.get("MyOrganization", False)),
            "OtherOrganization": bool(permission.get("OtherOrganization", False)),
        }
        merged_by_info_id[info_id] = item
        merged_permissions.append(item)

    for desired_permission in desired_payload.get("Permissions", []):
        info_id = _existing_role_permission_info_id(desired_permission)
        if info_id is None or info_id not in imported_permission_ids:
            continue
        target = merged_by_info_id.get(info_id)
        if target is None:
            target = {
                "Id": desired_permission.get("Id"),
                "PermissionsInfoId": info_id,
                "My": False,
                "MyOrganization": False,
                "OtherOrganization": False,
            }
            merged_by_info_id[info_id] = target
            merged_permissions.append(target)
        target["My"] = bool(desired_permission.get("My", False))
        target["MyOrganization"] = bool(desired_permission.get("MyOrganization", False))
        target["OtherOrganization"] = bool(desired_permission.get("OtherOrganization", False))

    return {
        "Id": existing_role.get("Id"),
        "Name": desired_payload.get("Name") or existing_role.get("Name"),
        "IsActive": bool(desired_payload.get("IsActive", existing_role.get("IsActive", True))),
        "Permissions": merged_permissions,
    }
