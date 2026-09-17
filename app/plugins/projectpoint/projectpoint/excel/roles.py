from __future__ import annotations

import re

from collections import defaultdict

import pandas as pd

from ..api.roles import *
from .objects import _normalize_excel_col

__all__ = ["ROLE_IMPORT_ROLE_COLUMNS", "ROLE_IMPORT_PERMISSION_COLUMNS", "_rename_import_columns", "parse_import_bool", "_parse_permission_id", "_find_header_row", "_cell_has_user_value", "_find_label_value", "_find_roles_permission_header_row", "read_roles_import_excel", "build_role_payloads_from_import"]


ROLE_IMPORT_ROLE_COLUMNS = {
    "rolename": "RoleName",
    "role name": "RoleName",
    "role": "RoleName",
    "роль": "RoleName",
    "название роли": "RoleName",
    "имя роли": "RoleName",
    "isactive": "IsActive",
    "is active": "IsActive",
    "active": "IsActive",
    "активна": "IsActive",
    "активный": "IsActive",
    "включена": "IsActive",
    "comment": "Comment",
    "комментарий": "Comment",
}


ROLE_IMPORT_PERMISSION_COLUMNS = {
    "rolename": "RoleName",
    "role name": "RoleName",
    "role": "RoleName",
    "роль": "RoleName",
    "название роли": "RoleName",
    "имя роли": "RoleName",
    "permissioninfoid": "PermissionInfoId",
    "permissionsinfoid": "PermissionInfoId",
    "permission id": "PermissionInfoId",
    "permissionid": "PermissionInfoId",
    "id привилегии": "PermissionInfoId",
    "id права": "PermissionInfoId",
    "ид привилегии": "PermissionInfoId",
    "ид права": "PermissionInfoId",
    "permissionname": "PermissionName",
    "permission name": "PermissionName",
    "название привилегии": "PermissionName",
    "название права": "PermissionName",
    "привилегия": "PermissionName",
    "право": "PermissionName",
    "module": "Module",
    "модуль": "Module",
    "my": "My",
    "мои объекты": "My",
    "свои объекты": "My",
    "myorganization": "MyOrganization",
    "my organization": "MyOrganization",
    "my organisation": "MyOrganization",
    "объекты моей организации": "MyOrganization",
    "моя организация": "MyOrganization",
    "otherorganization": "OtherOrganization",
    "other organization": "OtherOrganization",
    "other organisation": "OtherOrganization",
    "объекты других организаций": "OtherOrganization",
    "другие организации": "OtherOrganization",
    "comment": "Comment",
    "комментарий": "Comment",
}


def _rename_import_columns(df, mapping):
    rename = {}
    for col in df.columns:
        normalized = _normalize_excel_col(col)
        if normalized in mapping:
            rename[col] = mapping[normalized]
    return df.rename(columns=rename)


def parse_import_bool(value, default=False):
    """Parse checkbox-like Excel values into bool."""
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except Exception:
        pass
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(int(value))
    s = str(value).strip().lower()
    if not s:
        return default
    truthy = {"1", "true", "yes", "y", "да", "д", "истина", "+", "✓", "✔", "x", "х"}
    falsy = {"0", "false", "no", "n", "нет", "н", "ложь", "-", "✗", "×"}
    if s in truthy:
        return True
    if s in falsy:
        return False
    return default


def _parse_permission_id(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    text_value = str(value).strip()
    if not text_value:
        return None
    try:
        return int(float(text_value))
    except Exception:
        raise Exception(f"Некорректный PermissionInfoId: '{value}'")


def _find_header_row(df, required_terms):
    for row_idx in range(min(len(df), 80)):
        values = [_normalize_excel_col(v) for v in list(df.iloc[row_idx].values)]
        if all(any(term in value for value in values) for term in required_terms):
            return row_idx
    return None


def _cell_has_user_value(value):
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except Exception:
        pass
    return str(value).strip() != ""


def _find_label_value(df, label_variants, default_col_offset=1, max_rows=12):
    labels = {_normalize_excel_col(label) for label in label_variants}
    for row_idx in range(min(len(df), max_rows)):
        row_values = list(df.iloc[row_idx].values)
        for col_idx, value in enumerate(row_values):
            if _normalize_excel_col(value) in labels:
                target_idx = col_idx + default_col_offset
                if target_idx < len(row_values):
                    return row_values[target_idx]
    return None


def _find_roles_permission_header_row(df):
    """Find header row for role permission matrices.

    Accepts both:
    - "Право" / "Название права";
    - "Привилегия" / "Название привилегии".
    """
    permission_terms = {
        "право",
        "название права",
        "права",
        "привилегия",
        "название привилегии",
        "permissionname",
        "permission name",
    }
    my_terms = {
        "свои элементы",
        "мои объекты",
        "my",
    }

    for row_idx in range(min(len(df), 80)):
        values = [_normalize_excel_col(v) for v in list(df.iloc[row_idx].values)]
        has_permission = any(value in permission_terms for value in values)
        has_my = any(value in my_terms for value in values)
        if has_permission and has_my:
            return row_idx
    return None


def read_roles_import_excel(file_path, include_sheets=None):
    """Read role import workbook.

    Supported layouts:
    1. Multi-sheet module layout:
       each module sheet has columns "Роль", "Привилегия", "Свои элементы",
       "Моей орг. единицы", "Другие орг. единицы".
       Module is resolved from sheet name, and selected sheets can be imported only.
    2. Flat table layout:
       columns "Роль", "Модуль", "Привилегия", "Свои элементы",
       "Моей орг. единицы", "Другие орг. единицы".
    3. Previous one-role layouts with a top field "Название роли".
    """
    xls = pd.ExcelFile(file_path)
    sheet_names = list(xls.sheet_names)
    # Release the workbook handle before pandas opens individual sheets. This
    # is required on Windows where an open ExcelFile locks the source file.
    xls.close()
    if not sheet_names:
        raise Exception("В Excel нет листов")

    global_role_name = ""
    global_active_raw = None
    roles_meta = {}
    permission_rows = []
    read_errors = []

    meta_sheet_names = {"настройка роли", "инструкция", "readme"}
    include_sheets_set = None
    if include_sheets is not None:
        include_sheets_set = {str(name) for name in include_sheets if str(name).strip()}


    def add_or_update_role(role_name, active_raw=None, row_num=None):
        role_name = _normalize_lookup_text(role_name)
        if not role_name:
            return
        if role_name not in roles_meta:
            roles_meta[role_name] = {
                "RoleName": role_name,
                "IsActive": parse_import_bool(active_raw, default=True),
                "Row": row_num,
            }
        elif _cell_has_user_value(active_raw):
            roles_meta[role_name]["IsActive"] = parse_import_bool(active_raw, default=True)

    for sheet_name in sheet_names:
        sheet_name_normalized_for_filter = _normalize_excel_col(sheet_name)
        if include_sheets_set is not None and sheet_name not in include_sheets_set and sheet_name_normalized_for_filter not in meta_sheet_names:
            continue
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        if df.empty:
            continue

        header_row = _find_roles_permission_header_row(df)
        headers = [_normalize_excel_col(v) for v in list(df.iloc[header_row].values)] if header_row is not None else []
        has_tabular_role_col = any(header in {"роль", "rolename", "role name", "название роли", "имя роли"} for header in headers)

        # In flat tables the A1 header is often "Роль" and B1 is "Модуль".
        # Do not interpret that as a global role label/value pair.
        if not has_tabular_role_col:
            if not global_role_name:
                global_role_name = _normalize_lookup_text(
                    _find_label_value(df, ["Название роли", "Роль", "RoleName", "Role Name"])
                )
            if global_active_raw is None:
                global_active_raw = _find_label_value(df, ["Активна", "Активен", "IsActive", "Is Active"])

            if global_role_name:
                add_or_update_role(global_role_name, global_active_raw, None)

        if header_row is None:
            continue

        def find_col(variants, required=True):
            normalized_variants = {_normalize_excel_col(v) for v in variants}
            for idx, header in enumerate(headers):
                if header in normalized_variants:
                    return idx
            if required:
                raise Exception(f"Лист '{sheet_name}': не найдена колонка: {variants[0]}")
            return None

        role_col = find_col(["Роль", "Название роли", "RoleName", "Role Name"], required=False)
        active_col = find_col(["Активна", "Активен", "IsActive", "Is Active"], required=False)
        module_col = find_col(["Модуль", "Module"], required=False)
        permission_col = find_col(["Право", "Привилегия", "Название права", "Название привилегии", "PermissionName", "Permission Name"])
        my_col = find_col(["Свои элементы", "Мои объекты", "My"])
        my_org_col = find_col(["Моей орг. единицы", "Моей орг. Единицы", "Моя организация", "Объекты моей организации", "MyOrganization", "My Organization"])
        other_org_col = find_col(["Другие орг. единицы", "Другие орг. Единицы", "Другие организации", "Объекты других организаций", "OtherOrganization", "Other Organization"])

        sheet_name_normalized = _normalize_excel_col(sheet_name)
        default_module_name = "" if sheet_name_normalized in meta_sheet_names else _normalize_lookup_text(sheet_name)
        current_role_name = global_role_name
        current_module_name = default_module_name

        for row_idx in range(header_row + 1, len(df)):
            row = df.iloc[row_idx]
            row_num = int(row_idx) + 1

            permission_name = _normalize_lookup_text(row.iloc[permission_col] if permission_col < len(row) else None)
            if not permission_name:
                continue

            if role_col is not None and role_col < len(row):
                role_value = _normalize_lookup_text(row.iloc[role_col])
                if role_value:
                    current_role_name = role_value

            role_name = current_role_name
            if not role_name:
                read_errors.append(
                    f"Лист '{sheet_name}', строка {row_num}: не заполнена роль. "
                    f"Заполните колонку 'Роль' или поле 'Название роли'."
                )
                continue

            if module_col is not None and module_col < len(row):
                module_value = _normalize_lookup_text(row.iloc[module_col])
                if module_value:
                    current_module_name = module_value

            module_name = current_module_name
            if not module_name:
                read_errors.append(
                    f"Лист '{sheet_name}', строка {row_num}: не заполнен модуль для права '{permission_name}'."
                )
                continue

            active_raw = None
            if active_col is not None and active_col < len(row):
                active_raw = row.iloc[active_col]
            if not _cell_has_user_value(active_raw):
                active_raw = global_active_raw
            add_or_update_role(role_name, active_raw, row_num)

            raw_values = {
                "My": row.iloc[my_col] if my_col < len(row) else None,
                "MyOrganization": row.iloc[my_org_col] if my_org_col < len(row) else None,
                "OtherOrganization": row.iloc[other_org_col] if other_org_col < len(row) else None,
            }

            permission_rows.append({
                "row_num": row_num,
                "sheet_name": sheet_name,
                "RoleName": role_name,
                "Module": module_name,
                "PermissionName": permission_name,
                "My": parse_import_bool(raw_values["My"], default=False),
                "MyOrganization": parse_import_bool(raw_values["MyOrganization"], default=False),
                "OtherOrganization": parse_import_bool(raw_values["OtherOrganization"], default=False),
                "_has_My": _cell_has_user_value(raw_values["My"]),
                "_has_MyOrganization": _cell_has_user_value(raw_values["MyOrganization"]),
                "_has_OtherOrganization": _cell_has_user_value(raw_values["OtherOrganization"]),
            })

    if read_errors:
        raise Exception("Ошибки чтения Excel:\n" + "\n".join(read_errors[:100]))

    if not roles_meta:
        raise Exception("Не заполнена колонка 'Роль' или поле 'Название роли' в книге Excel")

    if not permission_rows:
        raise Exception(
            "В Excel не найдено ни одной таблицы прав. Нужны колонки: "
            "'Роль', 'Модуль', 'Привилегия', 'Свои элементы', "
            "'Моей орг. единицы', 'Другие орг. единицы'."
        )

    return roles_meta, permission_rows


def build_role_payloads_from_import(roles_meta, permission_rows, all_permissions, active_modules=None):
    """Build RoleService/Create payloads from user-facing one-sheet Excel rows."""
    module_title_map = build_permission_module_title_map(active_modules)
    visible_permissions = [
        p for p in all_permissions
        if p.get("Id") is not None
        and is_role_permission_visible(p)
        and is_permission_module_visible(p, active_modules)
    ]

    permissions_by_key = {}
    permissions_by_name = {}
    duplicate_names = set()

    for perm in visible_permissions:
        perm_name = _normalize_lookup_text(perm.get("Name"))
        module_name = permission_module_display(perm, module_title_map)
        if not perm_name:
            continue
        key = (module_name.lower(), perm_name.lower())
        permissions_by_key[key] = perm
        name_key = perm_name.lower()
        if name_key in permissions_by_name:
            duplicate_names.add(name_key)
        else:
            permissions_by_name[name_key] = perm

    payloads_by_role = {}
    for role_name, meta in roles_meta.items():
        payloads_by_role[role_name] = {
            "Name": role_name,
            "IsActive": bool(meta.get("IsActive", True)),
            "Permissions": [
                {
                    "Id": None,
                    "PermissionsInfoId": int(perm["Id"]),
                    "My": False,
                    "MyOrganization": False,
                    "OtherOrganization": False,
                }
                for perm in visible_permissions
            ],
        }

    perm_index_by_role = {
        role_name: {p["PermissionsInfoId"]: p for p in payload["Permissions"]}
        for role_name, payload in payloads_by_role.items()
    }

    errors = []
    duplicate_tracker = set()
    duplicate_count = 0
    imported_permission_ids_by_role = {role_name: set() for role_name in payloads_by_role}

    for row in permission_rows:
        role_name = row["RoleName"]
        module_name = _normalize_lookup_text(row.get("Module"))
        perm_name = _normalize_lookup_text(row.get("PermissionName"))

        perm = None
        if module_name:
            perm = permissions_by_key.get((module_name.lower(), perm_name.lower()))
        if not perm:
            name_key = perm_name.lower()
            if name_key in duplicate_names:
                errors.append(
                    f"Строка {row['row_num']}: право '{perm_name}' встречается в нескольких модулях. "
                    f"Укажите корректный модуль."
                )
                continue
            perm = permissions_by_name.get(name_key)

        if not perm:
            module_part = f" в модуле '{module_name}'" if module_name else ""
            errors.append(f"Строка {row['row_num']}: право '{perm_name}'{module_part} не найдено в Project Point")
            continue

        perm_id = int(perm["Id"])
        allowed_flags = get_permission_allowed_flags(perm_id)
        for flag, label in ROLE_PERMISSION_COLUMNS.items():
            if not allowed_flags.get(flag, True) and row.get(f"_has_{flag}"):
                errors.append(
                    f"Строка {row['row_num']}: для права '{perm_name}' колонка '{label}' недоступна в интерфейсе. "
                    f"Очистите эту ячейку."
                )

        duplicate_key = (role_name.lower(), perm_id)
        if duplicate_key in duplicate_tracker:
            duplicate_count += 1
        duplicate_tracker.add(duplicate_key)
        if role_name in imported_permission_ids_by_role:
            imported_permission_ids_by_role[role_name].add(perm_id)

        role_perm_index = perm_index_by_role.get(role_name)
        if role_perm_index is None:
            errors.append(f"Строка {row['row_num']}: роль '{role_name}' не найдена")
            continue

        perm_payload = role_perm_index.get(perm_id)
        if perm_payload is None:
            errors.append(f"Строка {row['row_num']}: право {perm_id} отсутствует в payload роли")
            continue

        for flag in ROLE_PERMISSION_COLUMNS:
            if allowed_flags.get(flag, True):
                perm_payload[flag] = bool(row.get(flag))
            else:
                perm_payload[flag] = False

    if errors:
        raise Exception("Ошибки валидации Excel:\n" + "\n".join(errors[:100]))

    changed_counts = {}
    for role_name, payload in payloads_by_role.items():
        changed_counts[role_name] = sum(
            1 for perm in payload.get("Permissions", [])
            if perm.get("My") or perm.get("MyOrganization") or perm.get("OtherOrganization")
        )

    return list(payloads_by_role.values()), changed_counts, duplicate_count, imported_permission_ids_by_role
