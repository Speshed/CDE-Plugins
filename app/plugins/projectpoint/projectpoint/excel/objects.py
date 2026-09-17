from __future__ import annotations

import re

from collections import defaultdict

import pandas as pd

from .common import read_excel_table

__all__ = ["_normalize_excel_col", "normalize_object_lookup_value", "read_objects_import_excel", "add_object_to_indexes", "build_object_indexes", "validate_and_order_objects_import_records", "resolve_object_parent_id"]


def _normalize_excel_col(name):
    return re.sub(r"\s+", " ", str(name or "").strip().lower())


def normalize_object_lookup_value(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def read_objects_import_excel(file_path, sheet_name):
    df = read_excel_table(file_path, sheet_name=sheet_name, required_headers=("Code", "Title"))
    col_map = {}
    for col in df.columns:
        normalized = _normalize_excel_col(col)
        if normalized == "code":
            col_map[col] = "Code"
        elif normalized == "title":
            col_map[col] = "Title"
        elif normalized == "parentcode":
            col_map[col] = "ParentCode"
        elif normalized == "parenttitle":
            col_map[col] = "ParentTitle"
    df = df.rename(columns=col_map)

    required_columns = {"Code", "Title"}
    if not required_columns.issubset(df.columns):
        missing = sorted(required_columns - set(df.columns))
        raise Exception(f"В Excel не хватает обязательных колонок: {', '.join(missing)}")

    records = []
    for idx, row in df.iterrows():
        code = normalize_object_lookup_value(row.get("Code"))
        title = normalize_object_lookup_value(row.get("Title"))
        parent_code = normalize_object_lookup_value(row.get("ParentCode")) if "ParentCode" in df.columns else ""
        parent_title = normalize_object_lookup_value(row.get("ParentTitle")) if "ParentTitle" in df.columns else ""
        records.append(
            {
                "RowNum": int(idx) + 2,
                "Code": code,
                "Title": title,
                "ParentCode": parent_code,
                "ParentTitle": parent_title,
            }
        )
    return records


def add_object_to_indexes(indexes, obj):
    obj_id = normalize_object_lookup_value(obj.get("Id"))
    if not obj_id:
        return
    code_key = normalize_object_lookup_value(obj.get("Code")).lower()
    title_key = normalize_object_lookup_value(obj.get("Title")).lower()
    if code_key:
        indexes["code_to_id"][code_key] = obj_id
    if title_key:
        indexes["title_to_objects"].setdefault(title_key, []).append(obj)


def build_object_indexes(objects):
    indexes = {"code_to_id": {}, "title_to_objects": {}}
    for obj in objects:
        add_object_to_indexes(indexes, obj)
    return indexes


def validate_and_order_objects_import_records(records, indexes, parent_mode, selected_parent_id):
    excel_parent_mode = parent_mode != "base"
    code_to_id = indexes.get("code_to_id", {})
    existing_codes = {normalize_object_lookup_value(code).lower() for code in code_to_id.keys() if normalize_object_lookup_value(code)}

    prepared_rows = []
    rows_by_code = defaultdict(list)
    errors = []

    for row in records:
        code = normalize_object_lookup_value(row.get("Code"))
        title = normalize_object_lookup_value(row.get("Title"))
        parent_code = normalize_object_lookup_value(row.get("ParentCode"))
        row_num = row.get("RowNum")
        code_key = code.lower()
        parent_key = parent_code.lower()

        prepared_rows.append(
            {
                "Row": row,
                "RowNum": row_num,
                "Code": code,
                "CodeKey": code_key,
                "ParentCode": parent_code,
                "ParentKey": parent_key,
            }
        )

        if code_key:
            rows_by_code[code_key].append(row)

        row_errors = []
        if not code:
            row_errors.append("Code должен быть заполнен")
        if not title:
            row_errors.append("Title должен быть заполнен")
        if excel_parent_mode and not parent_code:
            row_errors.append("ParentCode должен быть заполнен")
        if excel_parent_mode and code and parent_code and code_key == parent_key:
            row_errors.append("Code и ParentCode не должны совпадать")
        if code_key and code_key in existing_codes:
            row_errors.append(f"Code '{code}' уже существует в проекте")

        for error_text in row_errors:
            errors.append(f"Строка {row_num}: {error_text}")

    for code_key, rows in rows_by_code.items():
        if len(rows) <= 1:
            continue
        row_nums = ", ".join(str(r.get("RowNum")) for r in rows)
        code_value = normalize_object_lookup_value(rows[0].get("Code"))
        for row in rows:
            errors.append(f"Строка {row.get('RowNum')}: Code '{code_value}' дублируется в Excel (строки: {row_nums})")

    if excel_parent_mode:
        excel_codes = set(rows_by_code.keys())
        for item in prepared_rows:
            parent_code = item["ParentCode"]
            if not parent_code:
                continue
            parent_key = item["ParentKey"]
            if parent_key not in excel_codes and parent_key not in existing_codes:
                errors.append(f"Строка {item['RowNum']}: ParentCode '{parent_code}' не найден")
    elif not selected_parent_id:
        errors.append("Не выбран базовый родительский объект")

    if errors:
        return [], errors

    if not excel_parent_mode:
        return records, []

    row_by_code = {item["CodeKey"]: item["Row"] for item in prepared_rows if item["CodeKey"]}
    children_by_parent = defaultdict(list)
    roots = []
    for item in prepared_rows:
        parent_key = item["ParentKey"]
        if parent_key and parent_key in row_by_code:
            children_by_parent[parent_key].append(item["Row"])
        else:
            roots.append(item["Row"])

    ordered_rows = []
    visited = set()
    visiting = set()

    def visit(row):
        code = normalize_object_lookup_value(row.get("Code"))
        code_key = code.lower()
        if not code_key or code_key in visited:
            return
        if code_key in visiting:
            errors.append(f"Строка {row.get('RowNum')}: обнаружен цикл ParentCode")
            return
        visiting.add(code_key)
        ordered_rows.append(row)
        for child in children_by_parent.get(code_key, []):
            visit(child)
        visiting.remove(code_key)
        visited.add(code_key)

    for row in roots:
        visit(row)

    if len(visited) != len(prepared_rows):
        for item in prepared_rows:
            if item["CodeKey"] and item["CodeKey"] not in visited:
                visit(item["Row"])

    if errors:
        return [], errors

    return ordered_rows, []


def resolve_object_parent_id(row, indexes, parent_mode, selected_parent_id):
    code_to_id = indexes.get("code_to_id", {})
    parent_code = normalize_object_lookup_value(row.get("ParentCode"))

    if parent_mode != "base" and parent_code:
        parent_id = code_to_id.get(parent_code.lower())
        if not parent_id:
            return None, f"ParentCode '{parent_code}' не найден"
        return parent_id, None

    if parent_mode == "base":
        if selected_parent_id:
            return selected_parent_id, None
        return None, "Не выбран базовый родительский объект"
    return None, "ParentCode обязателен"
