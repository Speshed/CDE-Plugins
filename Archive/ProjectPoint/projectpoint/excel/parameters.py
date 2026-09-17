from __future__ import annotations

import re

import pandas as pd

from ..api.custom_fields import COLUMN_MAP, DATA_TYPE_ALIASES, DATA_TYPE_MAP
from ..errors import SourceGapError
from .common import read_excel_table


def _norm_col(value):
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def _clean_scalar(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def _normalize_data_type(value):
    value = _clean_scalar(value).lower()
    value = DATA_TYPE_ALIASES.get(value, value)
    return value


def read_excel(file_path, sheet_name=0):
    """Read the generated/custom parameter workbook into importer records.

    The reader deliberately auto-detects the real header row, so both the
    generated template (title in row 1, header in row 2) and plain tables with
    headers in row 1 are accepted.
    """
    df = read_excel_table(
        file_path,
        sheet_name=sheet_name,
        required_headers=("Наименование атрибута", "Внутреннее имя атрибута", "Тип данных"),
    )

    rename_map = {}
    for col in df.columns:
        normalized = _norm_col(col)
        target = COLUMN_MAP.get(normalized)
        if target:
            rename_map[col] = target
    df = df.rename(columns=rename_map)

    required = {"Name", "SystemName", "DataType"}
    missing = sorted(required - set(df.columns))
    if missing:
        human = {
            "Name": "Наименование атрибута",
            "SystemName": "Внутреннее имя атрибута",
            "DataType": "Тип данных",
        }
        raise Exception(
            "В Excel не хватает обязательных колонок: "
            + ", ".join(human.get(name, name) for name in missing)
        )

    for optional in ("Description", "ListValuesRaw"):
        if optional not in df.columns:
            df[optional] = ""

    records = []
    for _, row in df.iterrows():
        record = {
            "Name": _clean_scalar(row.get("Name")),
            "SystemName": _clean_scalar(row.get("SystemName")),
            "Description": _clean_scalar(row.get("Description")),
            "DataType": _normalize_data_type(row.get("DataType")),
            "ListValuesRaw": _clean_scalar(row.get("ListValuesRaw")),
        }
        # Ignore fully empty trailing/template rows.
        if not any(record.values()):
            continue
        records.append(record)
    return records


def validate_row(row, row_number):
    errors = []
    prefix = f"Строка {row_number}"

    name = _clean_scalar(row.get("Name"))
    system_name = _clean_scalar(row.get("SystemName"))
    data_type = _normalize_data_type(row.get("DataType"))
    list_values = _clean_scalar(row.get("ListValuesRaw"))

    if not name:
        errors.append(f"{prefix}: не заполнено 'Наименование атрибута'")
    if not system_name:
        errors.append(f"{prefix}: не заполнено 'Внутреннее имя атрибута'")
    if not data_type:
        errors.append(f"{prefix}: не заполнено 'Тип данных'")
    elif data_type not in DATA_TYPE_MAP:
        allowed = ", ".join(DATA_TYPE_MAP.keys())
        errors.append(f"{prefix}: неизвестный тип данных '{row.get('DataType', '')}'. Допустимо: {allowed}")

    if data_type == "список":
        values = [v.strip() for v in re.split(r"[;\n]+", list_values) if v.strip()]
        if not values:
            errors.append(f"{prefix}: для типа 'Список' заполните 'Значения списка'")

    return errors


def map_excel_row_to_payload(row):
    """Map a validated row to CustomFieldService/Create payload.

    The exact write DTO, especially the list-values structure, is still not
    present in the supplied HAR/source captures. Parsing/validation is fully
    implemented, but the state-changing payload remains intentionally blocked
    until a real Create request is captured.
    """
    raise SourceGapError(
        "Формат Excel параметров распознан и проверен, но точный JSON DTO для "
        "CustomFieldService/Create отсутствует в предоставленных HAR. Нужен один "
        "реальный запрос создания параметра из браузера Project Point, чтобы не "
        "угадывать структуру изменяющего запроса."
    )


__all__ = ["read_excel", "validate_row", "map_excel_row_to_payload"]
