from __future__ import annotations

import math
from typing import Any

import pandas as pd

from .common import read_excel_table

__all__ = [
    "read_project_stages_excel",
    "normalize_project_stage_row",
    "project_stage_payload_equal",
]

_HEADER_ALIASES = {
    "code": "Code",
    "код": "Code",
    "title": "Title",
    "наименование": "Title",
    "название": "Title",
    "displayparameters": "DisplayParameters",
    "отображать параметры": "DisplayParameters",
    "отображение параметров": "DisplayParameters",
    "isactive": "IsActive",
    "активен": "IsActive",
    "активный": "IsActive",
    "активна": "IsActive",
}


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value).strip()


def _normalize_header(value: Any) -> str:
    return " ".join(_clean_text(value).lower().replace("_", " ").split())


def _parse_bool(value: Any, *, default: bool) -> tuple[bool | None, str | None]:
    text = _clean_text(value)
    if not text:
        return default, None
    normalized = text.lower()
    if normalized in {"1", "true", "yes", "y", "да", "д", "активен", "активна", "вкл", "включен", "включено"}:
        return True, None
    if normalized in {"0", "false", "no", "n", "нет", "н", "неактивен", "неактивна", "выкл", "выключен", "выключено"}:
        return False, None
    return None, f"не удалось распознать логическое значение '{text}'"


def _parse_display_parameters(value: Any) -> tuple[int | None, str | None]:
    text = _clean_text(value)
    if not text:
        return 1, None
    normalized = text.lower()
    if normalized in {"да", "true", "yes", "вкл", "включено"}:
        return 1, None
    if normalized in {"нет", "false", "no", "выкл", "выключено"}:
        return 0, None
    try:
        number = int(float(text.replace(",", ".")))
    except ValueError:
        return None, f"DisplayParameters должен быть целым числом (обычно 0 или 1), получено '{text}'"
    if number < 0:
        return None, "DisplayParameters не может быть отрицательным"
    return number, None


def normalize_project_stage_row(row: dict[str, Any], row_number: int) -> tuple[dict[str, Any] | None, list[str]]:
    errors: list[str] = []
    code = _clean_text(row.get("Code"))
    title = _clean_text(row.get("Title"))

    if not code:
        errors.append(f"Строка {row_number}: не заполнен Code")
    if not title:
        errors.append(f"Строка {row_number}: не заполнен Title")

    display_parameters, display_error = _parse_display_parameters(row.get("DisplayParameters"))
    if display_error:
        errors.append(f"Строка {row_number}: {display_error}")

    is_active, active_error = _parse_bool(row.get("IsActive"), default=True)
    if active_error:
        errors.append(f"Строка {row_number}: IsActive — {active_error}")

    if errors:
        return None, errors

    return {
        "Code": code,
        "Title": title,
        "DisplayParameters": display_parameters,
        "IsActive": is_active,
    }, []


def read_project_stages_excel(file_path: str, sheet_name=0) -> list[dict[str, Any]]:
    df = read_excel_table(file_path, sheet_name=sheet_name, required_headers=("Code", "Title"))
    if df.empty:
        return []

    rename_map: dict[Any, str] = {}
    for column in df.columns:
        normalized = _normalize_header(column)
        canonical = _HEADER_ALIASES.get(normalized)
        if canonical:
            rename_map[column] = canonical
    df = df.rename(columns=rename_map)

    missing = [name for name in ("Code", "Title") if name not in df.columns]
    if missing:
        raise ValueError("В Excel отсутствуют обязательные колонки: " + ", ".join(missing))

    records: list[dict[str, Any]] = []
    all_errors: list[str] = []
    seen_codes: dict[str, int] = {}

    for excel_index, raw in df.iterrows():
        row_number = int(excel_index) + 2
        row = raw.to_dict()
        # Полностью пустые строки не считаем ошибкой.
        if not any(_clean_text(value) for value in row.values()):
            continue

        normalized, errors = normalize_project_stage_row(row, row_number)
        all_errors.extend(errors)
        if normalized is None:
            continue

        code_key = normalized["Code"].casefold()
        if code_key in seen_codes:
            all_errors.append(
                f"Строка {row_number}: Code '{normalized['Code']}' уже указан в строке {seen_codes[code_key]}"
            )
            continue
        seen_codes[code_key] = row_number
        records.append(normalized)

    if all_errors:
        raise ValueError("\n".join(all_errors))
    return records


def project_stage_payload_equal(existing: dict[str, Any], desired: dict[str, Any]) -> bool:
    return (
        _clean_text(existing.get("Code")) == _clean_text(desired.get("Code"))
        and _clean_text(existing.get("Title")) == _clean_text(desired.get("Title"))
        and int(existing.get("DisplayParameters") or 0) == int(desired.get("DisplayParameters") or 0)
        and bool(existing.get("IsActive")) == bool(desired.get("IsActive"))
    )
