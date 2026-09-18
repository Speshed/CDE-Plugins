# -*- coding: utf-8 -*-
"""
Larix User Manager — пользователи, проекты и ролевая матрица
Версия 4.0 — добавлена синхронизация прав папок Larix из Excel
"""
from __future__ import annotations
import sys
import json
import time
import base64
import os
import re
import traceback
import shutil
import hashlib
from io import BytesIO
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path
import requests
import pandas as pd
from datetime import datetime
from typing import Optional, List, Dict, Any, Callable

from PySide6 import QtWidgets, QtGui, QtCore
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PySide6.QtCore import Signal, QThread
from larix_ui_theme import (
    ACCENT_HOVER,
    STATUS_DANGER,
    STATUS_SUCCESS,
    ThemeToggle,
    build_qss,
    themed_icon,
    shared_asset_dir,
    initial_dark_theme,
    persist_dark_theme,
    apply_windows_titlebar_theme,
)
from shared.theme_core import shared_style_overrides, install_window_state
from shared.ui_components import add_standard_header_controls, install_status_bar, mark_destructive_buttons, make_preview_header
from shared.excel_style import ORANGE, DARK, LIGHT, BORDER, apply_excel_style

BASE_URL = "https://platform-api.larix.ru"
REQUEST_TIMEOUT = 30


# =====================================================================
#  Ролевая матрица Larix
# =====================================================================
# Права Larix независимы. Код в Excel означает точный набор прав:
# приложение НЕ добавляет «Просмотр» автоматически к скачиванию/загрузке.
MATRIX_ACCESS_CODE_TO_ID = {
    "ПР": 8,     # Просмотр
    "С": 9,      # Скачивание
    "З": 10,     # Загрузка/Создание
    "Пер": 11,   # Перемещение
    "У": 12,     # Удаление
    "ПД": 13,    # Полный доступ
}
MATRIX_ACCESS_ID_TO_CODE = {value: key for key, value in MATRIX_ACCESS_CODE_TO_ID.items()}
MATRIX_ACCESS_LABELS = {
    8: "Просмотр",
    9: "Скачивание",
    10: "Загрузка/Создание",
    11: "Перемещение",
    12: "Удаление",
    13: "Полный доступ",
}
MATRIX_ACCESS_IDS = tuple(sorted(MATRIX_ACCESS_ID_TO_CODE))


def _matrix_clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).replace("\xa0", " ").replace("\r", " ").replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return "" if text.lower() in {"nan", "none", "null"} else text


def _matrix_norm(value: Any) -> str:
    text = _matrix_clean_text(value).lower().replace("ё", "е")
    return re.sub(r"\s+", " ", text).strip()


def _matrix_norm_path(parts: List[str]) -> str:
    return "/".join(_matrix_norm(part) for part in parts if _matrix_clean_text(part))


def _matrix_display_path(parts: List[str]) -> str:
    return " / ".join(_matrix_clean_text(part) for part in parts if _matrix_clean_text(part))


def _matrix_codes_text(access_ids: set[int]) -> str:
    ordered = [MATRIX_ACCESS_ID_TO_CODE[item] for item in MATRIX_ACCESS_IDS if item in access_ids]
    return "; ".join(ordered) if ordered else "—"


def _parse_matrix_rights(value: Any) -> Dict[str, Any]:
    """Разбирает одну ячейку матрицы.

    Пусто -> не трогать.
    '-'   -> точный пустой набор (снять все шесть прав).
    Иначе допускаются комбинации: 'ПР; С; З', 'ПР,С', 'Пер / У'.
    """
    raw = _matrix_clean_text(value)
    if not raw:
        return {"touch": False, "access_ids": set(), "error": ""}
    if raw in {"-", "—", "–"}:
        return {"touch": True, "access_ids": set(), "error": ""}

    tokens = [
        _matrix_clean_text(part)
        for part in re.split(r"[,;/\\|\n\r]+", raw)
        if _matrix_clean_text(part)
    ]
    access_ids: set[int] = set()
    unknown: List[str] = []
    for token in tokens:
        norm = _matrix_norm(token).upper().replace(" ", "")
        aliases = {
            "ПР": "ПР", "ПРОСМОТР": "ПР", "VIEW": "ПР",
            "С": "С", "СКАЧИВАНИЕ": "С", "DOWNLOAD": "С",
            "З": "З", "ЗАГРУЗКА": "З", "ЗАГРУЗКА/СОЗДАНИЕ": "З", "UPLOAD": "З", "CREATE": "З",
            "ПЕР": "Пер", "ПЕРЕМЕЩЕНИЕ": "Пер", "MOVE": "Пер",
            "У": "У", "УДАЛЕНИЕ": "У", "DELETE": "У",
            "ПД": "ПД", "ПОЛНЫЙДОСТУП": "ПД", "FULLACCESS": "ПД",
        }
        code = aliases.get(norm)
        if not code:
            unknown.append(token)
            continue
        access_ids.add(MATRIX_ACCESS_CODE_TO_ID[code])
    if unknown:
        return {
            "touch": True,
            "access_ids": access_ids,
            "error": "Неизвестные коды: " + ", ".join(unknown),
        }
    if not access_ids:
        return {"touch": True, "access_ids": set(), "error": f"Не удалось разобрать значение '{raw}'"}
    return {"touch": True, "access_ids": access_ids, "error": ""}


def _parse_role_matrix_excel(excel_path: str, sheet_name: str = "") -> Dict[str, Any]:
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel-файл не найден: {excel_path}")
    book = pd.ExcelFile(excel_path)
    sheets = [str(item) for item in book.sheet_names]
    target = sheet_name if sheet_name in sheets else ("2. Ролевая матрица" if "2. Ролевая матрица" in sheets else (sheets[0] if sheets else ""))
    if not target:
        raise RuntimeError("В Excel нет листов")
    raw_df = pd.read_excel(book, sheet_name=target, header=None, dtype=object)
    rows = raw_df.values.tolist()

    header_index = None
    level_columns: List[tuple[int, int]] = []
    for row_idx, row in enumerate(rows[:40]):
        found: List[tuple[int, int]] = []
        for col_idx, value in enumerate(row):
            match = re.fullmatch(r"Уровень\s*(\d+)", _matrix_clean_text(value), flags=re.IGNORECASE)
            if match:
                found.append((int(match.group(1)), col_idx))
        if len(found) >= 2:
            header_index = row_idx
            level_columns = sorted(found, key=lambda item: item[0])
            break
    if header_index is None or not level_columns:
        raise RuntimeError("Не найдена строка заголовков с колонками 'Уровень N'")

    header = rows[header_index]
    last_level_col = max(col_idx for _level, col_idx in level_columns)
    principal_columns: List[tuple[int, str]] = []
    for col_idx in range(last_level_col + 1, len(header)):
        name = _matrix_clean_text(header[col_idx])
        if not name:
            continue
        norm = _matrix_norm(name)
        if norm.startswith("легенда"):
            continue
        # Колонки легенды справа не считаем получателями.
        if re.match(r"^(пр|с|з|пер|у|пд)\s*[-—–]", norm, flags=re.IGNORECASE):
            continue
        principal_columns.append((col_idx, name))
    if not principal_columns:
        raise RuntimeError("После колонок 'Уровень N' не найдены роли/пользователи")

    path_stack: List[Optional[str]] = [None] * len(level_columns)
    entries: List[Dict[str, Any]] = []
    parse_errors: List[str] = []
    for row_idx in range(header_index + 1, len(rows)):
        row = rows[row_idx]
        changed_path = False
        for stack_idx, (_level, col_idx) in enumerate(level_columns):
            value = _matrix_clean_text(row[col_idx] if col_idx < len(row) else None)
            if value:
                path_stack[stack_idx] = value
                changed_path = True
                for clear_idx in range(stack_idx + 1, len(path_stack)):
                    path_stack[clear_idx] = None
        path_parts = [part for part in path_stack if part]
        if not path_parts:
            continue

        commands: Dict[str, set[int]] = {}
        row_errors: List[str] = []
        for col_idx, principal_name in principal_columns:
            parsed = _parse_matrix_rights(row[col_idx] if col_idx < len(row) else None)
            if not parsed["touch"]:
                continue
            if parsed["error"]:
                row_errors.append(f"{principal_name}: {parsed['error']}")
            commands[principal_name] = set(parsed["access_ids"])
        if commands or row_errors:
            entry = {
                "row_number": row_idx + 1,
                "path_parts": list(path_parts),
                "path_text": _matrix_display_path(path_parts),
                "commands": commands,
                "errors": row_errors,
            }
            entries.append(entry)
            parse_errors.extend(f"Строка {row_idx + 1}: {error}" for error in row_errors)

    if not entries:
        raise RuntimeError("В Excel не найдено ни одной команды ролевой матрицы")
    return {
        "sheet_name": target,
        "entries": entries,
        "principal_names": [name for _col, name in principal_columns],
        "parse_errors": parse_errors,
    }


def _build_matrix_folder_index(tree_payload: Any) -> Dict[str, Dict[str, Any]]:
    roots = tree_payload
    if isinstance(tree_payload, dict):
        roots = tree_payload.get("data", tree_payload)
    if isinstance(roots, dict):
        roots = [roots]
    if not isinstance(roots, list):
        return {}

    result: Dict[str, Dict[str, Any]] = {}

    def walk(node: Dict[str, Any], parent_parts: List[str], skip_title: bool = False) -> None:
        if not isinstance(node, dict):
            return
        title = _matrix_clean_text(node.get("title") or node.get("name"))
        current = list(parent_parts)
        if title and not skip_title:
            current.append(title)
            key = _matrix_norm_path(current)
            if key and key not in result:
                result[key] = {
                    "id": _to_int(node.get("id") or node.get("folderId")),
                    "title": title,
                    "path_parts": current,
                    "path_text": _matrix_display_path(current),
                    "access": node.get("access") if isinstance(node.get("access"), list) else [],
                }
        for child in node.get("children") or []:
            if isinstance(child, dict):
                walk(child, current, False)

    for root in roots:
        if not isinstance(root, dict):
            continue
        # Ответ /projectRoles/folders/{projectId} содержит корень проекта.
        # Его заголовок не присутствует в Excel и поэтому исключается из пути.
        walk(root, [], True)
    return result


def _matrix_principal_access(folder: Dict[str, Any], principal_type: str, principal_id: int) -> set[int]:
    found: set[int] = set()
    for access in folder.get("access") or []:
        if not isinstance(access, dict):
            continue
        access_id = _to_int(access.get("ac_role_id") or access.get("accessRoleId") or access.get("id"))
        if access_id not in MATRIX_ACCESS_IDS:
            continue
        if principal_type == "role":
            ids = {
                _to_int(item.get("id") or item.get("roleId"))
                for item in (access.get("project_roles") or access.get("projectRoles") or [])
                if isinstance(item, dict)
            }
        else:
            ids = {_to_int(item) for item in (access.get("ac_user_ids") or access.get("userIds") or [])}
            for item in access.get("users") or []:
                if isinstance(item, dict):
                    ids.add(_to_int(item.get("id") or item.get("userId")))
        ids.discard(None)
        if int(principal_id) in ids:
            found.add(int(access_id))
    return found


def _matrix_access_buckets(folder: Dict[str, Any]) -> Dict[int, Dict[str, set[int]]]:
    result = {access_id: {"users": set(), "roles": set()} for access_id in MATRIX_ACCESS_IDS}
    for access in folder.get("access") or []:
        if not isinstance(access, dict):
            continue
        access_id = _to_int(access.get("ac_role_id") or access.get("accessRoleId") or access.get("id"))
        if access_id not in result:
            continue
        users = {_to_int(item) for item in (access.get("ac_user_ids") or access.get("userIds") or [])}
        for item in access.get("users") or []:
            if isinstance(item, dict):
                users.add(_to_int(item.get("id") or item.get("userId")))
        roles = {
            _to_int(item.get("id") or item.get("roleId"))
            for item in (access.get("project_roles") or access.get("projectRoles") or [])
            if isinstance(item, dict)
        }
        result[access_id]["users"] = {int(item) for item in users if item is not None}
        result[access_id]["roles"] = {int(item) for item in roles if item is not None}
    return result


def _matrix_build_principal_catalog(payload: Any) -> Dict[str, List[Dict[str, Any]]]:
    data = payload.get("data", payload) if isinstance(payload, dict) else {}
    roles = data.get("roles") if isinstance(data, dict) else []
    users = data.get("users") if isinstance(data, dict) else []
    catalog: Dict[str, List[Dict[str, Any]]] = {}

    def register(key: Any, item: Dict[str, Any]) -> None:
        norm = _matrix_norm(key)
        if norm:
            catalog.setdefault(norm, []).append(item)

    for role in roles or []:
        if not isinstance(role, dict):
            continue
        role_id = _to_int(role.get("id") or role.get("roleId"))
        title = _matrix_clean_text(role.get("title") or role.get("name"))
        if role_id and title:
            item = {"type": "role", "id": role_id, "name": title}
            register(title, item)
    for user in users or []:
        if not isinstance(user, dict):
            continue
        user_id = _to_int(user.get("id") or user.get("userId") or user.get("ac_user_id"))
        name = _matrix_clean_text(user.get("name") or user.get("fullName") or user.get("full_name"))
        email = _matrix_clean_text(user.get("email") or user.get("userEmail"))
        display = name or email
        if user_id and display:
            item = {"type": "user", "id": user_id, "name": display}
            register(display, item)
            if email:
                register(email, item)
    return catalog


def _matrix_resolve_principal(catalog: Dict[str, List[Dict[str, Any]]], requested_name: str) -> Dict[str, Any]:
    matches = catalog.get(_matrix_norm(requested_name), [])
    unique = {(item["type"], item["id"]): item for item in matches}
    matches = list(unique.values())
    if len(matches) == 1:
        return {"ok": True, "principal": matches[0], "message": "точное совпадение"}
    if len(matches) > 1:
        variants = ", ".join(f"{item['name']} ({item['type']})" for item in matches)
        return {"ok": False, "principal": None, "message": f"Неоднозначное имя: {variants}"}
    return {"ok": False, "principal": None, "message": "Роль/пользователь не найден в выбранном проекте"}

# =====================================================================
#  ВСТРОЕННЫЕ EXCEL-ШАБЛОНЫ
#
#  Оба утверждённых .xlsx-файла сохранены ниже как Base64. Поэтому рядом
#  со скриптом или собранным exe не требуется папка templates и не нужны
#  отдельные файлы. При скачивании программа восстанавливает точную копию
#  исходного Excel-файла, включая листы, оформление и данные.
# =====================================================================
APP_FOLDER_NAME = "LarixUserManager"
SETTINGS_FILE_NAME = "settings.json"
STRUCTURE_TEMPLATE_NAME = "Larix Platform Папки и проект.xlsx"
USER_IMPORT_TEMPLATE_NAME = "Larix Platform Импорт пользователей.xlsx"
ROLE_MATRIX_TEMPLATE_NAME = "Larix Platform Ролевая матрица.xlsx"
STRUCTURE_SHEET_NAME = "1. Папочная структура"

# SHA-256 исходных файлов используется для проверки целостности после декодирования.
STRUCTURE_TEMPLATE_SHA256 = "56214f2e595a15a30a0cc10a9978c5c8c2e528e8ced81c18306645bce36266b3"
USER_IMPORT_TEMPLATE_SHA256 = "b66715ba8fb602d48d5fb6a4c1d4bd8c517fc21a4c6b2ca8ca2f9a48a971bb09"
ROLE_MATRIX_TEMPLATE_SHA256 = "f16cc10b32fa5cc30ddd5f929caf91ab9d2a557e6dea61bc3a0d52656ebc64c2"
REMARK_TYPE_TEMPLATE_NAME = "Larix Platform Типы замечаний.xlsx"
TASK_TYPE_TEMPLATE_NAME = "Larix Platform Типы задач.xlsx"
REMARK_TYPE_TEMPLATE_SHA256 = "bc53ced4e2c0cc2875b519312409e5c7f7ce33e69ad9a75d625a3a27856c7a5d"
TASK_TYPE_TEMPLATE_SHA256 = "4d2ab55c2eb4c39645d85a2e6c64ae3efa2b5cf220945b4d3abfc552db09740e"
REMARK_TYPE_TEMPLATE_B64 = (
    'UEsDBBQAAAAIAKiIGl0KwygzHQEAAGwCAAAPAAAAeGwvd29ya2Jvb2sueG1svdIxasMwFIDhqwjtjRzHMbaJk6VL195AkqVYxJKM' +
    'pLQeQ9f2DoWeIBRCWwrJFZ5vVJKUpJClS7s9/uHxwXuTWacbdCecV9aUeDiIMBKG20qZeYmXQV5leDaddMW9dQtm7QJ1ujG+6Epc' +
    'h9AWhHheC039wLbCdLqR1mka/MC6OfGtE7TytRBBNySOopRoqgze7ztUf5qQoVqUGF7gHXb9I0aHelOVeIiRK1RV4tuM5amkcZ6P' +
    'KEvGnOJvi/uNxUqpuLi2fKmFCUeMEw0Nyhpfq9ZjRC40z/0K1vAKa9SvYAufsIGPH7T4ROOSMZZkXCRJmqQy/08a7Pa0/gneYLsP' +
    '/QNsLqijE7WSo7GMYypYlSWRjP6ASs7nJefPmX4BUEsDBBQAAAAIAKiIGl0MwPjBoAIAAB4ZAAANAAAAeGwvc3R5bGVzLnhtbOVZ' +
    'TW/iMBD9K5bv25BQCFRNqxZqaS+9tIe9huAES2M7ckw39NevbIeQftCStlmg5ZKZkef5OX4eR8P5ZckBPVBVMCki7J/0MKIikXMm' +
    'sggvdfprhC8vzsuzQq+A3i0o1ajkIIqzMsILrfMzzyuSBeVxcSJzKkoOqVQ81sWJVJlX5IrG88KkcfCCXm/o8ZgJbBBTKXSBErkU' +
    'OsKjOmQne0QPMUTY9zHyTEDEnLrQJFbAtLRxb5Oxfs7c+BpgWAEkEqRCKptFmFS/ttDsGXTvNWifBON++FnW/qvQYX84mLaFfgez' +
    'Q7rdQX/N/n2adWUUdhIGUIs5dGJmAOaZx1pTJQgDQJV9v8pphIUUtEasBr+blKl45QeD1nmFBDZ3vLLJFll5T/K/CJ8MyIhcd4c/' +
    'Hd9ckbBD/oQEk0l3+DcB6U3Hb+NXhlXaTKo5VbXWArwJOtU621hAU41s4Y6wXlRl94ncr8PJeOpqpOfGmyGKZYtWiTbBjNEyb5On' +
    'Ze4Yay15m0SXUZluubVp31JCAe4M3p+0flW+RS1TJJaccP17HuEeRuYIr00GUJkOqnLcnE3I9RQN9GD0Ufgy3cyzE0CwDSDOc1jd' +
    'LvmMKmJvYbNoFyVSND0GsPGuLZj136Tg/xQK1r8ClglON8qJ1wG0kIo9SqHN1WAOzVoiZXqM3B+o0iwxfkKFpmq31TT02N+TGIKf' +
    'QuG9Pf2r4vyelg5ql807DNovhdd6IQ0Vnu5JAv3DpeB/NwrdFOZD5f6xwny6w5dGx2LYRsH/bhTaV7iWRW2wpx0c/BQKX361Dvf/' +
    'kfx/KLSpZ23Uf7jsP12Ph8dWiU63nMEj494/Nu7h/stveMjlt2rHNDoxtjPzrNVTx5Hp5Eb41nCEpw2XZmOnsO7mn46Lf1BLAwQU' +
    'AAAACACoiBpd+lwBWQMDAADaDQAAEwAAAHhsL3RoZW1lL3RoZW1lMS54bWy9V9tymzAU/BVG7w03c/OEZBLHbh/SaafJD8ggQI0Q' +
    'HkmOnb/vIG4CjOM0duwHS2LP2UXnsMLXt/ucaK+IcVzQEJhXBtAQjYoY0zQEW5F888HtzTWciwzlSKMwRyFYZFB8//0MtH1OKJ/D' +
    'EGRCbOa6zqMM5ZBfFRtE9zlJCpZDwa8KluoxgztM05zolmG4eg4xBW3eJUE5ooKXCxFhT9EBsvJa/GKWP/yNLwjTXiEJwQ7TuNg9' +
    'o70AGoFcLAgLgSE/QNNvrvU2ioiJYCVwJT9NYB0Rv1gykKXrNtJYWv7M7BgkgogxcOmX3y6jRMAoQrSWo4JNxzV8qwErqGp4IHvg' +
    'mfYgQGGwxwyBe2/N+gESVQ1n4xtdBcsHpx8gUdXQGQXcGdZ9YPcDJKoauqOA2fLOs5b9AInKCKYvY7jr+b7bwFtMUpAfB/GB6xre' +
    'Q4PvYLrSalUCKnqN9ytJcIRk3+Xwb8FWBRWyylBgqom3DUpgVDYoJHjNsPaI00xIHjhH8B1AxI8C9AFnjum7Ao5QHyFt6ToGXd0M' +
    'uTW5mHwkE0zIk3gj6JFLcbwgOF5hQuRERrWl2GQLwhrCHjBlsBvzOlXKtU3BQ2CAyVzSQTAV1ZrrNU89nJNt/rOI66Y3WzuAcw5F' +
    'd8FwFJ9oGeQs5aqGEneyDs+e0NHRDXXYJ+qQd3KyEN/8sJDgqBBdKQ/BVIPlKeHMarvlESQoLgtWJ+iV9SwlDmZTd2R9dmtPKDHP' +
    'YIyavMaUkqlm67rwDEVWpHj+YSVBMCGk3KpLFFkf2wGh/Zm2K/m95u7+yyw2jIsHyLMKJy+15ytVaALD+QIaq9yZy9Howz1ESYIi' +
    'MbHSTR+5qLMcvPxZdDkptgKxpyzeaWuyZX9gHALHMx0DaDHmoimAFmPWtc/4/aJbh2STwdrJew9thZfjllMRK+UMpffnteJ1ujrL' +
    'cfV+1MC1puzWm34SL3A+Bsq5pPhH4H/UUyurPPexqepQ5U0arT0hz76Q0XZd+XWGOmzZ0mOb1zE5G/yBalZu/gFQSwMEFAAAAAgA' +
    'qIgaXQ0euehlAAAAcwAAABQAAAB4bC9zaGFyZWRTdHJpbmdzLnhtbAXBUQrDIAwA0KtI/mfcPsaQ2p5F2rQKJhaTDY+/95ZtcnM/' +
    'Glq7JHj6AI5k70eVK8HXzscHtnWZUdXc5CYaZ4JidkdE3QtxVt9vksnt7IOzqe/jQr0H5UMLkXHDVwhv5FwFHK5/UEsDBBQAAAAI' +
    'AKiIGl2+JF43PQQAAO0MAAAYAAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1slVfLbttGFP2VAfcRJcqyZdVy0ETuA2iBoot2zUgj' +
    'iYgoCiRtaalHUy8cwEARoEBa94FuuqQVMaYfon/hzC/0S4o7pGhbHjr2SpzHPffec8/cGe08H9k9dsBdz3L6da1UKGqM95tOy+p3' +
    '6tq+335W1Z7v7oxqQ8d97XU599nI7vW92qiudX1/UNN1r9nltukVnAHvj+xe23Ft0/cKjtvRvYHLzZY0s3u6USxu6rZp9TUClLM/' +
    'WHzo3Rkxr+sMv3St1jdWn3t1ragxcv3KcV7T8tctOaXv7uhKiC+k9+9c1uJtc7/nf+8Mv+JWp+vXtVJF2o1qTacnDZpOj9kWZa0x' +
    '2xzJ36HV8rt1rbyhsa7VavG+dNfc93zH/jFZK93AJOZGam5k5hXjCebl1Lx84/0J1hup9UZmXXpK7JXUvJKZG5VPmes3FErOG6Zv' +
    '0sB1hsyVm4jucmacFUDWvUl7Pi9pzJNB+3XN8125crCLfxHgFJeIsWRiigjXiDFnOEOAK4TiEAGWiHBOQRwkoWSgLzJQPZt7qZhr' +
    'KOb2bs/pMpdbKVE5/bpmVB9KyZAI1fWU/sACSwQMS4S4FjMxEVME4pjRhxgjxgUC9t/4HUOMBaIs8wLDCQKcYZ4mHdKOU3FMbIgp' +
    'QlyKt1gi/oxhIWaSt4jAxFQcMWKN8JaIccXETwhwjkvCoDgWiLNYIlyJo4KS0CylW4Qq5hqKub3bc/cIraSEbjxEaEUilMrrjK6z' +
    'kgolUKaQB/KnGFPmCMWY6Yx0Jo7FRJIjUVVgjQTMMNbBfkOID9JyQZW+kMxL4V6LMQLMVWh7N2gqijYfQ9Fmkt3mekDvKTviRRxK' +
    'BYT3j5Aywxd5gO8Q4oqOZaraQPyMiMakbaWYU3/ki8RKuguYmGCOGBEtk+FhuuFYSfdmTu3eE79KSvMsfr3tbS31e8xvpcwbDzG/' +
    'lUPUCUIxQYxYTDFHKKbyjM2T40tqiBHiQkzFTMl/HuzviPGRSsmksk5JVLgUx2Iq3hK1MebiiM7CTR0imo/wMUfLiSfj/sHAX0pq' +
    'c/eTOzEhdZDrh6mtPobaag4Hv8imJb0xXCQKu6BeIA6p8eUwmod2QgWJxEReMdRJz8QM51KqYowQQdJeJau4pj4h67q8u7CQp0kO' +
    'lCwn3ukQ3/X+t5Lj3N2IcZb6+qR8tx/BcWM7v/o5sW0/rv4MEZ3yJ8RbKj5GFLSL6rh9XxVBcjLUT4I8s39keJer5rTI72+KcyXl' +
    'QNdoIsIPYixmOJN9/5whIlhFA4zEm6wB4op4Si8uGigv4EYaPr3J7oT/TFmivN2yKU3Vl9GqHKvn9OoxZ3O3w1/yXvLOy0bM5W2q' +
    'Ro0eTPK2X18yanT1q5YalRrdeclldxd/YHb4t6bbsfoe6/G2X9eKhS2NuYkO5LfvDORXRWOvHN937NWoy80Wd2lU1ljbcfxskHjK' +
    '/rns/g9QSwMEFAAAAAgAqIgaXZFztYtfBAAA2w4AABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0Mi54bWytV91uGkcUfpXRXNssLAZj' +
    'ZIga27SREqnqRXq9YQdYZX/o7mLoHZALXzhSpDZSqlSRnTfYErbGxKxf4cwr9EmqM7tefjzr4io3wJw535n5vnNm5nD4ZGCZ5JS5' +
    'nuHYNVrI5SlhdtPRDbtdoz2/tVuhT+qHg2rfcV97HcZ8MrBM26sOarTj+92qonjNDrM0L+d0mT2wzJbjWprv5Ry3rXhdl2m6gFmm' +
    'oubzZcXSDJtiQGF9abC+tzYiXsfpf+8a+nPDZl6N5inBpV85zmucfqYLk1I/VKQhGmL1H12is5bWM/2fnP4PzGh3/BotlARuUG06' +
    'pgA0HZNYBrKmxNIG4rtv6H6nRot7lHQMXWe2WK7Z83zH+jmeKyzDxHA1gaspXN1/BLyYwIspvPCY1fcS+N5y849AlxJ06f8tXk7g' +
    '5SX10n/BlWUGRMqONV/Dgev0iSucMFvFFJzmT5RNE32+K1DiCcp+jXq+K2ZO63DBhxDABALChxDBVwjhmvwzfE/4GGZwy88JXEEA' +
    'NxDyMwhgATO4xv2cxrtK4z9N4yup7UhiO5bYTiS2xqpNEVRXGKsxY7XyEGNVRKhsMr5Emvwtgang+zcsICAwQcIRP4MFRHyE5Akf' +
    'QQQTuIUAphDwMX9L+IigO1yhYkKMEG4I3ArtQpjzMeLh+k7MGXmuucYgR+A9fOXvCEQwFR5fIEoUxsVnS3+UPvEKCG5EzMwhyElV' +
    'TzmuqC6xHUtsJxJbY9V2T/VSovreQ6qXRIRCcVP2T+uypeyltLKCJLmTYY6yMGmFy1AnMUpVN1F/QghfIIQF5p7AXOQkgkmc7AAm' +
    'smiNZTSZfuVt9CvHNMqbG/rIhzBD0eIahfD+wQylWmYEfPrsxa6gFfEhTGGGBSeOwFCqbta2LuASPkuVLWfk4yNKKVUvC/FB7O1M' +
    'ZGOT5T2R9xOR1YdE3v/WImcF/B3+gAupoFmIC7iUyhn7q8Ut/RuZ/nih8BHcQISXy8NiVrYRs5LB5BOE4gqN+BgmEPKxuFcn4vSv' +
    'Xpn8jVTSyjep26zNyWs29sYDuoV3I9MbIrjCe2Obej3YRuKDDBa/QRQriQ/TPC7WOV60/AxmMM9QNivah3W0VM8M7K5UzoPsms2Q' +
    '9GC7qhVP5ugxMhfyW+h8gl64fmkbeo0sb1H5Y/lLcX9ja01a3PVgM7ypWARTPorTg5/yHiwLjpm9ie9PPr47fjNxrw2xxRPNxhBC' +
    'CPB5hQU/R6e1Toe/W+0QYcLP4S/Bb5E2NGt9UJC0PtIqWtnoykOpbLS3FnPb7IiZceebjojLWqhVFXtERTalVrGRkU2dlKr4SMeL' +
    'rsfXNV97qZmGrvmGY3uk6fTstDzWJ4n/a5fVqGl4PiXeLyLyUbl6VMjnhTf+oeuZWqFOsdR34PNOXPI7uxTXTadxsB5ZYhKb62pt' +
    '9kJz24btEZO1/BrN5/YpceMSFr99pyt+lSh55fi+Y92NOkzTmYujIiUtx/HTQSxD+j+1/i9QSwMEFAAAAAgAqIgaXWWbGPyPBAAA' +
    'Ug8AABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0My54bWytV91uGkcUfpXRXtssLAbbyDhtMbSVUqnqRXq9YQdYZZehu4OhdxhLsapE' +
    'slRFipQqcvIGyGFr/MP6Fc68Qp+kOjObBZtZG1e9sXfOnO/8fOfMzGHv2dD3yCENQpd1q0YhlzcI7TaZ43bbVaPPW5s7xrP9vWFl' +
    'wIJXYYdSToa+1w0rw6rR4bxXMc2w2aG+HeZYj3aHvtdigW/zMMeCthn2Amo7EuZ7ppXPl03fdrsGGpTSFy4dhHdWJOywwfeB6zx3' +
    'uzSsGnmDoOuXjL3C7R8dKTL390ytiYb0/nNAHNqy+x7/hQ1+oG67w6tGoSRxw0qTeRLQZB7xXczaIL49lP8HrsM7VaO4ZZCO6zi0' +
    'K901+yFn/q9qr7Awo+BWArcWcOsJ8GICL6bwwlO8byXwrYX3J6BLCbr035yXE3g5hVulx+DmogKyZAc2t3ERsAEJpBJWq5iC0/rJ' +
    'tmmizrcFg4QyZV41Qh7IncN9OBMjmMA5TAjcQgzX4i1cQIwCMYYIriGCS/LP6B0RY5jBrXhD4AImcAOROIEJzGEGlxjfoYoy9fdd' +
    '6s9MZTWN7EAjq2tkjWWZKVNfYgCbh1cNa+chBixpYeceA/VN33Y9AlPMHv6GCOYEYpjCHGK4gDlMxAl+E3EEMXIkjsRYUnYtTsVY' +
    'vBVH4pSIo2z+bpC7CK7EsfgDIvgCMYFbMVKWJP1zaROL8NwO3GGOwDu0vogDMaoAWKhZhi9xKiuVoCYksR/DFUxy2iKllCwVSSM7' +
    '0MjqGlljWbZSpFJSpK2HilSSFgrF+336ESZwganKlotSNrRpZRj5WuoM8nSmalnxpOdGh6orlGXdR/0l6x/BHKZYyCtZqhjOVUNM' +
    '4FxnrbGwpqO1vA6tZZVG+X5AH8QIZsilanKIVo93pKU4w6BDfZbrhzT4hg5tv+fRXJP5WmKzIjqDT/BZS2o5oxQfkEUtcVmI98mx' +
    'jjQJrvC7nfBrPcTv9v/Nb4bBgB66dLAGvVkBncEnLblK3yquqd/I1FcXG9xAjJfPw9TurEPtTkYmHyGSV3Isr85IjJNLFG8Hdb/K' +
    'S3csjrUEZ5i1+7zDHqc3Kyh95yptPKFraDcyteWDNNW2zQq1u+tQu5uRxZ/40qlCErhSLXuFF7A4gRk+ZFpGM6y1+zTkjxKaAd7U' +
    '8rmb3awZnO6u1674tsqmWpvnQn4Nouuohf5L66TXyNKWLT/WvxWrgd0Z/tT0ZG2vMhbDVByp6uJf/SyXBX+v7jIkCl9RnKIi5FJN' +
    'NBimeA0zXMNcvIGI0OQJluPJCCKYyJdYbcIEpigURw+Po3JIQoVImpiqBv0iRuJYJhHDpXbcqS3lsfSSmvemap8GbVqjnhq40xUJ' +
    'aAuprOAoauq2rAoOQLqteqmCr7hyete+Y3P7he25js1d1g1Jk/W7affc3ST89x6tGp4bcoOEv0nLtXKlVsjnpTb+jux7dmHfwJOw' +
    'AZ831InY2DTQb7qNi7uWNSIZXM9u05/soO12Q+LRFq8a+dy2QQLV4fKbs578KhnkJeOc+V9XHWo7NMBV0SAtxni6UDSkP4/3/wVQ' +
    'SwMEFAAAAAgAqIgaXeRMn8zNAAAA+gAAABUAAAB4bC9wZXJzb25zL3BlcnNvbi54bWxdz09OhDAUgPGrNG8PLQRmgFAmjMpq4h1I' +
    'eUiT/iF91WCMe4+ja6NnwBuZWer2t/nytafNGvaEgbR3ErJUAEOn/KTdg4THOCcVnLp2M1E1Kwby7qIpss0aR81VJSwxrg3npBa0' +
    'I6VWq+DJzzFV3nI/z1ohpzXgONGCGK3hucgqHpcr4aS8tegiwZ8KmzStZny+Hy1KuIxBb+znbX/fP/bP/Xv/AqYnCS/lTTWIw22Z' +
    'nI9iSIqhPyR1KYpE1HlxzM93dd9nr8B41/J/C90vUEsDBBQAAAAAAKiIGl3al9bUKAEAACgBAAALAAAAX3JlbHMvLnJlbHPvu788' +
    'P3htbCB2ZXJzaW9uPSIxLjAiIGVuY29kaW5nPSJ1dGYtOCI/PjxSZWxhdGlvbnNoaXBzIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5v' +
    'cGVueG1sZm9ybWF0cy5vcmcvcGFja2FnZS8yMDA2L3JlbGF0aW9uc2hpcHMiPjxSZWxhdGlvbnNoaXAgVHlwZT0iaHR0cDovL3Nj' +
    'aGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcy9vZmZpY2VEb2N1bWVudCIg' +
    'VGFyZ2V0PSIveGwvd29ya2Jvb2sueG1sIiBJZD0iUmI1ODkwYmU2Mjk2YTQxOTMiIC8+PC9SZWxhdGlvbnNoaXBzPlBLAwQUAAAA' +
    'CACoiBpdk5K7QFsBAAC/BAAAGgAAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzzdRNTsQgGAbgqzTsLZRSSs1UN27c6lyAwkfb' +
    'TIEGUDtnc+GRvILxN61x4WaS2bB4v+TNAwRen19214udskcIcfSuRUVOUAZOeT26vkUPyVwIdH21u4NJptG7OIxzzBY7udiiIaX5' +
    'EuOoBrAy5n4Gt9jJ+GBlirkPPZ6lOsgeMCWE47DuQNvObH+c4T+N3phRwY1XDxZc+qMYx3ScIKJsL0MPqUV4mb6yfLETym51i+7A' +
    'UAV1TZjiHROGoQyfDJQGsLD1fESfa7FSMaZlyUkNtNIMRHVKVRxkAH2fwuj636e1Hq14uhGd0UWhoa4ZB35K3pMPhzgApC3tJ37f' +
    'AEBan57oGm4kbZpSdqxS8gx4dMVTpus6JhQwxhk3zRnwyvXlmrIylErotGDEkH/y7KiCj96kXHn7JcOUFDUuyC/UDCF6txV9Zt+z' +
    '9VMATiQlTLCyYbQSHxy8+Yau3gBQSwMEFAAAAAgAqIgaXYaoNIUvAQAAxQQAABMAAABbQ29udGVudF9UeXBlc10ueG1szZTBSgMx' +
    'EIZfZclVNmkriEi3PahXFfQFQnZ2NzSZhMy0ps/mwUfyFaTZUkSEKhb0krnMfP/3X/L28jpfZu+qDSSyARsxlRNRAZrQWuwbseau' +
    'vhTLxfxpG4Gq7B1SIwbmeKUUmQG8JhkiYPauC8lrJhlSr6I2K92Dmk0mF8oEZECueccQi/kNdHrtuLrNDDjGZu9EdT3u7aIaoWN0' +
    '1mi2AdUG208hdeg6a6ANZu0BWVJMoFsaANg7Wab02uJZAasvMxM4+lnovpVM4MoODTbSIeJ+AynZFqoHnfhOe2iEyk4Rbx2QPHHD' +
    'Aj0WzQN4GN/prwUK5mjZQSdoHzlZ7E/e+SP7mMhzSKtySKqM6YllDvyfisz+i8j5X4lESBSQ9vMbFp5qyAacHC8OfFU+pcU7UEsB' +
    'AhQDFAAAAAgAqIgaXQrDKDMdAQAAbAIAAA8AAAAAAAAAAAAAAKSBAAAAAHhsL3dvcmtib29rLnhtbFBLAQIUAxQAAAAIAKiIGl0M' +
    'wPjBoAIAAB4ZAAANAAAAAAAAAAAAAACkgUoBAAB4bC9zdHlsZXMueG1sUEsBAhQDFAAAAAgAqIgaXfpcAVkDAwAA2g0AABMAAAAA' +
    'AAAAAAAAAKSBFQQAAHhsL3RoZW1lL3RoZW1lMS54bWxQSwECFAMUAAAACACoiBpdDR656GUAAABzAAAAFAAAAAAAAAAAAAAApIFJ' +
    'BwAAeGwvc2hhcmVkU3RyaW5ncy54bWxQSwECFAMUAAAACACoiBpdviReNz0EAADtDAAAGAAAAAAAAAAAAAAApIHgBwAAeGwvd29y' +
    'a3NoZWV0cy9zaGVldDEueG1sUEsBAhQDFAAAAAgAqIgaXZFztYtfBAAA2w4AABgAAAAAAAAAAAAAAKSBUwwAAHhsL3dvcmtzaGVl' +
    'dHMvc2hlZXQyLnhtbFBLAQIUAxQAAAAIAKiIGl1lmxj8jwQAAFIPAAAYAAAAAAAAAAAAAACkgegQAAB4bC93b3Jrc2hlZXRzL3No' +
    'ZWV0My54bWxQSwECFAMUAAAACACoiBpd5EyfzM0AAAD6AAAAFQAAAAAAAAAAAAAApIGtFQAAeGwvcGVyc29ucy9wZXJzb24ueG1s' +
    'UEsBAhQDFAAAAAAAqIgaXdqX1tQoAQAAKAEAAAsAAAAAAAAAAAAAAKSBrRYAAF9yZWxzLy5yZWxzUEsBAhQDFAAAAAgAqIgaXZOS' +
    'u0BbAQAAvwQAABoAAAAAAAAAAAAAAKSB/hcAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzUEsBAhQDFAAAAAgAqIgaXYaoNIUv' +
    'AQAAxQQAABMAAAAAAAAAAAAAAKSBkRkAAFtDb250ZW50X1R5cGVzXS54bWxQSwUGAAAAAAsACwDSAgAA8RoAAAAA'
)

TASK_TYPE_TEMPLATE_B64 = (
    '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' +
    'Al4DCZyk9+uxHMo3KokZnGtHpNpSHkuvr3VnAg9Y1GY15qfDebYiEWshlRUcXy3dll3BoUm3VS9V8OVPnd627zrCeen4nusIj4cx' +
    'afJemHXP7U0ifuuyKvW9WFAS/6os18qVWiGfV9r4m7PnO4VdiidhDb6upSdibZ2i32wbF7cta0QquK7TZj85UdsLY+KzlqjSfG6T' +
    'kijtcPUteFd9lSh5xYXgwc2qwxyXRbgqUtLiXGSLlIbsp/Tuv1BLAwQUAAAACACoiBpduBYUYc4AAAD6AAAAFQAAAHhsL3BlcnNv' +
    'bnMvcGVyc29uLnhtbF3PT06EMBSA8as0bw+FkUFCKJMUZ1YT70DahzTpH8LrTDDGvcfRtdEz1BuZWer2t/nydYfNWXbFlUzwAsq8' +
    'AIZeBW38k4BLnLIGDn232ajaBVcK/mwoss1ZT+1NBcwxLi3npGZ0I+XOqDVQmGKuguNhmoxCTsuKo6YZMTrLd0XZ8DjfCLUKzqGP' +
    'BH8qTBta7Pj8ODoUcB5Xs7Gft/SePtJn+k5fwIwW8PJQ1rvhJMtMnopjVjXHIZPVUGd1td/fy7uilLJ5Bcb7jv9b6H8BUEsDBBQA' +
    'AAAAAKiIGl1NxcbBKAEAACgBAAALAAAAX3JlbHMvLnJlbHPvu788P3htbCB2ZXJzaW9uPSIxLjAiIGVuY29kaW5nPSJ1dGYtOCI/' +
    'PjxSZWxhdGlvbnNoaXBzIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvcGFja2FnZS8yMDA2L3JlbGF0' +
    'aW9uc2hpcHMiPjxSZWxhdGlvbnNoaXAgVHlwZT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3Vt' +
    'ZW50LzIwMDYvcmVsYXRpb25zaGlwcy9vZmZpY2VEb2N1bWVudCIgVGFyZ2V0PSIveGwvd29ya2Jvb2sueG1sIiBJZD0iUjAwNTM3' +
    'YmU5YWQwODRlMmUiIC8+PC9SZWxhdGlvbnNoaXBzPlBLAwQUAAAACACoiBpdURaU71sBAAC/BAAAGgAAAHhsL19yZWxzL3dvcmti' +
    'b29rLnhtbC5yZWxzzdRPTsQgFAbwqzTsLdC+ttTYmY0bt+oFgHm0jQUaQK1nc+GRvILxb1rjws0ks2HxveTLj0fC6/PLxX6xU/aA' +
    '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'
)

ROLE_MATRIX_TEMPLATE_B64 = (
    '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'
    'LyaDv3xw8c5ffnjxzsGB6kJyYLcX485vkRY6ZgvdbAt3LqWFttlCL9vC7Utp4dBsYT/80hZ+XKAF9V+KmHL08d0PT1Z3zz68/+i89eDkg4tbB4dvDw9aZ/c/'
    'VDcHu/9fnH68+1//oPW/Ti8uTh8mf310cvfeyZn6q3vQ+uD09CL9Yz9oPj09+7vd2P7R/wBQSwMEFAAAAAgAV2sLXQtmf60fAQAAfwEAABUAAAB4bC9wZXJz'
    'b25zL3BlcnNvbi54bWx9kD9qwzAcha8itMuSU/lfiBOCHU+ldzCyXBssy1hqSSmF0LlDoB269gYhIVAoyRl+vlEx6dIOnd7jG94Hb7ZYqwbdy97Uuo2x6zCM'
    'ZCt0Ube3Mb6zJQnxYj5bN1ZMO9kb3V7XxqK1alozHWmMK2u7KaVGVFLlxlG16LXRpXWEVlSXZS0kNV0v88JUUlrV0AlzQ2qrEclCaKVkaw3+ZUFFbbomf7jJ'
    'lYwxvMMBzsNmeEGwdRBshw3s4AvOsMeoLmL8mAU8CrwoIUvOUpIkqU+8jPvEj8KE+ROeBlH2hBH9T/IK58sm7IbnSyJ4g0/YwWmsH3AcNnCAIxxhP6Ift5tE'
    '7IqHEeHLlBEvCD3iZWlCMrZaJX6acua5Fzf98+P8G1BLAwQUAAAAAABXawtdcQrUICgBAAAoAQAACwAAAF9yZWxzLy5yZWxz77u/PD94bWwgdmVyc2lvbj0i'
    'MS4wIiBlbmNvZGluZz0idXRmLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAw'
    'Ni9yZWxhdGlvbnNoaXBzIj48UmVsYXRpb25zaGlwIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2'
    'L3JlbGF0aW9uc2hpcHMvb2ZmaWNlRG9jdW1lbnQiIFRhcmdldD0iL3hsL3dvcmtib29rLnhtbCIgSWQ9IlJiNzlmYTU4ODExZWI0OGU1IiAvPjwvUmVsYXRp'
    'b25zaGlwcz5QSwMEFAAAAAgAV2sLXZ66xrE7AQAAgQMAABoAAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc7XTP07DMBQG8KtE3omd2Ekq1LQLC2vpBfzn'
    'ObEa25HtQno2Bo7EFRBtQUnFwNLFw/ekTz8/6X2+f6y3kx2yVwjReNeiIicoAye9Mq5r0THphxXabtY7GHgy3sXejDGb7OBii/qUxkeMo+zB8pj7EdxkB+2D'
    '5SnmPnR45PLAO8AlITUO8w607Mz2pxH+0+i1NhKevDxacOmPYhzTaYCIsj0PHaQW4Wm4ZvlkB5Q9qxbtKiJpLUohFK1YWVOU4buBUg8Wlp5zdHmLmUo2qtJE'
    'iIavGKOqvKcq9jyAeknBuO52W/PRjCeaigEwWQstWUX1PXlvPhxiD5CWtN/4+wMAab69pipkQ3jNaMWYFvBPnjUy+Oh1yqW3VxkuSdHggtygRgjRu6Xokv3M'
    'ZhwuKV1x1XBQihWKnTl4cUibL1BLAwQUAAAACABXawtd/5xeBygBAACzAwAAEwAAAFtDb250ZW50X1R5cGVzXS54bWytk0FOwzAQRa8SeYtitywQQk27ALaA'
    'BBewnEli1R5bnmlxz8aCI3EFVKeqACEF1G48m5n3/l/44+19screVVtIZAM2Yi5nogI0obXYN2LDXX0tVsvFyy4CVdk7pEYMzPFGKTIDeE0yRMDsXReS10wy'
    'pF5Fbda6B3U5m10pE5ABueY9QywXd9DpjePqPjPgqM3eiep23NurGqFjdNZotgHVFtsfkjp0nTXQBrPxgCwpJtAtDQDsnSxTem3xooDVr84Ejv4nPbSSCVzZ'
    'ocFGOioet5CSbaF60okftIdGqOwU8c4ByTM3LNApNQ/gYXznJwcomMmyg07QPnOy2J+98/CFPRXkNaR1OSRVxun9v4c58qeCREgUkA7zDyk81ZANODleHPmq'
    'fMHlJ1BLAQIUAxQAAAAIAFdrC11obo6jVQEAANUCAAAPAAAAAAAAAAAAAACkgQAAAAB4bC93b3JrYm9vay54bWxQSwECFAMUAAAACABXawtd88a6AnQFAABJ'
    'cwAADQAAAAAAAAAAAAAApIGCAQAAeGwvc3R5bGVzLnhtbFBLAQIUAxQAAAAIAFdrC12F8OlgQAMAAEwPAAATAAAAAAAAAAAAAACkgSEHAAB4bC90aGVtZS90'
    'aGVtZTEueG1sUEsBAhQDFAAAAAgAV2sLXTr0WLayAAAAJAIAABQAAAAAAAAAAAAAAKSBkgoAAHhsL3NoYXJlZFN0cmluZ3MueG1sUEsBAhQDFAAAAAgAV2sL'
    'XWfC76CfKQAAYZsBABgAAAAAAAAAAAAAAKSBdgsAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbFBLAQIUAxQAAAAIAFdrC10LZn+tHwEAAH8BAAAVAAAAAAAA'
    'AAAAAACkgUs1AAB4bC9wZXJzb25zL3BlcnNvbi54bWxQSwECFAMUAAAAAABXawtdcQrUICgBAAAoAQAACwAAAAAAAAAAAAAApIGdNgAAX3JlbHMvLnJlbHNQ'
    'SwECFAMUAAAACABXawtdnrrGsTsBAACBAwAAGgAAAAAAAAAAAAAApIHuNwAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHNQSwECFAMUAAAACABXawtd/5xe'
    'BygBAACzAwAAEwAAAAAAAAAAAAAApIFhOQAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLBQYAAAAACQAJAEYCAAC6OgAAAAA='
)

STRUCTURE_TEMPLATE_B64 = (
    "UEsDBBQABgAIAAAAIQBBN4LPbgEAAAQFAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACsVMluwjAQvVfqP0S+Vomhh6qqCBy6HFsk6AeYeJJY"
    "JLblGSj8fSdmUVWxCMElUWzPWybzPBit2iZZQkDjbC76WU8kYAunja1y8T39SJ9FgqSsVo2zkIs1oBgN7+8G07UHTLjaYi5qIv8i"
    "JRY1tAoz58HyTulCq4g/QyW9KuaqAvnY6z3JwlkCSyl1GGI4eINSLRpK3le8vFEyM1Ykr5tzHVUulPeNKRSxULm0+h9J6srSFKBd"
    "sWgZOkMfQGmsAahtMh8MM4YJELExFPIgZ4AGLyPdusq4MgrD2nh8YOtHGLqd4662dV/8O4LRkIxVoE/Vsne5auSPC/OZc/PsNMil"
    "rYktylpl7E73Cf54GGV89W8spPMXgc/oIJ4xkPF5vYQIc4YQad0A3rrtEfQcc60C6Anx9FY3F/AX+5QOjtQ4OI+c2gCXd2EXka46"
    "9QwEgQzsQ3Jo2PaMHPmr2w7dnaJBH+CW8Q4b/gIAAP//AwBQSwMEFAAGAAgAAAAhALVVMCP0AAAATAIAAAsACAJfcmVscy8ucmVs"
    "cyCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACskk1P"
    "wzAMhu9I/IfI99XdkBBCS3dBSLshVH6ASdwPtY2jJBvdvyccEFQagwNHf71+/Mrb3TyN6sgh9uI0rIsSFDsjtnethpf6cXUHKiZy"
    "lkZxrOHEEXbV9dX2mUdKeSh2vY8qq7iooUvJ3yNG0/FEsRDPLlcaCROlHIYWPZmBWsZNWd5i+K4B1UJT7a2GsLc3oOqTz5t/15am"
    "6Q0/iDlM7NKZFchzYmfZrnzIbCH1+RpVU2g5abBinnI6InlfZGzA80SbvxP9fC1OnMhSIjQS+DLPR8cloPV/WrQ08cudecQ3CcOr"
    "yPDJgosfqN4BAAD//wMAUEsDBBQABgAIAAAAIQAX/MY50AMAAJEJAAAPAAAAeGwvd29ya2Jvb2sueG1stFbNbttGEL4X6Dsw7CEn"
    "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"
    "9qgLH/km6s63oCQubieh2VUipbaMTJaqx4OUP1XexYhmWFrrICPTRYz4M9QT3mVoUEJDE081E1cX0Y8lNI7TEZylWcUuNHcahv5s"
    "hNZM/PuIBRu3Nr9rpj58bKifpZDVn5buGD7pOhL3pOtH6Ur609KX09/G3bIGS2yNO9/LKlcClW3E2flSLVlDFNZNvZytO1kSVrIF"
    "D17Ozq0oH0aIG42si6zYgozGx6GJjsRaB17SGh4YDUk0qbxW+cBk56y8jSrra5XrQMeBqTwhnons2bplrdgza52Nu3S0W72aLVrx"
    "z8sJlSqYit6WFiRnz6AsX5XW/Wo90NPmfJ915u49qy8t1QFZ1Zvo1PpjY8NJZbGeVAOTGZrvmiPzrDg+PSltxdiP8kzPXitvz3Ot"
    "G9lIzUdzdWw2OuOROPO1HmlmxzoL0Uzl1DxTjWZjI/Z5Yv3suK6hd4Vk/Df5Tr4UbZ/iPzh9fi2/yKfyjWm/RfsIX+15kuKltmju"
    "6PkQx1a8j9kYtZRXcTIkkZlE8L3IkvmjpTmjLpg/PPXPHdyD0xEfR2Y7mnlrZl17z7g/R6sdeW/OSNHUYT0ZLLPYUtecFfY25IFD"
    "DX7NEtEZejboDWpiMtzKaX9yysf+vvw5P0J8tLGeJdFNeTPxtr17JFKJx7vmjpCcwGmilmV27vwVXFRYvdtmK0b6LmP1xfc7PavU"
    "dp3xb/Ira99WxjfZ8bUzvktbbm/o1p6CuRsnvonPySC6r2vm6g3L7tozfuS/4iRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRA"
    "AiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRA"
    "AiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiRAAiTwLBH4GwAA//8DAFBLAwQUAAYACAAAACEA"
    "t7HoX2gBAACqAgAAEQAIAWRvY1Byb3BzL2NvcmUueG1sIKIEASigAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "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"
)

USER_IMPORT_TEMPLATE_B64 = (
    "UEsDBBQABgAIAAAAIQDky1wPnAEAAKcGAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACslctu2zAQRfcF+g8Ct4VFp4sgKCxnkbbLJkBcoFuG"
    "HFuE+SpnnNh/nyGdFwLHQmBtREnk3HtIkVezy613zT1ktDF04qydigaCjsaGVSf+Ln5PLkSDpIJRLgboxA5QXM6/fpktdgmw4eqA"
    "neiJ0g8pUffgFbYxQeCeZcxeET/mlUxKr9UK5Pfp9FzqGAgCTahoiPnsJyzVxlHza8uv9yQZHIrmaj+weHVCpeSsVsSk8j6Ydy6T"
    "J4eWK+sY7G3Cb4wh5EGH0vOxwVPdNS9NtgaaG5Xpj/KMIbdOPsS8votx3R4XOUAZl0urwUS98bwCLaYMymAPQN61tW29suGZ+4h/"
    "HYyyNmcjg5T5VeEBDuLvDbJeT0eoMgOGvHMC6LIFcOQpv1EeYEDaORjbfi865NyrDOaWMp/O0QHeag99dnXHKyCpNGNvvSo64P9/"
    "A3m32EO83o9N8qp8DEdvkKL/5520BP4mx4Sng7yIFj3IZOElyQ4lAsdJNeZgzfD5c/GcnKV6kj7nyKF88kGEEvsGzAFvWX8z80cA"
    "AAD//wMAUEsDBBQABgAIAAAAIQC1VTAj9AAAAEwCAAALAAgCX3JlbHMvLnJlbHMgogQCKKAAAgAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArJJNT8MwDIbvSPyHyPfV3ZAQQkt3QUi7IVR+gEncD7WN"
    "oyQb3b8nHBBUGoMDR3+9fvzK2908jerIIfbiNKyLEhQ7I7Z3rYaX+nF1ByomcpZGcazhxBF21fXV9plHSnkodr2PKqu4qKFLyd8j"
    "RtPxRLEQzy5XGgkTpRyGFj2ZgVrGTVneYviuAdVCU+2thrC3N6Dqk8+bf9eWpukNP4g5TOzSmRXIc2Jn2a58yGwh9fkaVVNoOWmw"
    "Yp5yOiJ5X2RswPNEm78T/XwtTpzIUiI0Evgyz0fHJaD1f1q0NPHLnXnENwnDq8jwyYKLH6jeAQAA//8DAFBLAwQUAAYACAAAACEA"
    "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"
    "zuImFEBtkxVYMjmw0Oq4gZ/v/WEbhSieyjzlxwDr7MrkXGmmdwWT6hUA64d1BZSZbkuuO4cnR5eX2htk2iChlZtJQivDGE1xUZ36"
    "YclV5rpdptSgJ0OxfhtKGs3Wh8i1FJEL2kBTXSlo6px33LARwHnYBC067gz2+nCZLKB2uFzqIjqHA7OJyPIX/l2UZZFx0Uc8zgOu"
    "RCdXg4QInDmUJB1XDn9TDTRVGqK41eogCB8tuTbIysdGDpJuJhnPZngi9LRrLTLS+S0ofK4V1qfK/N3B0pItId0n8fTcGdNldg9B"
    "iQXNmgzglHA48qnl0ZwSOMPcCFlZfxcmpkJ29UNEVUN5O6KLGBUzii7mOVyJ6IaOutvEQLsrxgwB3Q7heC4n2PeedS+fqmXkNNEs"
    "50xDVeSsaRfTDzfJa6zKSdRglUu32jbwUuvaa62DQrXOEpfMum8wIWjUys4MapLxtgxLzS5aTWpXuCDQIhHuiNtmjrBG4l1nfrC7"
    "WLVyglivK1Xhq48d+vcINn4E4tGHk98lFVylEr42ZAgWffnZcS4b8Io8FsUaEa6cZUY67hMv6PpRPYgqXisYVPyG71VaQbdR6QZB"
    "ozYIal6/V38KE4uIk1qQf2gZwhEUXRWfW1T71ieXZH3Kdm3CkipTn1Sqirj65FKr7/7k4hAQnSdhfdhutHthpd3oDit+v9eqtKOw"
    "V+mHUbM/7EdBqz186jpnCux3G5EfDlqVsBZFFT/0JP1Wu9L06/Wu3+y2Bn73abGMgZHn8lHEAsKreO3/CwAA//8DAFBLAwQUAAYA"
    "CAAAACEAyZ4hWzIEAABwEQAADQAAAHhsL3N0eWxlcy54bWzsWM1u4zYQvhfoOwi6K/qJ5dqGpEUcR8UC22KBuECvlETZRCjSoKis"
    "3WKB9tTDHnvvM7SHAv1B0Vdw3qhDSrLkTeL89lA0ECCTo5nhfDPDGdLBq3VBjUssSsJZaLpHjmlglvKMsEVofjWPrZFplBKxDFHO"
    "cGhucGm+ij79JCjlhuLzJcbSABWsDM2llKuJbZfpEheoPOIrzOBLzkWBJEzFwi5XAqOsVEIFtT3HGdoFIsysNUyK9D5KCiQuqpWV"
    "8mKFJEkIJXKjdZlGkU5eLxgXKKFg6todoNRYu0PhGWvRLqKp19YpSCp4yXN5BHptnuckxdfNHdtjG6WdJtD8OE2ubzveHva1eKSm"
    "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"
    "UEsFBgAAAAASABIAvwQAAPtAAAAAAA=="
)

BG = "#FFFFFF"
FG = "#222222"
ACCENT = "#F7921E"
ACCENT_HOVER = "#FFA74B"
GRAY = "#D9D9D9"
BORDER = "#dcdcdc"
SUCCESS = "#4CAF50"
ERROR = "#F44336"


# =====================================================================
#  Helpers: настройки и встроенные Excel-шаблоны
# =====================================================================
def _user_config_dir() -> Path:
    """Папка пользовательских настроек. Нужна, чтобы выбранные шаблоны сохранялись между запусками."""
    if os.name == "nt" and os.getenv("APPDATA"):
        return Path(os.getenv("APPDATA", "")) / APP_FOLDER_NAME
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / APP_FOLDER_NAME
    xdg_config = os.getenv("XDG_CONFIG_HOME")
    if xdg_config:
        return Path(xdg_config) / APP_FOLDER_NAME
    return Path.home() / ".config" / APP_FOLDER_NAME


def _settings_file_path() -> Path:
    return _user_config_dir() / SETTINGS_FILE_NAME


def _load_template_settings() -> Dict[str, str]:
    """Читает сохранённые пути к Excel-шаблонам из settings.json."""
    path = _settings_file_path()
    try:
        if not path.exists():
            return {}
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {}
        return {str(k): str(v) for k, v in data.items() if isinstance(v, str)}
    except Exception:
        return {}


def _save_template_settings(structure_template_path: str, user_import_template_path: str) -> None:
    """Сохраняет выбранные пользователем пути к шаблонам."""
    path = _settings_file_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    data = {
        "structure_template_path": structure_template_path or "",
        "user_import_template_path": user_import_template_path or "",
    }
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _existing_file_or_empty(path: str) -> str:
    try:
        return path if path and Path(path).exists() and Path(path).is_file() else ""
    except OSError:
        return ""


def _normalize_xlsx_destination(destination_path: str) -> Path:
    """Возвращает путь назначения с расширением .xlsx и создаёт родительскую папку."""
    destination = Path(destination_path)
    if destination.suffix.lower() != ".xlsx":
        destination = destination.with_suffix(".xlsx")
    destination.parent.mkdir(parents=True, exist_ok=True)
    return destination


def _decode_embedded_template(encoded_data: str, expected_sha256: str) -> bytes:
    """Декодирует встроенный шаблон и проверяет, что его содержимое не повреждено."""
    try:
        data = base64.b64decode(encoded_data.encode("ascii"), validate=True)
    except Exception as exc:
        raise RuntimeError(f"Не удалось декодировать встроенный Excel-шаблон: {exc}") from exc

    actual_sha256 = hashlib.sha256(data).hexdigest()
    if actual_sha256 != expected_sha256:
        raise RuntimeError(
            "Встроенный Excel-шаблон повреждён: контрольная сумма не совпадает"
        )
    return data


REMARK_TYPE_TEMPLATE_SHA256 = "dacb7850157c922e64013fb4e44f3c95c05867d5236a6821dca0ba6ce23f0c9c"
TASK_TYPE_TEMPLATE_SHA256 = "2aaaca38abf6386a063f1ab69d93bda3cf7fb28b3e7c7ed8ea72c687f8b2fa75"
REMARK_TYPE_TEMPLATE_B64 = "UEsDBBQAAAAIACGSGl3kCSBW4gAAADwBAAAPAAAAeGwvd29ya2Jvb2sueG1sjZBNTsMwFISvYr09cQIVP1GSbtiw5QaO89xYtf0i24XcgAW34ATdIIoE4gqvN0ItiG7ZjUajmU/TLGfvxAPGZCm0UBUlCAyaBhtWLWyyObuGZdfM9SPFdU+0FrN3IdVzC2POUy1l0iN6lQqaMMzeGYpe5VRQXMk0RVRDGhGzd/K8LC+lVzbAoe/opj8lgvJ42Bf8wjv+2j8LfuMtf/Dr/om3/Mk7fgdxzN4NLVQgYm2HFu5LrW4W6sL0CqvFVY/wSxj/Q0jGWI23pDceQ/5BjOhUthTSaKcEQnaNPOHK0xPdN1BLAwQUAAAACAAhkhpdi09UBzgCAADQDwAADQAAAHhsL3N0eWxlcy54bWzlV0tv2zAM/iuC7ovtZE3WoG7RJjOwSy/tYVfFlm0BlGRISur01w+S/Oq6rvGQtVmXi0mC/PgMTV9c1RzQjirNpIhxNAkxoiKVGRNFjLcm//QFX11e1Ett9kDvSkoNqjkIvaxjXBpTLYNApyXlRE9kRUXNIZeKE6MnUhWBrhQlmbZmHIJpGM4DTpjAFjGXwmiUyq0wMZ51IufsEe0IxDiKMAqsQBBOvWhFFDAjnTzoLdrnxus/A0glSIVUsYlxkkSL2fxs/Veg/e8w6IbQzgkD6Irx2ReDAdhnRYyhSiQMADX0/b6iMRZS0A6xUX7VqFBkH03PRttpCSzzcRWrYcbr86/XyaLFG9gfCX/QrBfxG8JVciNVRlVXyynuhb4rnrYU0NwgN9gxNmUzlk/aebNYna/nrXOrb1UUK8pRhs7A6hhZjbEzsvIRGyP5GENv0ZA+3Y50VUopwJ3F+553pYocap0jseUJN9+yGIcY2RFtSQbQkB6qYbzPIWTrYoi++FP4Ou/9HAQwfQmAVBXsb7d8Q1XitpRN2ksTKYYcA+i5Gwfm+N+GEJ1uCNFHC8Hx18AKwWk/vKQVoAdFqntaeyg/oHV++mHvqDIsta+MlApDFR6fyOCPMHunKZz+LyG81s5SKvYohRk29JAenm70zwd05EyG77SW3iaEo6+l0wj7uGvppBIZ+W74mE14270yOux/aR0+/Dq55kAe3MbuVv7p+O7kyH47xvjWZgJPT+Dhqa0d23+bX/4AUEsDBBQAAAAIACGSGl36XAFZAwMAANoNAAATAAAAeGwvdGhlbWUvdGhlbWUxLnhtbL1X23KbMBT8FUbvDTdz84RkEsduH9Jpp8kPyCBAjRAeSY6dv+8gbgKM4zR27AdLYs/ZReewwte3+5xor4hxXNAQmFcG0BCNihjTNARbkXzzwe3NNZyLDOVIozBHIVhkUHz//Qy0fU4on8MQZEJs5rrOowzlkF8VG0T3OUkKlkPBrwqW6jGDO0zTnOiWYbh6DjEFbd4lQTmigpcLEWFP0QGy8lr8YpY//I0vCNNeIQnBDtO42D2jvQAagVwsCAuBIT9A02+u9TaKiIlgJXAlP01gHRG/WDKQpes20lha/szsGCSCiDFw6ZffLqNEwChCtJajgk3HNXyrASuoangge+CZ9iBAYbDHDIF7b836ARJVDWfjG10FywenHyBR1dAZBdwZ1n1g9wMkqhq6o4DZ8s6zlv0AicoIpi9juOv5vtvAW0xSkB8H8YHrGt5Dg+9gutJqVQIqeo33K0lwhGTf5fBvwVYFFbLKUGCqibcNSmBUNigkeM2w9ojTTEgeOEfwHUDEjwL0AWeO6bsCjlAfIW3pOgZd3Qy5NbmYfCQTTMiTeCPokUtxvCA4XmFC5ERGtaXYZAvCGsIeMGWwG/M6Vcq1TcFDYIDJXNJBMBXVmus1Tz2ck23+s4jrpjdbO4BzDkV3wXAUn2gZ5CzlqoYSd7IOz57Q0dENddgn6pB3crIQ3/ywkOCoEF0pD8FUg+Up4cxqu+URJCguC1Yn6JX1LCUOZlN3ZH12a08oMc9gjJq8xpSSqWbruvAMRVakeP5hJUEwIaTcqksUWR/bAaH9mbYr+b3m7v7LLDaMiwfIswonL7XnK1VoAsP5Ahqr3JnL0ejDPURJgiIxsdJNH7mosxy8/Fl0OSm2ArGnLN5pa7Jlf2AcAsczHQNoMeaiKYAWY9a1z/j9oluHZJPB2sl7D22Fl+OWUxEr5Qyl9+e14nW6Ostx9X7UwLWm7NabfhIvcD4Gyrmk+Efgf9RTK6s897Gp6lDlTRqtPSHPvpDRdl35dYY6bNnSY5vXMTkb/IFqVm7+AVBLAwQUAAAACAAhkhpdDR656GUAAABzAAAAFAAAAHhsL3NoYXJlZFN0cmluZ3MueG1sBcFRCsMgDADQq0j+Z9w+xpDankXatAomFpMNj7/3lm1ycz8aWrskePoAjmTvR5UrwdfOxwe2dZlR1dzkJhpngmJ2R0TdC3FW32+Sye3sg7Op7+NCvQflQwuRccNXCG/kXAUcrn9QSwMEFAAAAAgAIZIaXYhtVX9cBwAAVy8AABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWyd2tlu20YUBuBXIXifaPcS1Akqcd9E9qbXqk0vqGUZEpP40nZQ+KIBugItuiVpX0B1o0Z1IvsVzrxRMZQok9Y/ikdXJj/xHA5n5gxJWZ88OekeKs/i/uCgd7SlVh6WVSU+2u7tHBztbalPk90HG+qTx5+cPHre63852I/jRDnpHh4NHp1sqftJcvyoVBps78fdzuBh7zg+Ouke7vb63U4yeNjr75UGx/24s5OGdQ9L1XJ5rdTtHBypPGGqRnpw2Fd24t3O08Pks95zKz7Y20+21EpDVUr8wO3e4WD2V+ke8EaqSrdzkv59frCT7G+ptaqq7B/s7MRHW2pZVbafDpJe9/PpZ5XbNNPw6iy8Og+vbEqE12bhtdXC67Pw+mrhjVl4Y7XwtVn42mrh67Pw9dXCN2bhG6uFb87CN1cLr5SzeVNeMcF84lVWTJBNPb6xUoJs8vGNlRJk049vZMUjE5/NP76RNUCm+CrZDOQbswTVukyCbA7yjSxB+WMJSrfLSLruaJ2kw3f6vedKPz2ILzn1ahY8X4TSpWqbH/NpRVUGaeElW+og6aefPHtMf9KYbhR6R0P6QCN2QUOa0Jh9w8/5bHrmeY4mztG0/Qd0Rdd0zU7pLY1pQkN2zvdQkpagIT/Qz/QKBWiCgJ9oSFc0pHfsgsZ0hUJ1Qehr9iJt8CVd8wazcxrRe/ZSoRt2Stc0oit2TkOU0RBk/Jadsq9mma7EF28Kwn+mCf1LI5rQiJ1OO3PCztg5O+UtFeezhKM64uH8EpVsZ76RXviYXaCEtiDhjzSiD3StzDbG9Da3uSyhI0j4hs8T3l3snI+eQm94NvYdTZS7ny1L7wrSv2KnNE5nNB/QSyUPhZ1lyT1B8t/Smcdbl6ZOd9kZz6gUPxMl9gWJ/2IXNKK/acK+pv+U2e5k2j/zzz7WJcE0+e0DCMf2DNfunvJXGtE/6Vne0lBJM6fjOa2FIV2iU4S32W5PERWwlK5OuUWKL7PJdLUUL1LVaYrq3Vb+wgeMzwZ2QRNeoQsrFo3giiVK+Ipe0xu4PC2JgKuT6HiYXZdujyGIeACXF8nGW9LNsWWa40j1jSvZeE+mKb70lQZVVEYzrN+zjaHwtHzlYWd8TeWrPIqNlsUO+b1L4YsAveWrDXuppIvlTSHTQhHW7lOENcF5f+e3j/Q2f06X6Un5zYQ3YVS4c7IXsBRr0qUoHaHVJKaELkqP63BJY2ApyjTFkmqKLd0vjkxjXMkr9aSb48s0J6ihOqwJ6hCeLxR2L13TO37jE99BotpHqzAtimmaS/7ky14uL8L6fYqwLjjt93Q9rTV2pvDHgnfFp2BB7dXlhrQleby25HjBjbAuMQUMmYNN6bZYMultUXpceDKpXcle92SS+9LdEtRR4dXFN0BB8dXvdwtUaJyvJHFBivL9QX/zJ9jZY+ud17eFImzcpwgbglP9SO/ZN9k9j98KT+mKxtOn5w90Q++XvUI2RVnxTbAhMchaQ24G6UuOF9wGZZpjSl2pJZPalrxSR/pKXZnmeJLN8WWSBw1UiA1ciDBD2FjySHe+9FUvEoXyr19uaMzOsu9N0qJT0mI+Yy+yp0P2gp9ieS2u5UpuDVxrE2ELoYZQR2ggNBFaCG2EDkIXoYfQRxggbCMMEUYFXOj59VzPr6OeR9hCqCHUERoITYQWQhuhg9BF6CH0EQYI2whDhNH60p7fyPX8Bup5hC2EGkIdoYHQRGghtBE6CF2EHkIfYYCwjTBEGG0s7fnNXM9vop5H2EKoIdQRGghNhBZCG6GD0EXoIfQRBgjbCEOE0ebSnq+U8/+PKKO+h9qCqkHVoRpQTagWVBuqA9WF6kH1oQZQ21BDqFFRF8ei8L8h9JVxE2oLqgZVh2pANaFaUG2oDlQXqgfVhxpAbUMNoUZFXRyLan4s0PeOTagtqBpUHaoB1YRqQbWhOlBdqB5UH2oAtQ01hBoVdXEsavmxQN89NaG2oGpQdagGVBOqBdWG6kB1oXpQfagB1DbUEGpU1MWxqOfHAn0d0YTagqpB1aEaUE2oFlQbqgPVhepB9aEGUNtQQ6hRURfHopEfC/RG2oTagqpB1aEaUE2oFlQbqgPVhepB9aEGUNtQQ6hRURfHIv/CzP/fCsYCvjJD1aDqUA2oJlQLqg3VgepC9aD6UAOobagh1Kioi2ORf4XmP+oBYwFfoqFqUHWoBlQTqgXVhupAdaF6UH2oAdQ21BBqVNTFsci/VFfgWzXUFlQNqg7VgGpCtaDaUB2oLlQPqg81gNqGGkKNiro4FvnXbP4jPzAW8EUbqgZVh2pANaFaUG2oDlQXqgfVhxpAbUMNoUZFXfzpTf69m//acHEsoLagalB1qAZUE6oF1YbqQHWhelB9qAHUNtQQalTUbCxKd3652Y37e3ErPpz+qHO+p/TjXT72j/iPqabhxSOPO3ux3+nvHRwNlMN4N9lSyw/XVaU//d9Vup30jtOthqp80UuSXjfb2487O3Gf79VUZbfXS+Y70zPNfyP/+H9QSwMEFAAAAAAAIZIaXZfDiq0oAQAAKAEAAAsAAABfcmVscy8ucmVsc++7vzw/eG1sIHZlcnNpb249IjEuMCIgZW5jb2Rpbmc9InV0Zi04Ij8+PFJlbGF0aW9uc2hpcHMgeG1sbnM9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9wYWNrYWdlLzIwMDYvcmVsYXRpb25zaGlwcyI+PFJlbGF0aW9uc2hpcCBUeXBlPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9yZWxhdGlvbnNoaXBzL29mZmljZURvY3VtZW50IiBUYXJnZXQ9Ii94bC93b3JrYm9vay54bWwiIElkPSJSYTFmNmI0OGI0MzE5NGM2MCIgLz48L1JlbGF0aW9uc2hpcHM+UEsDBBQAAAAIACGSGl2xKEHIEAEAAPICAAAaAAAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHO1kj1OxDAQRq9iuSd2EstJ0Ga3oaFd9gKOM46j9U9keyF7NgqOxBUQC0IJoqBJM8U30tObT/P++rY7zNagZwhx9K7FeUYxAid9P7qhxZek7mp82O+OYEQavYt6nCKarXGxxTql6Z6QKDVYETM/gZutUT5YkWLmw0AmIc9iAFJQyklYMvCaiU7XCf5D9EqNEh68vFhw6Q8wielqIGJ0EmGA1GIym+8sm63B6LFv8bEplFSC8YarhpU9x4hsJpQ0WFj73KKvmS+sqorygitoqrpjnFZbWkUtAvRPKYxu+N3WcrUsTfC6hL7KS1ozybot9V58OEcNkNZqP/HnAQBp2R6VomGiVJ2AnFUd3PTI6nP3H1BLAwQUAAAACAAhkhpdjYLZqRYBAABTAwAAEwAAAFtDb250ZW50X1R5cGVzXS54bWytk0FOwzAQRa8SeYtqpywQQkm7ALaABBewnEli1R5bnmlIz8aCI3EFVAdFgJAi1G48m/F7/y/m4+292o7eFQMksgFrsZalKABNaCx2tdhzu7oW2031cohAxegdUi165nijFJkevCYZIuDoXRuS10wypE5FbXa6A3VZllfKBGRAXvGRITbVHbR677i4Hxlw0o7eieJ22juqaqFjdNZotgHVgM0vySq0rTXQBLP3gCwpJtAN9QDsncxTem3xIoPVn84Ejv4n/WolE7i8Q72NNCseB0jJNlA86cQP2kMt1OgU8cEByTM3zNAlNffgYXrXJwfImMWyvU7QPHOy2J2983f2UpDXkHb5I6k8Tu//M8zMn4OofCKbT1BLAQIUAxQAAAAIACGSGl3kCSBW4gAAADwBAAAPAAAAAAAAAAAAAACkgQAAAAB4bC93b3JrYm9vay54bWxQSwECFAMUAAAACAAhkhpdi09UBzgCAADQDwAADQAAAAAAAAAAAAAApIEPAQAAeGwvc3R5bGVzLnhtbFBLAQIUAxQAAAAIACGSGl36XAFZAwMAANoNAAATAAAAAAAAAAAAAACkgXIDAAB4bC90aGVtZS90aGVtZTEueG1sUEsBAhQDFAAAAAgAIZIaXQ0euehlAAAAcwAAABQAAAAAAAAAAAAAAKSBpgYAAHhsL3NoYXJlZFN0cmluZ3MueG1sUEsBAhQDFAAAAAgAIZIaXYhtVX9cBwAAVy8AABgAAAAAAAAAAAAAAKSBPQcAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbFBLAQIUAxQAAAAAACGSGl2Xw4qtKAEAACgBAAALAAAAAAAAAAAAAACkgc8OAABfcmVscy8ucmVsc1BLAQIUAxQAAAAIACGSGl2xKEHIEAEAAPICAAAaAAAAAAAAAAAAAACkgSAQAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc1BLAQIUAxQAAAAIACGSGl2NgtmpFgEAAFMDAAATAAAAAAAAAAAAAACkgWgRAABbQ29udGVudF9UeXBlc10ueG1sUEsFBgAAAAAIAAgAAwIAAK8SAAAAAA=="
TASK_TYPE_TEMPLATE_B64 = "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"


REMARK_TYPE_TEMPLATE_SHA256 = "252bcab17d40946f317f81a9e6a9442344222ff8fdf1f131de3f9d46312a0e31"
TASK_TYPE_TEMPLATE_SHA256 = "7784c7539e3be171654e3dd1adf6a49363f0f863b6364a50857dd16d8aa6419c"
REMARK_TYPE_TEMPLATE_B64 = "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"
TASK_TYPE_TEMPLATE_B64 = "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"


REMARK_TYPE_TEMPLATE_SHA256 = "7398610b323b709c4383c50f4dad6fe007ed45a81c7806dd10b3c74a7867aca5"
TASK_TYPE_TEMPLATE_SHA256 = "f857b8b4d2f3f61972dc7231a10ba5682ecced37fe6a3ea10f40c00be541df38"
REMARK_TYPE_TEMPLATE_B64 = "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"
TASK_TYPE_TEMPLATE_B64 = "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"


STRUCTURE_TEMPLATE_SHA256 = "3576bd8c06e439dca2703095f9c118d1aef4ab5101d3229a35aac1fa1fda1bbd"
STRUCTURE_TEMPLATE_B64 = "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"
USER_IMPORT_TEMPLATE_SHA256 = "472ee76a00fcfee98ac18ad871784ed0f062218b22b6a3e8a9f9499377211c78"
USER_IMPORT_TEMPLATE_B64 = "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"
ROLE_MATRIX_TEMPLATE_SHA256 = "cc43768ab7a450b799d8ac80992b44b701f2dfdad41f7061dcacd62677e07d74"
ROLE_MATRIX_TEMPLATE_B64 = "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"
REMARK_TYPE_TEMPLATE_SHA256 = "b49d0f43e33d7d61613c0ff886e7feb8b490c311cc60a2aba8df7082dc8b7b3b"
REMARK_TYPE_TEMPLATE_B64 = "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"
TASK_TYPE_TEMPLATE_SHA256 = "260d277deeef1986b9c52015e8d5f9fdde36028663bb8b79b26cf358cb1fb164"
TASK_TYPE_TEMPLATE_B64 = "UEsDBBQAAAAIAKRRG11EdsPz3AAAADQBAAAPAAAAeGwvd29ya2Jvb2sueG1sjZAxTgMxFESvYv2e9W4Skmi13jQ0tNzAWX9nrdj+K9sB3yAFt+AEKZCg4wrOjVACIi3daDSaeZpuk51lzxiiIS+gqWpg6AdSxu8EHJK+W8Om73L7QmG/Jdqz7KyPbRYwpjS1nMdhRCdjRRP67Kym4GSKFYUdj1NAqeKImJzls7pecieNh0vf1Y1/innp8LLPylv5LF/nV1Y+yqm8l9P5COyaeVQCGmChNUrAk1w1a1Xr1RaXs8Vc38MvWfgPGWltBnyg4eDQpx+0gFYmQz6OZorAeN/xGya/PdB/A1BLAwQUAAAACACkURtd/r/XpDQDAABELwAADQAAAHhsL3N0eWxlcy54bWztWstuozAU/RXk/ZRXSJqqtGoeSLPppl3MlhDzkIyNjNMh/fqRbRJgGqahTVJwJ5tcWz7H93Af9sK390WKtBdI84RgF5hXBtAgDsg6wZELNiz8cQ3u726Lm5xtEXyKIWRakSKc3xQuiBnLbnQ9D2KY+vkVySAuUhQSmvosvyI00vOMQn+dc1iKdMswxnrqJxhwxpBglmsB2WDmgsl+Smz2qr34yAWmCTSdT2A/hXJq7lOUMCLm9Qqx+1/J9W8IAoII1Wi0coHnmRN77CzOQi1/n6U2DnrtWVN78ob62Y9J6rcxn4HyOGcPf+ITMLd84UPMpZGLPRKE9tlmGjLdEoT4f+YzBin2EoS00n7eZtAFmGC4pywXvwuKqL81LaczLicoWUu/onld8WK6fPBklPQG/kT8tVidhX/o/p+bf7n0HG823O9fq8hW/tIQpbgidA3pvhgdUE3KqpY2txAMmSaOHhewuDw4Gu1gNplPF+Pd5nw9X0KTKO4EFAC+hpGsC46RTHrMGEm7ACWiNHdyuwiferOHhfEB4U3g8cKbuA7Cm8BPC59PFuOl9QHhTeDxwpu4DsKbwE8Lb/ahDsKbwOOFN3EdhDeBB4TvTdEQAojQE+f7FVZHtGAtQg1vUi9lP9cuMIDGj/OdmSBUmpKqHMg965S7LWrso+uP0hdhtc9RBFYbgZ9laPu4SVeQeuLKzEXLWY/g+ihBqBrNBJkY/9MFs78umKq5IMYPKIlwCqvk9XcT2m/qZ8+wkFQyQYuw/26/QMqSgN+uA4gZpKC7kFoh2F+UhdZ3ceG9cMaEJq8Es3pAj4lhf71/m6Adc9L4orZ0GRdO3pb64fZp21KvhHQ8G9QMwmX7Sme3h9QOO4uzVbl4jIaZXLbKydUWk8Ell62gkNGghTiqC7GGJqSt2O0B14ijXEsef0dxIxXE2arUl6OKkLa78+Bad9slVIm6aTuXLhulk8Xksm4rHZMzNuUhCxkPWkjtDvBfyBAawGTQUeqHEKWjVBfHH0Gc+jatXnXVxF2rUl3DFuIoGJFp727T5Tug2hMg8STorzdG+3mNvyZ2wSP3ETVf+tRfFOViWL2Hv/sDUEsDBBQAAAAIAKRRG136XAFZAwMAANoNAAATAAAAeGwvdGhlbWUvdGhlbWUxLnhtbL1X23KbMBT8FUbvDTdz84RkEsduH9Jpp8kPyCBAjRAeSY6dv+8gbgKM4zR27AdLYs/ZReewwte3+5xor4hxXNAQmFcG0BCNihjTNARbkXzzwe3NNZyLDOVIozBHIVhkUHz//Qy0fU4on8MQZEJs5rrOowzlkF8VG0T3OUkKlkPBrwqW6jGDO0zTnOiWYbh6DjEFbd4lQTmigpcLEWFP0QGy8lr8YpY//I0vCNNeIQnBDtO42D2jvQAagVwsCAuBIT9A02+u9TaKiIlgJXAlP01gHRG/WDKQpes20lha/szsGCSCiDFw6ZffLqNEwChCtJajgk3HNXyrASuoangge+CZ9iBAYbDHDIF7b836ARJVDWfjG10FywenHyBR1dAZBdwZ1n1g9wMkqhq6o4DZ8s6zlv0AicoIpi9juOv5vtvAW0xSkB8H8YHrGt5Dg+9gutJqVQIqeo33K0lwhGTf5fBvwVYFFbLKUGCqibcNSmBUNigkeM2w9ojTTEgeOEfwHUDEjwL0AWeO6bsCjlAfIW3pOgZd3Qy5NbmYfCQTTMiTeCPokUtxvCA4XmFC5ERGtaXYZAvCGsIeMGWwG/M6Vcq1TcFDYIDJXNJBMBXVmus1Tz2ck23+s4jrpjdbO4BzDkV3wXAUn2gZ5CzlqoYSd7IOz57Q0dENddgn6pB3crIQ3/ywkOCoEF0pD8FUg+Up4cxqu+URJCguC1Yn6JX1LCUOZlN3ZH12a08oMc9gjJq8xpSSqWbruvAMRVakeP5hJUEwIaTcqksUWR/bAaH9mbYr+b3m7v7LLDaMiwfIswonL7XnK1VoAsP5Ahqr3JnL0ejDPURJgiIxsdJNH7mosxy8/Fl0OSm2ArGnLN5pa7Jlf2AcAsczHQNoMeaiKYAWY9a1z/j9oluHZJPB2sl7D22Fl+OWUxEr5Qyl9+e14nW6Ostx9X7UwLWm7NabfhIvcD4Gyrmk+Efgf9RTK6s897Gp6lDlTRqtPSHPvpDRdl35dYY6bNnSY5vXMTkb/IFqVm7+AVBLAwQUAAAACACkURtdDR656GUAAABzAAAAFAAAAHhsL3NoYXJlZFN0cmluZ3MueG1sBcFRCsMgDADQq0j+Z9w+xpDankXatAomFpMNj7/3lm1ycz8aWrskePoAjmTvR5UrwdfOxwe2dZlR1dzkJhpngmJ2R0TdC3FW32+Sye3sg7Op7+NCvQflQwuRccNXCG/kXAUcrn9QSwMEFAAAAAgApFEbXT3Q4S0fBwAAiCsAABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWyd2ktv21YWwPGvQnCf6O0XohQT8Sk+RM1iulYt2hZqmYakxF7aCoouamAG6KIFipk28wEGGteauE4sf4XLbzS41MNU9KdraZOQP/EcUvfccylaevXVefdYeRf2+p3opKoWXuZVJTzZj9qdk8Oq+nZw8GJH/er1q/O9s6j3bf8oDAfKeff4pL93XlWPBoPTvVyuv38Udlv9l9FpeHLePT6Iet3WoP8y6h3m+qe9sNVOwrrHuWI+v5XrtjonqkyY6N864Vl/aU/pH0VnZq/TdjsnYb+q5lVFnvqbKPpWvmy3E8q9fpXDFEZy9qCntMOD1tvjwV+jMyvsHB4NqmqhksSd7+1Hx0nAfnSsdDvyXatKt3We/H/WaQ+OqmoprypHnXY7PElOt/+2P4i6X09fKzymmYYXZ+HFRXhhe43w0iy8tFl4eRZe3iy8MguvbBa+NQvf2ix8exa+vVn4zix8Z7Pw3Vn47mbhhfx83uQ3TLCYeIUNE8ynntzYKMF88smNjRLMp5/cmDfPOvHz+Sc35hewTvMV5jNQbswSFIvrJJjPQbkxT7DzZwlyj8tIsu5orUFL7vSiM6WXHCSXnFJ5HrxYhJK1b18e85eCqvSramlXVQZVtT/oJS+9ey3+LW7FgyI+ipG4EaP4e3ErT/duetJF+JuM8De290J8FmNxL8biRvxPjOMLiq9lnf4f4kZ8FrfiXtzGl/EwvhCjeCgm8YUiHuILMRFjcRcPxYiSaplJ44v4O3EbD2fRE74oPSv+R/Gz+JUijKyID8mlT2bn/BRfifv4B/GHIu7ERNzPXvwUX1FSM7s04/i7ZGS+F+P4UtyJW5nyJzESd2IkPspiiTtKaWWl/E98IW7FzZcpb8RYPIhRfCGHelrPeEiJ7cfEuQXWCR1Cl9Aj9KdY3k5jI+tt/SLG4vdkFt6I0XTQb8REXE8n0Uhc03sJ6LzNJcwlPZZqNdnrg6pafLLVitNLL8IkEaPpNSYjfS9G8d/TrYez/M0sXeHLdL+K38QHbLYnIrCPso7H7HrG0S+wYda+eHOd9Naa79R+PD41fQkdQpfQI/SLNH1nuPXMqw1mx1dWjpfLSXwpPouJXFootvlU7Ehci4kirpOlYBwP4yslHsr7wFKmlclfes7kL2VM/uSaxbW8QchVMr6atuln8SA+TZdpJVmTZGOM5Y1E/IHdUFqv4rXSuhNQK60xAfU1L8fIOp57Ye2Lt9a5eLtE3UDoELqEHqFfom4oZXQDvrGglDGfP4iJ+CjXT3mjFGPshdKf9kJ8uUhzLT+BfHGnXmmF8nNaoZzVCsnt6XfZvElLzNpB7sr+GE5vCalbfPweW6G8Rq1r5TXvCU8cn3FfyIrAo411Lt5c+2KsNd+uXaZWIHQIXUKP0C9TK5SzbwwZ7VB+3q1BEbfpuZ3dIln5/iX+Kz9RzT5GffEJZaUtKs9pi8qiLR5H4E0FxqpGqBHqhAahSWgR2oR1QofQJfQI/QpNiQpPCeyPoJJRvX9Ol5MnPgQ3s0LlU8ZD8lB2lSp/skrFl/H75N+huI7fy1M8PSu2UsXfSk5X2FkqPmGNUCPUCQ1Ck9AitAnrhA6hS+gR+oQNwoCwuYQrI7+dGvltGnnCGqFGqBMahCahRWgT1gkdQpfQI/QJG4QBYXP7yZHfSY38Do08YY1QI9QJDUKT0CK0CeuEDqFL6BH6hA3CgLC58+TI76ZGfpdGnrBGqBHqhAahSWgR2oR1QofQJfQIfcIGYUDY3H1y5Av59N8b8zT2qDVUDVVHNVBNVAvVRq2jOqguqofqozZQA9Tmsq7WYulvvwWsBWkNVUPVUQ1UE9VCtVHrqA6qi+qh+qgN1AC1uayrtSima1HEWpDWUDVUHdVANVEtVBu1juqguqgeqo/aQA1Qm8u6WotSuhYlrAVpDVVD1VENVBPVQrVR66gOqovqofqoDdQAtbmsq7Uop2tRxlqQ1lA1VB3VQDVRLVQbtY7qoLqoHqqP2kANUJvLulqLSroWFawFaQ1VQ9VRDVQT1UK1UeuoDqqL6qH6qA3UALW5rKu1SD8wy2+aoRb4yIyqoeqoBqqJaqHaqHVUB9VF9VB91AZqgNpc1tVapB+h5Zf2UAt8iEbVUHVUA9VEtVBt1Dqqg+qieqg+agM1QG0u62ot0g/V8iCoBT5Wo2qoOqqBaqJaqDZqHdVBdVE9VB+1gRqgNpd1tRbpx+wCPmej1lA1VB3VQDVRLVQbtY7qoLqoHqqP2kANUJvLuvqjhPRzdxGfu1FrqBqqjmqgmqgWqo1aR3VQXVQP1UdtoAaozWWd12L++9D5L7O6Ye8wrIXH0x9tLfaUXngga78nf2YyDV8+8rR1GHqt3mHnpK8chweDqpp/ua0qvem3KMn2IDpNtiqq8k00GETd+d5R2GqHPblXUpWDKBosdqZnWvyo9vX/AVBLAwQUAAAAAACkURtdZfYoGSgBAAAoAQAACwAAAF9yZWxzLy5yZWxz77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIj48UmVsYXRpb25zaGlwIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvb2ZmaWNlRG9jdW1lbnQiIFRhcmdldD0iL3hsL3dvcmtib29rLnhtbCIgSWQ9IlIxODcwMTY0ZTZjNjQ0NTI4IiAvPjwvUmVsYXRpb25zaGlwcz5QSwMEFAAAAAgApFEbXSoQQUsRAQAA8gIAABoAAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc7WSO07EMBCGr2K5J3beCdrsNjS0y17AscdxtH5EtheyZ6PgSFwBsSCUIAqaNFP8I3365te8v77tDrPR6Bl8GJ3tcJpQjMByJ0Y7dPgS5V2DD/vdETSLo7NBjVNAs9E2dFjFON0TErgCw0LiJrCz0dJ5w2JInB/IxPiZDUAySivilwy8ZqLTdYL/EJ2UI4cHxy8GbPwDTEK8aggYnZgfIHaYzPo7S2ajMXoUHT4KqDNZir6VjSzaPMOIbCYUFRhY+9yir5kurGRWFnWfNnlDeSGZ2NIqKOZBPEU/2uF3W8vVQo9XvM+BirSmbdGm7ZZ6L86fgwKIa7Wf+PMAgLhsj9VpI6ise6iyIpflTY+sPnf/AVBLAwQUAAAACACkURtdjYLZqRYBAABTAwAAEwAAAFtDb250ZW50X1R5cGVzXS54bWytk0FOwzAQRa8SeYtqpywQQkm7ALaABBewnEli1R5bnmlIz8aCI3EFVAdFgJAi1G48m/F7/y/m4+292o7eFQMksgFrsZalKABNaCx2tdhzu7oW2031cohAxegdUi165nijFJkevCYZIuDoXRuS10wypE5FbXa6A3VZllfKBGRAXvGRITbVHbR677i4Hxlw0o7eieJ22juqaqFjdNZotgHVgM0vySq0rTXQBLP3gCwpJtAN9QDsncxTem3xIoPVn84Ejv4n/WolE7i8Q72NNCseB0jJNlA86cQP2kMt1OgU8cEByTM3zNAlNffgYXrXJwfImMWyvU7QPHOy2J2983f2UpDXkHb5I6k8Tu//M8zMn4OofCKbT1BLAQIUAxQAAAAIAKRRG11EdsPz3AAAADQBAAAPAAAAAAAAAAAAAACkgQAAAAB4bC93b3JrYm9vay54bWxQSwECFAMUAAAACACkURtd/r/XpDQDAABELwAADQAAAAAAAAAAAAAApIEJAQAAeGwvc3R5bGVzLnhtbFBLAQIUAxQAAAAIAKRRG136XAFZAwMAANoNAAATAAAAAAAAAAAAAACkgWgEAAB4bC90aGVtZS90aGVtZTEueG1sUEsBAhQDFAAAAAgApFEbXQ0euehlAAAAcwAAABQAAAAAAAAAAAAAAKSBnAcAAHhsL3NoYXJlZFN0cmluZ3MueG1sUEsBAhQDFAAAAAgApFEbXT3Q4S0fBwAAiCsAABgAAAAAAAAAAAAAAKSBMwgAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbFBLAQIUAxQAAAAAAKRRG11l9igZKAEAACgBAAALAAAAAAAAAAAAAACkgYgPAABfcmVscy8ucmVsc1BLAQIUAxQAAAAIAKRRG10qEEFLEQEAAPICAAAaAAAAAAAAAAAAAACkgdkQAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc1BLAQIUAxQAAAAIAKRRG12NgtmpFgEAAFMDAAATAAAAAAAAAAAAAACkgSISAABbQ29udGVudF9UeXBlc10ueG1sUEsFBgAAAAAIAAgAAwIAAGkTAAAAAA=="
def _write_embedded_template(
    encoded_data: str,
    expected_sha256: str,
    destination_path: str,
    template_id: Optional[str] = None,
) -> str:
    """Сохраняет встроенный шаблон с единым визуальным оформлением."""
    destination = _normalize_xlsx_destination(destination_path)
    data = _decode_embedded_template(encoded_data, expected_sha256)
    workbook = load_workbook(BytesIO(data), data_only=False)
    _style_download_workbook(workbook, template_id or destination.name)
    workbook.save(destination)
    return str(destination)


def _style_download_workbook(workbook, template_name: str) -> None:
    """Apply the common restrained CDE workbook style to every Larix template."""
    apply_excel_style(workbook)


def _embedded_template_cache_dir() -> Path:
    """Служебная папка шаблонов в профиле пользователя, не рядом с программой."""
    return _user_config_dir() / "embedded_templates"


def _ensure_embedded_template_cache(
    file_name: str,
    encoded_data: str,
    expected_sha256: str,
) -> str:
    """Создаёт служебную копию шаблона для внутренней работы программы.

    Копия хранится в профиле пользователя. Пользователю не нужно размещать
    никакие файлы рядом с .py или .exe.
    """
    destination = _embedded_template_cache_dir() / file_name
    expected_data = _decode_embedded_template(encoded_data, expected_sha256)

    rewrite = True
    try:
        if destination.exists() and destination.is_file():
            rewrite = hashlib.sha256(destination.read_bytes()).hexdigest() != expected_sha256
    except OSError:
        rewrite = True

    if rewrite:
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(expected_data)
    return str(destination)



APPROVAL_TEMPLATE_NAME = "Larix Platform Маршруты согласований.xlsx"
APPROVAL_TEMPLATE_SHA256 = "e24ae63ba5785ca62ea236fae21d154b2e17650cf196d5b97f4aea80cbcacb55"
APPROVAL_TEMPLATE_B64 = "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"


def _ensure_builtin_approval_template() -> str:
    """Возвращает путь к встроенному шаблону маршрутов согласований."""
    return _ensure_embedded_template_cache(
        APPROVAL_TEMPLATE_NAME, APPROVAL_TEMPLATE_B64, APPROVAL_TEMPLATE_SHA256
    )


def _ensure_builtin_structure_template() -> str:
    """Возвращает путь к встроенному эталонному шаблону структуры."""
    return _ensure_embedded_template_cache(
        STRUCTURE_TEMPLATE_NAME,
        STRUCTURE_TEMPLATE_B64,
        STRUCTURE_TEMPLATE_SHA256,
    )


def _ensure_builtin_user_template() -> str:
    """Возвращает путь к встроенному эталонному шаблону пользователей."""
    return _ensure_embedded_template_cache(
        USER_IMPORT_TEMPLATE_NAME,
        USER_IMPORT_TEMPLATE_B64,
        USER_IMPORT_TEMPLATE_SHA256,
    )


def _ensure_builtin_role_matrix_template() -> str:
    """Возвращает путь к встроенному шаблону ролевой матрицы Larix."""
    return _ensure_embedded_template_cache(
        ROLE_MATRIX_TEMPLATE_NAME,
        ROLE_MATRIX_TEMPLATE_B64,
        ROLE_MATRIX_TEMPLATE_SHA256,
    )


def _ensure_builtin_remark_type_template() -> str:
    return _ensure_embedded_template_cache(
        REMARK_TYPE_TEMPLATE_NAME, REMARK_TYPE_TEMPLATE_B64, REMARK_TYPE_TEMPLATE_SHA256
    )


def _ensure_builtin_task_type_template() -> str:
    return _ensure_embedded_template_cache(
        TASK_TYPE_TEMPLATE_NAME, TASK_TYPE_TEMPLATE_B64, TASK_TYPE_TEMPLATE_SHA256
    )


# =====================================================================
class NoWheelComboBox(QtWidgets.QComboBox):
    """QComboBox, который не переключает значение колесом в закрытом состоянии.

    Это защищает проект/пространство/лист Excel от случайного изменения, когда
    пользователь просто прокручивает страницу. После открытия popup колесо
    обрабатывается самим списком как обычно.
    """

    def wheelEvent(self, event: QtGui.QWheelEvent) -> None:
        view = self.view()
        if view is not None and view.isVisible():
            super().wheelEvent(event)
            return
        event.ignore()


class NoWheelTabBar(QtWidgets.QTabBar):
    """Keep tab selection deliberate; wheel scrolling belongs to tab content."""
    def wheelEvent(self, event: QtGui.QWheelEvent) -> None:
        event.accept()


class PasswordLineEdit(QtWidgets.QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._eye_btn = QtWidgets.QToolButton(self)
        self._eye_btn.setObjectName("passwordEye")
        self._eye_btn.setCursor(QtCore.Qt.PointingHandCursor)
        self._eye_btn.setFixedSize(30, 28)
        self._eye_btn.setIconSize(QtCore.QSize(18, 18))
        self.setTextMargins(0, 0, 38, 0)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        x = self.width() - self._eye_btn.width() - 8
        y = (self.height() - self._eye_btn.height()) // 2
        self._eye_btn.move(x, y)

    def set_eye_clicked(self, callback) -> None:
        self._eye_btn.clicked.connect(callback)

    def set_eye_tooltip(self, text: str) -> None:
        self._eye_btn.setToolTip(text)


class RoundedComboDelegate(QtWidgets.QStyledItemDelegate):
    def paint(self, painter: QtGui.QPainter, option: QtWidgets.QStyleOptionViewItem, index: QtCore.QModelIndex) -> None:
        opt = QtWidgets.QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        # Popup QComboBox создаётся Qt как отдельное служебное окно, поэтому
        # opt.widget.window() не является MainWindow и не содержит is_dark_theme.
        # Тему берём из общего свойства QApplication, которое обновляется в
        # MainWindow._apply_styles(). Это сохраняет читаемый текст и при hover.
        app = QtWidgets.QApplication.instance()
        dark = bool(app.property("larix_dark")) if app is not None else False
        text_color = QtGui.QColor("#e0e0e0" if dark else "#222222")
        hover_color = QtGui.QColor(247, 146, 30, 46 if dark else 28)
        selected_color = QtGui.QColor(247, 146, 30, 82 if dark else 64)
        rect = opt.rect.adjusted(4, 2, -4, -2)
        is_selected = bool(opt.state & QtWidgets.QStyle.State_Selected)
        is_hover = bool(opt.state & QtWidgets.QStyle.State_MouseOver)
        painter.save()
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
        painter.setPen(QtCore.Qt.NoPen)
        if is_selected:
            painter.setBrush(selected_color)
            painter.drawRoundedRect(rect, 8, 8)
        elif is_hover:
            painter.setBrush(hover_color)
            painter.drawRoundedRect(rect, 8, 8)
        painter.setPen(text_color)
        text_rect = opt.rect.adjusted(10, 0, -10, 0)
        text = opt.fontMetrics.elidedText(opt.text, QtCore.Qt.ElideRight, text_rect.width())
        painter.drawText(text_rect, QtCore.Qt.AlignVCenter | QtCore.Qt.AlignLeft, text)
        painter.restore()

    def sizeHint(self, option: QtWidgets.QStyleOptionViewItem, index: QtCore.QModelIndex) -> QtCore.QSize:
        size = super().sizeHint(option, index)
        size.setHeight(max(size.height(), 28))
        return size


class TableRowOverlayDelegate(QtWidgets.QStyledItemDelegate):
    """Рисует мягкую смысловую подсветку поверх строки таблицы."""

    COLOR_ROLE = QtCore.Qt.UserRole + 317

    def paint(
        self,
        painter: QtGui.QPainter,
        option: QtWidgets.QStyleOptionViewItem,
        index: QtCore.QModelIndex,
    ) -> None:
        super().paint(painter, option, index)
        color = index.data(self.COLOR_ROLE)
        if isinstance(color, QtGui.QColor) and color.isValid() and color.alpha() > 0:
            painter.save()
            painter.setCompositionMode(QtGui.QPainter.CompositionMode_SourceOver)
            painter.fillRect(option.rect, color)
            painter.restore()


# =====================================================================
#  Utility helpers
# =====================================================================
def _decode_jwt_payload(token: str) -> Optional[Dict]:
    try:
        if not token or not isinstance(token, str):
            return None
        parts = token.split('.')
        if len(parts) != 3:
            return None
        payload_b64 = parts[1]
        padding = 4 - len(payload_b64) % 4
        if padding != 4:
            payload_b64 += '=' * padding
        payload_json = base64.urlsafe_b64decode(payload_b64)
        data = json.loads(payload_json)
        return data if isinstance(data, dict) else None
    except Exception:
        return None


def _extract_workspace_id(ws: Dict) -> Optional[int]:
    if not isinstance(ws, dict):
        return None
    for key in ["id", "workspace_id", "workspaceId"]:
        value = ws.get(key)
        if value is not None:
            try:
                return int(value)
            except (ValueError, TypeError):
                continue
    return None


def _extract_workspace_name(ws: Dict) -> str:
    if not isinstance(ws, dict):
        return "Без названия"
    for key in ["name", "title", "workspace_name", "workspaceName"]:
        value = ws.get(key)
        if value and isinstance(value, str) and value.strip():
            return value.strip()
    ws_id = _extract_workspace_id(ws)
    if ws_id:
        return f"Workspace {ws_id}"
    return "Без названия"


def _extract_data_list(payload) -> List[Dict]:
    if isinstance(payload, list):
        return [x for x in payload if isinstance(x, dict)]
    if not isinstance(payload, dict):
        return []
    direct_keys = ["data", "items", "workspaces", "users", "projects", "roles", "result", "list"]
    for key in direct_keys:
        value = payload.get(key)
        if isinstance(value, list):
            return [x for x in value if isinstance(x, dict)]
    data = payload.get("data")
    if isinstance(data, dict):
        for key in ["items", "list", "users", "projects", "roles", "workspaces", "result"]:
            value = data.get(key)
            if isinstance(value, list):
                return [x for x in value if isinstance(x, dict)]
        if any(key in data for key in ["id", "projectId", "workspaceId", "ac_user_id", "acsUserId"]):
            return [data]
    if any(key in payload for key in ["id", "projectId", "workspaceId", "workspace_id", "ac_user_id", "acsUserId"]):
        return [payload]
    return []


def _extract_project_id(project: Dict) -> Optional[int]:
    if not isinstance(project, dict):
        return None
    for key in ["id", "projectId", "project_id"]:
        value = project.get(key)
        if value is not None:
            try:
                return int(value)
            except (TypeError, ValueError):
                continue
    return None


def _extract_project_name(project: Dict) -> str:
    if not isinstance(project, dict):
        return "Без названия"
    for key in ["title", "name", "projectTitle", "project_name"]:
        value = project.get(key)
        if value and isinstance(value, str) and value.strip():
            return value.strip()
    project_id = _extract_project_id(project)
    return f"Project {project_id}" if project_id else "Без названия"


def _to_int(value: Any) -> Optional[int]:
    try:
        if value is None or value == "":
            return None
        return int(value)
    except (TypeError, ValueError):
        return None


def _extract_by_keys(data: Any, keys: List[str]) -> Optional[int]:
    if not isinstance(data, dict):
        return None
    for key in keys:
        value = _to_int(data.get(key))
        if value is not None:
            return value
    return None


def _extract_ac_user_id(data: Any, allow_plain_id: bool = False) -> Optional[int]:
    preferred_keys = [
        "ac_user_id", "acsUserId", "acs_user_id", "acUserId",
        "userId", "user_id", "acsId", "acs_id"
    ]
    fallback_keys = preferred_keys + (["id", "Id"] if allow_plain_id else [])
    value = _extract_by_keys(data, fallback_keys)
    if value is not None:
        return value
    if isinstance(data, dict):
        for nested_key in ["workspace_user", "workspaceUser", "user", "data", "account"]:
            nested = data.get(nested_key)
            value = _extract_by_keys(nested, fallback_keys)
            if value is not None:
                return value
    return None


def _looks_like_missing_user_error(data: Any) -> bool:
    if isinstance(data, dict):
        if data.get("need_create") is True or data.get("to_be_added") is True:
            return True
        inner = data.get("data")
        if isinstance(inner, dict) and inner.get("to_be_added") is True:
            return True
    text = str(data or "").lower()
    return any(marker in text for marker in [
        "не существует", "не найден", "not found", "not exist", "to_be_added", "user not"
    ])


def _looks_like_already_in_workspace_error(data: Any) -> bool:
    """Проверяет, что API ругается не на ошибку данных, а на уже добавленного пользователя."""
    text = str(data or "").lower()
    return any(marker in text for marker in [
        "уже добав", "уже существует", "уже есть", "уже состоит", "уже находится",
        "already added", "already exists", "already in", "already assigned", "already a member",
        "user already", "duplicate"
    ])


def _extract_folder_id(data: Dict) -> Optional[int]:
    if not isinstance(data, dict):
        return None
    for source in [data, data.get("data") if isinstance(data.get("data"), dict) else None]:
        if not isinstance(source, dict):
            continue
        for key in ["id", "Id", "folderId", "folder_id"]:
            value = _to_int(source.get(key))
            if value is not None:
                return value
    return None


# =====================================================================
#  Larix API Client
# =====================================================================
TYPE_IMPORT_SHEETS = ("Типы", "Права ролей", "Права пользователей")
TYPE_IMPORT_RIGHT_CODES = {"ПР", "С", "ПРС", "-"}


class TypeImportValidationError(ValueError):
    def __init__(self, errors: List[str]):
        self.errors = list(errors)
        super().__init__("; ".join(self.errors))


@dataclass
class TypeAccessPlan:
    name: str
    role_view: set[str] = field(default_factory=set)
    role_create: set[str] = field(default_factory=set)
    user_view: set[str] = field(default_factory=set)
    user_create: set[str] = field(default_factory=set)

    def as_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "role_view": sorted(self.role_view), "role_create": sorted(self.role_create),
            "user_view": sorted(self.user_view), "user_create": sorted(self.user_create),
        }


@dataclass
class TypeImportPlan:
    kind: str
    source_path: str
    types: List[TypeAccessPlan]

    @property
    def type_names(self) -> List[str]:
        return [item.name for item in self.types]

    def as_dict(self) -> Dict[str, Any]:
        return {"kind": self.kind, "source_path": self.source_path, "type_names": self.type_names,
                "types": [item.as_dict() for item in self.types]}


def _type_import_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return re.sub(r"\s+", " ", str(value).strip())


def _type_import_header(df: pd.DataFrame, required: List[str], sheet: str) -> tuple[int, Dict[str, int]]:
    wanted = [_type_import_text(item).lower() for item in required]
    for row_index, row in df.iterrows():
        headers = [_type_import_text(value).lower() for value in row.tolist()]
        if all(item in headers for item in wanted):
            return int(row_index), {item: headers.index(item) for item in wanted}
    raise TypeImportValidationError([f"Лист «{sheet}»: не найдены заголовки: {', '.join(required)}"])


def _type_import_data_rows(df: pd.DataFrame, header: int, columns: List[int]) -> List[tuple[int, List[str]]]:
    result: List[tuple[int, List[str]]] = []
    for index in range(header + 1, len(df)):
        values = [_type_import_text(df.iat[index, col]) for col in columns]
        if not any(values):
            break
        result.append((index + 1, values))
    return result


def validate_type_import_plan(plan: TypeImportPlan, known_roles: Optional[List[str]] = None,
                              known_users: Optional[List[str]] = None) -> List[str]:
    errors: List[str] = []
    if known_roles is not None:
        role_set = {_type_import_text(value) for value in known_roles}
        for item in plan.types:
            for role in item.role_view | item.role_create:
                if role not in role_set:
                    errors.append(f"Тип «{item.name}»: роль «{role}» не найдена в выбранном проекте Larix")
    if known_users is not None:
        user_set = {_type_import_text(value).lower() for value in known_users}
        for item in plan.types:
            for email in item.user_view | item.user_create:
                if email.lower() not in user_set:
                    errors.append(f"Тип «{item.name}»: пользователь «{email}» не найден в текущем пространстве Larix")
    return errors


def parse_type_import_excel(excel_path: str, kind: str = "", known_roles: Optional[List[str]] = None,
                           known_users: Optional[List[str]] = None) -> TypeImportPlan:
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel-файл не найден: {excel_path}")
    try:
        book = pd.ExcelFile(excel_path)
        sheets = list(book.sheet_names)
        missing = [name for name in TYPE_IMPORT_SHEETS if name not in sheets]
        unknown = [name for name in sheets if name not in TYPE_IMPORT_SHEETS]
        if missing or unknown:
            messages = []
            if missing:
                messages.append(f"Отсутствуют обязательные листы: {', '.join(missing)}")
            if unknown:
                messages.append(f"Неизвестные листы: {', '.join(unknown)}")
            raise TypeImportValidationError(messages)

        types_df = pd.read_excel(book, sheet_name="Типы", header=None, dtype=object)
        type_header, type_cols = _type_import_header(types_df, ["Название типа"], "Типы")
        type_rows = _type_import_data_rows(types_df, type_header, [type_cols["название типа"]])
        names: List[str] = []
        errors: List[str] = []
        for row_number, values in type_rows:
            name = values[0]
            if not name:
                errors.append(f"Типы, строка {row_number}: название типа обязательно")
            elif name in names:
                errors.append(f"Типы, строка {row_number}: дубликат типа «{name}»")
            else:
                names.append(name)
        if not names:
            errors.append("Лист «Типы»: не указан ни один тип")
        plans = {name: TypeAccessPlan(name) for name in names}
        role_df = pd.read_excel(book, sheet_name="Права ролей", header=None, dtype=object)
        user_df = pd.read_excel(book, sheet_name="Права пользователей", header=None, dtype=object)
        role_header, role_cols = _type_import_header(role_df, ["Название типа", "Роль", "Права"], "Права ролей")
        user_header, user_cols = _type_import_header(user_df, ["Название типа", "E-mail пользователя", "Права"], "Права пользователей")
        seen_roles: set[tuple[str, str]] = set()
        seen_users: set[tuple[str, str]] = set()
        for row_number, values in _type_import_data_rows(role_df, role_header, [role_cols[item] for item in ("название типа", "роль", "права")]):
            type_name, role, code = values
            if type_name not in plans:
                errors.append(f"Права ролей, строка {row_number}: неизвестный тип «{type_name}»"); continue
            key = (type_name, role)
            if not role or key in seen_roles:
                errors.append(f"Права ролей, строка {row_number}: пустая роль или дубликат записи"); continue
            seen_roles.add(key)
            code = code.upper().replace(" ", "")
            if code not in TYPE_IMPORT_RIGHT_CODES:
                errors.append(f"Права ролей, строка {row_number}: неизвестный код прав «{code}»"); continue
            if code in {"ПР", "ПРС"}: plans[type_name].role_view.add(role)
            if code in {"С", "ПРС"}: plans[type_name].role_create.add(role)
        for row_number, values in _type_import_data_rows(user_df, user_header, [user_cols[item] for item in ("название типа", "e-mail пользователя", "права")]):
            type_name, email, code = values
            if type_name not in plans:
                errors.append(f"Права пользователей, строка {row_number}: неизвестный тип «{type_name}»"); continue
            key = (type_name, email.lower())
            if not email or key in seen_users:
                errors.append(f"Права пользователей, строка {row_number}: пустой e-mail или дубликат записи"); continue
            seen_users.add(key)
            if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
                errors.append(f"Права пользователей, строка {row_number}: некорректный e-mail «{email}»"); continue
            code = code.upper().replace(" ", "")
            if code not in TYPE_IMPORT_RIGHT_CODES:
                errors.append(f"Права пользователей, строка {row_number}: неизвестный код прав «{code}»"); continue
            if code in {"ПР", "ПРС"}: plans[type_name].user_view.add(email)
            if code in {"С", "ПРС"}: plans[type_name].user_create.add(email)
        if errors:
            raise TypeImportValidationError(errors)
        plan = TypeImportPlan(kind=kind, source_path=os.path.abspath(excel_path), types=list(plans.values()))
        catalog_errors = validate_type_import_plan(plan, known_roles, known_users)
        if catalog_errors:
            raise TypeImportValidationError(catalog_errors)
        return plan
    except TypeImportValidationError:
        raise
    except (ValueError, KeyError, OSError) as exc:
        raise TypeImportValidationError([f"Не удалось прочитать Excel: {exc}"]) from exc


def _type_catalog_display_name(value: Any) -> str:
    """Return the exact role/user label exposed by a Larix catalog row."""
    if isinstance(value, dict):
        for key in ("fullName", "full_name", "name", "title", "surname"):
            candidate = _type_import_text(value.get(key))
            if candidate:
                return candidate
        for key in ("user", "profile", "account"):
            nested = value.get(key)
            candidate = _type_catalog_display_name(nested)
            if candidate:
                return candidate
    return _type_import_text(value)


def _type_catalog_names(values: Optional[List[Any]]) -> set[str]:
    return {_type_catalog_display_name(item) for item in (values or []) if _type_catalog_display_name(item)}


def validate_type_import_plan(plan: TypeImportPlan, known_roles: Optional[List[Any]] = None,
                              known_users: Optional[List[Any]] = None) -> List[str]:
    errors: List[str] = []
    role_set = _type_catalog_names(known_roles) if known_roles is not None else None
    user_set = _type_catalog_names(known_users) if known_users is not None else None
    for item in plan.types:
        if role_set is not None:
            for role in item.role_view | item.role_create:
                if role not in role_set:
                    errors.append(f"Тип «{item.name}»: неизвестная роль «{role}»")
        if user_set is not None:
            for user in item.user_view | item.user_create:
                if user not in user_set:
                    errors.append(f"Тип «{item.name}»: неизвестный пользователь «{user}»")
    return errors


def parse_type_import_excel(excel_path: str, kind: str = "", known_roles: Optional[List[Any]] = None,
                            known_users: Optional[List[Any]] = None, sheet_name: Optional[str] = None) -> TypeImportPlan:
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel-файл не найден: {excel_path}")
    expected_prefix = "Тип задачи" if str(kind).lower() in {"tasks", "task", "задачи", "задача"} else "Тип замечания"
    try:
        book = pd.ExcelFile(excel_path)
        sheets = list(book.sheet_names)
        if sheet_name is not None and sheet_name not in sheets:
            raise TypeImportValidationError([f"Лист «{sheet_name}» не найден. Доступны: {', '.join(sheets)}"])
        if sheet_name is None and len(sheets) != 1:
            raise TypeImportValidationError([f"Ожидается один рабочий лист, найдены: {', '.join(sheets)}"])
        sheet = sheet_name or sheets[0]
        df = pd.read_excel(book, sheet_name=sheet, header=None, dtype=object)
        header_row = header_col = None
        for row_index in range(len(df)):
            for col_index in range(df.shape[1]):
                text = _type_import_text(df.iat[row_index, col_index])
                if text.lower().startswith(expected_prefix.lower()):
                    header_row, header_col = row_index, col_index
                    break
            if header_row is not None:
                break
        if header_row is None:
            raise TypeImportValidationError([f"Не найден заголовок «{expected_prefix}» на листе «{sheet}»"])

        recipient_cols: List[tuple[int, str]] = []
        for col_index in range(header_col + 1, df.shape[1]):
            label = _type_import_text(df.iat[header_row, col_index])
            if not label:
                break
            recipient_cols.append((col_index, label))
        if not recipient_cols:
            raise TypeImportValidationError([f"Лист «{sheet}»: после первого столбца нет получателей"])

        errors: List[str] = []
        seen_recipients: set[str] = set()
        role_set = _type_catalog_names(known_roles) if known_roles is not None else None
        user_set = _type_catalog_names(known_users) if known_users is not None else None
        for _, recipient in recipient_cols:
            if recipient in seen_recipients:
                errors.append(f"Лист «{sheet}»: дубликат получателя «{recipient}»")
            seen_recipients.add(recipient)
            if role_set is not None and user_set is not None and recipient in role_set and recipient in user_set:
                errors.append(f"Лист «{sheet}»: получатель «{recipient}» совпадает с ролью и пользователем")
            elif role_set is not None and recipient not in role_set and user_set is not None and recipient not in user_set:
                errors.append(f"Лист «{sheet}»: неизвестный получатель «{recipient}»")
            elif role_set is not None and user_set is None and recipient not in role_set:
                errors.append(f"Лист «{sheet}»: неизвестная роль «{recipient}»")
            elif user_set is not None and role_set is None and recipient not in user_set:
                errors.append(f"Лист «{sheet}»: неизвестный пользователь «{recipient}»")

        plans: Dict[str, TypeAccessPlan] = {}
        for row_index in range(header_row + 1, len(df)):
            name = _type_import_text(df.iat[row_index, header_col])
            values = [_type_import_text(df.iat[row_index, col]) for col, _ in recipient_cols]
            if not name and not any(values):
                continue
            if name.strip().lower().rstrip(":") in {"важно", "примечание", "подсказка"}:
                continue
            if not name:
                errors.append(f"Лист «{sheet}», строка {row_index + 1}: пустое название типа")
                continue
            if name in plans:
                errors.append(f"Лист «{sheet}», строка {row_index + 1}: дубликат типа «{name}»")
                continue
            item = TypeAccessPlan(name)
            plans[name] = item
            for (_, recipient), raw_code in zip(recipient_cols, values):
                code = raw_code.upper().replace(" ", "")
                if not code:
                    errors.append(f"Лист «{sheet}», строка {row_index + 1}: пустой код для «{recipient}»")
                    continue
                if code not in TYPE_IMPORT_RIGHT_CODES:
                    errors.append(f"Лист «{sheet}», строка {row_index + 1}: неизвестный код прав «{raw_code}»")
                    continue
                if code in {"ПР", "ПРС"}:
                    if role_set is not None and recipient in role_set:
                        item.role_view.add(recipient)
                    else:
                        item.user_view.add(recipient)
                if code in {"С", "ПРС"}:
                    if role_set is not None and recipient in role_set:
                        item.role_create.add(recipient)
                    else:
                        item.user_create.add(recipient)
        if errors:
            raise TypeImportValidationError(errors)
        plan = TypeImportPlan(kind=kind, source_path=os.path.abspath(excel_path), types=list(plans.values()))
        catalog_errors = validate_type_import_plan(plan, known_roles, known_users)
        if catalog_errors:
            raise TypeImportValidationError(catalog_errors)
        return plan
    except TypeImportValidationError:
        raise
    except (ValueError, KeyError, OSError) as exc:
        raise TypeImportValidationError([f"Не удалось прочитать Excel: {exc}"])



# =====================================================================
#  Маршруты согласований Larix
# =====================================================================
APPROVAL_ACCESS_VIEW_ID = 8
APPROVAL_ACCESS_CREATE_ID = 10
APPROVAL_MAIN_SHEET = "5. Маршруты согласований"
APPROVAL_SETTINGS_SHEET = "5.1 Настройка_Согласований"
APPROVAL_FLAGS_SHEET = "5.2 Настройка_Согласований"
APPROVAL_MAX_STEPS = 7
APPROVAL_USERS_PER_STEP = 2


class ApprovalImportValidationError(ValueError):
    def __init__(self, errors: List[str]):
        self.errors = [str(item) for item in errors if str(item).strip()]
        super().__init__("\n".join(self.errors))


def _approval_clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).replace("\xa0", " ").replace("\r", " ").replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return "" if text.lower() in {"nan", "none", "null"} else text


def _approval_norm(value: Any) -> str:
    text = _approval_clean_text(value).lower().replace("ё", "е")
    text = re.sub(r"[^\w@.+-]+", " ", text, flags=re.UNICODE)
    return re.sub(r"\s+", " ", text).strip()


def _approval_user_norm(value: Any) -> str:
    text = _approval_clean_text(value)
    # В Larix UI к ФИО иногда визуально дописывается "Личный".
    # В Excel это может попасть слитно с именем: "НикитаЛичный".
    if text.lower().endswith("личный") and len(text) > len("личный"):
        text = text[:-len("личный")].rstrip(" -–—()")
    return _approval_norm(text)


def _approval_split_principals(value: Any) -> List[str]:
    text = _approval_clean_text(value)
    if not text:
        return []
    return [
        item.strip()
        for item in re.split(r"[,;\n|]+", text)
        if item and item.strip()
    ]


def _approval_bool(value: Any, location: str, errors: List[str]) -> bool:
    text = _approval_norm(value)
    if not text:
        return False
    if text in {"да", "yes", "true", "1", "д", "y"}:
        return True
    if text in {"нет", "no", "false", "0", "н", "n"}:
        return False
    errors.append(f"{location}: ожидается «Да»/«Нет», получено «{_approval_clean_text(value)}»")
    return False


def _approval_duration(value: Any, location: str, errors: List[str]) -> Optional[int]:
    text = _approval_clean_text(value)
    if not text:
        errors.append(f"{location}: не указана продолжительность этапа")
        return None
    try:
        number = float(str(value).replace(",", "."))
        if not number.is_integer() or number <= 0:
            raise ValueError
        return int(number)
    except (TypeError, ValueError):
        errors.append(f"{location}: продолжительность должна быть целым числом больше 0")
        return None


def _approval_find_sheet(sheet_names: List[str], expected: str) -> Optional[str]:
    if expected in sheet_names:
        return expected
    expected_norm = _approval_norm(expected)
    for item in sheet_names:
        if _approval_norm(item) == expected_norm:
            return item
    return None


def parse_approval_workflows_excel(
    excel_path: str,
    main_sheet_name: Optional[str] = None,
    settings_sheet_name: Optional[str] = None,
    flags_sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Читает выбранные три листа маршрутов и возвращает независимый план без ID Larix."""
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel-файл не найден: {excel_path}")
    errors: List[str] = []
    try:
        book = pd.ExcelFile(excel_path)
        sheet_names = list(book.sheet_names)

        def resolve_manual(requested: Optional[str], expected: str, fallback_index: int) -> Optional[str]:
            requested_text = _approval_clean_text(requested)
            if requested_text:
                if requested_text not in sheet_names:
                    raise ApprovalImportValidationError([
                        f"Выбранный лист «{requested_text}» не найден. Доступны: {', '.join(sheet_names)}"
                    ])
                return requested_text
            auto = _approval_find_sheet(sheet_names, expected)
            if auto:
                return auto
            if 0 <= fallback_index < len(sheet_names):
                return sheet_names[fallback_index]
            return None

        main_sheet = resolve_manual(main_sheet_name, APPROVAL_MAIN_SHEET, 0)
        settings_sheet = resolve_manual(settings_sheet_name, APPROVAL_SETTINGS_SHEET, 1)
        flags_sheet = resolve_manual(flags_sheet_name, APPROVAL_FLAGS_SHEET, 2)
        missing = [
            label for label, value in (
                (APPROVAL_MAIN_SHEET, main_sheet),
                (APPROVAL_SETTINGS_SHEET, settings_sheet),
                (APPROVAL_FLAGS_SHEET, flags_sheet),
            ) if value is None
        ]
        if missing:
            raise ApprovalImportValidationError(
                [f"Не удалось определить лист «{name}»" for name in missing]
            )
        if len({main_sheet, settings_sheet, flags_sheet}) < 3:
            raise ApprovalImportValidationError([
                "Для маршрутов, доступа/длительности и настроек согласующих нужно выбрать три разных листа Excel"
            ])

        main_df = pd.read_excel(book, sheet_name=main_sheet, header=None, dtype=object)
        settings_df = pd.read_excel(book, sheet_name=settings_sheet, header=None, dtype=object)
        flags_df = pd.read_excel(book, sheet_name=flags_sheet, header=None, dtype=object)

        settings_rows: Dict[str, int] = {}
        for row_index in range(3, len(settings_df)):
            title = _approval_clean_text(settings_df.iat[row_index, 0] if settings_df.shape[1] else "")
            if not title:
                continue
            key = _approval_norm(title)
            if key in settings_rows:
                errors.append(f"Лист «{settings_sheet}»: дубликат маршрута «{title}»")
            else:
                settings_rows[key] = row_index

        flags_rows: Dict[str, int] = {}
        for row_index in range(4, len(flags_df)):
            title = _approval_clean_text(flags_df.iat[row_index, 0] if flags_df.shape[1] else "")
            if not title:
                continue
            key = _approval_norm(title)
            if key in flags_rows:
                errors.append(f"Лист «{flags_sheet}»: дубликат маршрута «{title}»")
            else:
                flags_rows[key] = row_index

        workflows: List[Dict[str, Any]] = []
        seen_titles: set[str] = set()
        for row_index in range(3, len(main_df)):
            title = _approval_clean_text(main_df.iat[row_index, 0] if main_df.shape[1] else "")
            row_has_users = any(
                _approval_clean_text(main_df.iat[row_index, col])
                for col in range(1, min(main_df.shape[1], 1 + APPROVAL_MAX_STEPS * APPROVAL_USERS_PER_STEP))
            )
            if not title and not row_has_users:
                continue
            if not title:
                errors.append(f"Лист «{main_sheet}», строка {row_index + 1}: пустое название маршрута")
                continue
            key = _approval_norm(title)
            if key in seen_titles:
                errors.append(f"Лист «{main_sheet}», строка {row_index + 1}: дубликат маршрута «{title}»")
                continue
            seen_titles.add(key)

            route_errors: List[str] = []
            settings_row = settings_rows.get(key)
            flags_row = flags_rows.get(key)
            if settings_row is None:
                route_errors.append(f"нет строки маршрута на листе «{settings_sheet}»")
            if flags_row is None:
                route_errors.append(f"нет строки маршрута на листе «{flags_sheet}»")

            view_principals: List[str] = []
            create_principals: List[str] = []
            if settings_row is not None:
                if settings_df.shape[1] > 1:
                    view_principals = _approval_split_principals(settings_df.iat[settings_row, 1])
                if settings_df.shape[1] > 2:
                    create_principals = _approval_split_principals(settings_df.iat[settings_row, 2])

            steps: List[Dict[str, Any]] = []
            gap_seen = False
            for step_index in range(APPROVAL_MAX_STEPS):
                base_col = 1 + step_index * APPROVAL_USERS_PER_STEP
                users_raw: List[str] = []
                for user_pos in range(APPROVAL_USERS_PER_STEP):
                    col = base_col + user_pos
                    value = _approval_clean_text(main_df.iat[row_index, col]) if col < main_df.shape[1] else ""
                    users_raw.append(value)
                present_users = [name for name in users_raw if name]
                if not present_users:
                    if steps:
                        gap_seen = True
                    continue
                if gap_seen:
                    route_errors.append(f"{step_index + 1} этап заполнен после пропущенного этапа")

                duration: Optional[int] = None
                if settings_row is not None:
                    duration_col = 3 + step_index
                    if duration_col < settings_df.shape[1]:
                        duration = _approval_duration(
                            settings_df.iat[settings_row, duration_col],
                            f"маршрут «{title}», {step_index + 1} этап",
                            route_errors,
                        )
                    else:
                        route_errors.append(f"маршрут «{title}», {step_index + 1} этап: нет колонки длительности")

                step_title = ""
                if main_df.shape[0] > 0 and base_col < main_df.shape[1]:
                    step_title = _approval_clean_text(main_df.iat[0, base_col])
                if not step_title:
                    step_title = f"{step_index + 1} Этап"

                users: List[Dict[str, Any]] = []
                seen_step_users: set[str] = set()
                for user_pos, user_name in enumerate(users_raw):
                    if not user_name:
                        continue
                    user_key = _approval_user_norm(user_name)
                    if user_key in seen_step_users:
                        route_errors.append(f"{step_index + 1} этап: согласующий «{user_name}» указан дважды")
                        continue
                    seen_step_users.add(user_key)

                    can_cancel = False
                    must_approve = False
                    if flags_row is not None:
                        flag_col = 1 + step_index * 4 + user_pos * 2
                        if flag_col < flags_df.shape[1]:
                            can_cancel = _approval_bool(
                                flags_df.iat[flags_row, flag_col],
                                f"маршрут «{title}», {step_index + 1} этап, {user_pos + 1} согласующий, отмена",
                                route_errors,
                            )
                        if flag_col + 1 < flags_df.shape[1]:
                            must_approve = _approval_bool(
                                flags_df.iat[flags_row, flag_col + 1],
                                f"маршрут «{title}», {step_index + 1} этап, {user_pos + 1} согласующий, обязательность",
                                route_errors,
                            )
                    users.append({
                        "source_name": user_name,
                        "can_cancel": bool(can_cancel),
                        "must_approve": bool(must_approve),
                    })

                steps.append({
                    "index": step_index,
                    "title": step_title,
                    "duration": duration,
                    "users": users,
                })

            if not steps:
                route_errors.append("не заполнен ни один этап")
            workflows.append({
                "title": title,
                "steps": steps,
                "view_principals": view_principals,
                "create_principals": create_principals,
                "errors": route_errors,
                "source_row": row_index + 1,
            })

        for key, row_index in settings_rows.items():
            if key not in seen_titles:
                extra = _approval_clean_text(settings_df.iat[row_index, 0])
                errors.append(f"Лист «{settings_sheet}»: маршрут «{extra}» отсутствует на основном листе")
        for key, row_index in flags_rows.items():
            if key not in seen_titles:
                extra = _approval_clean_text(flags_df.iat[row_index, 0])
                errors.append(f"Лист «{flags_sheet}»: маршрут «{extra}» отсутствует на основном листе")

        if errors:
            raise ApprovalImportValidationError(errors)
        if not workflows:
            raise ApprovalImportValidationError(["В шаблоне не найдено ни одного маршрута"])
        return {
            "source_path": os.path.abspath(excel_path),
            "workflows": workflows,
            "sheets": {
                "main": main_sheet,
                "settings": settings_sheet,
                "flags": flags_sheet,
            },
        }
    except ApprovalImportValidationError:
        raise
    except (ValueError, KeyError, OSError) as exc:
        raise ApprovalImportValidationError([f"Не удалось прочитать Excel: {exc}"]) from exc


def _approval_catalog_data(result: Dict[str, Any]) -> Dict[str, Any]:
    payload = result.get("data")
    if not isinstance(payload, dict):
        return {}
    nested = payload.get("data")
    return nested if isinstance(nested, dict) else payload


def _approval_role_aliases(title: str) -> set[str]:
    norm = _approval_norm(title)
    aliases = {norm} if norm else set()
    words = re.findall(r"[A-Za-zА-Яа-яЁё0-9]+", _approval_clean_text(title))
    if len(words) >= 2:
        acronym = "".join(word[0] for word in words if word)
        if acronym:
            aliases.add(_approval_norm(acronym))
    if norm == "руководитель проекта":
        aliases.update({"рп", "рук проекта"})
    return {item for item in aliases if item}


def _approval_build_catalogs(roles_result: Dict[str, Any]) -> Dict[str, Any]:
    data = _approval_catalog_data(roles_result)
    roles = [item for item in data.get("roles", []) if isinstance(item, dict)]
    users = [item for item in data.get("users", []) if isinstance(item, dict)]

    role_index: Dict[str, List[Dict[str, Any]]] = {}
    for item in roles:
        title = _approval_clean_text(item.get("title"))
        try:
            item_id = int(item.get("id"))
        except (TypeError, ValueError):
            continue
        normalized_item = {"id": item_id, "title": title}
        for alias in _approval_role_aliases(title):
            role_index.setdefault(alias, []).append(normalized_item)

    user_index: Dict[str, List[Dict[str, Any]]] = {}
    for item in users:
        name = _approval_clean_text(item.get("name") or item.get("title"))
        try:
            item_id = int(item.get("id"))
        except (TypeError, ValueError):
            continue
        normalized_item = {"id": item_id, "name": name}
        key = _approval_user_norm(name)
        if key:
            user_index.setdefault(key, []).append(normalized_item)
    return {
        "roles": roles,
        "users": users,
        "role_index": role_index,
        "user_index": user_index,
    }


def _approval_resolve_user(name: str, catalogs: Dict[str, Any]) -> Dict[str, Any]:
    key = _approval_user_norm(name)
    matches = catalogs.get("user_index", {}).get(key, [])
    if len(matches) == 1:
        return {"ok": True, "item": matches[0]}
    if len(matches) > 1:
        return {"ok": False, "error": f"неоднозначный пользователь «{name}»"}
    return {"ok": False, "error": f"пользователь «{name}» не найден в выбранном проекте"}


def _approval_resolve_access_principal(name: str, catalogs: Dict[str, Any]) -> Dict[str, Any]:
    role_matches = catalogs.get("role_index", {}).get(_approval_norm(name), [])
    user_matches = catalogs.get("user_index", {}).get(_approval_user_norm(name), [])
    total = len(role_matches) + len(user_matches)
    if total == 1:
        if role_matches:
            return {"ok": True, "kind": "role", "item": role_matches[0]}
        return {"ok": True, "kind": "user", "item": user_matches[0]}
    if total > 1:
        return {"ok": False, "error": f"получатель доступа «{name}» найден неоднозначно"}
    return {"ok": False, "error": f"роль/пользователь доступа «{name}» не найден в выбранном проекте"}


class LarixAPIClient:
    def __init__(self, base_url: str = BASE_URL, log_func: Optional[Callable[[str], None]] = None):
        self.base_url = base_url.rstrip("/")
        self.token: Optional[str] = None
        self.username: Optional[str] = None
        self.workspace_id: Optional[int] = None
        self.selected_workspace_id: Optional[int] = None
        self._log_func = log_func

    def _log(self, msg: str):
        if self._log_func:
            try:
                self._log_func(msg)
            except Exception:
                pass


    @staticmethod
    def _type_api_meta(kind: str) -> Dict[str, str]:
        value = (kind or "").strip().lower()
        if value in {"remark", "remarks", "замечание", "замечания"}:
            return {"kind": "remarks", "type_id": "remarkTypeId", "access_url": "/api/remark-types/access"}
        if value in {"task", "tasks", "задача", "задачи"}:
            return {"kind": "tasks", "type_id": "taskTypeId", "access_url": "/api/task-types/access"}
        raise ValueError(f"Неизвестный вид типов: {kind}")

    @staticmethod
    def _type_api_items(payload: Any) -> List[Dict[str, Any]]:
        if isinstance(payload, list):
            return [item for item in payload if isinstance(item, dict)]
        if not isinstance(payload, dict):
            return []
        for key in ("items", "content", "rows", "results"):
            if isinstance(payload.get(key), list):
                return [item for item in payload[key] if isinstance(item, dict)]
        for key in ("data", "result"):
            nested = payload.get(key)
            items = LarixAPIClient._type_api_items(nested)
            if items:
                return items
        return []

    def list_type_types(self, project_id: int, kind: str, page: int = 1, size: int = 25, max_pages: int = 100) -> Dict[str, Any]:
        """Получает все страницы типов замечаний/задач через POST list endpoint из HAR."""
        meta = self._type_api_meta(kind)
        if not self.token:
            return {"success": False, "error": "Не авторизован", "items": []}
        url = f"{self.base_url}/api/{meta['kind']}/types/list"
        size = max(1, min(int(size), 100))
        max_pages = max(1, min(int(max_pages), 100))
        try:
            all_items: List[Dict[str, Any]] = []
            first_data: Any = None
            for offset in range(max_pages):
                current_page = int(page) + offset
                payload = {"filters": [], "sorts": [], "page": current_page, "size": size}
                response = requests.post(url, params={"projectId": int(project_id)}, json=payload,
                                         headers={**self._headers(), "accept": "application/json", "Content-Type": "application/json"},
                                         timeout=REQUEST_TIMEOUT)
                if response.status_code != 200:
                    return {"success": False, "error": f"HTTP {response.status_code}: {response.text[:300]}", "items": all_items}
                data = response.json()
                if first_data is None:
                    first_data = data
                items = self._type_api_items(data)
                all_items.extend(items)
                if len(items) < size:
                    break
            return {"success": True, "data": first_data, "items": all_items}
        except requests.RequestException as exc:
            return {"success": False, "error": f"Network: {exc}", "items": []}
        except ValueError as exc:
            return {"success": False, "error": f"Некорректный JSON: {exc}", "items": []}

    def list_remark_types(self, project_id: int, **kwargs) -> Dict[str, Any]:
        return self.list_type_types(project_id, "remarks", **kwargs)

    def list_task_types(self, project_id: int, **kwargs) -> Dict[str, Any]:
        return self.list_type_types(project_id, "tasks", **kwargs)

    def create_type(self, project_id: int, title: str, kind: str) -> Dict[str, Any]:
        meta = self._type_api_meta(kind)
        if not self.token:
            return {"success": False, "error": "Не авторизован"}
        title = str(title or "").strip()
        if not title:
            return {"success": False, "error": "Название типа обязательно"}
        url = f"{self.base_url}/api/{meta['kind']}/types"
        try:
            response = requests.post(url, json={"title": title, "projectId": int(project_id)},
                                     headers={**self._headers(), "accept": "application/json", "Content-Type": "application/json"},
                                     timeout=REQUEST_TIMEOUT)
            if response.status_code not in (200, 201):
                return {"success": False, "error": f"HTTP {response.status_code}: {response.text[:300]}"}
            return {"success": True, "data": response.json()}
        except requests.RequestException as exc:
            return {"success": False, "error": f"Network: {exc}"}
        except ValueError as exc:
            return {"success": False, "error": f"Некорректный JSON: {exc}"}

    def create_remark_type(self, project_id: int, title: str) -> Dict[str, Any]:
        return self.create_type(project_id, title, "remarks")

    def create_task_type(self, project_id: int, title: str) -> Dict[str, Any]:
        return self.create_type(project_id, title, "tasks")

    def list_project_roles(self, project_id: int) -> Dict[str, Any]:
        if not self.token:
            return {"success": False, "error": "Не авторизован", "items": []}
        url = f"{self.base_url}/api/projectRoles/select/{int(project_id)}"
        try:
            response = requests.get(url, headers=self._headers(), timeout=REQUEST_TIMEOUT)
            if response.status_code != 200:
                return {"success": False, "error": f"HTTP {response.status_code}: {response.text[:300]}", "items": []}
            data = response.json()
            return {"success": True, "data": data, "items": self._type_api_items(data)}
        except requests.RequestException as exc:
            return {"success": False, "error": f"Network: {exc}", "items": []}
        except ValueError as exc:
            return {"success": False, "error": f"Некорректный JSON: {exc}", "items": []}

    def list_workspace_users(self, workspace_id: Optional[int] = None, size: int = 100, max_pages: int = 100) -> Dict[str, Any]:
        if not self.token:
            return {"success": False, "error": "Не авторизован", "items": []}
        workspace_id = workspace_id or self.selected_workspace_id or self.workspace_id
        if workspace_id is None:
            return {"success": False, "error": "Не выбрано пространство", "items": []}
        url = f"{self.base_url}/api/settings/users"
        size = max(1, min(int(size), 100))
        max_pages = max(1, min(int(max_pages), 100))
        try:
            all_items: List[Dict[str, Any]] = []
            first_data: Any = None
            for page in range(1, max_pages + 1):
                payload = {"filters": [], "sorts": [], "page": page, "size": size}
                response = requests.post(url, params={"workspaceId": int(workspace_id), "restricted": "true"}, json=payload,
                                         headers={**self._headers(), "accept": "application/json", "Content-Type": "application/json"},
                                         timeout=REQUEST_TIMEOUT)
                if response.status_code != 200:
                    return {"success": False, "error": f"HTTP {response.status_code}: {response.text[:300]}", "items": all_items}
                data = response.json()
                if first_data is None:
                    first_data = data
                items = self._type_api_items(data)
                all_items.extend(items)
                if len(items) < size:
                    break
            return {"success": True, "data": first_data, "items": all_items}
        except requests.RequestException as exc:
            return {"success": False, "error": f"Network: {exc}", "items": []}
        except ValueError as exc:
            return {"success": False, "error": f"Некорректный JSON: {exc}", "items": []}

    def set_type_access(self, type_id: int, kind: str, user_ids: List[int], project_role_ids: List[int], access_role_id: int) -> Dict[str, Any]:
        """Точно задаёт один независимый набор доступа, включая пустые списки."""
        meta = self._type_api_meta(kind)
        if not self.token:
            return {"success": False, "error": "Не авторизован"}
        payload = {meta["type_id"]: int(type_id), "userId": sorted({int(item) for item in user_ids}),
                   "projectRoleId": sorted({int(item) for item in project_role_ids}), "accessRoleId": int(access_role_id)}
        try:
            response = requests.put(f"{self.base_url}{meta['access_url']}", json=payload,
                                    headers={**self._headers(), "accept": "application/json", "Content-Type": "application/json"},
                                    timeout=REQUEST_TIMEOUT)
            if response.status_code not in (200, 201, 204):
                return {"success": False, "error": f"HTTP {response.status_code}: {response.text[:300]}", "payload": payload}
            data = None if response.status_code == 204 else response.json()
            return {"success": True, "data": data, "payload": payload}
        except requests.RequestException as exc:
            return {"success": False, "error": f"Network: {exc}", "payload": payload}
        except ValueError as exc:
            return {"success": False, "error": f"Некорректный JSON: {exc}", "payload": payload}



    def list_approval_workflows(self, project_id: int, workspace_id: Optional[int] = None,
                                size: int = 25, max_pages: int = 100) -> Dict[str, Any]:
        """Получает все маршруты согласований выбранного проекта."""
        if not self.token:
            return {"success": False, "error": "Не авторизован", "items": []}
        workspace_id = workspace_id or self.selected_workspace_id or self.workspace_id
        if workspace_id is None:
            return {"success": False, "error": "Не выбрано пространство", "items": []}
        size = max(1, min(int(size), 100))
        max_pages = max(1, min(int(max_pages), 100))
        url = f"{self.base_url}/api/approvals/workflow/list"
        all_items: List[Dict[str, Any]] = []
        try:
            for page in range(1, max_pages + 1):
                payload = {"filters": [], "sorts": [], "page": page, "size": size}
                response = requests.post(
                    url,
                    params={"workspaceId": int(workspace_id), "projectId": int(project_id)},
                    json=payload,
                    headers={**self._headers(), "accept": "application/json", "Content-Type": "application/json"},
                    timeout=REQUEST_TIMEOUT,
                )
                if response.status_code != 200:
                    return {"success": False, "error": f"HTTP {response.status_code}: {response.text[:500]}", "items": all_items}
                data = response.json()
                items = data.get("items", []) if isinstance(data, dict) else []
                items = [item for item in items if isinstance(item, dict)]
                all_items.extend(items)
                if len(items) < size:
                    break
            return {"success": True, "items": all_items}
        except requests.RequestException as exc:
            return {"success": False, "error": f"Network: {exc}", "items": all_items}
        except ValueError as exc:
            return {"success": False, "error": f"Некорректный JSON: {exc}", "items": all_items}

    def create_approval_workflow(self, project_id: int, title: str, steps: List[Dict[str, Any]],
                                 workspace_id: Optional[int] = None) -> Dict[str, Any]:
        """Создаёт маршрут согласования по схеме, зафиксированной в HAR."""
        if not self.token:
            return {"success": False, "error": "Не авторизован"}
        workspace_id = workspace_id or self.selected_workspace_id or self.workspace_id
        if workspace_id is None:
            return {"success": False, "error": "Не выбрано пространство"}
        clean_title = _approval_clean_text(title)
        if not clean_title:
            return {"success": False, "error": "Название маршрута обязательно"}

        api_steps: List[Dict[str, Any]] = []
        for index, source_step in enumerate(steps or []):
            users: List[Dict[str, Any]] = []
            for source_user in source_step.get("users", []):
                try:
                    user_id = int(source_user["user_id"])
                except (KeyError, TypeError, ValueError):
                    return {"success": False, "error": f"{index + 1} этап: некорректный ID согласующего"}
                users.append({
                    "user_id": user_id,
                    "can_cancel": bool(source_user.get("can_cancel", False)),
                    "must_approve": bool(source_user.get("must_approve", False)),
                })
            try:
                duration = int(source_step.get("duration"))
            except (TypeError, ValueError):
                return {"success": False, "error": f"{index + 1} этап: некорректная длительность"}
            api_steps.append({
                "duration": duration,
                "title": _approval_clean_text(source_step.get("title")) or f"{index + 1} Этап",
                "users": users,
                "workspace_id": int(workspace_id),
                "index": int(source_step.get("index", index)),
            })

        payload = {
            "workspace_id": int(workspace_id),
            "project_id": int(project_id),
            "title": clean_title,
            "created_by": "",
            "modified_by": "",
            "steps": api_steps,
        }
        try:
            response = requests.post(
                f"{self.base_url}/api/approvals/workflow/create",
                json=payload,
                headers={**self._headers(), "accept": "application/json", "Content-Type": "application/json"},
                timeout=REQUEST_TIMEOUT,
            )
            if response.status_code not in (200, 201):
                return {
                    "success": False,
                    "error": f"HTTP {response.status_code}: {response.text[:500]}",
                    "status_code": response.status_code,
                    "payload": payload,
                }
            return {"success": True, "data": response.json(), "payload": payload}
        except requests.RequestException as exc:
            return {"success": False, "error": f"Network: {exc}", "payload": payload}
        except ValueError as exc:
            return {"success": False, "error": f"Некорректный JSON: {exc}", "payload": payload}

    def set_approval_workflow_access(self, workflow_id: int, user_ids: List[int],
                                     project_role_ids: List[int], access_role_id: int) -> Dict[str, Any]:
        """Назначает просмотр (8) или создание (10) маршрута согласования."""
        if not self.token:
            return {"success": False, "error": "Не авторизован"}
        payload = {
            "approvalWorkflowId": int(workflow_id),
            "userId": sorted({int(item) for item in (user_ids or [])}),
            "projectRoleId": sorted({int(item) for item in (project_role_ids or [])}),
            "accessRoleId": int(access_role_id),
        }
        try:
            response = requests.put(
                f"{self.base_url}/api/approvals/workflow/access",
                json=payload,
                headers={**self._headers(), "accept": "application/json", "Content-Type": "application/json"},
                timeout=REQUEST_TIMEOUT,
            )
            if response.status_code not in (200, 201, 204):
                return {"success": False, "error": f"HTTP {response.status_code}: {response.text[:500]}", "payload": payload}
            data = None if response.status_code == 204 else response.json()
            return {"success": True, "data": data, "payload": payload}
        except requests.RequestException as exc:
            return {"success": False, "error": f"Network: {exc}", "payload": payload}
        except ValueError as exc:
            return {"success": False, "error": f"Некорректный JSON: {exc}", "payload": payload}



# =====================================================================
    def _headers(self) -> Dict[str, str]:
        headers = {"accept": "*/*"}
        if self.token:
            headers["Authorization"] = f"Bearer {self.token}"
        return headers

    def _sync_identity_from_token(self) -> None:
        if not self.token:
            return
        decoded = _decode_jwt_payload(self.token)
        if not decoded:
            return
        ws = decoded.get("workspace_id") or decoded.get("workspaceId")
        if ws is not None:
            try:
                self.workspace_id = int(ws)
                if self.selected_workspace_id is None:
                    self.selected_workspace_id = self.workspace_id
            except (ValueError, TypeError):
                pass

    def login(self, username: str, password: str) -> bool:
        url = f"{self.base_url}/api/admin/login"
        payload = {"username": username, "password": password, "app_code": ""}
        try:
            r = requests.post(url, json=payload, headers={
                "accept": "*/*", "Content-Type": "application/json"
            }, timeout=12)
            if r.status_code == 401:
                return False
            r.raise_for_status()
            data = r.json()
            token = data.get("token") or data.get("accessToken") or data.get("access_token")
            if not token:
                return False
            self.token = token
            self.username = username
            self._sync_identity_from_token()
            return True
        except requests.RequestException:
            return False

    def list_workspaces(self) -> List[Dict]:
        if not self.token:
            return []
        endpoints = [
            f"{self.base_url}/api/workspace/list",
            f"{self.base_url}/api/admin/workspace/list"
        ]
        for url in endpoints:
            try:
                r = requests.get(url, headers=self._headers(), timeout=12)
                if r.status_code == 401:
                    continue
                if r.status_code == 200:
                    data = r.json()
                    workspaces = _extract_data_list(data)
                    if workspaces:
                        return [
                            {"id": _extract_workspace_id(ws), "name": _extract_workspace_name(ws)}
                            for ws in workspaces if _extract_workspace_id(ws)
                        ]
            except requests.RequestException:
                continue
        return []

    def change_workspace(self, workspace_id: int) -> bool:
        if not self.token:
            return False
        if str(self.selected_workspace_id) == str(workspace_id):
            return True
        url = f"{self.base_url}/api/admin/workspace/change"
        try:
            r = requests.put(url, headers=self._headers(), params={"workspaceId": workspace_id}, timeout=12)
            if r.status_code == 401:
                return False
            r.raise_for_status()
            data = r.json()
            token_block = data.get("data") if isinstance(data.get("data"), dict) else data
            if isinstance(token_block, dict):
                new_token = token_block.get("access") or token_block.get("token")
                if new_token:
                    self.token = new_token
            self.selected_workspace_id = workspace_id
            self.workspace_id = workspace_id
            self._sync_identity_from_token()
            return True
        except requests.RequestException:
            return False

    def list_projects(self) -> List[Dict[str, Any]]:
        if not self.token:
            return []
        url = f"{self.base_url}/api/workspace/projects/show"
        try:
            r = requests.get(url, headers=self._headers(), timeout=REQUEST_TIMEOUT)
            if r.status_code != 200:
                return []
            projects = _extract_data_list(r.json())
            result = []
            for project in projects:
                project_id = _extract_project_id(project)
                if project_id:
                    result.append({"id": project_id, "name": _extract_project_name(project)})
            return result
        except (requests.RequestException, ValueError):
            return []

    def list_roles(self) -> List[Dict[str, Any]]:
        if not self.token:
            return []
        url = f"{self.base_url}/api/projectRoles/list"
        payload = {"filters": [], "sorts": [], "page": 1, "size": 100}
        headers = self._headers()
        headers["Content-Type"] = "application/json"
        try:
            r = requests.post(url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT)
            if r.status_code != 200:
                self._log(f"⚠️ Роли: HTTP {r.status_code}")
                return []
            data = r.json()
            roles = []
            if isinstance(data, dict):
                inner_data = data.get("data")
                if isinstance(inner_data, dict):
                    roles = inner_data.get("items", [])
                    if not roles:
                        for key in ["list", "roles", "data"]:
                            if isinstance(inner_data.get(key), list):
                                roles = inner_data[key]
                                break
                elif isinstance(inner_data, list):
                    roles = inner_data
            if isinstance(data, list):
                roles = data
            if not roles:
                return []
            result = []
            for role in roles:
                if not isinstance(role, dict):
                    continue
                role_id = None
                for key in ["roleId", "id", "role_id", "ID"]:
                    if role.get(key) is not None:
                        try:
                            role_id = int(role[key])
                            break
                        except (TypeError, ValueError):
                            continue
                title = ""
                for key in ["title", "name", "roleName", "role_name", "Title"]:
                    if role.get(key) and isinstance(role[key], str):
                        title = role[key].strip()
                        break
                if role_id and title:
                    result.append({"id": role_id, "title": title})
            return result
        except Exception:
            return []


    def get_folder_access_types(self) -> List[Dict[str, Any]]:
        if not self.token:
            return []
        try:
            r = requests.get(f"{self.base_url}/api/projectRoles/select", headers=self._headers(), timeout=REQUEST_TIMEOUT)
            if r.status_code != 200:
                return []
            data = r.json()
            items = data.get("data", {}).get("items", []) if isinstance(data, dict) else []
            return [item for item in items if isinstance(item, dict)]
        except (requests.RequestException, ValueError):
            return []

    def get_project_folder_permissions(self, project_id: int) -> Dict[str, Any]:
        if not self.token:
            return {"success": False, "error": "Не авторизован"}
        try:
            r = requests.get(
                f"{self.base_url}/api/projectRoles/folders/{int(project_id)}",
                headers=self._headers(), timeout=REQUEST_TIMEOUT,
            )
            if r.status_code != 200:
                return {"success": False, "error": f"HTTP {r.status_code}: {r.text[:300]}"}
            return {"success": True, "data": r.json()}
        except requests.RequestException as exc:
            return {"success": False, "error": f"Network: {exc}"}
        except ValueError as exc:
            return {"success": False, "error": f"Некорректный JSON дерева папок: {exc}"}

    def get_project_folder_principals(self, project_id: int) -> Dict[str, Any]:
        if not self.token:
            return {"success": False, "error": "Не авторизован"}
        try:
            r = requests.get(
                f"{self.base_url}/api/projectRoles/select/{int(project_id)}",
                headers=self._headers(), timeout=REQUEST_TIMEOUT,
            )
            if r.status_code != 200:
                return {"success": False, "error": f"HTTP {r.status_code}: {r.text[:300]}"}
            return {"success": True, "data": r.json()}
        except requests.RequestException as exc:
            return {"success": False, "error": f"Network: {exc}"}
        except ValueError as exc:
            return {"success": False, "error": f"Некорректный JSON ролей/пользователей: {exc}"}

    def assign_folder_access(self, project_id: int, folder_id: int, access_role_id: int,
                             user_ids: List[int], project_role_ids: List[int]) -> Dict[str, Any]:
        """Полностью задаёт список получателей одного типа права на папке.

        API Larix работает не как add/remove одного principal: запрос содержит
        целиком userId/projectRoleId для accessRoleId. Поэтому вызывающий код
        обязан сохранить всех текущих получателей, которых не меняет Excel.
        """
        if not self.token:
            return {"success": False, "error": "Не авторизован"}
        payload = {
            "projectId": str(int(project_id)),
            "folderId": int(folder_id),
            "userId": sorted({int(item) for item in user_ids}),
            "projectRoleId": sorted({int(item) for item in project_role_ids}),
            "accessRoleId": int(access_role_id),
        }
        headers = self._headers()
        headers["Content-Type"] = "application/json"
        try:
            r = requests.post(
                f"{self.base_url}/api/projectRoles/assign",
                headers=headers, json=payload, timeout=REQUEST_TIMEOUT,
            )
            if r.status_code not in (200, 201):
                return {"success": False, "error": f"HTTP {r.status_code}: {r.text[:300]}", "payload": payload}
            try:
                data = r.json()
            except ValueError:
                data = {"raw": r.text}
            return {"success": True, "data": data, "payload": payload}
        except requests.RequestException as exc:
            return {"success": False, "error": f"Network: {exc}", "payload": payload}

    def set_project_admin(self, user_id: int, project_id: int, is_admin: bool) -> Dict[str, Any]:
        if not self.token:
            return {"success": False, "error": "Не авторизован"}
        url = f"{self.base_url}/api/projectRoles/set/admin"
        payload = {"set": bool(is_admin), "userId": int(user_id), "projectId": int(project_id)}
        headers = self._headers()
        headers["Content-Type"] = "application/json"
        try:
            r = requests.put(url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT)
            if r.status_code not in (200, 201):
                return {"success": False, "error": f"HTTP {r.status_code}: {r.text[:200]}"}
            return {"success": True, "message": "Админ установлен" if is_admin else "Админ снят"}
        except requests.RequestException as e:
            return {"success": False, "error": f"Network: {e}"}

    def create_project(self, workspace_id: int, title: str, description: str = "") -> Dict[str, Any]:
        if not self.token:
            return {"success": False, "error": "Не авторизован"}
        if not title or not title.strip():
            return {"success": False, "error": "Пустое название проекта"}
        url = f"{self.base_url}/api/project/add"
        files = {"file": ("", b"", "application/octet-stream")}
        form_data = {"title": title.strip(), "description": (description or "").strip(), "id": "0"}
        headers = self._headers()
        try:
            r = requests.post(url, headers=headers, data=form_data, files=files, timeout=REQUEST_TIMEOUT)
            if r.status_code not in (200, 201):
                return {"success": False, "error": f"HTTP {r.status_code}: {r.text[:300]}"}
            data = r.json()
            project = data.get("data") if isinstance(data, dict) and isinstance(data.get("data"), dict) else data
            project_id = _extract_project_id(project) if isinstance(project, dict) else None
            if not project_id:
                return {"success": False, "error": "Проект создан, но ID не найден в ответе", "raw": data}
            return {"success": True, "project_id": project_id, "title": project.get("title", title), "data": project}
        except requests.RequestException as e:
            return {"success": False, "error": f"Network: {e}"}

    def get_root_folder_id(self, project_id: int, retries: int = 5, delay: float = 0.6) -> Dict[str, Any]:
        """Получает ID корневой папки проекта.

        Larix требует получить корень через /api/folder/root-folder/{projectId}
        ДО создания пользовательских папок и передавать его ID как родителя
        для папок первого уровня.
        """
        if not self.token:
            return {"success": False, "error": "Не авторизован"}
        if not project_id:
            return {"success": False, "error": "Не указан project_id"}

        url = f"{self.base_url}/api/folder/root-folder/{int(project_id)}"
        last_error = "Корневая папка не найдена"

        for attempt in range(1, max(1, int(retries)) + 1):
            try:
                r = requests.get(url, headers=self._headers(), timeout=REQUEST_TIMEOUT)
                if r.status_code == 200:
                    try:
                        data = r.json()
                    except ValueError:
                        data = r.text.strip()

                    root_id: Optional[int] = None
                    if isinstance(data, dict):
                        # Актуальный Larix API возвращает root ID как скаляр:
                        # {"success": true, "data": 15562}. Поддерживаем также
                        # старый/альтернативный вариант, где data является объектом.
                        raw_root = data.get("data")
                        if isinstance(raw_root, dict):
                            root_id = _extract_folder_id(raw_root)
                        else:
                            root_id = _to_int(raw_root)
                        if root_id is None:
                            root_id = _extract_folder_id(data)
                    elif isinstance(data, list) and data:
                        first = data[0]
                        if isinstance(first, dict):
                            root_id = _extract_folder_id(first)
                        else:
                            root_id = _to_int(first)
                    else:
                        root_id = _to_int(data)

                    if root_id:
                        return {"success": True, "root_folder_id": root_id, "data": data}
                    last_error = f"HTTP 200, но ID корневой папки не найден в ответе: {str(data)[:250]}"
                else:
                    last_error = f"HTTP {r.status_code}: {r.text[:250]}"
            except requests.RequestException as exc:
                last_error = f"Network: {exc}"

            if attempt < max(1, int(retries)) and delay:
                time.sleep(delay)

        return {"success": False, "error": last_error}

    def create_folder(self, project_id: int, name: str, parent_folder_id: Optional[int] = None) -> Dict[str, Any]:
        if not self.token:
            return {"success": False, "error": "Не авторизован"}
        if not name or not name.strip():
            return {"success": False, "error": "Пустое имя папки"}
        if not project_id:
            return {"success": False, "error": "Не указан project_id"}
        url = f"{self.base_url}/api/folder/add"
        # В сетевых запросах Larix параметр родителя называется parentFolderId.
        # Для первого уровня сюда теперь передаётся ID, полученный через root-folder/{projectId}.
        payload = {"projectId": int(project_id), "name": name.strip(), "id": 0, "parentFolderId": parent_folder_id}
        try:
            r = requests.post(url, headers=self._headers(), json=payload, timeout=REQUEST_TIMEOUT)
            if r.status_code not in (200, 201):
                return {"success": False, "error": f"HTTP {r.status_code}: {r.text[:200]}"}
            data = r.json()
            folder_id = _extract_folder_id(data) if isinstance(data, dict) else None
            if not folder_id:
                return {"success": True, "warning": "ID папки не получен", "raw": data, "folder_id": None}
            return {"success": True, "folder_id": folder_id}
        except requests.RequestException as e:
            return {"success": False, "error": f"Network: {e}"}

    def create_folders_from_excel(self, excel_path: str, sheet_name: str, project_id: int,
                                  log: Optional[Callable[[str], None]] = None, delay: float = 0.3) -> Dict[str, Any]:
        def _log(msg: str):
            if log:
                try:
                    log(msg)
                except Exception:
                    pass
        created_ids: Dict[str, int] = {}
        created_count = 0
        errors: List[str] = []
        try:
            if not os.path.exists(excel_path):
                return {"success": False, "error": f"Файл не найден: {excel_path}"}
            excel_file = pd.ExcelFile(excel_path)
            sheets = [str(s) for s in excel_file.sheet_names]
            _log(f"Доступные листы: {sheets}")
            target = None
            for s in sheets:
                if s.strip() == sheet_name.strip():
                    target = s
                    break
            if not target:
                return {"success": False, "error": f"Лист '{sheet_name}' не найден. Доступные: {sheets}"}
            df = pd.read_excel(excel_file, sheet_name=target).fillna("")
            df.columns = [f"Уровень_{i}" for i in range(len(df.columns))]
            _log(f"Загружено строк: {len(df)}")

            # По актуальному API Larix сначала обязательно получаем корневую
            # папку проекта. Все папки первого уровня создаются именно в ней.
            _log(f"📂 Получение корневой папки проекта ID={project_id}...")
            root_result = self.get_root_folder_id(project_id)
            if not root_result.get("success"):
                return {
                    "success": False,
                    "error": "Не удалось получить корневую папку проекта: "
                             + str(root_result.get("error", "неизвестная ошибка")),
                }
            root_folder_id = _to_int(root_result.get("root_folder_id"))
            if not root_folder_id:
                return {"success": False, "error": "API не вернул ID корневой папки проекта"}
            _log(f"✅ Корневая папка проекта: ID={root_folder_id}")

            for idx, row in df.iterrows():
                folders = [
                    str(row[col]).strip()
                    for col in df.columns
                    if col.startswith("Уровень") and str(row[col]).strip().lower() not in ("", "nan")
                ]
                if not folders:
                    continue
                # Первый уровень дерева создаём внутри корневой папки проекта.
                parent_id = root_folder_id
                current_path = ""
                for folder_name in folders:
                    current_path = f"{current_path}/{folder_name}" if current_path else folder_name
                    if current_path in created_ids:
                        parent_id = created_ids[current_path]
                        continue
                    result = self.create_folder(project_id, folder_name, parent_id)
                    if not result.get("success"):
                        msg = f"Ошибка папки '{current_path}': {result.get('error')}"
                        errors.append(msg)
                        _log("❌ " + msg)
                        break
                    folder_id = result.get("folder_id")
                    if folder_id:
                        created_ids[current_path] = folder_id
                        parent_id = folder_id
                        created_count += 1
                        _log(f"✅ Создано: {current_path} (ID: {folder_id})")
                    else:
                        _log(f"⚠️ Папка '{current_path}' создана, но ID не получен")
                    if delay:
                        time.sleep(delay)
        except Exception as e:
            return {"success": False, "error": f"Исключение: {e}\n{traceback.format_exc()}"}
        return {"success": True, "created": created_count, "errors": errors, "path_to_id": created_ids}

    def create_user_in_system(self, email: str, last_name: str, first_name: str) -> Dict[str, Any]:
        if not self.token or not email:
            return {"success": False, "error": "Missing parameters"}
        endpoint = f"{self.base_url}/api/user/system/add"
        payload = {
            "email": email.strip().lower(),
            "lastName": last_name.strip() if last_name else "",
            "firstName": first_name.strip() if first_name else "",
        }
        try:
            r = requests.post(endpoint, headers=self._headers(), json=payload, timeout=REQUEST_TIMEOUT)
            if r.status_code in [200, 201]:
                data = r.json()
                inner = data.get("data", {}) if isinstance(data, dict) else {}
                acs_user_id = _extract_ac_user_id(inner, allow_plain_id=False)
                return {
                    "success": True,
                    "email": email,
                    "acs_user_id": acs_user_id,
                    "message": data.get("message", "Пользователь создан") if isinstance(data, dict) else "Пользователь создан",
                    "data": inner
                }
            if r.status_code == 409:
                return {"success": False, "error": "Пользователь уже существует", "email": email, "code": 409}
            return {"success": False, "error": f"HTTP {r.status_code}: {r.text[:200]}", "email": email}
        except requests.RequestException as e:
            return {"success": False, "error": f"Network: {e}", "email": email}

    def add_user_to_workspace(self, workspace_id: int, email: str) -> Dict[str, Any]:
        if not self.token or not workspace_id or not email:
            return {"success": False, "error": "Missing parameters"}
        url = f"{self.base_url}/api/settings/users/add"
        clean_email = email.strip().lower()
        payload = {"workspaceId": int(workspace_id), "email": clean_email}
        try:
            r = requests.post(url, headers=self._headers(), json=payload, timeout=REQUEST_TIMEOUT)
            if r.status_code == 409:
                return {
                    "success": True,
                    "email": clean_email,
                    "already_in_workspace": True,
                    "message": "Уже добавлен в пространство",
                    "code": 409,
                }
            if r.status_code not in [200, 201]:
                raw_text = r.text[:500]
                try:
                    raw_payload = r.json()
                except ValueError:
                    raw_payload = raw_text
                # Некоторые версии API возвращают 400/422 вместо 409, если пользователь уже есть в пространстве.
                # Это не должно блокировать дальнейшее назначение пользователя в проект.
                if _looks_like_already_in_workspace_error(raw_payload) or _looks_like_already_in_workspace_error(raw_text):
                    return {
                        "success": True,
                        "email": clean_email,
                        "already_in_workspace": True,
                        "message": "Уже добавлен в пространство",
                        "code": r.status_code,
                        "raw": raw_payload,
                    }
                return {"success": False, "error": f"HTTP {r.status_code}: {r.text[:200]}", "email": clean_email}
            data = r.json()
            inner_data = data.get("data", {}) if isinstance(data, dict) else {}
            if isinstance(inner_data, dict) and inner_data.get("to_be_added") is True:
                return {
                    "success": True,
                    "need_create": True,
                    "email": clean_email,
                    "message": data.get("message", "Пользователь не существует в ACS") if isinstance(data, dict) else "Пользователь не существует в ACS",
                    "data": inner_data
                }
            ac_user_id = _extract_ac_user_id(inner_data, allow_plain_id=False)

            # В разных версиях API ID пользователя находится либо прямо в
            # data.acsUserId, либо внутри data.workspace_user.user.id.
            # Поле workspace_user.id может быть ID записи пространства, поэтому
            # его как ACS ID не используем без проверки вложенного user.
            workspace_user = inner_data.get("workspace_user") if isinstance(inner_data, dict) else None
            if ac_user_id is None and isinstance(workspace_user, dict):
                ac_user_id = _extract_ac_user_id(workspace_user, allow_plain_id=False)
                nested_user = workspace_user.get("user")
                if ac_user_id is None and isinstance(nested_user, dict):
                    nested_email = str(nested_user.get("email", "")).strip().lower()
                    if not nested_email or nested_email == clean_email:
                        ac_user_id = _extract_by_keys(nested_user, ["id", "Id"])

            return {
                "success": True,
                "email": clean_email,
                "ac_user_id": ac_user_id,
                "message": data.get("message", "Добавлен в пространство") if isinstance(data, dict) else "Добавлен в пространство",
                "data": inner_data
            }
        except requests.RequestException as e:
            return {"success": False, "error": f"Network: {e}", "email": email}

    def find_workspace_user_id(self, workspace_id: int, email: str,
                                project_id: Optional[int] = None,
                                retries: int = 5, delay: float = 1.0) -> Optional[int]:
        """Возвращает ACS ID пользователя из выбранного пространства.

        Актуальный API Larix получает список пользователей через POST
        /api/settings/users с параметрами в query string и пагинацией в JSON-теле.
        Ранее здесь использовался GET, из-за чего уже существующий пользователь
        мог добавиться в пространство, но программа не находила его ID и не
        продолжала назначение в проект.
        """
        if not self.token or not workspace_id or not email:
            return None

        target_email = email.strip().lower()
        params: Dict[str, Any] = {
            "workspaceId": int(workspace_id),
            "fullList": "true",
            "restricted": "false",
        }
        if project_id:
            params["projectId"] = int(project_id)

        url = f"{self.base_url}/api/settings/users"
        payload = {"page": 1, "size": 1000}
        headers = self._headers()
        headers["Content-Type"] = "application/json"

        for attempt in range(1, retries + 1):
            try:
                # Формат подтверждён сетевым HAR интерфейса Larix.
                r = requests.post(
                    url,
                    headers=headers,
                    params=params,
                    json=payload,
                    timeout=REQUEST_TIMEOUT,
                )

                # Совместимость со старыми версиями API, если там ещё был GET.
                if r.status_code in (404, 405):
                    r = requests.get(
                        url,
                        headers=self._headers(),
                        params=params,
                        timeout=REQUEST_TIMEOUT,
                    )

                if r.status_code != 200:
                    self._log(
                        f"   Поиск пользователя: HTTP {r.status_code} "
                        f"(попытка {attempt}/{retries})"
                    )
                    if attempt < retries:
                        time.sleep(delay)
                    continue

                users = _extract_data_list(r.json())
                for item in users:
                    if not isinstance(item, dict):
                        continue

                    user_block = item.get("user") if isinstance(item.get("user"), dict) else item
                    if not isinstance(user_block, dict):
                        continue

                    user_email = str(
                        user_block.get("email", item.get("email", ""))
                    ).strip().lower()
                    if user_email != target_email:
                        continue

                    # Сначала ищем явный acsUserId/ac_user_id, затем обычный
                    # user.id — именно так список пользователей выглядит в Larix.
                    value = _extract_ac_user_id(item, allow_plain_id=False)
                    if value is None:
                        value = _extract_ac_user_id(user_block, allow_plain_id=False)
                    if value is None:
                        value = _extract_by_keys(user_block, ["id", "Id"])
                    if value is None and isinstance(item.get("workspace_user"), dict):
                        value = _extract_ac_user_id(item["workspace_user"], allow_plain_id=False)

                    if value is not None:
                        self._log(f"   Найден ACS ID пользователя: {value}")
                        return value

                if attempt < retries:
                    self._log(
                        f"   (попытка {attempt}/{retries}) пользователь пока не найден, "
                        f"жду {delay}с..."
                    )
                    time.sleep(delay)

            except (requests.RequestException, ValueError) as exc:
                self._log(
                    f"   Ошибка поиска пользователя: {exc} "
                    f"(попытка {attempt}/{retries})"
                )
                if attempt < retries:
                    time.sleep(delay)

        return None

    def assign_user_to_projects(self, ac_user_id: int, project_ids: List[int]) -> Dict[str, Any]:
        if not self.token or not ac_user_id:
            return {"success": False, "error": "Нет ac_user_id"}
        project_ids = sorted({int(pid) for pid in project_ids if pid})
        if not project_ids:
            return {"success": True, "message": "Проекты не выбраны"}
        url = f"{self.base_url}/api/user/project/assign"
        payload = [{"ac_user_id": int(ac_user_id), "project_id": int(project_id)} for project_id in project_ids]
        headers = self._headers()
        headers["Content-Type"] = "application/json"
        try:
            r = requests.post(url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT)
            if r.status_code == 409:
                return {"success": True, "message": "Уже добавлен в проект(ы)", "code": 409}
            if r.status_code not in [200, 201]:
                raw_text = r.text[:500]
                try:
                    raw_payload = r.json()
                except ValueError:
                    raw_payload = raw_text
                if _looks_like_already_in_workspace_error(raw_payload) or _looks_like_already_in_workspace_error(raw_text):
                    return {"success": True, "message": "Уже добавлен в проект(ы)", "code": r.status_code, "raw": raw_payload}
                return {"success": False, "error": f"Назначение в проект: HTTP {r.status_code} - {r.text[:300]}", "code": r.status_code, "raw": r.text[:500]}
            data = r.json()
            if isinstance(data, dict) and data.get("success") is False:
                return {"success": False, "error": data.get("message", "Не удалось назначить в проект"), "raw": data}
            return {"success": True, "message": f"Добавлен в проект(ы): {len(project_ids)}", "data": data}
        except requests.RequestException as e:
            return {"success": False, "error": f"Назначение в проект: {e}"}

    def list_user_projects(self, ac_user_id: int, retries: int = 3, delay: float = 0.7) -> List[Dict[str, Any]]:
        if not self.token or not ac_user_id:
            return []
        url = f"{self.base_url}/api/workspace/user/projects"
        for attempt in range(1, retries + 1):
            try:
                r = requests.get(url, headers=self._headers(), params={"userId": int(ac_user_id)}, timeout=REQUEST_TIMEOUT)
                if r.status_code == 200:
                    projects = _extract_data_list(r.json())
                    if projects or attempt == retries:
                        return projects
                if attempt < retries:
                    time.sleep(delay)
            except (requests.RequestException, ValueError):
                if attempt < retries:
                    time.sleep(delay)
        return []

    def update_user_project_roles(self, ac_user_id: int, target_project_ids: List[int],
                                  role_items: List[Dict[str, Any]],
                                  is_project_admin: Optional[bool] = None) -> Dict[str, Any]:
        if not self.token or not ac_user_id or not target_project_ids:
            return {"success": True, "message": "Роли не выбраны"}
        role_items = role_items or []
        target_set = {int(project_id) for project_id in target_project_ids if project_id}
        if not target_set:
            return {"success": True, "message": "Проекты не выбраны"}
        user_projects = self.list_user_projects(ac_user_id)
        existing_project_ids = {
            int(p.get("projectId") or p.get("project_id") or 0)
            for p in user_projects if isinstance(p, dict) and (p.get("projectId") or p.get("project_id"))
        }
        for pid in sorted(target_set - existing_project_ids):
            user_projects.append({
                "projectId": int(pid),
                "workspaceId": self.selected_workspace_id,
                "projectTitle": "",
                "isProjectAdmin": False,
                "roles": [],
                "roleIds": [],
            })
        payload_projects = []
        updated_count = 0
        for project in user_projects:
            if not isinstance(project, dict):
                continue
            project_id = int(project.get("projectId") or project.get("project_id") or 0)
            if not project_id:
                continue
            roles = list(project.get("roles") or [])
            if project_id in target_set:
                by_id = {
                    int(role.get("id")): {"id": int(role.get("id")), "title": str(role.get("title", ""))}
                    for role in roles if isinstance(role, dict) and role.get("id") is not None
                }
                for role in role_items:
                    if not isinstance(role, dict) or role.get("id") is None:
                        continue
                    by_id[int(role["id"])] = {"id": int(role["id"]), "title": str(role["title"])}
                roles = sorted(by_id.values(), key=lambda r: int(r["id"]))
                updated_count += 1
            existing_role_ids = [int(v) for v in (project.get("roleIds") or []) if _to_int(v) is not None]
            role_ids = sorted({*existing_role_ids, *[int(role["id"]) for role in roles if role.get("id") is not None]})
            admin_state = bool(project.get("isProjectAdmin", False))
            if project_id in target_set and is_project_admin is not None:
                admin_state = bool(is_project_admin)
            payload_projects.append({
                "workspaceId": project.get("workspaceId") or project.get("workspace_id") or self.selected_workspace_id,
                "projectId": project_id,
                "projectTitle": project.get("projectTitle") or project.get("title") or project.get("name") or "",
                "isProjectAdmin": admin_state,
                "roleIds": role_ids,
                "roles": roles,
            })
        if updated_count == 0:
            return {"success": False, "error": "Пользователь не назначен в проекты для установки роли"}
        url = f"{self.base_url}/api/projectRoles/user/update"
        payload = {"userId": int(ac_user_id), "userProjectWithRoles": payload_projects}
        headers = self._headers()
        headers["Content-Type"] = "application/json"
        try:
            r = requests.put(url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT)
            if r.status_code not in [200, 201]:
                return {"success": False, "error": f"Назначение роли: HTTP {r.status_code} - {r.text[:300]}"}
            data = r.json()
            if isinstance(data, dict) and data.get("success") is False:
                return {"success": False, "error": data.get("message", "Не удалось назначить роль")}
            return {"success": True, "message": f"Роли назначены: {len(role_items)}"}
        except requests.RequestException as e:
            return {"success": False, "error": f"Назначение роли: {e}"}

    def add_user_with_creation(self, workspace_id: int, email: str,
                               last_name: str = "", first_name: str = "",
                               project_ids: Optional[List[int]] = None,
                               role_items: Optional[List[Dict[str, Any]]] = None,
                               is_project_admin: bool = False) -> Dict[str, Any]:
        project_ids = sorted({int(pid) for pid in (project_ids or []) if pid})
        role_items = role_items or []
        clean_email = email.strip().lower()
        messages = []
        ac_user_id: Optional[int] = None
        result = self.add_user_to_workspace(workspace_id, clean_email)
        ac_user_id = result.get("ac_user_id")
        if result.get("success") and result.get("need_create"):
            self._log("   Пользователь не существует в ACS, создаю...")
            create_result = self.create_user_in_system(clean_email, last_name, first_name)
            if not create_result.get("success") and create_result.get("code") != 409:
                return {"success": False, "email": clean_email, "error": f"Не удалось создать: {create_result.get('error')}"}
            if create_result.get("success"):
                ac_user_id = create_result.get("acs_user_id") or ac_user_id
                messages.append("Создан в ACS")
                if ac_user_id:
                    self._log(f"   Получен acsUserId = {ac_user_id}")
            time.sleep(0.7)
            result = self.add_user_to_workspace(workspace_id, clean_email)
            if not result.get("success"):
                return {"success": False, "email": clean_email, "error": f"Не удалось добавить в пространство после создания: {result.get('error')}"}
            ac_user_id = ac_user_id or result.get("ac_user_id")
            messages.append("Добавлен в пространство")
        elif result.get("success"):
            messages.append(result.get("message", "В пространстве"))
        elif _looks_like_missing_user_error(result):
            self._log("   Пользователь не найден, создаю через /api/user/system/add...")
            create_result = self.create_user_in_system(clean_email, last_name, first_name)
            if not create_result.get("success") and create_result.get("code") != 409:
                return {"success": False, "email": clean_email, "error": f"Не удалось создать: {create_result.get('error')}"}
            ac_user_id = create_result.get("acs_user_id") or ac_user_id
            time.sleep(0.7)
            result = self.add_user_to_workspace(workspace_id, clean_email)
            if not result.get("success"):
                return {"success": False, "email": clean_email, "error": f"Не удалось добавить в пространство: {result.get('error')}"}
            ac_user_id = ac_user_id or result.get("ac_user_id")
            messages.append("Создан и добавлен в пространство")
        else:
            return {"success": False, "email": clean_email, "error": result.get("error", "Ошибка добавления в пространство")}
        if not ac_user_id:
            self._log("   Поиск ac_user_id в пространстве...")
            # Важно: сначала ищем пользователя по всему пространству, без projectId.
            # Если пользователь уже есть в пространстве, но ещё не назначен в целевой проект,
            # фильтр projectId может вернуть пустой список и сорвать назначение в проект.
            ac_user_id = self.find_workspace_user_id(
                workspace_id, clean_email,
                project_id=None,
                retries=5, delay=1.0
            )
            if not ac_user_id and project_ids:
                self._log("   Поиск ac_user_id с фильтром по проекту не дал результата, пробую дополнительную проверку...")
                ac_user_id = self.find_workspace_user_id(
                    workspace_id, clean_email,
                    project_id=project_ids[0],
                    retries=2, delay=0.8
                )
        if not ac_user_id:
            return {"success": False, "email": clean_email, "error": "Не удалось найти ac_user_id/acsUserId после добавления в пространство"}
        if project_ids:
            assign_result = self.assign_user_to_projects(ac_user_id, project_ids)
            if not assign_result.get("success") and _looks_like_missing_user_error(assign_result):
                self._log("   Проект не принял пользователя: повторно добавляю в пространство и пробую ещё раз...")
                self.add_user_to_workspace(workspace_id, clean_email)
                time.sleep(0.8)
                ac_user_id = self.find_workspace_user_id(
                    workspace_id, clean_email,
                    project_id=None,
                    retries=5, delay=1.0
                ) or ac_user_id
                assign_result = self.assign_user_to_projects(ac_user_id, project_ids)
            if not assign_result.get("success"):
                return {"success": False, "email": clean_email, "error": assign_result.get("error", "Не удалось назначить в проект")}
            messages.append(f"Проектов: {len(project_ids)}")
        if project_ids and (role_items or is_project_admin):
            time.sleep(0.5)
            role_result = self.update_user_project_roles(
                ac_user_id, project_ids, role_items,
                is_project_admin=True if is_project_admin else None
            )
            if role_result.get("success"):
                messages.append(role_result.get("message", "роли назначены"))
            else:
                messages.append(f"⚠️ роли/admin update: {role_result.get('error')}")
        if is_project_admin and project_ids:
            admin_ok = 0
            admin_fail = 0
            for pid in project_ids:
                admin_result = self.set_project_admin(ac_user_id, pid, True)
                if admin_result.get("success"):
                    admin_ok += 1
                else:
                    admin_fail += 1
                    self._log(f"   ⚠️ set/admin для проекта {pid}: {admin_result.get('error')}")
            if admin_ok:
                messages.append(f"Админ в {admin_ok} проектах")
            if admin_fail:
                messages.append(f"⚠️ не удалось сделать админом в {admin_fail} проектах")
        return {"success": True, "email": clean_email, "ac_user_id": ac_user_id, "message": "; ".join(messages) if messages else "OK"}


# =====================================================================
#  Workers: ролевая матрица
# =====================================================================
class RoleMatrixPreviewWorker(QThread):
    log = Signal(str)
    finished_ok = Signal(dict)
    failed = Signal(str)

    def __init__(self, client: LarixAPIClient, project_id: int, excel_path: str, sheet_name: str):
        super().__init__()
        self.client = client
        self.project_id = int(project_id)
        self.excel_path = excel_path
        self.sheet_name = sheet_name

    def run(self):
        try:
            self.log.emit("Ролевая матрица: читаю Excel...")
            parsed = _parse_role_matrix_excel(self.excel_path, self.sheet_name)
            self.log.emit(
                f"Лист '{parsed['sheet_name']}': строк с командами {len(parsed['entries'])}; "
                f"колонок ролей/пользователей {len(parsed['principal_names'])}"
            )

            principal_result = self.client.get_project_folder_principals(self.project_id)
            if not principal_result.get("success"):
                raise RuntimeError("Не удалось загрузить роли/пользователей: " + str(principal_result.get("error")))
            catalog = _matrix_build_principal_catalog(principal_result.get("data"))

            tree_result = self.client.get_project_folder_permissions(self.project_id)
            if not tree_result.get("success"):
                raise RuntimeError("Не удалось загрузить дерево прав: " + str(tree_result.get("error")))
            folder_index = _build_matrix_folder_index(tree_result.get("data"))
            self.log.emit(f"Папок в дереве Larix: {len(folder_index)}")

            resolution_cache: Dict[str, Dict[str, Any]] = {}
            plan: List[Dict[str, Any]] = []
            for entry in parsed["entries"]:
                folder = folder_index.get(_matrix_norm_path(entry["path_parts"]))
                status = "К применению"
                messages: List[str] = list(entry.get("errors") or [])
                resolved_commands: List[Dict[str, Any]] = []
                changes: List[str] = []
                unchanged: List[str] = []

                if not folder:
                    status = "Ошибка"
                    messages.append("Папка не найдена по полному пути")
                for principal_name, desired_ids in entry["commands"].items():
                    resolution = resolution_cache.get(principal_name)
                    if resolution is None:
                        resolution = _matrix_resolve_principal(catalog, principal_name)
                        resolution_cache[principal_name] = resolution
                    if not resolution.get("ok"):
                        status = "Ошибка"
                        messages.append(f"{principal_name}: {resolution.get('message')}")
                        continue
                    principal = dict(resolution["principal"])
                    command = {
                        "requested_name": principal_name,
                        "principal": principal,
                        "desired_ids": set(desired_ids),
                    }
                    resolved_commands.append(command)
                    if folder:
                        current_ids = _matrix_principal_access(folder, principal["type"], principal["id"])
                        command["current_ids"] = current_ids
                        if current_ids == set(desired_ids):
                            unchanged.append(principal_name)
                        else:
                            changes.append(
                                f"{principal_name}: {_matrix_codes_text(current_ids)} → {_matrix_codes_text(set(desired_ids))}"
                            )

                if entry.get("errors"):
                    status = "Ошибка"
                if status != "Ошибка":
                    status = "К применению" if changes else "Без изменений"
                    if unchanged:
                        messages.append(f"Уже соответствует Excel: {len(unchanged)}")
                plan.append({
                    "row_number": entry["row_number"],
                    "path_parts": list(entry["path_parts"]),
                    "path_text": entry["path_text"],
                    "folder_id": folder.get("id") if folder else None,
                    "folder_text": folder.get("path_text") if folder else "",
                    "status": status,
                    "message": "; ".join(messages),
                    "changes": changes,
                    "commands": resolved_commands,
                })

            pending = sum(1 for row in plan if row["status"] == "К применению")
            unchanged = sum(1 for row in plan if row["status"] == "Без изменений")
            errors = sum(1 for row in plan if row["status"] == "Ошибка")
            self.log.emit(f"Предпросмотр: изменить {pending}; без изменений {unchanged}; ошибок {errors}")
            self.finished_ok.emit({
                "plan": plan,
                "sheet_name": parsed["sheet_name"],
                "pending": pending,
                "unchanged": unchanged,
                "errors": errors,
            })
        except Exception:
            self.failed.emit(traceback.format_exc())


class RoleMatrixApplyWorker(QThread):
    log = Signal(str)
    finished_ok = Signal(dict)
    failed = Signal(str)

    def __init__(self, client: LarixAPIClient, project_id: int, plan: List[Dict[str, Any]]):
        super().__init__()
        self.client = client
        self.project_id = int(project_id)
        # Родители раньше детей: изменения родителя Larix наследует вниз.
        self.plan = sorted(
            [row for row in plan if row.get("status") in {"К применению", "Без изменений"}],
            key=lambda row: (len(row.get("path_parts") or []), int(row.get("row_number") or 0)),
        )

    def _verify_row(self, row: Dict[str, Any]) -> bool:
        tree_result = self.client.get_project_folder_permissions(self.project_id)
        if not tree_result.get("success"):
            return False
        folder = _build_matrix_folder_index(tree_result.get("data")).get(_matrix_norm_path(row["path_parts"]))
        if not folder:
            return False
        for command in row.get("commands") or []:
            principal = command["principal"]
            current = _matrix_principal_access(folder, principal["type"], principal["id"])
            if current != set(command["desired_ids"]):
                return False
        return True

    def run(self):
        try:
            stats = {"folders": 0, "buckets": 0, "unchanged": 0, "errors": 0, "skipped": 0, "skipped_buckets": 0}
            rows = [row for row in self.plan if row.get("commands")]
            for idx, row in enumerate(rows, 1):
                self.log.emit(f"[{idx}/{len(rows)}] {row['path_text']}")

                # Перед КАЖДОЙ папкой читаем актуальное состояние. Это важно,
                # потому что Larix наследует изменения родителя на потомков.
                tree_result = self.client.get_project_folder_permissions(self.project_id)
                if not tree_result.get("success"):
                    stats["errors"] += 1
                    self.log.emit("   ❌ дерево прав: " + str(tree_result.get("error")))
                    continue
                folder = _build_matrix_folder_index(tree_result.get("data")).get(_matrix_norm_path(row["path_parts"]))
                if not folder or not folder.get("id"):
                    stats["errors"] += 1
                    self.log.emit("   ❌ папка не найдена")
                    continue

                current = _matrix_access_buckets(folder)
                target = {
                    access_id: {
                        "users": set(bucket["users"]),
                        "roles": set(bucket["roles"]),
                    }
                    for access_id, bucket in current.items()
                }

                for command in row.get("commands") or []:
                    principal = command["principal"]
                    desired = set(command["desired_ids"])
                    bucket_key = "roles" if principal["type"] == "role" else "users"
                    pid = int(principal["id"])
                    for access_id in MATRIX_ACCESS_IDS:
                        if access_id in desired:
                            target[access_id][bucket_key].add(pid)
                        else:
                            target[access_id][bucket_key].discard(pid)

                changed_access_ids = [
                    access_id for access_id in MATRIX_ACCESS_IDS
                    if current[access_id]["users"] != target[access_id]["users"]
                    or current[access_id]["roles"] != target[access_id]["roles"]
                ]
                skipped_access_ids = [
                    access_id for access_id in MATRIX_ACCESS_IDS
                    if access_id not in changed_access_ids
                ]
                stats["skipped_buckets"] += len(skipped_access_ids)
                if skipped_access_ids:
                    self.log.emit(
                        "   = пропуск уже актуальных прав: "
                        + ", ".join(MATRIX_ACCESS_LABELS[item] for item in skipped_access_ids)
                    )
                if not changed_access_ids:
                    stats["unchanged"] += 1
                    self.log.emit("   = папка уже полностью соответствует Excel")
                    continue

                row_ok = True
                for access_id in changed_access_ids:
                    result = self.client.assign_folder_access(
                        self.project_id,
                        int(folder["id"]),
                        access_id,
                        sorted(target[access_id]["users"]),
                        sorted(target[access_id]["roles"]),
                    )
                    if result.get("success"):
                        stats["buckets"] += 1
                        self.log.emit(
                            f"   ✓ {MATRIX_ACCESS_LABELS[access_id]}: "
                            f"users={len(target[access_id]['users'])}, roles={len(target[access_id]['roles'])}"
                        )
                    else:
                        row_ok = False
                        self.log.emit(f"   ❌ {MATRIX_ACCESS_LABELS[access_id]}: {result.get('error')}")
                        break

                if row_ok:
                    stats["folders"] += 1
                else:
                    # Изменяющий POST мог успеть примениться до сетевой ошибки.
                    # Проверяем фактическое состояние, а не повторяем запрос вслепую.
                    if self._verify_row(row):
                        stats["folders"] += 1
                        self.log.emit("   ✓ повторная проверка подтвердила применение")
                    else:
                        stats["errors"] += 1

            stats["skipped"] = sum(1 for row in self.plan if not row.get("commands"))
            self.finished_ok.emit(stats)
        except Exception:
            self.failed.emit(traceback.format_exc())


# =====================================================================
#  Worker: создание проекта
# =====================================================================
def _type_plan_item_id(item: Dict[str, Any]) -> Optional[int]:
    for key in ("id", "typeId", "remarkTypeId", "taskTypeId"):
        try:
            if item.get(key) is not None:
                return int(item[key])
        except (TypeError, ValueError):
            continue
    return None


def _type_plan_item_name(item: Dict[str, Any]) -> str:
    for key in ("title", "name", "typeName", "description"):
        value = _type_import_text(item.get(key))
        if value:
            return value
    return ""


def _type_plan_user_email(item: Dict[str, Any]) -> str:
    return _type_import_text(item.get("email") or item.get("userEmail") or item.get("login")).lower()


def _type_plan_user_display_name(item: Any) -> str:
    """Get the exact full-name label used by the wide Excel matrix."""
    return _type_catalog_display_name(item)


def _type_api_items_robust(payload: Any) -> List[Dict[str, Any]]:
    """Flatten Larix catalog envelopes, including data.roles JSON strings."""
    if isinstance(payload, str):
        text = payload.strip()
        if text.startswith(("{", "[")):
            try:
                return _type_api_items_robust(json.loads(text))
            except (TypeError, ValueError, json.JSONDecodeError):
                return []
        return []
    if isinstance(payload, list):
        result: List[Dict[str, Any]] = []
        for item in payload:
            if isinstance(item, dict):
                result.append(item)
            elif isinstance(item, str):
                result.extend(_type_api_items_robust(item))
        return result
    if not isinstance(payload, dict):
        return []
    for key in ("items", "content", "rows", "results", "roles", "users"):
        if key in payload:
            items = _type_api_items_robust(payload[key])
            if items:
                return items
    for key in ("data", "result"):
        if key in payload:
            items = _type_api_items_robust(payload[key])
            if items:
                return items
    if any(key in payload for key in ("id", "roleId", "title", "name")):
        return [payload]
    return []


LarixAPIClient._type_api_items = staticmethod(_type_api_items_robust)


class TypeImportPreviewWorker(QThread):
    log = Signal(str)
    progress = Signal(int)
    finished_ok = Signal(dict)
    failed = Signal(str)

    def __init__(self, client: LarixAPIClient, project_id: int, excel_path: str, kind: str, sheet_name: Optional[str] = None):
        super().__init__()
        self.client, self.project_id = client, int(project_id)
        self.excel_path, self.kind, self.sheet_name = excel_path, kind, sheet_name

    def run(self):
        try:
            self.log.emit("Предпросмотр типов: загружаю каталог Larix...")
            types_result = self.client.list_type_types(self.project_id, self.kind)
            roles_result = self.client.list_project_roles(self.project_id)
            users_result = self.client.list_workspace_users()
            for label, result in (("типов", types_result), ("ролей", roles_result), ("пользователей", users_result)):
                if not result.get("success"):
                    raise RuntimeError(f"Не удалось загрузить каталог {label}: {result.get('error', 'неизвестная ошибка')}")
            role_items = roles_result.get("items", [])
            user_items = users_result.get("items", [])
            known_roles = {_type_import_text(_type_plan_item_name(item)): _type_plan_item_id(item) for item in role_items}
            known_users = {_type_plan_user_email(item): _type_plan_item_id(item) for item in user_items}
            parsed = parse_type_import_excel(self.excel_path, kind=self.kind,
                                             known_roles=list(known_roles), known_users=list(known_users))
            existing = {_type_plan_item_name(item): _type_plan_item_id(item) for item in types_result.get("items", [])}
            plan_rows: List[Dict[str, Any]] = []
            for index, item in enumerate(parsed.types, 1):
                role_ids_view = sorted(known_roles[name] for name in item.role_view if known_roles.get(name) is not None)
                role_ids_create = sorted(known_roles[name] for name in item.role_create if known_roles.get(name) is not None)
                user_ids_view = sorted(known_users[email.lower()] for email in item.user_view if known_users.get(email.lower()) is not None)
                user_ids_create = sorted(known_users[email.lower()] for email in item.user_create if known_users.get(email.lower()) is not None)
                type_id = existing.get(item.name)
                plan_rows.append({"name": item.name, "type_id": type_id,
                                  "action": "update" if type_id is not None else "create",
                                  "role_view_ids": role_ids_view, "role_create_ids": role_ids_create,
                                  "user_view_ids": user_ids_view, "user_create_ids": user_ids_create,
                                  "source": item.as_dict()})
                self.progress.emit(int(index * 100 / max(1, len(parsed.types))))
            result = {"kind": self.kind, "project_id": self.project_id, "source_path": parsed.source_path,
                      "plan": plan_rows, "types": parsed.as_dict(), "errors": []}
            self.log.emit(f"Предпросмотр готов: {len(plan_rows)} типов")
            self.finished_ok.emit(result)
        except Exception:
            self.failed.emit(traceback.format_exc())


def _type_import_preview_run(self) -> None:
    try:
        self.log.emit("Предпросмотр: загрузка каталогов ролей и пользователей Larix...")
        types_result = self.client.list_type_types(self.project_id, self.kind)
        roles_result = self.client.list_project_roles(self.project_id)
        users_result = self.client.list_workspace_users()
        for label, result in (("типов", types_result), ("ролей", roles_result), ("пользователей", users_result)):
            if not result.get("success"):
                raise RuntimeError(f"Не удалось получить каталог {label}: {result.get('error', 'неизвестная ошибка')}")
        role_items = roles_result.get("items", [])
        user_items = users_result.get("items", [])
        known_roles = {_type_import_text(_type_plan_item_name(item)): _type_plan_item_id(item) for item in role_items}
        known_users = {_type_plan_user_display_name(item): _type_plan_item_id(item) for item in user_items}
        known_users = {name: uid for name, uid in known_users.items() if name and uid is not None}
        parsed = parse_type_import_excel(self.excel_path, kind=self.kind,
                                         known_roles=list(known_roles), known_users=list(known_users),
                                         sheet_name=self.sheet_name)
        existing = {_type_plan_item_name(item): _type_plan_item_id(item) for item in types_result.get("items", [])}
        plan_rows: List[Dict[str, Any]] = []
        for index, item in enumerate(parsed.types, 1):
            role_ids_view = sorted(known_roles[name] for name in item.role_view if known_roles.get(name) is not None)
            role_ids_create = sorted(known_roles[name] for name in item.role_create if known_roles.get(name) is not None)
            user_ids_view = sorted(known_users[name] for name in item.user_view if known_users.get(name) is not None)
            user_ids_create = sorted(known_users[name] for name in item.user_create if known_users.get(name) is not None)
            type_id = existing.get(item.name)
            plan_rows.append({"name": item.name, "type_id": type_id,
                              "action": "update" if type_id is not None else "create",
                              "role_view_ids": role_ids_view, "role_create_ids": role_ids_create,
                              "user_view_ids": user_ids_view, "user_create_ids": user_ids_create,
                              "role_view_names": sorted(item.role_view), "role_create_names": sorted(item.role_create),
                              "user_view_names": sorted(item.user_view), "user_create_names": sorted(item.user_create),
                              "source": item.as_dict()})
            self.progress.emit(int(index * 100 / max(1, len(parsed.types))))
        result = {"kind": self.kind, "project_id": self.project_id, "sheet_name": self.sheet_name,
                  "source_path": parsed.source_path,
                  "plan": plan_rows, "parsed": parsed.as_dict()}
        self.log.emit(f"Предпросмотр готов: {len(plan_rows)} типов.")
        self.finished_ok.emit(result)
    except Exception:
        self.failed.emit(traceback.format_exc())


TypeImportPreviewWorker.run = _type_import_preview_run


class TypeImportApplyWorker(QThread):
    log = Signal(str)
    progress = Signal(int)
    finished_ok = Signal(dict)
    failed = Signal(str)

    def __init__(self, client: LarixAPIClient, project_id: int, kind: str, plan: List[Dict[str, Any]]):
        super().__init__()
        self.client, self.project_id, self.kind = client, int(project_id), kind
        self.plan = list(plan or [])

    def run(self):
        report: List[Dict[str, Any]] = []
        total = max(1, len(self.plan))
        for index, row in enumerate(self.plan, 1):
            item_report = {"name": row.get("name", ""), "action": row.get("action", "update"), "status": "ok", "errors": []}
            try:
                type_id = row.get("type_id")
                if type_id is None:
                    created = self.client.create_type(self.project_id, row["name"], self.kind)
                    if not created.get("success"):
                        raise RuntimeError(created.get("error", "не удалось создать тип"))
                    data = created.get("data")
                    type_id = _type_plan_item_id(data) if isinstance(data, dict) else None
                    if type_id is None and isinstance(data, dict):
                        type_id = _type_plan_item_id(data.get("data", {}))
                    if type_id is None:
                        raise RuntimeError("API не вернул ID созданного типа")
                    item_report["type_id"] = type_id
                else:
                    item_report["type_id"] = int(type_id)
                for access_role_id, users_key, roles_key in ((8, "user_view_ids", "role_view_ids"), (10, "user_create_ids", "role_create_ids")):
                    result = self.client.set_type_access(type_id, self.kind, row.get(users_key, []), row.get(roles_key, []), access_role_id)
                    if not result.get("success"):
                        raise RuntimeError(f"accessRoleId={access_role_id}: {result.get('error', 'ошибка API')}")
                self.log.emit(f"✓ {row.get('name', '')}: обработан")
            except Exception as exc:
                item_report["status"] = "error"
                item_report["errors"].append(str(exc))
                self.log.emit(f"✗ {row.get('name', '')}: {exc}")
            report.append(item_report)
            self.progress.emit(int(index * 100 / total))
        stats = {"total": len(report), "success": sum(item["status"] == "ok" for item in report),
                 "errors": sum(item["status"] == "error" for item in report)}
        self.finished_ok.emit({"kind": self.kind, "project_id": self.project_id, "report": report, "stats": stats})



class ApprovalPreviewWorker(QThread):
    log = Signal(str)
    progress = Signal(int)
    finished_ok = Signal(dict)
    failed = Signal(str)

    def __init__(
        self, client: LarixAPIClient, project_id: int, excel_path: str,
        main_sheet_name: str = "", settings_sheet_name: str = "", flags_sheet_name: str = "",
    ):
        super().__init__()
        self.client = client
        self.project_id = int(project_id)
        self.excel_path = excel_path
        self.main_sheet_name = _approval_clean_text(main_sheet_name)
        self.settings_sheet_name = _approval_clean_text(settings_sheet_name)
        self.flags_sheet_name = _approval_clean_text(flags_sheet_name)

    def run(self):
        try:
            self.log.emit("Маршруты согласований: читаю Excel...")
            parsed = parse_approval_workflows_excel(
                self.excel_path,
                main_sheet_name=self.main_sheet_name,
                settings_sheet_name=self.settings_sheet_name,
                flags_sheet_name=self.flags_sheet_name,
            )
            self.progress.emit(15)

            self.log.emit("Маршруты согласований: загружаю роли и пользователей проекта...")
            roles_result = self.client.list_project_roles(self.project_id)
            if not roles_result.get("success"):
                raise RuntimeError(f"Не удалось получить роли/пользователей проекта: {roles_result.get('error', 'неизвестная ошибка')}")
            catalogs = _approval_build_catalogs(roles_result)
            self.progress.emit(35)

            self.log.emit("Маршруты согласований: проверяю существующие маршруты...")
            workflow_result = self.client.list_approval_workflows(self.project_id)
            if not workflow_result.get("success"):
                raise RuntimeError(f"Не удалось получить маршруты проекта: {workflow_result.get('error', 'неизвестная ошибка')}")
            existing = {
                _approval_norm(item.get("title")): item
                for item in workflow_result.get("items", [])
                if _approval_clean_text(item.get("title"))
            }
            self.progress.emit(50)

            plan: List[Dict[str, Any]] = []
            total = max(1, len(parsed["workflows"]))
            for index, workflow in enumerate(parsed["workflows"], 1):
                row = {
                    "title": workflow["title"],
                    "steps": [],
                    "view_role_ids": [],
                    "view_user_ids": [],
                    "create_role_ids": [],
                    "create_user_ids": [],
                    "view_names": list(workflow.get("view_principals", [])),
                    "create_names": list(workflow.get("create_principals", [])),
                    "errors": list(workflow.get("errors", [])),
                    "action": "create",
                    "existing_id": None,
                    "source_row": workflow.get("source_row"),
                }

                existing_item = existing.get(_approval_norm(workflow["title"]))
                if existing_item is not None:
                    row["action"] = "skip"
                    row["existing_id"] = existing_item.get("id")

                for step in workflow.get("steps", []):
                    resolved_step = {
                        "index": int(step.get("index", len(row["steps"]))),
                        "title": step.get("title") or f"{len(row['steps']) + 1} Этап",
                        "duration": step.get("duration"),
                        "users": [],
                    }
                    for participant in step.get("users", []):
                        source_name = participant.get("source_name", "")
                        resolved = _approval_resolve_user(source_name, catalogs)
                        if not resolved.get("ok"):
                            row["errors"].append(f"{resolved_step['title']}: {resolved.get('error')}")
                            continue
                        user = resolved["item"]
                        resolved_step["users"].append({
                            "user_id": int(user["id"]),
                            "user_name": user["name"],
                            "source_name": source_name,
                            "can_cancel": bool(participant.get("can_cancel", False)),
                            "must_approve": bool(participant.get("must_approve", False)),
                        })
                    if not resolved_step["users"]:
                        row["errors"].append(f"{resolved_step['title']}: не найден ни один согласующий")
                    row["steps"].append(resolved_step)

                for principal_names, role_key, user_key in (
                    (workflow.get("view_principals", []), "view_role_ids", "view_user_ids"),
                    (workflow.get("create_principals", []), "create_role_ids", "create_user_ids"),
                ):
                    for principal_name in principal_names:
                        resolved = _approval_resolve_access_principal(principal_name, catalogs)
                        if not resolved.get("ok"):
                            row["errors"].append(resolved.get("error", f"не найден получатель «{principal_name}»"))
                            continue
                        item = resolved["item"]
                        target_key = role_key if resolved["kind"] == "role" else user_key
                        row[target_key].append(int(item["id"]))

                for key in ("view_role_ids", "view_user_ids", "create_role_ids", "create_user_ids"):
                    row[key] = sorted(set(row[key]))

                if row["errors"]:
                    row["action"] = "error"
                plan.append(row)
                self.progress.emit(50 + int(index * 50 / total))

            stats = {
                "total": len(plan),
                "create": sum(item["action"] == "create" for item in plan),
                "skip": sum(item["action"] == "skip" for item in plan),
                "errors": sum(item["action"] == "error" for item in plan),
            }
            self.log.emit(
                f"Предпросмотр маршрутов готов: создать {stats['create']}, "
                f"уже существуют {stats['skip']}, ошибок {stats['errors']}."
            )
            self.finished_ok.emit({
                "project_id": self.project_id,
                "source_path": parsed["source_path"],
                "plan": plan,
                "stats": stats,
                "sheets": parsed["sheets"],
            })
        except Exception:
            self.failed.emit(traceback.format_exc())


class ApprovalApplyWorker(QThread):
    log = Signal(str)
    progress = Signal(int)
    finished_ok = Signal(dict)
    failed = Signal(str)

    def __init__(self, client: LarixAPIClient, project_id: int, plan: List[Dict[str, Any]]):
        super().__init__()
        self.client = client
        self.project_id = int(project_id)
        self.plan = list(plan or [])

    def run(self):
        try:
            report: List[Dict[str, Any]] = []
            candidates = [item for item in self.plan if item.get("action") == "create"]
            total = max(1, len(candidates))
            done = 0
            for row in self.plan:
                if row.get("action") == "skip":
                    report.append({
                        "title": row.get("title", ""),
                        "status": "skipped",
                        "workflow_id": row.get("existing_id"),
                        "errors": [],
                    })
                    continue
                if row.get("action") == "error":
                    report.append({
                        "title": row.get("title", ""),
                        "status": "error",
                        "workflow_id": None,
                        "errors": list(row.get("errors", [])),
                    })
                    continue

                item_report = {
                    "title": row.get("title", ""),
                    "status": "ok",
                    "workflow_id": None,
                    "errors": [],
                }
                try:
                    created = self.client.create_approval_workflow(
                        self.project_id, row["title"], row.get("steps", [])
                    )
                    if not created.get("success"):
                        raise RuntimeError(created.get("error", "не удалось создать маршрут"))
                    data = created.get("data", {})
                    workflow_id = data.get("id") if isinstance(data, dict) else None
                    if workflow_id is None:
                        raise RuntimeError("API не вернул ID созданного маршрута")
                    workflow_id = int(workflow_id)
                    item_report["workflow_id"] = workflow_id

                    for access_role_id, users_key, roles_key, label in (
                        (APPROVAL_ACCESS_VIEW_ID, "view_user_ids", "view_role_ids", "Просмотр"),
                        (APPROVAL_ACCESS_CREATE_ID, "create_user_ids", "create_role_ids", "Создание"),
                    ):
                        if not row.get(users_key) and not row.get(roles_key):
                            continue
                        access_result = self.client.set_approval_workflow_access(
                            workflow_id,
                            row.get(users_key, []),
                            row.get(roles_key, []),
                            access_role_id,
                        )
                        if not access_result.get("success"):
                            raise RuntimeError(f"{label}: {access_result.get('error', 'ошибка назначения прав')}")
                    self.log.emit(f"✓ {row.get('title', '')}: маршрут создан")
                except Exception as exc:
                    item_report["status"] = "error"
                    item_report["errors"].append(str(exc))
                    self.log.emit(f"✗ {row.get('title', '')}: {exc}")
                report.append(item_report)
                done += 1
                self.progress.emit(int(done * 100 / total))

            stats = {
                "total": len(report),
                "success": sum(item["status"] == "ok" for item in report),
                "skipped": sum(item["status"] == "skipped" for item in report),
                "errors": sum(item["status"] == "error" for item in report),
            }
            self.finished_ok.emit({
                "project_id": self.project_id,
                "report": report,
                "stats": stats,
            })
        except Exception:
            self.failed.emit(traceback.format_exc())


class CreateProjectWorker(QThread):
    log = Signal(str)
    finished_ok = Signal(dict)
    failed = Signal(str)

    def __init__(self, client: LarixAPIClient, workspace_id: int, title: str, description: str,
                 create_with_structure: bool, default_structure_path: str):
        super().__init__()
        self.client = client
        self.workspace_id = int(workspace_id)
        self.title = title
        self.description = description or ""
        self.create_with_structure = create_with_structure
        self.default_structure_path = default_structure_path

    def run(self):
        try:
            self.log.emit(f"Создание проекта '{self.title}' в пространстве {self.workspace_id}...")
            proj = self.client.create_project(self.workspace_id, self.title, self.description)
            if not proj.get("success"):
                self.failed.emit(f"Не удалось создать проект: {proj.get('error')}")
                return
            project_id = proj["project_id"]
            self.log.emit(f"✅ Проект создан. ID = {project_id}")
            folders_created = 0
            if self.create_with_structure and self.default_structure_path:
                if os.path.exists(self.default_structure_path):
                    self.log.emit("📁 Создание типового дерева папок...")
                    result = self.client.create_folders_from_excel(
                        self.default_structure_path, STRUCTURE_SHEET_NAME,
                        project_id, log=lambda m: self.log.emit(m), delay=0.3
                    )
                    if result.get("success"):
                        folders_created = result.get("created", 0)
                        self.log.emit(f"✅ Создано папок: {folders_created}")
                    else:
                        self.log.emit(f"⚠️ Ошибка при создании папок: {result.get('error')}")
                else:
                    self.log.emit(f"⚠️ Файл структуры не найден: {self.default_structure_path}")
            self.finished_ok.emit({
                "project_id": project_id,
                "title": self.title,
                "folders_created": folders_created,
                "structure_requested": bool(self.create_with_structure),
            })
        except Exception as e:
            self.failed.emit(f"Исключение: {e}\n{traceback.format_exc()}")


# =====================================================================
#  Worker: добавление структуры в существующий проект
# =====================================================================
class AddStructureWorker(QThread):
    log = Signal(str)
    finished_ok = Signal(dict)
    failed = Signal(str)

    def __init__(self, client: LarixAPIClient, project_id: int, project_title: str,
                 structure_path: str):
        super().__init__()
        self.client = client
        self.project_id = int(project_id)
        self.project_title = project_title
        self.structure_path = structure_path

    def run(self):
        try:
            if not self.structure_path or not os.path.exists(self.structure_path):
                self.failed.emit("Excel-файл структуры не найден")
                return

            self.log.emit(
                f"📁 Добавление структуры в существующий проект "
                f"[{self.project_id}] {self.project_title}..."
            )
            result = self.client.create_folders_from_excel(
                self.structure_path,
                STRUCTURE_SHEET_NAME,
                self.project_id,
                log=lambda m: self.log.emit(m),
                delay=0.3,
            )
            if not result.get("success"):
                self.failed.emit(result.get("error", "Не удалось создать структуру"))
                return

            self.finished_ok.emit({
                "project_id": self.project_id,
                "title": self.project_title,
                "folders_created": int(result.get("created", 0)),
                "errors": list(result.get("errors") or []),
            })
        except Exception as exc:
            self.failed.emit(f"Исключение: {exc}\n{traceback.format_exc()}")


# =====================================================================
#  Диалог создания проекта
# =====================================================================
class CreateProjectDialog(QtWidgets.QDialog):
    def __init__(self, parent: QtWidgets.QWidget, client: LarixAPIClient, workspace_id: int,
                 workspace_display: str, default_structure_path: str = ""):
        super().__init__(parent)
        self.setWindowTitle("Создание проекта")
        self.resize(520, 380)
        self.setStyleSheet(parent.styleSheet())
        self.client = client
        self.workspace_id = workspace_id
        self.workspace_display = workspace_display
        self.default_structure_path = default_structure_path
        self._worker: Optional[CreateProjectWorker] = None
        self.result_data: Optional[Dict[str, Any]] = None
        self._build_ui()

    def _build_ui(self):
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 18)
        root.setSpacing(10)
        info = QtWidgets.QLabel(f"Пространство: <b>{self.workspace_display}</b>")
        root.addWidget(info)
        form = QtWidgets.QGridLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        self.ed_title = QtWidgets.QLineEdit()
        self.ed_title.setPlaceholderText("Например: Проект_API")
        self.ed_description = QtWidgets.QTextEdit()
        self.ed_description.setPlaceholderText("Необязательное описание")
        self.ed_description.setMaximumHeight(60)
        self.cb_create_structure = QtWidgets.QCheckBox("Создать с типовым деревом проекта")
        self.cb_create_structure.setChecked(True)
        self.cb_create_structure.stateChanged.connect(self._on_structure_changed)
        self.lbl_structure_status = QtWidgets.QLabel()
        self.lbl_structure_status.setWordWrap(True)
        self.btn_structure = QtWidgets.QPushButton("Выбрать другой...")
        self.btn_structure.clicked.connect(self._pick_structure_file)
        form.addWidget(QtWidgets.QLabel("Название проекта*:"), 0, 0)
        form.addWidget(self.ed_title, 0, 1, 1, 2)
        form.addWidget(QtWidgets.QLabel("Описание:"), 1, 0, QtCore.Qt.AlignTop)
        form.addWidget(self.ed_description, 1, 1, 1, 2)
        form.addWidget(self.cb_create_structure, 2, 0, 1, 3)
        form.addWidget(QtWidgets.QLabel("Структура:"), 3, 0)
        form.addWidget(self.lbl_structure_status, 3, 1)
        form.addWidget(self.btn_structure, 3, 2)
        self._refresh_structure_status()
        form.setColumnStretch(1, 1)
        root.addLayout(form)
        self.log = QtWidgets.QPlainTextEdit(self)
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("Журнал создания...")
        self.log.setFixedHeight(100)
        root.addWidget(self.log)
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch(1)
        self.btn_cancel = QtWidgets.QPushButton("Закрыть")
        self.btn_cancel.clicked.connect(self.close)
        self.btn_create = QtWidgets.QPushButton("Создать проект")
        self.btn_create.setObjectName("accent")
        self.btn_create.setMinimumHeight(38)
        if hasattr(self.parent(), "_set_button_icon"):
            self.parent()._set_button_icon(self.btn_create, "free-icon-plus-3303893.png")
            self.parent()._set_button_icon(self.btn_structure, "folder_icon_variant_1.png")
        self.btn_create.clicked.connect(self._on_create)
        btn_row.addWidget(self.btn_cancel)
        btn_row.addWidget(self.btn_create)
        root.addLayout(btn_row)

    def _refresh_structure_status(self) -> None:
        if self.default_structure_path and os.path.exists(self.default_structure_path):
            text = "Шаблон структуры загружен. При включённой галочке проект будет создан с папками."
            tooltip = self.default_structure_path
        else:
            text = "Шаблон структуры не выбран. Проект можно создать, но папки автоматически не создадутся."
            tooltip = "Выберите Excel-файл структуры проекта"
        self.lbl_structure_status.setText(text)
        self.lbl_structure_status.setToolTip(tooltip)

    def _on_structure_changed(self, state):
        enabled = state == QtCore.Qt.Checked
        self.lbl_structure_status.setEnabled(enabled)
        self.btn_structure.setEnabled(enabled)

    def _pick_structure_file(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Выберите Excel-файл со структурой проекта",
            "", "Excel files (*.xlsx *.xls *.xlsm);;All files (*.*)"
        )
        if path:
            self.default_structure_path = path
            self._refresh_structure_status()

    def _log(self, msg: str):
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
        self.log.appendPlainText(line)
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def _on_create(self):
        title = self.ed_title.text().strip()
        if not title:
            QtWidgets.QMessageBox.warning(self, "Ошибка", "Введите название проекта")
            return
        description = self.ed_description.toPlainText().strip()
        create_structure = self.cb_create_structure.isChecked()
        structure_path = self.default_structure_path
        self.btn_create.setEnabled(False)
        self.btn_cancel.setEnabled(False)
        self.log.clear()
        self._worker = CreateProjectWorker(
            self.client, self.workspace_id, title, description, create_structure, structure_path
        )
        self._worker.log.connect(self._log)
        self._worker.finished_ok.connect(self._on_finished_ok)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(lambda: (self.btn_create.setEnabled(True), self.btn_cancel.setEnabled(True)))
        self._worker.start()

    def _on_finished_ok(self, data: Dict[str, Any]):
        self.result_data = data
        self._log("")
        self._log("=" * 40)
        self._log(f"✅ Готово! Проект ID={data['project_id']} '{data['title']}'")
        self._log(f"📁 Папок создано: {data['folders_created']}")
        QtWidgets.QMessageBox.information(
            self, "Успех",
            f"Проект создан.\nID: {data['project_id']}\nНазвание: {data['title']}\n"
            f"Папок создано: {data['folders_created']}"
        )
        self.accept()

    def _on_failed(self, err: str):
        self._log(f"❌ Ошибка: {err}")
        QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось создать проект:\n{err}")


# =====================================================================
#  Диалог выбора файла типовой структуры
# =====================================================================
class TemplateSettingsDialog(QtWidgets.QDialog):
    def __init__(self, parent: QtWidgets.QWidget, structure_path: str):
        super().__init__(parent)
        self.setWindowTitle("Настройка шаблона структуры")
        self.resize(680, 170)
        self.setStyleSheet(parent.styleSheet())
        self.structure_path = structure_path or ""
        self._build_ui()

    def _build_ui(self) -> None:
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(18, 16, 18, 16)
        root.setSpacing(10)

        hint = QtWidgets.QLabel(
            "Здесь выбирается только Excel-файл типовой структуры. Шаблон пользователей не настраивается: "
            "его можно скачать и заполнить, а готовую таблицу загрузить ниже в поле 'Excel (пользователи)'."
        )
        hint.setWordWrap(True)
        root.addWidget(hint)

        row = QtWidgets.QHBoxLayout()
        self.lbl_structure = QtWidgets.QLabel()
        self.lbl_structure.setWordWrap(True)
        self.btn_pick_structure = QtWidgets.QPushButton("Выбрать файл...")
        self.btn_pick_structure.clicked.connect(self._pick_structure)
        row.addWidget(self.lbl_structure, 1)
        row.addWidget(self.btn_pick_structure)
        root.addLayout(row)
        self._refresh_status()

        btn_row = QtWidgets.QHBoxLayout()
        self.btn_auto = QtWidgets.QPushButton("Встроенный шаблон")
        self.btn_auto.setToolTip("Вернуть встроенный эталонный шаблон структуры")
        self.btn_auto.clicked.connect(self._auto_find)
        self.btn_clear = QtWidgets.QPushButton("Очистить")
        self.btn_clear.clicked.connect(self._clear_path)
        btn_row.addWidget(self.btn_auto)
        btn_row.addWidget(self.btn_clear)
        btn_row.addStretch(1)

        self.btn_cancel = QtWidgets.QPushButton("Отмена")
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_save = QtWidgets.QPushButton("Сохранить")
        self.btn_save.setObjectName("accent")
        if hasattr(self.parent(), "_set_button_icon"):
            self.parent()._set_button_icon(self.btn_pick_structure, "folder_icon_variant_1.png")
            self.parent()._set_button_icon(self.btn_save, "krug_galka.png")
        self.btn_save.clicked.connect(self._validate_and_accept)
        btn_row.addWidget(self.btn_cancel)
        btn_row.addWidget(self.btn_save)
        root.addLayout(btn_row)

    def _refresh_status(self) -> None:
        if self.structure_path and os.path.exists(self.structure_path):
            self.lbl_structure.setText("Шаблон структуры выбран")
            self.lbl_structure.setToolTip(self.structure_path)
        else:
            self.lbl_structure.setText("Шаблон структуры не выбран")
            self.lbl_structure.setToolTip("По умолчанию используется встроенный эталонный шаблон")

    def _pick_structure(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Выберите Excel-файл типовой структуры проекта",
            "", "Excel files (*.xlsx *.xls *.xlsm);;All files (*.*)"
        )
        if path:
            self.structure_path = path
            self._refresh_status()

    def _auto_find(self) -> None:
        self.structure_path = _ensure_builtin_structure_template()
        self._refresh_status()

    def _clear_path(self) -> None:
        self.structure_path = ""
        self._refresh_status()

    def _validate_and_accept(self) -> None:
        if self.structure_path:
            if not os.path.exists(self.structure_path):
                QtWidgets.QMessageBox.warning(self, "Ошибка", f"Файл не найден:\n{self.structure_path}")
                return
            try:
                sheets = [str(s) for s in pd.ExcelFile(self.structure_path).sheet_names]
            except Exception as exc:
                QtWidgets.QMessageBox.warning(self, "Ошибка", f"Не удалось прочитать файл структуры:\n{exc}")
                return
            if STRUCTURE_SHEET_NAME not in sheets:
                QtWidgets.QMessageBox.warning(
                    self, "Ошибка",
                    f"В файле структуры должен быть лист '{STRUCTURE_SHEET_NAME}'.\n"
                    f"Найдены листы: {', '.join(sheets) if sheets else 'нет листов'}"
                )
                return
        self.accept()


# =====================================================================
#  Главное окно
# =====================================================================
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Larix CDE — Larix Platform")
        self.resize(1200, 700)
        self.setMinimumSize(1020, 620)
        self.asset_dir = shared_asset_dir(Path(__file__).resolve().parent)
        app_icon_path = self.asset_dir / "icon.ico"
        if app_icon_path.exists():
            self.setWindowIcon(QtGui.QIcon(str(app_icon_path)))
        self.client = LarixAPIClient(log_func=self._log)
        self.workspace_map: Dict[str, int] = {}
        self.project_map: Dict[str, Optional[int]] = {}
        self.project_name_map: Dict[str, int] = {}
        self.role_map: Dict[str, Dict[str, Any]] = {}
        self.loaded_users: List[Dict[str, Any]] = []
        self.excel_path: Optional[str] = None
        self.is_dark_theme = initial_dark_theme(False)
        self._password_visible = False
        self._log_dialog: Optional[QtWidgets.QDialog] = None
        self._log_view: Optional[QtWidgets.QPlainTextEdit] = None
        self._project_worker: Optional[CreateProjectWorker] = None
        self._structure_worker: Optional[AddStructureWorker] = None
        self._syncing_workspace_combos = False
        self.role_matrix_path: Optional[str] = None
        self.role_matrix_plan: List[Dict[str, Any]] = []
        # Последняя успешная сверка хранится отдельно от плана применения.
        # После применения plan очищается, но отчет по выполненной сверке
        # остается доступен для просмотра и выгрузки в Excel.
        self.role_matrix_last_report: List[Dict[str, Any]] = []
        self._matrix_preview_worker: Optional[RoleMatrixPreviewWorker] = None
        self._matrix_apply_worker: Optional[RoleMatrixApplyWorker] = None
        self._matrix_apply_after_preview = False
        self.type_import_state: Dict[str, Dict[str, Any]] = {}
        self.approval_state: Dict[str, Any] = {}

        # =============================================================
        #  Эталонные Excel-шаблоны вшиты в этот .py и используются только
        #  кнопками «Скачать шаблон». Они НЕ считаются загруженными файлами.
        #
        #  Файл структуры для фактического создания папок пользователь
        #  выбирает вручную при каждом запуске программы.
        # =============================================================
        self.default_structure_path: str = ""
        self.user_import_template_path: str = ""

        # Удаляем сохранённый старой версией путь, чтобы после обновления
        # программа также запускалась без автоматически выбранной структуры.
        try:
            _save_template_settings("", "")
        except Exception:
            pass

        self._setup_ui()
        install_status_bar(self)
        mark_destructive_buttons(self)
        self._apply_styles()
        install_window_state(self, "cde_tool")

    def _toggle_password_visibility(self) -> None:
        self._password_visible = not self._password_visible
        mode = QtWidgets.QLineEdit.Normal if self._password_visible else QtWidgets.QLineEdit.Password
        self.ed_password.setEchoMode(mode)
        self._update_password_eye_text()

    def _update_password_eye_text(self) -> None:
        if not hasattr(self, "ed_password"):
            return
        icon_name = "free-icon-hide-11238328.png" if self._password_visible else "free-icon-eye-2455724.png"
        self.ed_password._eye_btn.setText("")
        self.ed_password._eye_btn.setIcon(themed_icon(self.asset_dir, icon_name, self.is_dark_theme))
        self.ed_password._eye_btn.setIconSize(QtCore.QSize(18, 18))
        self.ed_password.set_eye_tooltip("Скрыть пароль" if self._password_visible else "Показать пароль")

    def _set_native_titlebar_theme(self, dark: bool) -> None:
        apply_windows_titlebar_theme(self, bool(dark))

    def _toggle_theme(self, dark: Optional[bool] = None) -> None:
        self.is_dark_theme = (not self.is_dark_theme) if dark is None else bool(dark)
        persist_dark_theme(self.is_dark_theme)
        self._apply_styles()

    def _update_theme_button(self) -> None:
        toggle = getattr(self, "theme_toggle", None)
        if toggle is None:
            return
        toggle.blockSignals(True)
        toggle.setChecked(self.is_dark_theme, animate=False)
        toggle.blockSignals(False)

    def _set_auth_status(self, authorized: bool) -> None:
        text = "Авторизован" if authorized else "Не авторизован"
        color = SUCCESS if authorized else "#777777"
        self.auth_text.setText(f"Статус: {text}")
        self.auth_text.setStyleSheet(f"color: {color}; font-weight: 600; background: transparent;")

    def _setup_combo_popup(self, combo: QtWidgets.QComboBox) -> None:
        view = QtWidgets.QListView(combo)
        view.setObjectName("comboPopupView")
        view.viewport().setObjectName("comboPopupViewport")
        view.setMouseTracking(True)
        view.setSpacing(1)
        view.setUniformItemSizes(False)
        # Никакой собственной рамки/скругления у QListView: иначе в углах
        # просвечивает системный QComboBoxPrivateContainer (белые треугольники
        # в dark mode). Скругляем только сами строки через delegate.
        view.setFrameShape(QtWidgets.QFrame.NoFrame)
        view.setLineWidth(0)
        view.setContentsMargins(0, 0, 0, 0)
        view.setItemDelegate(RoundedComboDelegate(view))
        combo.setView(view)

    def _show_log(self) -> None:
        if self._log_dialog is not None and self._log_dialog.isVisible():
            self._log_dialog.raise_()
            self._log_dialog.activateWindow()
            return
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Log")
        dlg.resize(720, 420)
        dlg.setStyleSheet(self.styleSheet())
        dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        layout = QtWidgets.QVBoxLayout(dlg)
        view = QtWidgets.QPlainTextEdit()
        view.setReadOnly(True)
        view.setPlainText(self.log.toPlainText())
        view.verticalScrollBar().setValue(view.verticalScrollBar().maximum())
        layout.addWidget(view)
        btn_close = QtWidgets.QPushButton("Закрыть")
        btn_close.clicked.connect(dlg.close)
        layout.addWidget(btn_close, 0, QtCore.Qt.AlignRight)
        self._log_dialog = dlg
        self._log_view = view
        def _reset_log_dialog():
            self._log_dialog = None
            self._log_view = None
        dlg.destroyed.connect(_reset_log_dialog)
        dlg.show()

    def _show_alert(self, title: str, text: str) -> None:
        self._show_message(title, text, QtWidgets.QMessageBox.Ok)

    def _show_success(self, title: str, text: str) -> None:
        self._show_message(title, text, QtWidgets.QMessageBox.Ok)

    def _ask_alert(self, title: str, text: str) -> bool:
        msg = self._make_message_box(title, text)
        yes_btn = msg.addButton("Да", QtWidgets.QMessageBox.YesRole)
        no_btn = msg.addButton("Нет", QtWidgets.QMessageBox.NoRole)
        msg.setDefaultButton(no_btn)
        msg.exec()
        return msg.clickedButton() == yes_btn

    def _show_message(self, title: str, text: str, buttons) -> None:
        msg = self._make_message_box(title, text)
        msg.setStandardButtons(buttons)
        msg.exec()

    def _make_message_box(self, title: str, text: str) -> QtWidgets.QMessageBox:
        msg = QtWidgets.QMessageBox(self)
        msg.setWindowTitle(title)
        msg.setText(text)
        msg.setStyleSheet(self.styleSheet())
        return msg

    def _set_button_icon(self, button: QtWidgets.QAbstractButton, file_name: str, size: int = 16) -> None:
        if button is None:
            return
        button.setProperty("iconAsset", file_name)
        icon = themed_icon(self.asset_dir, file_name, self.is_dark_theme)
        button.setIcon(icon)
        button.setIconSize(QtCore.QSize(size, size))

    def _refresh_action_icons(self) -> None:
        mappings = [
            ("btn_login", "free-icon-login-2623062.png"),
            ("btn_download_matrix_template", "free-icon-download-126488.png"),
            ("btn_matrix_file", "upload.png"),
            ("btn_matrix_details", "information.png"),
            ("btn_matrix_export", "free-icon-download-126488.png"),
            ("btn_matrix_preview", "preview.png"),
            ("btn_matrix_apply", "krug_galka.png"),
            ("btn_download_structure_template", "free-icon-download-126488.png"),
            ("btn_upload_structure_template", "upload.png"),
            ("btn_create_project", "free-icon-plus-3303893.png"),
            ("btn_refresh_projects", "free-icon-refresh-5234214.png"),
            ("btn_download_user_template_top", "free-icon-download-126488.png"),
            ("btn_upload_user_file", "upload.png"),
            ("btn_run", "import.png"),
            ("btn_log", "information.png"),
        ]
        for attr, file_name in mappings:
            button = getattr(self, attr, None)
            if button is not None:
                if attr in {"btn_matrix_file", "btn_upload_structure_template", "btn_upload_user_file"} and bool(button.property("fileLoaded")):
                    self._set_button_icon(button, "krug_galka.png")
                else:
                    self._set_button_icon(button, file_name)
        for kind, state in getattr(self, "type_import_state", {}).items():
            buttons = ((state.get("download"), "free-icon-download-126488.png"),
                       (state.get("upload"), "upload.png"),
                       (state.get("preview"), "preview.png"),
                       (state.get("details"), "information.png"),
                       (state.get("apply"), "import.png"))
            for button, file_name in buttons:
                if button is not None:
                    try:
                        loaded = button is state.get("upload") and bool(button.property("fileLoaded"))
                        self._set_button_icon(button, "krug_galka.png" if loaded else file_name)
                    except RuntimeError:
                        # A stale card can briefly remain in state during style rebuild.
                        continue

        approval_state = getattr(self, "approval_state", {}) or {}
        approval_buttons = (
            (approval_state.get("download"), "free-icon-download-126488.png"),
            (approval_state.get("upload"), "upload.png"),
            (approval_state.get("preview"), "preview.png"),
            (approval_state.get("details"), "information.png"),
            (approval_state.get("apply"), "import.png"),
        )
        for button, file_name in approval_buttons:
            if button is not None:
                try:
                    loaded = button is approval_state.get("upload") and bool(button.property("fileLoaded"))
                    self._set_button_icon(button, "krug_galka.png" if loaded else file_name)
                except RuntimeError:
                    continue

    def _refresh_tab_icons(self) -> None:
        # Верхняя навигация намеренно без иконок: так три режима читаются
        # как самостоятельные разделы, а не как набор разнотипных действий.
        tabs = getattr(self, "main_tabs", None)
        if tabs is None:
            return
        for index in range(tabs.count()):
            tabs.setTabIcon(index, QtGui.QIcon())

    def _refresh_static_icons(self) -> None:
        for label in self.findChildren(QtWidgets.QLabel):
            asset = label.property("iconAsset")
            if not asset:
                continue
            try:
                size = int(label.property("iconSize") or 18)
            except Exception:
                size = 18
            icon = themed_icon(self.asset_dir, str(asset), self.is_dark_theme)
            label.setPixmap(icon.pixmap(QtCore.QSize(size, size)))

    def _make_status_chip(self, text: str, tone: str) -> QtWidgets.QLabel:
        chip = QtWidgets.QLabel(text)
        chip.setObjectName("statusChip")
        chip.setProperty("tone", tone)
        chip.setAlignment(QtCore.Qt.AlignCenter)
        return chip

    def _set_table_row_tone(self, table: QtWidgets.QTableWidget, row: int, tone: str) -> None:
        """Подсвечивает всю строку таблицы тем же цветом, что и статус в легенде."""
        colors = {
            "success": (STATUS_SUCCESS, 52 if self.is_dark_theme else 30),
            "pending": (ACCENT_HOVER, 56 if self.is_dark_theme else 32),
            "danger": (STATUS_DANGER, 58 if self.is_dark_theme else 34),
        }
        color_spec = colors.get(tone)
        if not color_spec:
            return
        color = QtGui.QColor(color_spec[0])
        color.setAlpha(color_spec[1])
        for column in range(table.columnCount()):
            item = table.item(row, column)
            if item is not None:
                item.setData(TableRowOverlayDelegate.COLOR_ROLE, color)

    def _make_badge(self, text: str, tone: str, size: int = 34) -> QtWidgets.QLabel:
        badge = QtWidgets.QLabel()
        badge.setObjectName("iconBadge")
        badge.setProperty("tone", tone)
        badge.setAlignment(QtCore.Qt.AlignCenter)
        badge.setFixedSize(size, size)
        candidate = Path(self.asset_dir) / str(text)
        if candidate.exists():
            icon = themed_icon(self.asset_dir, str(text), self.is_dark_theme)
            badge.setPixmap(icon.pixmap(QtCore.QSize(18, 18)))
            badge.setProperty("iconAsset", str(text))
            badge.setProperty("iconSize", 18)
        else:
            badge.setText(str(text))
        return badge

    def _add_card_header(self, layout: QtWidgets.QVBoxLayout, title: str,
                         icon_text: str, tone: str,
                         right_widget: Optional[QtWidgets.QWidget] = None) -> None:
        row = QtWidgets.QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)
        if icon_text:
            row.addWidget(self._make_badge(icon_text, tone, 32))
        label = QtWidgets.QLabel(title)
        label.setObjectName("cardTitle")
        row.addWidget(label)
        row.addStretch(1)
        if right_widget is not None:
            row.addWidget(right_widget)
        layout.addLayout(row)

    def _make_inline_icon(self, asset_name: str, size: int = 16) -> QtWidgets.QLabel:
        label = QtWidgets.QLabel()
        label.setObjectName("inlineIcon")
        label.setFixedSize(size + 4, size + 4)
        label.setAlignment(QtCore.Qt.AlignCenter)
        icon = themed_icon(self.asset_dir, asset_name, self.is_dark_theme)
        label.setPixmap(icon.pixmap(QtCore.QSize(size, size)))
        label.setProperty("iconAsset", asset_name)
        label.setProperty("iconSize", size)
        return label

    def _make_file_row(self, title: str, subtitle: str, icon_text: str,
                       tone: str, download_button: QtWidgets.QPushButton,
                       upload_button: QtWidgets.QPushButton,
                       is_last: bool = False) -> QtWidgets.QWidget:
        row_widget = QtWidgets.QWidget()
        row_widget.setObjectName("fileRowLast" if is_last else "fileRow")
        row = QtWidgets.QHBoxLayout(row_widget)
        row.setContentsMargins(0, 10, 0, 10)
        row.setSpacing(12)
        if icon_text:
            row.addWidget(self._make_badge(icon_text, tone, 32))

        text_col = QtWidgets.QVBoxLayout()
        text_col.setContentsMargins(0, 0, 0, 0)
        text_col.setSpacing(2)
        title_label = QtWidgets.QLabel(title)
        title_label.setObjectName("rowTitle")
        subtitle_label = QtWidgets.QLabel(subtitle)
        subtitle_label.setObjectName("rowSubtitle")
        text_col.addWidget(title_label)
        text_col.addWidget(subtitle_label)
        row.addLayout(text_col, 1)

        download_button.setObjectName("downloadButton")
        download_button.setFixedWidth(190)
        download_button.setFixedHeight(34)
        upload_button.setObjectName("uploadButton")
        upload_button.setProperty("fileLoaded", False)
        upload_button.setFixedWidth(190)
        upload_button.setFixedHeight(34)
        row.addWidget(download_button)
        row.addWidget(upload_button)
        return row_widget

    @staticmethod
    def _configure_interactive_table_columns(
        table: QtWidgets.QTableWidget,
        widths: List[int],
    ) -> None:
        """Задаёт стартовую ширину, сохраняя ручное изменение всех колонок."""
        header = table.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        header.setMinimumSectionSize(54)
        header.setStretchLastSection(True)
        table.setItemDelegate(TableRowOverlayDelegate(table))
        table.setShowGrid(False)
        table.setHorizontalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        for column, width in enumerate(widths[:table.columnCount()]):
            table.setColumnWidth(column, width)

    @staticmethod
    def _repolish(widget: QtWidgets.QWidget) -> None:
        widget.style().unpolish(widget)
        widget.style().polish(widget)
        widget.update()

    def _on_project_workspace_changed(self, display: str) -> None:
        if self._syncing_workspace_combos or not display:
            return
        if hasattr(self, "cb_workspace") and self.cb_workspace.currentText() != display:
            self.cb_workspace.setCurrentText(display)

    def _create_project_inline(self) -> None:
        title = self.ed_project_title.text().strip()
        if not title:
            self._show_alert("Ошибка", "Введите название проекта")
            self.ed_project_title.setFocus()
            return

        ws_display = self.cb_project_workspace.currentText()
        if ws_display not in self.workspace_map:
            self._show_alert("Ошибка", "Сначала войдите в систему и выберите пространство")
            return

        workspace_id = self.workspace_map[ws_display]
        description = self.ed_project_description.toPlainText().strip()

        # Галочка имеет прямой смысл: отмечена — создаём только проект,
        # не отмечена — создаём проект и папки из вручную загруженного Excel.
        create_without_folders = bool(
            getattr(self, "cb_create_without_folders", None)
            and self.cb_create_without_folders.isChecked()
        )
        use_structure = not create_without_folders

        if use_structure and not (
            self.default_structure_path and os.path.exists(self.default_structure_path)
        ):
            self._show_alert(
                "Не загружена структура",
                "Загрузите Excel-файл структуры проекта или отметьте "
                "«Создать проект без папок»."
            )
            return

        self.btn_create_project.setEnabled(False)
        self._project_worker = CreateProjectWorker(
            self.client,
            workspace_id,
            title,
            description,
            use_structure,
            self.default_structure_path if use_structure else "",
        )
        self._project_worker.log.connect(self._log)
        self._project_worker.finished_ok.connect(self._on_inline_project_created)
        self._project_worker.failed.connect(self._on_inline_project_failed)
        self._project_worker.finished.connect(self._update_project_action_state)
        self._project_worker.start()

    def _on_inline_project_created(self, data: Dict[str, Any]) -> None:
        self.ed_project_title.clear()
        self.ed_project_description.clear()
        self._log(
            f"✅ Проект создан: ID={data['project_id']}, '{data['title']}', "
            f"папок: {data['folders_created']}"
        )
        structure_line = (
            f"Папок создано: {data['folders_created']}"
            if data.get("structure_requested")
            else "Проект создан без папочной структуры"
        )
        self._show_success(
            "Проект создан",
            f"Название: {data['title']}\nID: {data['project_id']}\n{structure_line}"
        )
        QtCore.QTimer.singleShot(900, self._load_projects)

    def _on_inline_project_failed(self, error: str) -> None:
        self._log(f"❌ Не удалось создать проект: {error}")
        self._show_alert("Ошибка", f"Не удалось создать проект:\n{error}")

    def _set_project_mode(self, mode: str) -> None:
        """Переключает одну карточку между созданием нового и заполнением существующего проекта."""
        existing = mode == "existing"
        self._project_mode = "existing" if existing else "new"

        if hasattr(self, "project_mode_stack"):
            self.project_mode_stack.setCurrentIndex(1 if existing else 0)
        if hasattr(self, "btn_mode_new"):
            self.btn_mode_new.setChecked(not existing)
        if hasattr(self, "btn_mode_existing"):
            self.btn_mode_existing.setChecked(existing)
        if hasattr(self, "btn_create_project"):
            self.btn_create_project.setText(
                "Добавить структуру" if existing else "Создать проект"
            )

        self._update_create_project_tooltip()
        self._update_project_action_state()

    def _run_project_card_action(self) -> None:
        if getattr(self, "_project_mode", "new") == "existing":
            self._add_structure_to_existing_project()
        else:
            self._create_project_inline()

    def _update_project_action_state(self) -> None:
        if not hasattr(self, "btn_create_project"):
            return

        # Пока один из воркеров работает, не даём запустить вторую операцию.
        if ((self._project_worker is not None and self._project_worker.isRunning()) or
                (self._structure_worker is not None and self._structure_worker.isRunning())):
            self.btn_create_project.setEnabled(False)
            return

        if getattr(self, "_project_mode", "new") == "existing":
            enabled = bool(
                self.client.token
                and self.project_map
                and hasattr(self, "cb_existing_project")
                and self.cb_existing_project.currentText().strip()
            )
        else:
            enabled = bool(
                self.client.token
                and hasattr(self, "cb_project_workspace")
                and self.cb_project_workspace.currentText().strip()
            )
        self.btn_create_project.setEnabled(enabled)

    def _add_structure_to_existing_project(self) -> None:
        display = self.cb_existing_project.currentText().strip()
        project_id = self.project_map.get(display)

        if not self.client.token:
            self._show_alert("Ошибка", "Сначала войдите в систему")
            return
        if not project_id:
            self._show_alert("Ошибка", "Выберите существующий проект")
            return
        if not self.default_structure_path or not os.path.exists(self.default_structure_path):
            self._show_alert(
                "Не загружена структура",
                "Сначала загрузите Excel-файл структуры проекта в блоке «Excel-файлы»."
            )
            return

        project_title = display
        match = re.match(r"^\[\d+\]\s*(.*)$", display)
        if match:
            project_title = match.group(1).strip() or display

        # Текущий импорт создаёт папки по Excel. Он не удаляет существующие данные
        # проекта, но и не выполняет автоматическое объединение одноимённых папок.
        if not self._ask_alert(
            "Добавить структуру",
            f"Добавить папочную структуру из загруженного Excel в проект:\n"
            f"{display}\n\n"
            "Существующие файлы и папки удаляться не будут.\n"
            "Если в проекте уже есть папки с такими же именами, API может создать дубли.\n\n"
            "Продолжить?"
        ):
            return

        self.btn_create_project.setEnabled(False)
        self.cb_existing_project.setEnabled(False)
        self._structure_worker = AddStructureWorker(
            self.client,
            int(project_id),
            project_title,
            self.default_structure_path,
        )
        self._structure_worker.log.connect(self._log)
        self._structure_worker.finished_ok.connect(self._on_existing_structure_finished)
        self._structure_worker.failed.connect(self._on_existing_structure_failed)
        self._structure_worker.finished.connect(self._restore_existing_structure_controls)
        self._structure_worker.start()

    def _restore_existing_structure_controls(self) -> None:
        enabled = bool(self.client.token and self.project_map)
        self.cb_existing_project.setEnabled(enabled)
        self._update_project_action_state()

    def _on_existing_structure_finished(self, data: Dict[str, Any]) -> None:
        errors = data.get("errors") or []
        folders_created = int(data.get("folders_created", 0))
        self._log(
            f"✅ Структура добавлена в проект ID={data['project_id']}. "
            f"Создано папок: {folders_created}"
        )
        if errors:
            self._log(f"⚠️ Ошибок при создании веток: {len(errors)}")
        message = (
            f"Проект: [{data['project_id']}] {data['title']}\n"
            f"Создано папок: {folders_created}"
        )
        if errors:
            message += f"\nОшибок: {len(errors)} — подробности в логе."
        self._show_success("Структура добавлена", message)

    def _on_existing_structure_failed(self, error: str) -> None:
        self._log(f"❌ Не удалось добавить структуру: {error}")
        self._show_alert("Ошибка", f"Не удалось добавить структуру в проект:\n{error}")

    def _setup_ui(self):
        central = QtWidgets.QWidget()
        central.setObjectName("centralRoot")
        self.setCentralWidget(central)
        central_layout = QtWidgets.QVBoxLayout(central)
        central_layout.setContentsMargins(0, 0, 0, 0)
        central_layout.setSpacing(0)

        scroll = QtWidgets.QScrollArea()
        scroll.setObjectName("mainScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        central_layout.addWidget(scroll)

        page = QtWidgets.QWidget()
        page.setObjectName("page")
        scroll.setWidget(page)
        layout = QtWidgets.QVBoxLayout(page)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(10)

        # -------------------------------------------------------------
        # Header — простой и функциональный, без брендовых изображений
        # -------------------------------------------------------------
        header = QtWidgets.QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 4)
        header.setSpacing(12)
        title_col = QtWidgets.QVBoxLayout()
        title_col.setSpacing(2)
        title_label = QtWidgets.QLabel("Larix Platform")
        title_label.setObjectName("pageTitle")
        subtitle_label = QtWidgets.QLabel(
            "Проекты, папочная структура, пользователи, роли и маршруты согласований."
        )
        subtitle_label.setObjectName("pageSubtitle")
        title_col.addWidget(title_label)
        title_col.addWidget(subtitle_label)
        header.addLayout(title_col, 1)
        self.theme_toggle = ThemeToggle(self.asset_dir, self)
        self.theme_toggle.setToolTip("Светлая / тёмная тема")
        self.theme_toggle.toggled.connect(self._toggle_theme)
        self.back_to_manager_button = add_standard_header_controls(header, self, self.theme_toggle)
        layout.addLayout(header)

        # -------------------------------------------------------------
        # Shared authorization card
        # -------------------------------------------------------------
        auth_card = QtWidgets.QFrame()
        auth_card.setObjectName("card")
        auth_layout = QtWidgets.QVBoxLayout(auth_card)
        auth_layout.setContentsMargins(14, 11, 14, 13)
        auth_layout.setSpacing(8)
        self.auth_text = QtWidgets.QLabel("Не авторизован")
        self.auth_text.setObjectName("authStatus")
        self._add_card_header(auth_layout, "Подключение к Larix", "free-icon-login-2623062.png", "purple", self.auth_text)

        auth_grid = QtWidgets.QGridLayout()
        auth_grid.setHorizontalSpacing(12)
        auth_grid.setVerticalSpacing(6)
        auth_grid.setColumnStretch(0, 1)
        auth_grid.setColumnStretch(1, 1)
        auth_grid.setColumnMinimumWidth(2, 112)

        self.ed_login = QtWidgets.QLineEdit()
        self.ed_login.setPlaceholderText("Логин")
        self.ed_password = PasswordLineEdit()
        self.ed_password.setPlaceholderText("Пароль")
        self.ed_password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.ed_password.set_eye_clicked(self._toggle_password_visibility)
        self._update_password_eye_text()
        self.cb_workspace = NoWheelComboBox()
        self.cb_workspace.setPlaceholderText("Пространство")
        self.cb_workspace.setCurrentIndex(-1)
        self.cb_workspace.setEnabled(False)
        self.cb_workspace.currentTextChanged.connect(self._on_workspace_changed)
        self._setup_combo_popup(self.cb_workspace)
        self.btn_login = QtWidgets.QPushButton("Войти")
        self.btn_login.setObjectName("loginButton")
        self.btn_login.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_login.setFixedHeight(36)
        self.btn_login.setMinimumWidth(112)
        self.btn_login.clicked.connect(self._do_login)

        login_label = QtWidgets.QLabel("Логин"); login_label.setObjectName("fieldLabel")
        password_label = QtWidgets.QLabel("Пароль"); password_label.setObjectName("fieldLabel")
        workspace_label = QtWidgets.QLabel("Пространство"); workspace_label.setObjectName("fieldLabel")
        auth_grid.addWidget(login_label, 0, 0)
        auth_grid.addWidget(password_label, 0, 1)
        auth_grid.addWidget(self.ed_login, 1, 0)
        auth_grid.addWidget(self.ed_password, 1, 1)
        auth_grid.addWidget(self.btn_login, 1, 2)
        auth_grid.addWidget(workspace_label, 2, 0, 1, 3)
        auth_grid.addWidget(self.cb_workspace, 3, 0, 1, 3)
        auth_layout.addLayout(auth_grid)
        layout.addWidget(auth_card)

        # -------------------------------------------------------------
        # Main workspace tabs
        # -------------------------------------------------------------
        self.main_tabs = QtWidgets.QTabWidget()
        self.main_tabs.setTabBar(NoWheelTabBar())
        self.main_tabs.setObjectName("workspaceTabs")
        self.main_tabs.setDocumentMode(True)
        self.main_tabs.setMovable(False)
        self.main_tabs.tabBar().setDrawBase(False)
        self.main_tabs.tabBar().setExpanding(True)
        self.main_tabs.tabBar().setUsesScrollButtons(False)
        self.main_tabs.tabBar().setElideMode(QtCore.Qt.ElideNone)
        layout.addWidget(self.main_tabs, 1)

        # =============================================================
        # TAB 1 — Role matrix
        # =============================================================
        matrix_tab = QtWidgets.QWidget()
        matrix_tab.setObjectName("tabPage")
        matrix_tab_layout = QtWidgets.QVBoxLayout(matrix_tab)
        matrix_tab_layout.setContentsMargins(2, 12, 2, 2)
        matrix_tab_layout.setSpacing(12)
        matrix_tab_layout.setAlignment(QtCore.Qt.AlignTop)

        matrix_file_card = QtWidgets.QFrame()
        matrix_file_card.setObjectName("card")
        matrix_file_card.setMinimumHeight(115)
        matrix_file_layout = QtWidgets.QVBoxLayout(matrix_file_card)
        matrix_file_layout.setContentsMargins(14, 11, 14, 8)
        matrix_file_layout.setSpacing(0)
        self.btn_download_matrix_template = QtWidgets.QPushButton("Скачать шаблон")
        self.btn_download_matrix_template.clicked.connect(self._download_role_matrix_template)
        self.btn_matrix_file = QtWidgets.QPushButton("Загрузить файл")
        self.btn_matrix_file.clicked.connect(self._pick_role_matrix_file)
        matrix_file_layout.addWidget(self._make_file_row(
            "Excel-файл с правами",
            "Уровни папок + колонки проектных ролей или пользователей. Права задаются кодами из легенды ниже.",
            "Excel.png", "purple", self.btn_download_matrix_template, self.btn_matrix_file, True,
        ))
        rights_box = QtWidgets.QFrame(); rights_box.setObjectName("rightsLegendBox")
        rights_layout = QtWidgets.QHBoxLayout(rights_box); rights_layout.setContentsMargins(10, 7, 10, 7); rights_layout.setSpacing(8)
        rights_legend = QtWidgets.QLabel(
            "ПР — Просмотр    ·    С — Скачивание    ·    З — Загрузка/Создание    ·    "
            "Пер — Перемещение    ·    У — Удаление    ·    ПД — Полный доступ"
        )
        rights_legend.setObjectName("rightsLegend")
        rights_legend.setWordWrap(True)
        rights_layout.addWidget(rights_legend, 1)
        matrix_file_layout.addWidget(rights_box)
        matrix_tab_layout.addWidget(matrix_file_card)

        matrix_card = QtWidgets.QFrame()
        matrix_card.setObjectName("card")
        matrix_layout = QtWidgets.QVBoxLayout(matrix_card)
        matrix_layout.setContentsMargins(14, 12, 14, 14)
        matrix_layout.setSpacing(9)
        matrix_layout.addWidget(make_preview_header("Предпросмотр прав"))
        matrix_grid = QtWidgets.QGridLayout()
        matrix_grid.setHorizontalSpacing(12)
        matrix_grid.setVerticalSpacing(6)
        matrix_grid.setColumnStretch(0, 1)
        matrix_grid.setColumnStretch(1, 1)

        project_label = QtWidgets.QLabel("Проект"); project_label.setObjectName("fieldLabel")
        self.cb_matrix_project = NoWheelComboBox()
        self.cb_matrix_project.setPlaceholderText("Выберите проект")
        self.cb_matrix_project.setCurrentIndex(-1)
        self.cb_matrix_project.setEnabled(False)
        self.cb_matrix_project.currentIndexChanged.connect(self._invalidate_role_matrix_plan)
        self._setup_combo_popup(self.cb_matrix_project)
        sheet_label = QtWidgets.QLabel("Лист Excel"); sheet_label.setObjectName("fieldLabel")
        self.cb_matrix_sheet = NoWheelComboBox()
        self.cb_matrix_sheet.setPlaceholderText("Сначала загрузите Excel")
        self.cb_matrix_sheet.setCurrentIndex(-1)
        self.cb_matrix_sheet.setEnabled(False)
        self.cb_matrix_sheet.currentIndexChanged.connect(self._invalidate_role_matrix_plan)
        self._setup_combo_popup(self.cb_matrix_sheet)
        self.ed_matrix_file = QtWidgets.QLineEdit()
        self.ed_matrix_file.setReadOnly(True)
        self.ed_matrix_file.setPlaceholderText("Файл матрицы не загружен")
        file_label = QtWidgets.QLabel("Загруженный файл"); file_label.setObjectName("fieldLabel")
        matrix_grid.addWidget(project_label, 0, 0)
        matrix_grid.addWidget(sheet_label, 0, 1)
        matrix_grid.addWidget(self.cb_matrix_project, 1, 0)
        matrix_grid.addWidget(self.cb_matrix_sheet, 1, 1)
        matrix_grid.addWidget(file_label, 2, 0, 1, 2)
        matrix_grid.addWidget(self.ed_matrix_file, 3, 0, 1, 2)
        matrix_layout.addLayout(matrix_grid)

        matrix_hint = QtWidgets.QLabel(
            "Строгий режим: приложение назначает только то, что записано в Excel. "
            "Пусто = не менять, '-' = снять все права. Уже совпадающие права при сверке и повторном запуске пропускаются."
        )
        matrix_hint.setObjectName("infoText")
        matrix_hint.setWordWrap(True)
        hint_box = QtWidgets.QFrame(); hint_box.setObjectName("infoBanner")
        hint_l = QtWidgets.QHBoxLayout(hint_box); hint_l.setContentsMargins(10, 7, 10, 7); hint_l.setSpacing(8)
        hint_l.addWidget(self._make_inline_icon("information.png", 16), 0, QtCore.Qt.AlignTop)
        hint_l.addWidget(matrix_hint, 1)
        matrix_layout.addWidget(hint_box)

        status_row = QtWidgets.QHBoxLayout()
        status_row.setSpacing(8)
        status_title = QtWidgets.QLabel("Статусы сверки:")
        status_title.setObjectName("fieldLabel")
        status_row.addWidget(status_title)
        status_row.addWidget(self._make_status_chip("К применению", "pending"))
        status_row.addWidget(self._make_status_chip("Без изменений", "success"))
        status_row.addWidget(self._make_status_chip("Ошибка", "danger"))
        status_row.addStretch(1)
        matrix_layout.addLayout(status_row)

        self.tbl_matrix = QtWidgets.QTableWidget(0, 6)
        self.tbl_matrix.setObjectName("matrixTable")
        self.tbl_matrix.setHorizontalHeaderLabels([
            "Статус", "Строка", "Путь Excel", "Папка Larix", "Изменения", "Комментарий"
        ])
        self.tbl_matrix.verticalHeader().setVisible(False)
        self.tbl_matrix.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.tbl_matrix.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tbl_matrix.setAlternatingRowColors(True)
        self.tbl_matrix.setWordWrap(False)
        # На основной вкладке оставляем компактный предпросмотр. Полную
        # таблицу пользователь открывает кнопкой «Подробнее».
        self.tbl_matrix.setMinimumHeight(180)
        self.tbl_matrix.setMaximumHeight(230)
        self._configure_interactive_table_columns(
            self.tbl_matrix, [120, 80, 260, 260, 320, 320]
        )
        matrix_layout.addWidget(self.tbl_matrix, 1)

        matrix_footer = QtWidgets.QHBoxLayout()
        self.lbl_matrix_summary = QtWidgets.QLabel("Матрица не проверена")
        self.lbl_matrix_summary.setObjectName("rowSubtitle")
        self.btn_matrix_details = QtWidgets.QPushButton("Подробнее")
        self.btn_matrix_details.setObjectName("modeSwitch")
        self.btn_matrix_details.setToolTip("Открыть полную таблицу сверки в отдельном окне")
        self.btn_matrix_details.clicked.connect(self._show_role_matrix_details)
        self.btn_matrix_export = QtWidgets.QPushButton("Скачать")
        self.btn_matrix_export.setObjectName("downloadButton")
        self.btn_matrix_export.setToolTip("Скачать полный отчет сверки в Excel")
        self.btn_matrix_export.clicked.connect(self._export_role_matrix_report)
        self.btn_matrix_preview = QtWidgets.QPushButton("Сверить с Larix")
        self.btn_matrix_preview.setObjectName("loginButton")
        self.btn_matrix_preview.clicked.connect(self._preview_role_matrix)
        self.btn_matrix_apply = QtWidgets.QPushButton("Применить права")
        self.btn_matrix_apply.setObjectName("orangeAction")
        self.btn_matrix_apply.clicked.connect(self._apply_role_matrix)
        self.btn_matrix_details.setEnabled(False)
        self.btn_matrix_export.setEnabled(False)
        self.btn_matrix_preview.setEnabled(False)
        self.btn_matrix_apply.setEnabled(False)
        matrix_footer.addWidget(self.lbl_matrix_summary, 1)
        matrix_footer.addWidget(self.btn_matrix_details)
        matrix_footer.addWidget(self.btn_matrix_export)
        matrix_footer.addWidget(self.btn_matrix_preview)
        matrix_footer.addWidget(self.btn_matrix_apply)
        matrix_layout.addLayout(matrix_footer)
        matrix_tab_layout.addWidget(matrix_card)
        self.main_tabs.addTab(matrix_tab, "Ролевая матрица")

        # =============================================================
        # TAB 2 — Project and folders
        # =============================================================
        project_tab = QtWidgets.QWidget()
        project_tab.setObjectName("tabPage")
        project_tab_layout = QtWidgets.QVBoxLayout(project_tab)
        project_tab_layout.setContentsMargins(2, 12, 2, 2)
        project_tab_layout.setSpacing(12)

        structure_file_card = QtWidgets.QFrame()
        structure_file_card.setObjectName("card")
        structure_file_layout = QtWidgets.QVBoxLayout(structure_file_card)
        structure_file_layout.setContentsMargins(14, 11, 14, 8)
        structure_file_layout.setSpacing(0)
        self.btn_download_structure_template = QtWidgets.QPushButton("Скачать шаблон")
        self.btn_download_structure_template.clicked.connect(self._download_structure_template)
        self.btn_upload_structure_template = QtWidgets.QPushButton("Загрузить файл")
        self.btn_upload_structure_template.clicked.connect(self._upload_structure_template)
        structure_file_layout.addWidget(self._make_file_row(
            "Excel-файл структуры",
            "Используется при создании нового проекта с папками или при добавлении структуры в существующий проект.",
            "Excel.png", "orange", self.btn_download_structure_template, self.btn_upload_structure_template, True,
        ))
        project_tab_layout.addWidget(structure_file_card)

        create_card = QtWidgets.QFrame()
        create_card.setObjectName("card")
        create_layout = QtWidgets.QVBoxLayout(create_card)
        create_layout.setContentsMargins(14, 10, 14, 12)
        create_layout.setSpacing(8)

        mode_row = QtWidgets.QHBoxLayout(); mode_row.setSpacing(8)
        self.btn_mode_new = QtWidgets.QPushButton("Новый проект"); self.btn_mode_new.setObjectName("projectModeSwitch")
        self.btn_mode_new.setCheckable(True); self.btn_mode_new.clicked.connect(lambda: self._set_project_mode("new"))
        self.btn_mode_existing = QtWidgets.QPushButton("Добавить папки в существующий"); self.btn_mode_existing.setObjectName("projectModeSwitch")
        self.btn_mode_existing.setCheckable(True); self.btn_mode_existing.clicked.connect(lambda: self._set_project_mode("existing"))
        self.btn_mode_new.setMinimumHeight(40); self.btn_mode_existing.setMinimumHeight(40)
        mode_row.addWidget(self.btn_mode_new, 1); mode_row.addWidget(self.btn_mode_existing, 1)
        create_layout.addLayout(mode_row)

        self.project_mode_stack = QtWidgets.QStackedWidget()
        self.project_mode_stack.setObjectName("projectModeStack")
        self.project_mode_stack.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Maximum)
        new_project_page = QtWidgets.QWidget()
        new_project_page.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Maximum)
        new_grid = QtWidgets.QGridLayout(new_project_page)
        new_grid.setContentsMargins(0, 0, 0, 0)
        new_grid.setHorizontalSpacing(10)
        new_grid.setVerticalSpacing(6)
        new_grid.setColumnMinimumWidth(0, 105)
        new_grid.setColumnStretch(1, 1)
        self.ed_project_title = QtWidgets.QLineEdit(); self.ed_project_title.setPlaceholderText("Название проекта")
        self.cb_project_workspace = NoWheelComboBox(); self.cb_project_workspace.setPlaceholderText("Пространство")
        self.cb_project_workspace.setCurrentIndex(-1); self.cb_project_workspace.setEnabled(False)
        self.cb_project_workspace.currentTextChanged.connect(self._on_project_workspace_changed)
        self.cb_project_workspace.currentTextChanged.connect(lambda _text: self._update_project_action_state())
        self._setup_combo_popup(self.cb_project_workspace)
        self.ed_project_description = QtWidgets.QTextEdit(); self.ed_project_description.setPlaceholderText("Краткое описание проекта")
        self.ed_project_description.setFixedHeight(58)
        for txt, row in (("Название", 0), ("Пространство", 1), ("Описание", 2)):
            lab = QtWidgets.QLabel(txt); lab.setObjectName("fieldLabel")
            if row == 2: lab.setAlignment(QtCore.Qt.AlignTop | QtCore.Qt.AlignLeft)
            new_grid.addWidget(lab, row, 0)
        new_grid.addWidget(self.ed_project_title, 0, 1)
        new_grid.addWidget(self.cb_project_workspace, 1, 1)
        new_grid.addWidget(self.ed_project_description, 2, 1)
        self.cb_create_without_folders = QtWidgets.QCheckBox("Создать проект без папок")
        self.cb_create_without_folders.setObjectName("structureCheck")
        self.cb_create_without_folders.setCursor(QtCore.Qt.PointingHandCursor)
        self.cb_create_without_folders.setChecked(True)
        self.cb_create_without_folders.toggled.connect(lambda _checked: self._update_create_project_tooltip())
        new_grid.addWidget(self.cb_create_without_folders, 3, 1)
        self.project_mode_stack.addWidget(new_project_page)

        existing_page = QtWidgets.QWidget()
        existing_layout = QtWidgets.QVBoxLayout(existing_page)
        existing_layout.setContentsMargins(0, 0, 0, 0)
        existing_layout.setSpacing(8)
        existing_row = QtWidgets.QHBoxLayout(); existing_row.setSpacing(10)
        el = QtWidgets.QLabel("Проект"); el.setObjectName("fieldLabel"); el.setMinimumWidth(105)
        self.cb_existing_project = NoWheelComboBox(); self.cb_existing_project.setPlaceholderText("Выберите существующий проект")
        self.cb_existing_project.setCurrentIndex(-1); self.cb_existing_project.setEnabled(False)
        self.cb_existing_project.currentTextChanged.connect(lambda _text: self._update_project_action_state())
        self._setup_combo_popup(self.cb_existing_project)
        existing_row.addWidget(el); existing_row.addWidget(self.cb_existing_project, 1)
        existing_layout.addLayout(existing_row)
        existing_hint = QtWidgets.QLabel(
            "В выбранный проект будет добавлена структура из Excel, загруженного в этом разделе. Существующие данные не удаляются."
        )
        existing_hint.setObjectName("rowSubtitle"); existing_hint.setWordWrap(True)
        existing_layout.addWidget(existing_hint)
        self.project_mode_stack.addWidget(existing_page)
        create_layout.addWidget(self.project_mode_stack, 0)

        self.btn_create_project = QtWidgets.QPushButton("Создать проект")
        self.btn_create_project.setObjectName("orangeAction")
        self.btn_create_project.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_create_project.setFixedHeight(38); self.btn_create_project.setMinimumWidth(190)
        self.btn_create_project.setEnabled(False); self.btn_create_project.clicked.connect(self._run_project_card_action)
        create_layout.addWidget(self.btn_create_project, 0, QtCore.Qt.AlignHCenter)
        # Служебная кнопка сохранена для совместимости с логикой обновления списка проектов.
        self.btn_refresh_projects = QtWidgets.QPushButton()
        self.btn_refresh_projects.setVisible(False)
        self.btn_refresh_projects.setEnabled(False)
        self.btn_refresh_projects.clicked.connect(self._load_projects)
        project_tab_layout.addWidget(create_card, 0, QtCore.Qt.AlignTop)
        project_tab_layout.addStretch(1)
        self.main_tabs.addTab(project_tab, "Папки и проект")

        # =============================================================
        # TAB 3 — User import
        # =============================================================
        users_tab = QtWidgets.QWidget()
        users_tab.setObjectName("tabPage")
        users_layout = QtWidgets.QVBoxLayout(users_tab)
        users_layout.setContentsMargins(2, 12, 2, 2)
        users_layout.setSpacing(12)
        users_layout.setAlignment(QtCore.Qt.AlignTop)

        user_file_card = QtWidgets.QFrame(); user_file_card.setObjectName("card")
        user_file_layout = QtWidgets.QVBoxLayout(user_file_card)
        user_file_layout.setContentsMargins(14, 11, 14, 8); user_file_layout.setSpacing(0)
        self.btn_download_user_template_top = QtWidgets.QPushButton("Скачать шаблон")
        self.btn_download_user_template_top.clicked.connect(self._download_user_import_template)
        self.btn_upload_user_file = QtWidgets.QPushButton("Загрузить файл")
        self.btn_upload_user_file.clicked.connect(self._load_excel)
        user_file_layout.addWidget(self._make_file_row(
            "Excel-файл пользователей",
            "Таблица пользователей, проектов и проектных ролей.",
            "Excel.png", "green", self.btn_download_user_template_top, self.btn_upload_user_file, True,
        ))
        users_layout.addWidget(user_file_card)

        import_card = QtWidgets.QFrame(); import_card.setObjectName("card")
        import_layout = QtWidgets.QVBoxLayout(import_card)
        import_layout.setContentsMargins(14, 12, 14, 14); import_layout.setSpacing(10)
        sheet_row = QtWidgets.QHBoxLayout(); sheet_row.setSpacing(10)
        sheet_label = QtWidgets.QLabel("Лист Excel"); sheet_label.setObjectName("fieldLabel"); sheet_label.setMinimumWidth(90)
        self.cb_sheet = NoWheelComboBox(); self.cb_sheet.setPlaceholderText("Сначала загрузите Excel")
        self.cb_sheet.setEnabled(False); self.cb_sheet.currentTextChanged.connect(self._on_sheet_changed)
        self._setup_combo_popup(self.cb_sheet)
        sheet_row.addWidget(sheet_label); sheet_row.addWidget(self.cb_sheet, 1)
        import_layout.addLayout(sheet_row)
        info_banner = QtWidgets.QFrame(); info_banner.setObjectName("infoBanner")
        info_layout = QtWidgets.QHBoxLayout(info_banner); info_layout.setContentsMargins(10, 8, 10, 8); info_layout.setSpacing(8)
        info_text = QtWidgets.QLabel(
            "Пользователи импортируются из выбранного листа. Проекты и роли сопоставляются с текущим пространством Larix."
        )
        info_text.setObjectName("infoText"); info_text.setWordWrap(True)
        info_layout.addWidget(self._make_inline_icon("information.png", 16), 0, QtCore.Qt.AlignTop)
        info_layout.addWidget(info_text, 1)
        import_layout.addWidget(info_banner)

        import_status_row = QtWidgets.QHBoxLayout()
        import_status_row.setSpacing(8)
        import_status_title = QtWidgets.QLabel("Статусы импорта:")
        import_status_title.setObjectName("fieldLabel")
        import_status_row.addWidget(import_status_title)
        import_status_row.addWidget(self._make_status_chip("Успешно", "success"))
        import_status_row.addWidget(self._make_status_chip("Уже добавлен", "pending"))
        import_status_row.addWidget(self._make_status_chip("Ошибка", "danger"))
        import_status_row.addStretch(1)
        import_layout.addLayout(import_status_row)

        self.lbl_user_import_summary = QtWidgets.QLabel("Импорт ещё не запускался")
        self.lbl_user_import_summary.setObjectName("rowSubtitle")
        import_layout.addWidget(self.lbl_user_import_summary)
        import_layout.addSpacing(4)
        self.btn_run = QtWidgets.QPushButton("Импортировать пользователей")
        self.btn_run.setObjectName("greenAction"); self.btn_run.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_run.setFixedHeight(38); self.btn_run.setMinimumWidth(210); self.btn_run.setEnabled(False)
        self.btn_run.clicked.connect(self._start_add_users)
        import_layout.addWidget(self.btn_run, 0, QtCore.Qt.AlignHCenter)
        users_layout.addWidget(import_card)
        self.main_tabs.addTab(users_tab, "Импорт пользователей")
        self._setup_type_import_tab()
        self._setup_approval_tab()

        # -------------------------------------------------------------
        # Shared log
        # -------------------------------------------------------------
        self.log = QtWidgets.QPlainTextEdit(self)
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("Журнал операций...")
        self.log.hide()
        log_bar = QtWidgets.QFrame(); log_bar.setObjectName("logBar")
        log_layout = QtWidgets.QHBoxLayout(log_bar); log_layout.setContentsMargins(10, 0, 8, 0); log_layout.setSpacing(6)
        self.btn_log = QtWidgets.QPushButton("Открыть лог"); self.btn_log.setObjectName("logButton")
        self.btn_log.setCursor(QtCore.Qt.PointingHandCursor); self.btn_log.clicked.connect(self._show_log)
        log_layout.addWidget(self.btn_log); log_layout.addStretch(1)
        chevron = QtWidgets.QLabel("›"); chevron.setObjectName("chevron"); log_layout.addWidget(chevron)
        layout.addWidget(log_bar)

        self._set_project_mode("new")
        self._refresh_template_path_fields()
        self._set_auth_status(False)

    def _apply_styles(self):
        self.setStyleSheet(build_qss(self.is_dark_theme, self.asset_dir) + shared_style_overrides(self.is_dark_theme))
        app = QtWidgets.QApplication.instance()
        if app is not None:
            app.setProperty("larix_dark", bool(self.is_dark_theme))

        # QSS не управляет системной строкой заголовка Windows, поэтому
        # синхронизируем её отдельно через DWM. Два отложенных вызова нужны
        # для случаев, когда native HWND только что был создан/перерисован.
        self._set_native_titlebar_theme(self.is_dark_theme)
        QtCore.QTimer.singleShot(0, lambda: self._set_native_titlebar_theme(self.is_dark_theme))
        QtCore.QTimer.singleShot(120, lambda: self._set_native_titlebar_theme(self.is_dark_theme))

        self._update_theme_button()
        self._refresh_action_icons()
        self._refresh_tab_icons()
        self._refresh_static_icons()
        if hasattr(self, "ed_password"):
            self._update_password_eye_text()

    def _log(self, text: str):
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {text}"
        self.log.appendPlainText(line)
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())
        if self._log_view is not None:
            self._log_view.appendPlainText(line)
            self._log_view.verticalScrollBar().setValue(self._log_view.verticalScrollBar().maximum())
        try:
            self.statusBar().showMessage(str(text).strip().splitlines()[-1][:180], 8000)
        except Exception:
            pass

    def _update_create_project_tooltip(self) -> None:
        if not hasattr(self, "btn_create_project"):
            return
        if getattr(self, "_project_mode", "new") == "existing":
            self.btn_create_project.setToolTip(
                "Добавить папки из загруженного Excel-файла в выбранный существующий проект"
            )
            return
        without_folders = bool(
            getattr(self, "cb_create_without_folders", None)
            and self.cb_create_without_folders.isChecked()
        )
        self.btn_create_project.setToolTip(
            "Создать пустой проект без папок"
            if without_folders
            else "Создать проект с папками из вручную загруженного Excel-файла"
        )

    def _refresh_template_path_fields(self) -> None:
        # Загруженным считается только файл, который пользователь выбрал вручную.
        # Встроенный шаблон служит лишь источником для кнопки «Скачать шаблон».
        structure_loaded = bool(
            self.default_structure_path and os.path.exists(self.default_structure_path)
        )
        if hasattr(self, "btn_upload_structure_template"):
            self.btn_upload_structure_template.setText(
                "Файл загружен" if structure_loaded else "Загрузить файл"
            )
            self.btn_upload_structure_template.setProperty("fileLoaded", structure_loaded)
            self.btn_upload_structure_template.setToolTip(
                self.default_structure_path if structure_loaded
                else "Выбрать заполненный Excel-файл структуры проекта"
            )
            self._set_button_icon(
                self.btn_upload_structure_template,
                "krug_galka.png" if structure_loaded else "upload.png"
            )
            self._repolish(self.btn_upload_structure_template)

        if hasattr(self, "btn_download_structure_template"):
            self.btn_download_structure_template.setEnabled(True)
            self.btn_download_structure_template.setToolTip(
                "Скачать встроенный эталонный Excel-шаблон папочной структуры"
            )

        if hasattr(self, "cb_create_without_folders"):
            self.cb_create_without_folders.setEnabled(True)
            self.cb_create_without_folders.setToolTip(
                "Если отмечено, Excel-файл структуры не требуется"
            )

        self._update_create_project_tooltip()

        if hasattr(self, "btn_download_user_template_top"):
            self.btn_download_user_template_top.setEnabled(True)
            self.btn_download_user_template_top.setToolTip(
                "Скачать встроенный эталонный Excel-шаблон импорта пользователей"
            )

        user_file_loaded = bool(self.excel_path and os.path.exists(self.excel_path))
        if hasattr(self, "btn_upload_user_file"):
            self.btn_upload_user_file.setText(
                "Файл загружен" if user_file_loaded else "Загрузить файл"
            )
            self.btn_upload_user_file.setProperty("fileLoaded", user_file_loaded)
            if user_file_loaded:
                self.btn_upload_user_file.setToolTip(self.excel_path)
            self._set_button_icon(
                self.btn_upload_user_file,
                "krug_galka.png" if user_file_loaded else "upload.png"
            )
            self._repolish(self.btn_upload_user_file)

    def _save_template_paths(self) -> None:
        # Путь намеренно не сохраняется: при следующем запуске пользователь
        # должен снова вручную выбрать актуальный файл структуры.
        try:
            _save_template_settings("", "")
        except Exception as exc:
            self._log(f"⚠️ Не удалось очистить сохранённый путь структуры: {exc}")

    def _open_template_settings_dialog(self) -> None:
        dlg = TemplateSettingsDialog(self, self.default_structure_path)
        if dlg.exec() != QtWidgets.QDialog.Accepted:
            return
        self.default_structure_path = dlg.structure_path
        self._refresh_template_path_fields()
        self._save_template_paths()
        self._log("✅ Настройка шаблона структуры сохранена")
        if self.default_structure_path:
            self._log(f"   Структура проекта: {self.default_structure_path}")
        else:
            self._log("   Структура проекта: файл не выбран")

    def _validate_structure_template(self, path: str) -> bool:
        if not path or not os.path.exists(path):
            self._show_alert("Ошибка", "Файл структуры не найден")
            return False
        try:
            sheets = self._excel_sheet_names(path)
        except Exception as exc:
            self._show_alert("Ошибка", f"Не удалось прочитать Excel-файл структуры:\n{exc}")
            return False
        if STRUCTURE_SHEET_NAME not in sheets:
            self._show_alert(
                "Ошибка",
                f"В файле структуры должен быть лист '{STRUCTURE_SHEET_NAME}'.\n"
                f"Найдены листы: {', '.join(sheets) if sheets else 'нет листов'}"
            )
            return False
        return True

    def _upload_structure_template(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Загрузите Excel-файл с типовой структурой проекта",
            "", "Excel files (*.xlsx *.xls *.xlsm);;All files (*.*)"
        )
        if not path:
            return
        if not self._validate_structure_template(path):
            return
        self.default_structure_path = path
        self._refresh_template_path_fields()
        self._save_template_paths()
        self._log(f"✅ Загружен новый шаблон типовой структуры: {path}")

    def _download_structure_template(self):
        save_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Скачать шаблон типовой структуры",
            STRUCTURE_TEMPLATE_NAME, "Excel files (*.xlsx);;All files (*.*)"
        )
        if not save_path:
            return
        try:
            saved_to = _write_embedded_template(
                STRUCTURE_TEMPLATE_B64,
                STRUCTURE_TEMPLATE_SHA256,
                save_path,
                "structure",
            )
            self._show_success("Готово", "Встроенный шаблон типовой структуры сохранён")
            self._log(f"⬇️ Встроенный шаблон типовой структуры сохранён: {saved_to}")
        except Exception as exc:
            self._show_alert("Ошибка", f"Не удалось сохранить шаблон:\n{exc}")
            self._log(f"❌ Не удалось сохранить шаблон структуры: {exc}")

    def _download_role_matrix_template(self):
        save_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Скачать шаблон ролевой матрицы",
            ROLE_MATRIX_TEMPLATE_NAME, "Excel files (*.xlsx);;All files (*.*)"
        )
        if not save_path:
            return
        try:
            saved_to = _write_embedded_template(
                ROLE_MATRIX_TEMPLATE_B64,
                ROLE_MATRIX_TEMPLATE_SHA256,
                save_path,
                "role",
            )
            self._show_success("Готово", "Шаблон ролевой матрицы сохранён")
            self._log(f"⬇️ Шаблон ролевой матрицы сохранён: {saved_to}")
        except Exception as exc:
            self._show_alert("Ошибка", f"Не удалось сохранить шаблон:\n{exc}")
            self._log(f"❌ Не удалось сохранить шаблон ролевой матрицы: {exc}")

    def _download_user_import_template(self):
        save_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Скачать шаблон импорта пользователей",
            USER_IMPORT_TEMPLATE_NAME, "Excel files (*.xlsx);;All files (*.*)"
        )
        if not save_path:
            return
        try:
            saved_to = _write_embedded_template(
                USER_IMPORT_TEMPLATE_B64,
                USER_IMPORT_TEMPLATE_SHA256,
                save_path,
                "users",
            )
            self._show_success("Готово", "Встроенный шаблон импорта пользователей сохранён")
            self._log(f"⬇️ Встроенный шаблон импорта пользователей сохранён: {saved_to}")
        except Exception as exc:
            self._show_alert("Ошибка", f"Не удалось сохранить шаблон:\n{exc}")
            self._log(f"❌ Не удалось сохранить шаблон пользователей: {exc}")

    def _do_login(self):
        login = self.ed_login.text().strip()
        password = self.ed_password.text().strip()
        if not login or not password:
            self._show_alert("Ошибка", "Введите логин и пароль")
            return
        self.btn_login.setEnabled(False)
        self._log("Авторизация...")
        if self.client.login(login, password):
            self._log("✅ Успешно")
            self._set_auth_status(True)
            self._load_workspaces()
        else:
            self._log("❌ Ошибка авторизации")
            self._show_alert("Ошибка", "Неверный логин или пароль")
            self._set_auth_status(False)
        self.btn_login.setEnabled(True)

    def _load_workspaces(self):
        self._log("Загрузка пространств...")
        workspaces = self.client.list_workspaces()
        if not workspaces:
            self._log("Пространства не найдены")
            return

        self.workspace_map.clear()
        displays: List[str] = []
        for ws in workspaces:
            display = f"[{ws['id']}] {ws['name']}"
            self.workspace_map[display] = ws['id']
            displays.append(display)

        self._syncing_workspace_combos = True
        try:
            for combo in (self.cb_workspace, self.cb_project_workspace):
                combo.blockSignals(True)
                combo.clear()
                combo.addItems(displays)
                combo.setEnabled(bool(displays))
                combo.setCurrentIndex(0 if displays else -1)
                combo.blockSignals(False)
        finally:
            self._syncing_workspace_combos = False

        if displays:
            self._on_workspace_changed(displays[0])
        self._log(f"Загружено пространств: {len(workspaces)}")

    def _on_workspace_changed(self, display: str):
        if self._syncing_workspace_combos or display not in self.workspace_map:
            return

        if hasattr(self, "cb_project_workspace") and self.cb_project_workspace.currentText() != display:
            self._syncing_workspace_combos = True
            try:
                self.cb_project_workspace.blockSignals(True)
                self.cb_project_workspace.setCurrentText(display)
                self.cb_project_workspace.blockSignals(False)
            finally:
                self._syncing_workspace_combos = False

        ws_id = self.workspace_map[display]
        self._log(f"Смена пространства: {display}")
        if self.client.change_workspace(ws_id):
            self._log("✅ Пространство активно")
            self._load_projects()
            self._load_roles()
            self._update_project_action_state()
            self.btn_refresh_projects.setEnabled(True)
            if self.loaded_users:
                self.btn_run.setEnabled(True)
        else:
            self._log("❌ Не удалось сменить пространство")
            self.btn_create_project.setEnabled(False)
            self.btn_refresh_projects.setEnabled(False)
            self.btn_run.setEnabled(False)

    def _setup_type_import_tab(self) -> None:
        """Create independent remark/task tabs with their own card state."""
        for kind, title in (("remarks", "Типы замечаний"), ("tasks", "Типы задач")):
            tab = QtWidgets.QWidget()
            tab.setObjectName("tabPage")
            layout = QtWidgets.QVBoxLayout(tab)
            layout.setContentsMargins(2, 12, 2, 2)
            layout.setSpacing(12)
            layout.addWidget(self._build_type_import_card(kind, title))
            layout.addSpacing(4)
            self.main_tabs.addTab(tab, title)


    def _build_type_import_card(self, kind: str, title: str) -> QtWidgets.QFrame:
        page = QtWidgets.QWidget()
        page.setObjectName("typeImportPage")
        page_layout = QtWidgets.QVBoxLayout(page)
        # Внешняя вкладка уже задаёт те же поля, что и у ролевой матрицы.
        # Повторные внутренние поля делали карточки типов уже и ниже.
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(12)
        page_layout.setAlignment(QtCore.Qt.AlignTop)

        # Верхняя карточка повторяет файловый блок ролевой матрицы: в ней
        # остаются только назначение шаблона, легенда и действия с Excel.
        file_card = QtWidgets.QFrame()
        file_card.setObjectName("card")
        file_card.setMinimumHeight(115)
        file_layout = QtWidgets.QVBoxLayout(file_card)
        file_layout.setContentsMargins(14, 11, 14, 8)
        file_layout.setSpacing(0)
        grid = QtWidgets.QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(6)
        project = NoWheelComboBox()
        project.setPlaceholderText("Выберите проект")
        project.setCurrentIndex(-1)
        project.setEnabled(False)
        self._setup_combo_popup(project)
        sheet_combo = NoWheelComboBox()
        sheet_combo.setPlaceholderText("Выберите лист")
        sheet_combo.setCurrentIndex(-1)
        sheet_combo.setEnabled(False)
        self._setup_combo_popup(sheet_combo)
        file_edit = QtWidgets.QLineEdit()
        file_edit.setReadOnly(True)
        file_edit.setPlaceholderText("Excel-файл не выбран")
        choose = QtWidgets.QPushButton("Загрузить файл")
        choose.setObjectName("modeSwitch")
        choose.setProperty("fileLoaded", False)
        choose.clicked.connect(lambda _checked=False, k=kind: self._pick_type_import_file(k))
        download = QtWidgets.QPushButton("Скачать шаблон")
        download.setObjectName("downloadButton")
        download.clicked.connect(lambda _checked=False, k=kind: self._download_type_import_template(k))
        file_layout.addWidget(self._make_file_row("Excel-файл типов", "Загрузите шаблон и проверьте права перед импортом", "Excel.png", "info", download, choose, True))
        rights_box = QtWidgets.QFrame()
        rights_box.setObjectName("rightsLegendBox")
        rights_layout = QtWidgets.QHBoxLayout(rights_box)
        rights_layout.setContentsMargins(10, 7, 10, 7)
        rights_layout.setSpacing(8)
        rights_label = QtWidgets.QLabel(
            "ПР — Просмотр    ·    С — Создание    ·    "
            "ПРС — Просмотр и создание    ·    «-» — Нет прав"
        )
        rights_label.setObjectName("rightsLegend")
        rights_label.setWordWrap(True)
        rights_layout.addWidget(rights_label, 1)
        file_layout.addWidget(rights_box)
        page_layout.addWidget(file_card)

        # Нижняя карточка содержит состояние операции и её компактный план.
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        outer = QtWidgets.QVBoxLayout(card)
        outer.setContentsMargins(14, 12, 14, 14)
        outer.setSpacing(10)
        outer.addWidget(make_preview_header(title))
        info_banner = QtWidgets.QFrame()
        info_banner.setObjectName("infoBanner")
        info_layout = QtWidgets.QHBoxLayout(info_banner)
        info_layout.setContentsMargins(10, 7, 10, 7)
        info_label = QtWidgets.QLabel("Строгий режим: назначаются только права из Excel; остальные права просмотра и создания для перечисленных типов удаляются.")
        info_label.setObjectName("infoText")
        info_label.setWordWrap(True)
        info_layout.addWidget(self._make_inline_icon("information.png", 16), 0, QtCore.Qt.AlignTop)
        info_layout.addWidget(info_label, 1)
        project_label = QtWidgets.QLabel("Проект", objectName="fieldLabel")
        sheet_label = QtWidgets.QLabel("Лист Excel", objectName="fieldLabel")
        file_label = QtWidgets.QLabel("Загруженный файл", objectName="fieldLabel")
        grid.addWidget(project_label, 0, 0)
        grid.addWidget(sheet_label, 0, 1)
        grid.addWidget(project, 1, 0)
        grid.addWidget(sheet_combo, 1, 1)
        grid.addWidget(file_label, 2, 0, 1, 2)
        grid.addWidget(file_edit, 3, 0, 1, 2)
        grid.setColumnStretch(0, 1); grid.setColumnStretch(1, 1)
        outer.addLayout(grid)
        outer.addWidget(info_banner)
        status_row = QtWidgets.QHBoxLayout()
        status_row.setSpacing(8)
        status_row.addWidget(QtWidgets.QLabel("Статусы импорта:", objectName="fieldLabel"))
        status_row.addWidget(self._make_status_chip("Создать", "success"))
        status_row.addWidget(self._make_status_chip("Обновить", "pending"))
        status_row.addWidget(self._make_status_chip("Ошибка", "danger"))
        status_row.addStretch(1)
        outer.addLayout(status_row)
        summary = QtWidgets.QLabel("Загрузите файл и выберите проект для предпросмотра")
        summary.setObjectName("rowSubtitle")
        summary.setWordWrap(True)
        table = QtWidgets.QTableWidget(0, 4)
        table.setObjectName("matrixTable")
        table.setHorizontalHeaderLabels(["Тип", "Действие", "Просмотр", "Создание"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        table.setMinimumHeight(150)
        table.setMaximumHeight(220)
        self._configure_interactive_table_columns(table, [220, 120, 320, 320])
        outer.addWidget(table)
        footer = QtWidgets.QHBoxLayout()
        preview = QtWidgets.QPushButton("Предпросмотр")
        preview.setObjectName("loginButton")
        preview.clicked.connect(lambda _checked=False, k=kind: self._preview_type_import(k))
        apply_btn = QtWidgets.QPushButton("Импортировать")
        apply_btn.setObjectName("greenAction")
        apply_btn.clicked.connect(lambda _checked=False, k=kind: self._apply_type_import(k))
        details = QtWidgets.QPushButton("Подробнее")
        details.setObjectName("modeSwitch")
        details.clicked.connect(lambda _checked=False, k=kind: self._show_type_import_details(k))
        footer.addWidget(summary, 1); footer.addWidget(details); footer.addWidget(preview); footer.addWidget(apply_btn)
        outer.addLayout(footer)
        page_layout.addWidget(card)
        state = {"kind": kind, "project": project, "sheet": sheet_combo, "file": file_edit, "table": table, "summary": summary,
                 "download": download, "upload": choose, "preview": preview, "apply": apply_btn, "details": details, "status_widgets": status_row, "path": None, "plan": None,
                 "preview_worker": None, "apply_worker": None}
        project.currentIndexChanged.connect(lambda _index, k=kind: self._invalidate_type_import_plan(k))
        sheet_combo.currentIndexChanged.connect(lambda _index, k=kind: self._invalidate_type_import_plan(k))
        self.type_import_state[kind] = state
        self._type_import_state_changed(kind)
        return page

    def _type_import_state_changed(self, kind: str) -> None:
        state = self.type_import_state.get(kind)
        if not state:
            return
        busy = any(worker is not None and worker.isRunning() for worker in (state.get("preview_worker"), state.get("apply_worker")))
        ready = bool(self.client.token and state["project"].currentData() and state.get("path"))
        has_plan = bool(state.get("plan"))
        state["preview"].setEnabled(ready and not busy)
        state["apply"].setEnabled(has_plan and not busy)
        state["details"].setEnabled(has_plan and not busy)
        state["project"].setEnabled(bool(self.client.token and self.project_map) and not busy)

    def _invalidate_type_import_plan(self, kind: str) -> None:
        state = self.type_import_state.get(kind)
        if not state:
            return
        state["plan"] = None
        state["table"].setRowCount(0)
        state["summary"].setText("Параметры изменены. Запустите предпросмотр заново.")
        self._type_import_state_changed(kind)

    def _update_type_import_projects(self) -> None:
        for state in self.type_import_state.values():
            combo = state["project"]
            current = combo.currentText()
            combo.blockSignals(True); combo.clear()
            for name, project_id in self.project_map.items(): combo.addItem(name, project_id)
            if current in self.project_map: combo.setCurrentText(current)
            elif combo.count(): combo.setCurrentIndex(0)
            else: combo.setCurrentIndex(-1)
            combo.blockSignals(False)
            self._type_import_state_changed(state["kind"])

    def _pick_type_import_file(self, kind: str) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Выберите Excel-файл", "", "Excel (*.xlsx *.xls *.xlsm)")
        if not path: return
        state = self.type_import_state[kind]
        state["path"], state["plan"] = path, None
        state["file"].setText(path); state["summary"].setText("Файл выбран. Нажмите «Предпросмотр» для проверки.")
        try:
            sheets = list(pd.ExcelFile(path).sheet_names)
        except Exception as exc:
            sheets = []
            state["summary"].setText(f"Не удалось прочитать листы Excel: {exc}")
        state["sheet"].blockSignals(True)
        state["sheet"].clear(); state["sheet"].addItems(sheets)
        state["sheet"].setEnabled(bool(sheets)); state["sheet"].setCurrentIndex(0 if sheets else -1)
        state["sheet"].blockSignals(False)
        state["upload"].setProperty("fileLoaded", True)
        state["upload"].setToolTip("Файл выбран: нажмите, чтобы заменить его")
        self._refresh_action_icons()
        state["table"].setRowCount(0); self._type_import_state_changed(kind)

    def _download_type_import_template(self, kind: str) -> None:
        try:
            path = _ensure_builtin_remark_type_template() if kind == "remarks" else _ensure_builtin_task_type_template()
            name = REMARK_TYPE_TEMPLATE_NAME if kind == "remarks" else TASK_TYPE_TEMPLATE_NAME
            target, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Сохранить шаблон", name, "Excel (*.xlsx)")
            if target:
                saved_to = _write_embedded_template(
                    REMARK_TYPE_TEMPLATE_B64 if kind == "remarks" else TASK_TYPE_TEMPLATE_B64,
                    REMARK_TYPE_TEMPLATE_SHA256 if kind == "remarks" else TASK_TYPE_TEMPLATE_SHA256,
                    target,
                    kind,
                )
                self._show_success("Готово", f"Шаблон сохранён: {saved_to}")
                self._log(f"⬇️ Встроенный шаблон {kind} сохранён: {saved_to}")
        except Exception as exc:
            self._show_alert("Ошибка", f"Не удалось сохранить шаблон:\n{exc}")
            self._log(f"❌ Не удалось сохранить шаблон {kind}: {exc}")

    def _preview_type_import(self, kind: str) -> None:
        state = self.type_import_state[kind]
        project_id = state["project"].currentData()
        if not project_id or not state.get("path"):
            self._show_alert("Предпросмотр", "Выберите проект и Excel-файл."); return
        state["plan"] = None; state["preview"].setEnabled(False)
        worker = TypeImportPreviewWorker(self.client, int(project_id), state["path"], kind,
                                         state["sheet"].currentText() or None)
        state["preview_worker"] = worker
        worker.log.connect(self._log); worker.progress.connect(lambda value, k=kind: self._type_import_progress(k, value))
        worker.finished_ok.connect(lambda data, k=kind: self._on_type_import_preview(k, data))
        worker.failed.connect(lambda error, k=kind: self._on_type_import_failed(k, error))
        worker.finished.connect(lambda k=kind: self._type_import_worker_finished(k, "preview")); worker.start()

    def _on_type_import_preview(self, kind: str, data: Dict[str, Any]) -> None:
        state = self.type_import_state[kind]; state["plan"] = data.get("plan", [])
        table = state["table"]; table.setRowCount(0)
        for row in state["plan"]:
            values = [row.get("name", ""), "Создать" if row.get("action") == "create" else "Обновить",
                      f"Роли: {len(row.get('role_view_ids', []))}; пользователи: {len(row.get('user_view_ids', []))}",
                      f"Роли: {len(row.get('role_create_ids', []))}; пользователи: {len(row.get('user_create_ids', []))}"]
            index = table.rowCount(); table.insertRow(index)
            for col, value in enumerate(values): table.setItem(index, col, QtWidgets.QTableWidgetItem(str(value)))
            tone = "danger" if row.get("errors") else ("success" if row.get("action") == "create" else "pending")
            self._set_table_row_tone(table, index, tone)
        state["summary"].setText(f"Готово к импорту: {len(state['plan'])} типов. Проверьте сводку и нажмите «Импортировать».")
        self._type_import_state_changed(kind)

    def _apply_type_import(self, kind: str) -> None:
        state = self.type_import_state[kind]
        if not state.get("plan"): return
        if not self._ask_alert("Подтверждение импорта", f"Применить изменения для {len(state['plan'])} типов?\nПрава будут полностью заменены."): return
        project_id = state["project"].currentData(); state["apply"].setEnabled(False)
        worker = TypeImportApplyWorker(self.client, int(project_id), kind, state["plan"]); state["apply_worker"] = worker
        worker.log.connect(self._log); worker.progress.connect(lambda value, k=kind: self._type_import_progress(k, value))
        worker.finished_ok.connect(lambda data, k=kind: self._on_type_import_applied(k, data))
        worker.failed.connect(lambda error, k=kind: self._on_type_import_failed(k, error))
        worker.finished.connect(lambda k=kind: self._type_import_worker_finished(k, "apply")); worker.start()

    def _on_type_import_applied(self, kind: str, data: Dict[str, Any]) -> None:
        stats = data.get("stats", {}); state = self.type_import_state[kind]
        state["summary"].setText(f"Импорт завершён: успешно {stats.get('success', 0)}, ошибок {stats.get('errors', 0)}.")
        self._show_success("Импорт типов", state["summary"].text())

    def _show_type_import_details(self, kind: str) -> None:
        state = self.type_import_state[kind]; plan = state.get("plan") or []
        if not plan:
            self._show_alert("Детали предпросмотра", "Нет данных")
            return
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Детали предпросмотра типов")
        dlg.resize(1080, 620); dlg.setMinimumSize(780, 460); dlg.setStyleSheet(self.styleSheet())
        layout = QtWidgets.QVBoxLayout(dlg); layout.setContentsMargins(16, 16, 16, 14); layout.setSpacing(10)
        title = QtWidgets.QLabel(f"Типы {('замечаний' if kind == 'remarks' else 'задач')}")
        title.setObjectName("pageTitle"); layout.addWidget(title)
        summary = QtWidgets.QLabel(f"План содержит {len(plan)} типов. Права будут заменены по данным предпросмотра.")
        summary.setObjectName("rowSubtitle"); summary.setWordWrap(True); layout.addWidget(summary)
        table = QtWidgets.QTableWidget(len(plan), 7); table.setObjectName("matrixTable")
        table.setHorizontalHeaderLabels(["Тип", "Действие", "Просмотр: роли", "Просмотр: пользователи", "Создание: роли", "Создание: пользователи", "Ошибки"])
        table.horizontalHeader().setDefaultAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        table.verticalHeader().setVisible(False); table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows); table.setWordWrap(True)
        display_plan = []
        for source_row in plan:
            row = dict(source_row)
            for names_key, ids_key in (("role_view_names", "role_view_ids"), ("user_view_names", "user_view_ids"), ("role_create_names", "role_create_ids"), ("user_create_names", "user_create_ids")):
                if row.get(names_key):
                    row[ids_key] = row[names_key]
            display_plan.append(row)
        for row_index, row in enumerate(display_plan):
            values = [row.get("name", ""), "Создать" if row.get("action") == "create" else "Обновить",
                      ", ".join(map(str, row.get("role_view_ids", []))) or "—",
                      ", ".join(map(str, row.get("user_view_ids", []))) or "—",
                      ", ".join(map(str, row.get("role_create_ids", []))) or "—",
                      ", ".join(map(str, row.get("user_create_ids", []))) or "—",
                      "; ".join(map(str, row.get("errors", []))) or "—"]
            for col, value in enumerate(values): table.setItem(row_index, col, QtWidgets.QTableWidgetItem(str(value)))
            tone = "danger" if row.get("errors") else ("success" if row.get("action") == "create" else "pending")
            self._set_table_row_tone(table, row_index, tone)
        self._configure_interactive_table_columns(
            table, [220, 120, 320, 280, 320, 280, 300]
        )
        table.resizeRowsToContents(); layout.addWidget(table, 1)
        close = QtWidgets.QPushButton("Закрыть"); close.setObjectName("modeSwitch"); close.clicked.connect(dlg.accept)
        layout.addWidget(close, 0, QtCore.Qt.AlignRight)
        dlg.exec()

    def _on_type_import_failed(self, kind: str, error: str) -> None:
        self._log(f"❌ Импорт типов ({kind}): {error}"); self._show_alert("Ошибка импорта типов", error)

    def _type_import_worker_finished(self, kind: str, mode: str) -> None:
        self.type_import_state[kind][f"{mode}_worker"] = None; self._type_import_state_changed(kind)

    def _type_import_progress(self, kind: str, value: int) -> None:
        self.type_import_state[kind]["summary"].setText(f"Выполняется операция: {value}%")


    def _setup_approval_tab(self) -> None:
        tab = QtWidgets.QWidget()
        tab.setObjectName("tabPage")
        layout = QtWidgets.QVBoxLayout(tab)
        layout.setContentsMargins(2, 12, 2, 2)
        layout.setSpacing(12)
        layout.setAlignment(QtCore.Qt.AlignTop)

        file_card = QtWidgets.QFrame()
        file_card.setObjectName("card")
        file_card.setMinimumHeight(115)
        file_layout = QtWidgets.QVBoxLayout(file_card)
        file_layout.setContentsMargins(14, 11, 14, 8)
        file_layout.setSpacing(0)

        download = QtWidgets.QPushButton("Скачать шаблон")
        download.setObjectName("downloadButton")
        download.clicked.connect(self._download_approval_template)
        upload = QtWidgets.QPushButton("Загрузить файл")
        upload.setObjectName("modeSwitch")
        upload.setProperty("fileLoaded", False)
        upload.clicked.connect(self._pick_approval_file)
        file_layout.addWidget(self._make_file_row(
            "Excel-файл маршрутов",
            "Маршруты, этапы, согласующие, длительности и права из трёх связанных листов шаблона.",
            "Excel.png", "info", download, upload, True,
        ))

        legend_box = QtWidgets.QFrame()
        legend_box.setObjectName("rightsLegendBox")
        legend_layout = QtWidgets.QHBoxLayout(legend_box)
        legend_layout.setContentsMargins(10, 7, 10, 7)
        legend_layout.setSpacing(8)
        legend = QtWidgets.QLabel(
            "5. Маршруты согласований — состав этапов    ·    "
            "5.1 — доступ и длительность    ·    "
            "5.2 — отмена процесса и обязательность проверки"
        )
        legend.setObjectName("rightsLegend")
        legend.setWordWrap(True)
        legend_layout.addWidget(legend, 1)
        file_layout.addWidget(legend_box)

        sheet_grid = QtWidgets.QGridLayout()
        sheet_grid.setHorizontalSpacing(10)
        sheet_grid.setVerticalSpacing(6)
        main_sheet = NoWheelComboBox()
        settings_sheet = NoWheelComboBox()
        flags_sheet = NoWheelComboBox()
        for combo in (main_sheet, settings_sheet, flags_sheet):
            combo.setPlaceholderText("Выберите лист")
            combo.setCurrentIndex(-1)
            combo.setEnabled(False)
            self._setup_combo_popup(combo)
        sheet_grid.addWidget(QtWidgets.QLabel("Маршруты / этапы", objectName="fieldLabel"), 0, 0)
        sheet_grid.addWidget(main_sheet, 0, 1)
        sheet_grid.addWidget(QtWidgets.QLabel("Доступ / длительность", objectName="fieldLabel"), 1, 0)
        sheet_grid.addWidget(settings_sheet, 1, 1)
        sheet_grid.addWidget(QtWidgets.QLabel("Настройки согласующих", objectName="fieldLabel"), 2, 0)
        sheet_grid.addWidget(flags_sheet, 2, 1)
        sheet_grid.setColumnStretch(1, 1)
        file_layout.addLayout(sheet_grid)
        layout.addWidget(file_card)

        card = QtWidgets.QFrame()
        card.setObjectName("card")
        outer = QtWidgets.QVBoxLayout(card)
        outer.setContentsMargins(14, 12, 14, 14)
        outer.setSpacing(10)

        outer.addWidget(make_preview_header("Маршруты согласований"))

        grid = QtWidgets.QGridLayout()
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(6)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)

        project_label = QtWidgets.QLabel("Проект")
        project_label.setObjectName("fieldLabel")
        file_label = QtWidgets.QLabel("Загруженный файл")
        file_label.setObjectName("fieldLabel")

        project = NoWheelComboBox()
        project.setPlaceholderText("Выберите проект")
        project.setCurrentIndex(-1)
        project.setEnabled(False)
        self._setup_combo_popup(project)

        file_edit = QtWidgets.QLineEdit()
        file_edit.setReadOnly(True)
        file_edit.setPlaceholderText("Excel-файл не выбран")

        grid.addWidget(project_label, 0, 0, 1, 2)
        grid.addWidget(project, 1, 0, 1, 2)
        grid.addWidget(file_label, 2, 0, 1, 2)
        grid.addWidget(file_edit, 3, 0, 1, 2)
        outer.addLayout(grid)

        info_banner = QtWidgets.QFrame()
        info_banner.setObjectName("infoBanner")
        info_layout = QtWidgets.QHBoxLayout(info_banner)
        info_layout.setContentsMargins(10, 7, 10, 7)
        info_layout.setSpacing(8)
        info = QtWidgets.QLabel(
            "Безопасный режим: существующие маршруты не изменяются и не удаляются. "
            "Создаются только новые маршруты. Если пользователь, роль или данные этапа не найдены, "
            "создание блокируется до исправления Excel."
        )
        info.setObjectName("infoText")
        info.setWordWrap(True)
        info_layout.addWidget(self._make_inline_icon("information.png", 16), 0, QtCore.Qt.AlignTop)
        info_layout.addWidget(info, 1)
        outer.addWidget(info_banner)

        status_row = QtWidgets.QHBoxLayout()
        status_row.setSpacing(8)
        status_row.addWidget(QtWidgets.QLabel("Статусы:", objectName="fieldLabel"))
        status_row.addWidget(self._make_status_chip("Создать", "success"))
        status_row.addWidget(self._make_status_chip("Уже существует", "pending"))
        status_row.addWidget(self._make_status_chip("Ошибка", "danger"))
        status_row.addStretch(1)
        outer.addLayout(status_row)

        summary = QtWidgets.QLabel("Загрузите Excel и выберите проект для предпросмотра")
        summary.setObjectName("rowSubtitle")
        summary.setWordWrap(True)
        outer.addWidget(summary)

        table = QtWidgets.QTableWidget(0, 5)
        table.setObjectName("matrixTable")
        table.setHorizontalHeaderLabels(["Маршрут", "Действие", "Этапов", "Согласующих", "Доступ"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        table.setMinimumHeight(160)
        table.setMaximumHeight(250)
        self._configure_interactive_table_columns(table, [260, 150, 90, 120, 360])
        outer.addWidget(table)

        footer = QtWidgets.QHBoxLayout()
        footer.setSpacing(8)
        details = QtWidgets.QPushButton("Подробнее")
        details.setObjectName("modeSwitch")
        details.clicked.connect(self._show_approval_details)
        preview = QtWidgets.QPushButton("Предпросмотр")
        preview.setObjectName("loginButton")
        preview.clicked.connect(self._preview_approvals)
        apply_btn = QtWidgets.QPushButton("Создать маршруты")
        apply_btn.setObjectName("greenAction")
        apply_btn.clicked.connect(self._apply_approvals)
        footer.addStretch(1)
        footer.addWidget(details)
        footer.addWidget(preview)
        footer.addWidget(apply_btn)
        outer.addLayout(footer)

        layout.addWidget(card)
        self.main_tabs.addTab(tab, "Маршруты согласований")

        self.approval_state = {
            "project": project,
            "file": file_edit,
            "table": table,
            "summary": summary,
            "download": download,
            "upload": upload,
            "main_sheet": main_sheet,
            "settings_sheet": settings_sheet,
            "flags_sheet": flags_sheet,
            "preview": preview,
            "apply": apply_btn,
            "details": details,
            "path": None,
            "plan": None,
            "preview_worker": None,
            "apply_worker": None,
        }
        project.currentIndexChanged.connect(self._invalidate_approval_plan)
        for combo in (main_sheet, settings_sheet, flags_sheet):
            combo.currentIndexChanged.connect(self._invalidate_approval_plan)
        self._approval_state_changed()

    def _approval_state_changed(self) -> None:
        state = self.approval_state
        if not state:
            return
        busy = any(
            worker is not None and worker.isRunning()
            for worker in (state.get("preview_worker"), state.get("apply_worker"))
        )
        sheets_ready = all(
            state.get(key) is not None and state[key].currentText().strip()
            for key in ("main_sheet", "settings_sheet", "flags_sheet")
        )
        ready = bool(self.client.token and state["project"].currentData() and state.get("path") and sheets_ready)
        plan = state.get("plan") or []
        has_errors = any(row.get("action") == "error" or row.get("errors") for row in plan)
        has_creates = any(row.get("action") == "create" for row in plan)
        state["preview"].setEnabled(ready and not busy)
        state["apply"].setEnabled(bool(plan) and has_creates and not has_errors and not busy)
        state["details"].setEnabled(bool(plan) and not busy)
        state["project"].setEnabled(bool(self.client.token and self.project_map) and not busy)
        for key in ("main_sheet", "settings_sheet", "flags_sheet"):
            combo = state.get(key)
            if combo is not None:
                combo.setEnabled(bool(state.get("path")) and combo.count() > 0 and not busy)

    def _invalidate_approval_plan(self, *_args) -> None:
        state = self.approval_state
        if not state:
            return
        state["plan"] = None
        state["table"].setRowCount(0)
        state["summary"].setText("Параметры изменены. Запустите предпросмотр заново.")
        self._approval_state_changed()

    def _update_approval_projects(self) -> None:
        state = self.approval_state
        if not state:
            return
        combo = state["project"]
        current = combo.currentText()
        combo.blockSignals(True)
        combo.clear()
        for name, project_id in self.project_map.items():
            combo.addItem(name, project_id)
        if current in self.project_map:
            combo.setCurrentText(current)
        elif combo.count():
            combo.setCurrentIndex(0)
        else:
            combo.setCurrentIndex(-1)
        combo.blockSignals(False)
        self._approval_state_changed()

    def _populate_approval_sheet_choices(self, path: str) -> None:
        state = self.approval_state
        sheets = [str(item) for item in pd.ExcelFile(path).sheet_names]
        expected = (APPROVAL_MAIN_SHEET, APPROVAL_SETTINGS_SHEET, APPROVAL_FLAGS_SHEET)
        fallback_indices = (0, 1, 2)
        for key, expected_name, fallback_index in zip(
            ("main_sheet", "settings_sheet", "flags_sheet"), expected, fallback_indices
        ):
            combo = state[key]
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(sheets)
            target = _approval_find_sheet(sheets, expected_name)
            if not target and 0 <= fallback_index < len(sheets):
                target = sheets[fallback_index]
            if target:
                combo.setCurrentText(target)
            else:
                combo.setCurrentIndex(-1)
            combo.setEnabled(bool(sheets))
            combo.blockSignals(False)

    def _pick_approval_file(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Выберите Excel-файл маршрутов согласований",
            "",
            "Excel (*.xlsx *.xlsm);;All files (*.*)",
        )
        if not path:
            return
        state = self.approval_state
        state["path"] = path
        state["plan"] = None
        state["file"].setText(path)
        try:
            self._populate_approval_sheet_choices(path)
        except Exception as exc:
            state["path"] = None
            state["file"].clear()
            self._show_alert("Ошибка", f"Не удалось прочитать листы Excel:\n{exc}")
            self._approval_state_changed()
            return
        state["table"].setRowCount(0)
        state["summary"].setText("Файл выбран. Проверьте листы и нажмите «Предпросмотр».")
        state["upload"].setProperty("fileLoaded", True)
        state["upload"].setToolTip("Файл выбран: нажмите, чтобы заменить его")
        self._refresh_action_icons()
        self._approval_state_changed()

    def _download_approval_template(self) -> None:
        try:
            path = _ensure_builtin_approval_template()
            target, _ = QtWidgets.QFileDialog.getSaveFileName(
                self, "Сохранить шаблон", APPROVAL_TEMPLATE_NAME, "Excel (*.xlsx)"
            )
            if target:
                saved_to = _write_embedded_template(
                    APPROVAL_TEMPLATE_B64, APPROVAL_TEMPLATE_SHA256, target, "approval"
                )
                self._show_success("Готово", f"Шаблон сохранён: {saved_to}")
                self._log(f"⬇️ Встроенный шаблон маршрутов сохранён: {saved_to}")
        except Exception as exc:
            self._show_alert("Ошибка", f"Не удалось сохранить шаблон:\n{exc}")
            self._log(f"❌ Не удалось сохранить шаблон маршрутов: {exc}")

    def _preview_approvals(self) -> None:
        state = self.approval_state
        project_id = state["project"].currentData()
        if not project_id or not state.get("path"):
            self._show_alert("Предпросмотр", "Выберите проект и Excel-файл.")
            return
        selected_sheets = (
            state["main_sheet"].currentText().strip(),
            state["settings_sheet"].currentText().strip(),
            state["flags_sheet"].currentText().strip(),
        )
        if not all(selected_sheets):
            self._show_alert("Предпросмотр", "Выберите все три листа Excel.")
            return
        if len(set(selected_sheets)) < 3:
            self._show_alert("Предпросмотр", "Для трёх блоков маршрута выберите три разных листа Excel.")
            return
        state["plan"] = None
        state["table"].setRowCount(0)
        state["summary"].setText("Проверяю маршруты...")
        worker = ApprovalPreviewWorker(
            self.client, int(project_id), state["path"],
            main_sheet_name=selected_sheets[0],
            settings_sheet_name=selected_sheets[1],
            flags_sheet_name=selected_sheets[2],
        )
        state["preview_worker"] = worker
        worker.log.connect(self._log)
        worker.progress.connect(self._approval_progress)
        worker.finished_ok.connect(self._on_approval_preview)
        worker.failed.connect(self._on_approval_failed)
        worker.finished.connect(lambda: self._approval_worker_finished("preview"))
        self._approval_state_changed()
        worker.start()

    def _on_approval_preview(self, data: Dict[str, Any]) -> None:
        state = self.approval_state
        state["plan"] = data.get("plan", [])
        table = state["table"]
        table.setRowCount(0)
        action_labels = {"create": "Создать", "skip": "Уже существует", "error": "Ошибка"}
        for row in state["plan"]:
            index = table.rowCount()
            table.insertRow(index)
            approvers = sum(len(step.get("users", [])) for step in row.get("steps", []))
            access_text = (
                f"Просмотр: {len(row.get('view_role_ids', [])) + len(row.get('view_user_ids', []))}; "
                f"создание: {len(row.get('create_role_ids', [])) + len(row.get('create_user_ids', []))}"
            )
            values = [
                row.get("title", ""),
                action_labels.get(row.get("action"), row.get("action", "")),
                len(row.get("steps", [])),
                approvers,
                access_text if not row.get("errors") else "; ".join(row.get("errors", [])),
            ]
            for col, value in enumerate(values):
                item = QtWidgets.QTableWidgetItem(str(value))
                if col == 4:
                    item.setToolTip(str(value))
                table.setItem(index, col, item)
            tone = "danger" if row.get("action") == "error" else ("success" if row.get("action") == "create" else "pending")
            self._set_table_row_tone(table, index, tone)

        stats = data.get("stats", {})
        state["summary"].setText(
            f"Предпросмотр: создать {stats.get('create', 0)}, "
            f"уже существуют {stats.get('skip', 0)}, ошибок {stats.get('errors', 0)}."
        )
        self._approval_state_changed()

    def _apply_approvals(self) -> None:
        state = self.approval_state
        plan = state.get("plan") or []
        create_count = sum(row.get("action") == "create" for row in plan)
        if not create_count:
            return
        if any(row.get("action") == "error" or row.get("errors") for row in plan):
            self._show_alert("Маршруты согласований", "В предпросмотре есть ошибки. Исправьте Excel и повторите проверку.")
            return
        if not self._ask_alert(
            "Создание маршрутов",
            f"Создать {create_count} новых маршрутов согласований?\n"
            "Существующие маршруты будут пропущены без изменений.",
        ):
            return
        project_id = state["project"].currentData()
        worker = ApprovalApplyWorker(self.client, int(project_id), plan)
        state["apply_worker"] = worker
        state["summary"].setText("Создаю маршруты...")
        worker.log.connect(self._log)
        worker.progress.connect(self._approval_progress)
        worker.finished_ok.connect(self._on_approval_applied)
        worker.failed.connect(self._on_approval_failed)
        worker.finished.connect(lambda: self._approval_worker_finished("apply"))
        self._approval_state_changed()
        worker.start()

    def _on_approval_applied(self, data: Dict[str, Any]) -> None:
        state = self.approval_state
        stats = data.get("stats", {})
        state["summary"].setText(
            f"Готово: создано {stats.get('success', 0)}, "
            f"пропущено {stats.get('skipped', 0)}, ошибок {stats.get('errors', 0)}."
        )
        state["plan"] = None
        if stats.get("errors", 0):
            self._show_alert("Маршруты согласований", state["summary"].text() + "\nПодробности сохранены в журнале.")
        else:
            self._show_success("Маршруты согласований", state["summary"].text())
        self._approval_state_changed()

    def _show_approval_details(self) -> None:
        state = self.approval_state
        plan = state.get("plan") or []
        if not plan:
            self._show_alert("Детали предпросмотра", "Нет данных")
            return

        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Детали маршрутов согласований")
        dlg.resize(1080, 620)
        dlg.setMinimumSize(780, 460)
        dlg.setStyleSheet(self.styleSheet())
        layout = QtWidgets.QVBoxLayout(dlg)
        layout.setContentsMargins(16, 16, 16, 14)
        layout.setSpacing(10)

        preview_title = make_preview_header("Предпросмотр маршрутов согласований")
        preview_title.title_label.setObjectName("pageTitle")
        layout.addWidget(preview_title)
        note = QtWidgets.QLabel(
            "Существующие маршруты не изменяются. Имена согласующих показаны так, как они разрешились в выбранном проекте Larix."
        )
        note.setObjectName("rowSubtitle")
        note.setWordWrap(True)
        layout.addWidget(note)

        table = QtWidgets.QTableWidget(len(plan), 8)
        table.setObjectName("matrixTable")
        table.setHorizontalHeaderLabels([
            "Маршрут", "Действие", "Этапы", "Согласующие", "Просмотр",
            "Создание", "Флаги", "Ошибки",
        ])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        table.setWordWrap(True)
        action_labels = {"create": "Создать", "skip": "Уже существует", "error": "Ошибка"}

        for row_index, row in enumerate(plan):
            step_text = "\n".join(
                f"{step.get('title')} — {step.get('duration')} р.д."
                for step in row.get("steps", [])
            ) or "—"
            approver_text = "\n".join(
                f"{step.get('title')}: " + ", ".join(user.get("user_name") or user.get("source_name", "") for user in step.get("users", []))
                for step in row.get("steps", [])
            ) or "—"
            flags_text = "\n".join(
                f"{step.get('title')}: " + ", ".join(
                    f"{user.get('user_name') or user.get('source_name')} "
                    f"[отмена={'Да' if user.get('can_cancel') else 'Нет'}, "
                    f"обяз.={'Да' if user.get('must_approve') else 'Нет'}]"
                    for user in step.get("users", [])
                )
                for step in row.get("steps", [])
            ) or "—"
            values = [
                row.get("title", ""),
                action_labels.get(row.get("action"), row.get("action", "")),
                step_text,
                approver_text,
                ", ".join(row.get("view_names", [])) or "—",
                ", ".join(row.get("create_names", [])) or "—",
                flags_text,
                "; ".join(row.get("errors", [])) or "—",
            ]
            for col, value in enumerate(values):
                table.setItem(row_index, col, QtWidgets.QTableWidgetItem(str(value)))
            tone = "danger" if row.get("action") == "error" else ("success" if row.get("action") == "create" else "pending")
            self._set_table_row_tone(table, row_index, tone)

        self._configure_interactive_table_columns(table, [210, 130, 180, 260, 160, 160, 340, 360])
        table.resizeRowsToContents()
        layout.addWidget(table, 1)
        close = QtWidgets.QPushButton("Закрыть")
        close.setObjectName("modeSwitch")
        close.clicked.connect(dlg.accept)
        layout.addWidget(close, 0, QtCore.Qt.AlignRight)
        dlg.exec()

    def _on_approval_failed(self, error: str) -> None:
        message = error.strip()
        self._log(f"❌ Маршруты согласований:\n{message}")
        if "ApprovalImportValidationError:" in message:
            user_message = message.split("ApprovalImportValidationError:", 1)[1].strip()
        else:
            user_message = message.splitlines()[-1] if message.splitlines() else message
        self._show_alert("Ошибка маршрутов согласований", user_message)

    def _approval_worker_finished(self, mode: str) -> None:
        if self.approval_state:
            self.approval_state[f"{mode}_worker"] = None
            self._approval_state_changed()

    def _approval_progress(self, value: int) -> None:
        if self.approval_state:
            self.approval_state["summary"].setText(f"Выполняется операция: {value}%")

    def _load_projects(self):
        self._log("🔄 Загрузка проектов...")
        projects = self.client.list_projects()
        self.project_map.clear()
        self.project_name_map.clear()
        self._log(f"📊 Получено проектов от API: {len(projects)}")
        for project in projects:
            name = str(project.get("name", "")).strip()
            project_id = int(project["id"])
            display = f"[{project_id}] {name}"
            self.project_map[display] = project_id
            self.project_name_map[self._normalize_project_name(name)] = project_id
            self.project_name_map[str(project_id)] = project_id
        self._log(f"✅ Загружено проектов: {len(self.project_map)}")

        if hasattr(self, "cb_existing_project"):
            current = self.cb_existing_project.currentText()
            self.cb_existing_project.blockSignals(True)
            self.cb_existing_project.clear()
            self.cb_existing_project.addItems(list(self.project_map.keys()))
            if current in self.project_map:
                self.cb_existing_project.setCurrentText(current)
            elif self.cb_existing_project.count() > 0:
                self.cb_existing_project.setCurrentIndex(0)
            else:
                self.cb_existing_project.setCurrentIndex(-1)
            self.cb_existing_project.setEnabled(bool(self.client.token and self.project_map))
            self.cb_existing_project.blockSignals(False)

        if hasattr(self, "cb_matrix_project"):
            current_matrix = self.cb_matrix_project.currentText()
            self.cb_matrix_project.blockSignals(True)
            self.cb_matrix_project.clear()
            self.cb_matrix_project.addItems(list(self.project_map.keys()))
            if current_matrix in self.project_map:
                self.cb_matrix_project.setCurrentText(current_matrix)
            elif self.cb_matrix_project.count() > 0:
                self.cb_matrix_project.setCurrentIndex(0)
            else:
                self.cb_matrix_project.setCurrentIndex(-1)
            self.cb_matrix_project.setEnabled(bool(self.client.token and self.project_map))
            self.cb_matrix_project.blockSignals(False)

        self._update_project_action_state()
        self._update_role_matrix_controls()
        self._update_type_import_projects()
        self._update_approval_projects()

        if len(self.project_map) <= 20:
            for display in self.project_map.keys():
                self._log(f"   • {display}")

    def _load_roles(self):
        self._log("🔄 Загрузка ролей через POST /api/projectRoles/list...")
        roles = self.client.list_roles()
        self.role_map.clear()
        self._log(f"📋 Получено ролей от API: {len(roles)}")
        for role in roles:
            title = str(role.get("title", "")).strip()
            role_item = {"id": int(role["id"]), "title": title}
            normalized = self._normalize_role_name(title)
            self.role_map[normalized] = role_item
            self._log(f"   • [{role['id']}] '{title}' → '{normalized}'")
            norm_lower = normalized.lower()
            if "руководитель" in norm_lower and "проект" in norm_lower:
                self.role_map["руководитель проекта"] = role_item
                self.role_map["рук проекта"] = role_item
                self.role_map["руководительпроекта"] = role_item
            if "специалист" in norm_lower and "пто" in norm_lower:
                self.role_map["специалист пто"] = role_item
                self.role_map["спец пто"] = role_item
                self.role_map["специалистпто"] = role_item
            if "администратор" in norm_lower and "проект" in norm_lower:
                self.role_map["администратор проекта"] = role_item
                self.role_map["админ проекта"] = role_item
        self._log(f"✅ Загружено ролей (с алиасами): {len(self.role_map)}")

    def _load_excel(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Выберите Excel-файл", "",
            "Excel files (*.xlsx *.xls *.xlsm);;All files (*.*)"
        )
        if not path:
            return
        try:
            sheets = self._excel_sheet_names(path)
            active_sheet = self._active_excel_sheet_name(path)
            if active_sheet not in sheets and sheets:
                active_sheet = sheets[0]
            users, sheet_name = self._read_users_from_excel(path, active_sheet)
        except Exception as exc:
            self._show_alert("Ошибка", f"Не удалось прочитать Excel:\n{exc}")
            self._log(f"Ошибка чтения Excel: {exc}")
            return
        self.excel_path = path
        if hasattr(self, "btn_upload_user_file"):
            self.btn_upload_user_file.setText("Файл загружен")
            self.btn_upload_user_file.setProperty("fileLoaded", True)
            self._set_button_icon(self.btn_upload_user_file, "krug_galka.png")
            self.btn_upload_user_file.setToolTip(path)
            self._repolish(self.btn_upload_user_file)
        self.cb_sheet.blockSignals(True)
        self.cb_sheet.clear()
        self.cb_sheet.addItems(sheets)
        if sheet_name:
            self.cb_sheet.setCurrentText(sheet_name)
        self.cb_sheet.setEnabled(bool(sheets))
        self.cb_sheet.blockSignals(False)
        if not users:
            self.loaded_users = []
            self.btn_run.setEnabled(False)
            self._show_alert("Ошибка", f"На листе '{sheet_name}' не найдены строки с email")
            self._log(f"На листе '{sheet_name}' не найдены строки с email")
            return
        self.loaded_users = users
        self.btn_run.setEnabled(bool(self.client.token and self.cb_workspace.currentText()))
        self._log(f"✅ Excel загружен: лист '{sheet_name}', пользователей: {len(users)}")

    def _on_sheet_changed(self, sheet_name: str) -> None:
        if not self.excel_path or not sheet_name:
            return
        try:
            users, loaded_sheet = self._read_users_from_excel(self.excel_path, sheet_name)
        except Exception as exc:
            self._show_alert("Ошибка", f"Не удалось прочитать лист Excel:\n{exc}")
            self._log(f"Ошибка чтения листа '{sheet_name}': {exc}")
            return
        if not users:
            self.loaded_users = []
            self.btn_run.setEnabled(False)
            self._show_alert("Ошибка", f"На листе '{sheet_name}' не найдены строки с email")
            return
        self.loaded_users = users
        self.btn_run.setEnabled(bool(self.client.token and self.cb_workspace.currentText()))
        self._log(f"Выбран лист '{loaded_sheet}', пользователей: {len(users)}")

    def _read_users_from_excel(self, path: str, sheet_name: Optional[str] = None) -> tuple[List[Dict[str, Any]], str]:
        if sheet_name is None:
            sheet_name = self._active_excel_sheet_name(path)
        df = pd.read_excel(path, sheet_name=sheet_name, dtype=str).fillna("")
        return self._extract_users_from_dataframe(df), str(sheet_name)

    def _excel_sheet_names(self, path: str) -> List[str]:
        book = pd.ExcelFile(path)
        return [str(name) for name in book.sheet_names]

    def _active_excel_sheet_name(self, path: str) -> str:
        book = pd.ExcelFile(path)
        first_sheet = book.sheet_names[0] if book.sheet_names else 0
        if str(path).lower().endswith((".xlsx", ".xlsm")):
            try:
                from openpyxl import load_workbook
                workbook = load_workbook(path, read_only=True, data_only=True)
                return workbook.active.title
            except Exception:
                return first_sheet
        return first_sheet

    def _extract_users_from_dataframe(self, df) -> List[Dict[str, Any]]:
        if df.empty:
            return []
        headers = [str(col).strip().lower() for col in df.columns]
        email_col = self._find_header_index(headers, ["почта", "email", "e-mail", "mail"])
        surname_col = self._find_header_index(headers, ["фамилия", "surname", "last"])
        name_col = self._find_header_index(headers, ["имя", "name", "first"])
        project_col = self._find_header_index(headers, ["проект", "проекты", "project", "projects"])
        role_col = self._find_header_index(headers, ["роль", "роли", "role", "roles"])
        admin_col = self._find_header_index(headers, ["администратор проекта", "администратор", "admin", "админ"])
        if email_col is None:
            for idx, col in enumerate(df.columns):
                if df[col].astype(str).str.contains("@", regex=False, na=False).any():
                    email_col = idx
                    break
        users: List[Dict[str, Any]] = []
        for _, row in df.iterrows():
            values = [str(value).strip() for value in row.tolist()]
            if not any(values):
                continue
            if email_col is None or email_col >= len(values):
                continue
            email = values[email_col].strip().lower()
            if "@" not in email:
                continue
            surname = values[surname_col].strip() if surname_col is not None and surname_col < len(values) else ""
            first_name = values[name_col].strip() if name_col is not None and name_col < len(values) else ""
            projects = self._split_project_cell(values[project_col]) if project_col is not None and project_col < len(values) else []
            roles = self._split_project_cell(values[role_col]) if role_col is not None and role_col < len(values) else []
            is_admin = False
            if admin_col is not None and admin_col < len(values):
                admin_value = values[admin_col].strip().lower()
                is_admin = admin_value in ["да", "yes", "true", "1", "y"]
            users.append({
                "email": email, "name": first_name, "surname": surname,
                "projects": projects, "roles": roles, "is_project_admin": is_admin
            })
        return users

    @staticmethod
    def _split_project_cell(value: str) -> List[str]:
        return [part.strip() for part in str(value or "").split(";") if part.strip()]

    @staticmethod
    def _normalize_project_name(value: str) -> str:
        return " ".join(str(value or "").strip().lower().split())

    @staticmethod
    def _normalize_role_name(value: str) -> str:
        if not value:
            return ""
        value = re.sub(r'\([^)]*\)', '', value)
        return " ".join(str(value).strip().lower().split())

    @staticmethod
    def _find_header_index(headers: List[str], keywords: List[str]) -> Optional[int]:
        for idx, header in enumerate(headers):
            if any(keyword in header for keyword in keywords):
                return idx
        return None

    def _all_project_ids(self) -> List[int]:
        return sorted({int(pid) for pid in self.project_map.values() if pid is not None})

    def _resolve_user_project_ids(self, users: List[Dict[str, Any]]) -> tuple[List[Dict[str, Any]], List[str], List[str], List[str]]:
        missing_projects: List[str] = []
        missing_roles: List[str] = []
        roles_without_projects: List[str] = []
        resolved_users: List[Dict[str, Any]] = []
        all_project_ids = self._all_project_ids()
        for user in users:
            resolved = dict(user)
            project_ids: List[int] = []
            for raw_name in user.get("projects", []):
                normalized = self._normalize_project_name(raw_name)
                if normalized in {"все", "все проекты", "all", "*"}:
                    project_ids.extend(all_project_ids)
                    continue
                project_id = self.project_name_map.get(normalized)
                if project_id is None:
                    missing_projects.append(str(raw_name))
                    continue
                project_ids.append(project_id)
            resolved["project_ids"] = sorted(set(project_ids))
            role_items: List[Dict[str, Any]] = []
            for raw_role in user.get("roles", []):
                normalized_role = self._normalize_role_name(raw_role)
                role = self.role_map.get(normalized_role)
                if role is None:
                    role = self.role_map.get(raw_role.strip().lower())
                if role is None:
                    missing_roles.append(str(raw_role))
                    continue
                role_items.append(role)
            if user.get("is_project_admin", False):
                admin_role = (
                    self.role_map.get("администратор проекта")
                    or self.role_map.get(self._normalize_role_name("Администратор проекта"))
                    or self.role_map.get("админ проекта")
                )
                if admin_role:
                    role_items.append(admin_role)
            resolved["role_items"] = sorted({int(r["id"]): r for r in role_items}.values(), key=lambda r: int(r["id"]))
            if (resolved["role_items"] or resolved.get("is_project_admin", False)) and not resolved["project_ids"]:
                roles_without_projects.append(str(user.get("email", "")))
            resolved_users.append(resolved)
        return resolved_users, sorted(set(missing_projects)), sorted(set(missing_roles)), sorted(set(roles_without_projects))

    def _invalidate_role_matrix_plan(self, *_args) -> None:
        self.role_matrix_plan = []
        self.role_matrix_last_report = []
        if hasattr(self, "tbl_matrix"):
            self.tbl_matrix.setRowCount(0)
        if hasattr(self, "lbl_matrix_summary"):
            self.lbl_matrix_summary.setText("Матрица не проверена")
        if hasattr(self, "btn_matrix_apply"):
            self.btn_matrix_apply.setEnabled(False)
        self._update_role_matrix_controls()

    def _update_role_matrix_controls(self) -> None:
        if not hasattr(self, "btn_matrix_preview"):
            return
        project_id = self._matrix_project_id()
        ready = bool(
            self.client.token
            and project_id
            and self.role_matrix_path
            and os.path.exists(self.role_matrix_path)
            and self.cb_matrix_sheet.currentText().strip()
        )
        busy = bool(
            (self._matrix_preview_worker and self._matrix_preview_worker.isRunning())
            or (self._matrix_apply_worker and self._matrix_apply_worker.isRunning())
        )
        self.btn_matrix_preview.setEnabled(ready and not busy)
        self.btn_matrix_apply.setEnabled(bool(self.role_matrix_plan) and ready and not busy)
        if hasattr(self, "btn_matrix_details"):
            self.btn_matrix_details.setEnabled(bool(self.role_matrix_last_report) and not busy)
        if hasattr(self, "btn_matrix_export"):
            self.btn_matrix_export.setEnabled(bool(self.role_matrix_last_report) and not busy)
        self.btn_matrix_file.setEnabled(not busy)
        self.cb_matrix_project.setEnabled(bool(self.client.token and self.project_map) and not busy)
        self.cb_matrix_sheet.setEnabled(bool(self.role_matrix_path) and not busy)

    def _matrix_project_id(self) -> Optional[int]:
        if not hasattr(self, "cb_matrix_project"):
            return None
        return self.project_map.get(self.cb_matrix_project.currentText())

    def _pick_role_matrix_file(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Выберите Excel-файл ролевой матрицы", "",
            "Excel files (*.xlsx *.xls *.xlsm);;All files (*.*)"
        )
        if not path:
            return
        try:
            sheets = [str(item) for item in pd.ExcelFile(path).sheet_names]
        except Exception as exc:
            self._show_alert("Ошибка", f"Не удалось прочитать Excel:\n{exc}")
            return
        self.role_matrix_path = path
        self.ed_matrix_file.setText(path)
        if hasattr(self, "btn_matrix_file"):
            self.btn_matrix_file.setText("Файл загружен")
            self.btn_matrix_file.setProperty("fileLoaded", True)
            self.btn_matrix_file.setToolTip(path)
            self._set_button_icon(self.btn_matrix_file, "krug_galka.png")
            self._repolish(self.btn_matrix_file)
        self.cb_matrix_sheet.blockSignals(True)
        self.cb_matrix_sheet.clear()
        self.cb_matrix_sheet.addItems(sheets)
        if "2. Ролевая матрица" in sheets:
            self.cb_matrix_sheet.setCurrentText("2. Ролевая матрица")
        elif sheets:
            self.cb_matrix_sheet.setCurrentIndex(0)
        self.cb_matrix_sheet.blockSignals(False)
        self._log(f"✅ Ролевая матрица загружена: {os.path.basename(path)}")
        self._invalidate_role_matrix_plan()

    def _fill_role_matrix_table(
        self,
        table: QtWidgets.QTableWidget,
        rows: List[Dict[str, Any]],
        wrap: bool = False,
    ) -> None:
        """Заполняет таблицу сверки. Используется и в основном окне, и в «Подробнее»."""
        table.setSortingEnabled(False)
        table.setRowCount(len(rows))
        table.setWordWrap(bool(wrap))
        for row_idx, row in enumerate(rows):
            status = row.get("status")
            tone = {
                "Ошибка": "danger",
                "К применению": "pending",
                "Без изменений": "success",
            }.get(status, "")
            changes = " | ".join(row.get("changes") or [])
            values = [
                status or "",
                str(row.get("row_number", "")),
                row.get("path_text", ""),
                row.get("folder_text", ""),
                changes,
                row.get("message", ""),
            ]
            for col_idx, value in enumerate(values):
                item = QtWidgets.QTableWidgetItem(str(value or ""))
                item.setToolTip(str(value or ""))
                if col_idx == 0:
                    if status == "Ошибка":
                        item.setForeground(QtGui.QBrush(QtGui.QColor("#D94A4A")))
                    elif status == "К применению":
                        item.setForeground(QtGui.QBrush(QtGui.QColor("#FF7A00")))
                    elif status == "Без изменений":
                        item.setForeground(QtGui.QBrush(QtGui.QColor("#3A9B66")))
                table.setItem(row_idx, col_idx, item)
            self._set_table_row_tone(table, row_idx, tone)
        if wrap:
            table.resizeRowsToContents()
        table.setSortingEnabled(True)

    def _show_role_matrix_details(self) -> None:
        rows = list(self.role_matrix_last_report or [])
        if not rows:
            self._show_alert("Ролевая матрица", "Сначала выполните сверку с Larix")
            return

        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Подробная сверка ролевой матрицы")
        dlg.resize(1080, 620)
        dlg.setMinimumSize(780, 460)
        dlg.setStyleSheet(self.styleSheet())
        layout = QtWidgets.QVBoxLayout(dlg)
        layout.setContentsMargins(16, 16, 16, 14)
        layout.setSpacing(10)

        top = QtWidgets.QHBoxLayout()
        title = QtWidgets.QLabel("Полная сверка данных")
        title.setObjectName("pageTitle")
        top.addWidget(title)
        top.addStretch(1)
        filter_label = QtWidgets.QLabel("Показать:")
        filter_label.setObjectName("fieldLabel")
        top.addWidget(filter_label)
        status_filter = NoWheelComboBox()
        status_filter.addItems(["Все", "Ошибки", "К применению", "Без изменений"])
        status_filter.setMinimumWidth(170)
        self._setup_combo_popup(status_filter)
        top.addWidget(status_filter)
        layout.addLayout(top)

        summary = QtWidgets.QLabel(
            f"Всего строк: {len(rows)} · "
            f"Изменить: {sum(1 for r in rows if r.get('status') == 'К применению')} · "
            f"Без изменений: {sum(1 for r in rows if r.get('status') == 'Без изменений')} · "
            f"Ошибок: {sum(1 for r in rows if r.get('status') == 'Ошибка')}"
        )
        summary.setObjectName("rowSubtitle")
        layout.addWidget(summary)

        table = QtWidgets.QTableWidget(0, 6)
        table.setObjectName("matrixTable")
        table.setHorizontalHeaderLabels([
            "Статус", "Строка", "Путь Excel", "Папка Larix", "Изменения", "Комментарий"
        ])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        table.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        table.setAlternatingRowColors(True)
        table.setSortingEnabled(True)
        self._configure_interactive_table_columns(
            table, [120, 80, 260, 260, 340, 340]
        )
        layout.addWidget(table, 1)

        def filtered_rows() -> List[Dict[str, Any]]:
            mode = status_filter.currentText()
            if mode == "Ошибки":
                return [r for r in rows if r.get("status") == "Ошибка"]
            if mode == "К применению":
                return [r for r in rows if r.get("status") == "К применению"]
            if mode == "Без изменений":
                return [r for r in rows if r.get("status") == "Без изменений"]
            return rows

        def refresh_table(*_args) -> None:
            current = filtered_rows()
            self._fill_role_matrix_table(table, current, wrap=True)
            summary.setText(
                f"Показано: {len(current)} из {len(rows)} · "
                f"Изменить: {sum(1 for r in rows if r.get('status') == 'К применению')} · "
                f"Без изменений: {sum(1 for r in rows if r.get('status') == 'Без изменений')} · "
                f"Ошибок: {sum(1 for r in rows if r.get('status') == 'Ошибка')}"
            )

        status_filter.currentTextChanged.connect(refresh_table)
        refresh_table()

        buttons = QtWidgets.QHBoxLayout()
        buttons.addStretch(1)
        btn_export = QtWidgets.QPushButton("Скачать Excel")
        btn_export.setObjectName("downloadButton")
        self._set_button_icon(btn_export, "free-icon-download-126488.png")
        btn_export.clicked.connect(self._export_role_matrix_report)
        buttons.addWidget(btn_export)
        btn_close = QtWidgets.QPushButton("Закрыть")
        btn_close.clicked.connect(dlg.accept)
        buttons.addWidget(btn_close)
        layout.addLayout(buttons)
        dlg.exec()

    def _export_role_matrix_report(self) -> None:
        rows = list(self.role_matrix_last_report or [])
        if not rows:
            self._show_alert("Ролевая матрица", "Нет данных для выгрузки. Сначала выполните сверку с Larix")
            return

        project_name = self.cb_matrix_project.currentText().strip() or "project"
        safe_project = re.sub(r"[^0-9A-Za-zА-Яа-яЁё._-]+", "_", project_name).strip("_") or "project"
        default_name = f"Сверка_ролевой_матрицы_{safe_project}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Скачать отчет сверки",
            default_name,
            "Excel (*.xlsx)",
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        try:
            summary_rows: List[Dict[str, Any]] = []
            detail_rows: List[Dict[str, Any]] = []
            error_rows: List[Dict[str, Any]] = []

            for row in rows:
                summary_item = {
                    "Статус": row.get("status", ""),
                    "Строка Excel": row.get("row_number", ""),
                    "Путь Excel": row.get("path_text", ""),
                    "Папка Larix": row.get("folder_text", ""),
                    "ID папки Larix": row.get("folder_id", ""),
                    "Изменения": " | ".join(row.get("changes") or []),
                    "Комментарий / ошибка": row.get("message", ""),
                }
                summary_rows.append(summary_item)
                if row.get("status") == "Ошибка":
                    error_rows.append(dict(summary_item))

                for command in row.get("commands") or []:
                    principal = command.get("principal") or {}
                    current_ids = set(command.get("current_ids") or set())
                    desired_ids = set(command.get("desired_ids") or set())
                    detail_rows.append({
                        "Строка Excel": row.get("row_number", ""),
                        "Путь Excel": row.get("path_text", ""),
                        "Статус строки": row.get("status", ""),
                        "Имя в Excel": command.get("requested_name", ""),
                        "Объект Larix": principal.get("name", ""),
                        "Тип": "Роль" if principal.get("type") == "role" else "Пользователь",
                        "ID": principal.get("id", ""),
                        "Текущие права": _matrix_codes_text(current_ids),
                        "Требуемые права": _matrix_codes_text(desired_ids),
                        "Нужно изменить": "Да" if current_ids != desired_ids else "Нет",
                    })

            overview = pd.DataFrame([
                {"Параметр": "Дата отчета", "Значение": datetime.now().strftime("%d.%m.%Y %H:%M:%S")},
                {"Параметр": "Проект", "Значение": project_name},
                {"Параметр": "Файл матрицы", "Значение": self.role_matrix_path or ""},
                {"Параметр": "Лист Excel", "Значение": self.cb_matrix_sheet.currentText().strip()},
                {"Параметр": "Всего строк", "Значение": len(rows)},
                {"Параметр": "К применению", "Значение": sum(1 for r in rows if r.get("status") == "К применению")},
                {"Параметр": "Без изменений", "Значение": sum(1 for r in rows if r.get("status") == "Без изменений")},
                {"Параметр": "Ошибок", "Значение": sum(1 for r in rows if r.get("status") == "Ошибка")},
            ])
            df_summary = pd.DataFrame(summary_rows)
            df_errors = pd.DataFrame(error_rows, columns=df_summary.columns)
            df_details = pd.DataFrame(detail_rows, columns=[
                "Строка Excel", "Путь Excel", "Статус строки", "Имя в Excel", "Объект Larix",
                "Тип", "ID", "Текущие права", "Требуемые права", "Нужно изменить"
            ])

            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                overview.to_excel(writer, sheet_name="Сводка", index=False)
                df_summary.to_excel(writer, sheet_name="Сверка", index=False)
                df_errors.to_excel(writer, sheet_name="Ошибки", index=False)
                df_details.to_excel(writer, sheet_name="Детали прав", index=False)

                for ws in writer.book.worksheets:
                    ws.freeze_panes = "A2"
                    if ws.max_row >= 1 and ws.max_column >= 1:
                        ws.auto_filter.ref = ws.dimensions
                    for column_cells in ws.columns:
                        values = [str(cell.value or "") for cell in column_cells[:200]]
                        width = min(max(max((len(v) for v in values), default=0) + 2, 10), 55)
                        ws.column_dimensions[column_cells[0].column_letter].width = width
                    for row_cells in ws.iter_rows():
                        for cell in row_cells:
                            alignment = copy(cell.alignment)
                            alignment.vertical = "top"
                            alignment.wrap_text = True
                            cell.alignment = alignment

            self._show_success(
                "Ролевая матрица",
                f"Отчет сохранен:\n{path}\n\nЛисты: Сводка, Сверка, Ошибки, Детали прав."
            )
        except Exception as exc:
            self._log("❌ Не удалось выгрузить отчет сверки: " + traceback.format_exc())
            self._show_alert("Ошибка выгрузки", f"Не удалось сохранить Excel:\n{exc}")

    def _preview_role_matrix(self) -> None:
        project_id = self._matrix_project_id()
        sheet_name = self.cb_matrix_sheet.currentText().strip() if hasattr(self, "cb_matrix_sheet") else ""
        if not self.client.token:
            self._show_alert("Ролевая матрица", "Сначала войдите в Larix")
            return
        if not project_id:
            self._show_alert("Ролевая матрица", "Выберите проект")
            return
        if not self.role_matrix_path or not os.path.exists(self.role_matrix_path):
            self._show_alert("Ролевая матрица", "Загрузите Excel-файл матрицы")
            return
        if not sheet_name:
            self._show_alert("Ролевая матрица", "Выберите лист Excel")
            return

        self.role_matrix_plan = []
        self.role_matrix_last_report = []
        self.tbl_matrix.setRowCount(0)
        self.lbl_matrix_summary.setText("Сверка с Larix...")
        self._matrix_preview_worker = RoleMatrixPreviewWorker(
            self.client, int(project_id), self.role_matrix_path, sheet_name
        )
        self._matrix_preview_worker.log.connect(self._log)
        self._matrix_preview_worker.finished_ok.connect(self._on_role_matrix_preview_done)
        self._matrix_preview_worker.failed.connect(self._on_role_matrix_failed)
        self._matrix_preview_worker.finished.connect(self._update_role_matrix_controls)
        self._matrix_preview_worker.start()
        self._update_role_matrix_controls()

    def _on_role_matrix_preview_done(self, data: Dict[str, Any]) -> None:
        self.role_matrix_plan = list(data.get("plan") or [])
        self.role_matrix_last_report = list(self.role_matrix_plan)
        self._fill_role_matrix_table(self.tbl_matrix, self.role_matrix_plan, wrap=False)
        self.lbl_matrix_summary.setText(
            f"Изменить: {data.get('pending', 0)} · Без изменений: {data.get('unchanged', 0)} · Ошибок: {data.get('errors', 0)}"
        )
        self._update_role_matrix_controls()
        if self._matrix_apply_after_preview:
            self._matrix_apply_after_preview = False
            if int(data.get("pending", 0)) > 0:
                self._apply_role_matrix()

    def _apply_role_matrix(self) -> None:
        if not self.role_matrix_plan:
            self._matrix_apply_after_preview = True
            self._preview_role_matrix()
            return
        project_id = self._matrix_project_id()
        if not project_id:
            self._show_alert("Ролевая матрица", "Выберите проект")
            return
        pending = sum(1 for row in self.role_matrix_plan if row.get("status") == "К применению")
        errors = sum(1 for row in self.role_matrix_plan if row.get("status") == "Ошибка")
        if pending == 0:
            self._show_success("Ролевая матрица", "Все проверенные права уже соответствуют Excel")
            return
        warning = f"\n\nСтрок с ошибками: {errors}. Они будут пропущены." if errors else ""
        if not self._ask_alert(
            "Применить права",
            f"Применить изменения ролевой матрицы к проекту?\nПапок к изменению: {pending}.{warning}"
        ):
            return

        self.lbl_matrix_summary.setText("Применение прав...")
        self._matrix_apply_worker = RoleMatrixApplyWorker(self.client, int(project_id), self.role_matrix_plan)
        self._matrix_apply_worker.log.connect(self._log)
        self._matrix_apply_worker.finished_ok.connect(self._on_role_matrix_apply_done)
        self._matrix_apply_worker.failed.connect(self._on_role_matrix_failed)
        self._matrix_apply_worker.finished.connect(self._update_role_matrix_controls)
        self._matrix_apply_worker.start()
        self._update_role_matrix_controls()

    def _on_role_matrix_apply_done(self, stats: Dict[str, Any]) -> None:
        text = (
            f"Изменено папок: {stats.get('folders', 0)} · "
            f"обновлено типов прав: {stats.get('buckets', 0)} · "
            f"пропущено актуальных: {stats.get('skipped_buckets', 0)} · "
            f"папок без изменений: {stats.get('unchanged', 0)} · "
            f"ошибок: {stats.get('errors', 0)}"
        )
        self.lbl_matrix_summary.setText(text)
        self.role_matrix_plan = []
        self.btn_matrix_apply.setEnabled(False)
        self._show_success("Ролевая матрица", text + "\n\nДля актуальной сверки нажмите «Сверить с Larix» ещё раз.")

    def _on_role_matrix_failed(self, error: str) -> None:
        self._matrix_apply_after_preview = False
        self._log("❌ Ролевая матрица:\n" + error)
        self.lbl_matrix_summary.setText("Ошибка операции")
        self._show_alert("Ролевая матрица", error.splitlines()[-1] if error else "Неизвестная ошибка")
        self._update_role_matrix_controls()

    def _open_create_project_dialog(self):
        ws_display = self.cb_workspace.currentText()
        if ws_display not in self.workspace_map:
            self._show_alert("Ошибка", "Сначала выберите пространство")
            return
        ws_id = self.workspace_map[ws_display]
        dlg = CreateProjectDialog(self, self.client, ws_id, ws_display, self.default_structure_path)
        if dlg.exec() == QtWidgets.QDialog.Accepted and dlg.result_data:
            project_id = dlg.result_data['project_id']
            project_title = dlg.result_data['title']
            self._log(f"✅ Проект создан: ID={project_id}, '{project_title}'")
            self._log("⏳ Ожидание синхронизации (1 сек)...")
            time.sleep(1.0)
            self._log("🔄 Обновление списка проектов...")
            self._load_projects()
            found = False
            for display, pid in self.project_map.items():
                if pid == project_id:
                    found = True
                    self._log(f"✅ Проект найден в списке: {display}")
                    break
            if not found:
                self._log(f"⚠️ Проект ID={project_id} не найден, повторяю через 2 сек...")
                time.sleep(2.0)
                self._load_projects()
                for display, pid in self.project_map.items():
                    if pid == project_id:
                        found = True
                        self._log(f"✅ Проект найден после повторной загрузки: {display}")
                        break
            if found:
                self._show_success(
                    "Проект создан",
                    f"ID: {project_id}\nНазвание: {project_title}\n"
                    f"Папок: {dlg.result_data['folders_created']}\n\nТеперь можно импортировать пользователей."
                )
            else:
                self._show_alert("Внимание",
                    f"Проект создан (ID={project_id}), но не найден в списке.\n"
                    f"Нажмите «Обновить проекты».")

    def _start_add_users(self):
        users = list(self.loaded_users)
        if not users:
            self._show_alert("Ошибка", "Нет данных")
            return
        ws_display = self.cb_workspace.currentText()
        users, missing_projects, missing_roles, roles_without_projects = self._resolve_user_project_ids(users)
        if missing_projects:
            self._show_alert("Ошибка",
                "В Excel указаны проекты, которых нет в выбранном пространстве:\n" + "\n".join(missing_projects))
            return
        if missing_roles:
            self._show_alert("Ошибка",
                "В Excel указаны роли, которых нет в справочнике ролей:\n" + "\n".join(missing_roles) +
                "\n\n💡 Посмотрите в Log, какие роли загрузились")
            return
        if roles_without_projects:
            self._show_alert("Ошибка",
                "Для этих пользователей указана роль, но не указан проект:\n" + "\n".join(roles_without_projects))
            return
        assignments_count = sum(len(u.get("project_ids", [])) for u in users)
        roles_count = sum(len(u.get("role_items", [])) for u in users)
        admin_count = sum(1 for u in users if u.get("is_project_admin", False))
        project_line = f"Назначений в проекты: {assignments_count}" if assignments_count else "Проекты не указаны"
        role_line = f"Назначений ролей: {roles_count}" if roles_count else "Роли не указаны"
        admin_line = f"Администраторов: {admin_count}" if admin_count else "Администраторов: нет"
        if not self._ask_alert("Подтверждение",
            f"Добавить {len(users)} пользователей в:\n{ws_display}\n\n"
            f"📋 Детали:\n• {project_line}\n• {role_line}\n• {admin_line}\n\nПродолжить?"):
            return
        self.log.clear()
        if self._log_view is not None:
            self._log_view.clear()
        self._log(f"Пространство: {ws_display}")
        self._log(f"• {project_line}")
        self._log(f"• {role_line}")
        self._log(f"• {admin_line}")
        self._log(f"Запланировано: {len(users)}")
        self._log("Начинаю...")
        self.btn_run.setEnabled(False)

        class Worker(QThread):
            log_signal = Signal(str)
            done_signal = Signal(list)
            def __init__(self, client, ws_id, users):
                super().__init__()
                self.client = client
                self.ws_id = ws_id
                self.users = users
            def run(self):
                results = []
                for i, user in enumerate(self.users, 1):
                    email = user["email"]
                    name = f"{user['surname']} {user['name']}".strip() or email
                    is_admin = user.get("is_project_admin", False)
                    admin_mark = " [АДМИН]" if is_admin else ""
                    self.log_signal.emit(f"[{i}/{len(self.users)}] {name} ({email}){admin_mark}")
                    result = self.client.add_user_with_creation(
                        self.ws_id, email,
                        last_name=user.get("surname", ""),
                        first_name=user.get("name", ""),
                        project_ids=user.get("project_ids", []),
                        role_items=user.get("role_items", []),
                        is_project_admin=is_admin
                    )
                    result["display"] = name
                    results.append(result)
                    if result.get("success"):
                        self.log_signal.emit(f"   ✅ {result.get('message', 'OK')}")
                    else:
                        self.log_signal.emit(f"   ❌ Ошибка: {result.get('error', 'Ошибка')}")
                    time.sleep(0.3)
                self.done_signal.emit(results)

        ws_id = self.workspace_map[ws_display]
        self.worker = Worker(self.client, ws_id, users)
        self.worker.log_signal.connect(self._log)
        self.worker.done_signal.connect(self._on_finished)
        self.worker.finished.connect(lambda: self.btn_run.setEnabled(True))
        self.worker.start()

    def _on_finished(self, results: List[Dict]):
        ok = sum(1 for r in results if r.get("success"))
        fail = len(results) - ok
        already = sum(
            1 for r in results
            if r.get("success") and "уже добавлен" in str(r.get("message", "")).lower()
        )
        changed = max(0, ok - already)
        self._log("\n" + "=" * 50)
        self._log(f"ГОТОВО: {len(results)}")
        self._log(f"   ✅ Успешно: {changed}")
        if already:
            self._log(f"   ↺ Уже добавлены: {already}")
        if fail:
            self._log(f"   ❌ Ошибки: {fail}")
        self._log("=" * 50)
        if hasattr(self, "lbl_user_import_summary"):
            self.lbl_user_import_summary.setText(
                f"Результат последнего импорта: успешно {changed} · уже добавлены {already} · ошибок {fail}"
            )
        msg = f"Готово!\nУспешно: {changed}"
        if already:
            msg += f"\nУже добавлены: {already}"
        if fail:
            msg += f"\nОшибок: {fail}"
        self._show_success("Результат", msg)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    font = QtGui.QFont("Segoe UI", 9)
    app.setFont(font)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
