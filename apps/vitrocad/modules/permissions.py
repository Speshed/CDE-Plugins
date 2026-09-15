# -*- coding: utf-8 -*-
"""
VitroCAD Permissions GUI
------------------------
Утилита для выдачи прав в VitroCAD по Excel-матрице.

Зависимости:
    pip install requests openpyxl PySide6
Если PySide6 недоступен:
    pip install requests openpyxl PyQt5

Порядок работы:
    1) Войти в VitroCAD.
    2) Выбрать Excel.
    3) Можно нажать «Проверить» для предварительного контроля.
    4) Можно сразу нажать «Применить права»: программа сама выполнит проверку,
       покажет лог, и если ошибок нет — применит права после одного подтверждения.
"""

from __future__ import annotations

import difflib
import json
import os
import re
import sys
import time
import traceback
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import requests
import urllib3
from openpyxl import load_workbook

try:
    from PySide6 import QtCore, QtGui, QtWidgets

    Signal = QtCore.Signal
    Slot = QtCore.Slot
except ImportError:  # pragma: no cover
    from PyQt5 import QtCore, QtGui, QtWidgets

    Signal = QtCore.pyqtSignal
    Slot = QtCore.pyqtSlot


# -----------------------------------------------------------------------------
# Базовые системные ID VitroCAD
# -----------------------------------------------------------------------------
# На большинстве инсталляций эти системные списки одинаковые. Если на другом
# сервере они отличаются, их можно поменять в блоке "Расширенные настройки".
DEFAULT_SERVER = "https://vitrocad.bim-info.ru"
DEFAULT_FILE_LIST_ID = "966e62c5-a803-49a0-a1be-e680d130c481"
DEFAULT_SCOPE_LIST_ID = "ecc16787-6b93-40b9-b563-704d9b090fe8"  # Разрывы прав
DEFAULT_USER_PERMISSION_CONTENT_TYPE_ID = "30a13e40-f715-438d-af7e-cc47cf3a5891"  # Доступ пользователя
DEFAULT_PRINCIPAL_LIST_ID = "e3a94bde-0ca9-456f-b338-4465d40389ee"  # Пользователи/группы

VERIFY_SSL_CERTIFICATE = False
if not VERIFY_SSL_CERTIFICATE:
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Индивидуальные права из Excel. В Excel обычно перечислены действия,
# а VitroCAD хранит готовый "уровень доступа". Если на сервере есть точное
# имя уровня, будет использовано оно. Если нет — сработает fallback ниже.
PERMISSION_RANK = [
    "Просмотр",
    "Скачать",
    "Загрузить",
    "Создать",
    "Удалить",
    "Изменить",
]

# Fallback для серверов, где уровни доступа называются не как отдельные действия,
# а укрупнённо: "Просмотр", "Совместная работа", "Изменить".
EXCEL_PERMISSION_FALLBACKS: Dict[str, List[str]] = {
    "Просмотр": ["Просмотр", "Read"],
    "Скачать": ["Скачать", "Download", "Совместная работа", "Изменить"],
    "Загрузить": ["Загрузить", "Загрузить файл", "Upload", "Совместная работа", "Изменить"],
    "Создать": ["Создать", "Create", "Совместная работа", "Изменить"],
    "Удалить": ["Удалить", "Delete", "Изменить", "Совместная работа"],
    "Изменить": ["Изменить", "Update", "Совместная работа"],
    "Без доступа": ["Без доступа", "No access"],
}


# -----------------------------------------------------------------------------
# Текстовые утилиты
# -----------------------------------------------------------------------------
def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").replace("\r", " ").replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)
    if text.casefold() in {"none", "nan", "null"}:
        return ""
    return text


def normalize_text(value: object) -> str:
    text = clean_text(value)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("ё", "е").replace("Ё", "Е")
    text = text.casefold()
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_path_parts(parts: Sequence[object]) -> str:
    return "/".join(normalize_text(part) for part in parts if clean_text(part))


def display_path(parts: Sequence[object]) -> str:
    return " / ".join(clean_text(part) for part in parts if clean_text(part))


def get_item_name(item: dict) -> str:
    field_map = item.get("fieldValueMap") or {}
    for key in ("name", "Name", "Название", "title", "Title", "code"):
        value = clean_text(field_map.get(key))
        if value:
            return value
    for key in ("name", "Name", "title", "Title"):
        value = clean_text(item.get(key))
        if value:
            return value
    return "Без названия"


def field_value_name(value: object) -> str:
    if isinstance(value, dict):
        return get_item_name(value)
    if isinstance(value, list) and value:
        return field_value_name(value[0])
    return clean_text(value)


def field_value_id(value: object) -> str:
    if isinstance(value, dict):
        return clean_text(value.get("id") or value.get("Id"))
    if isinstance(value, list) and value:
        return field_value_id(value[0])
    return clean_text(value)


def content_type_id_from_item(item: dict) -> str:
    return clean_text(item.get("contentTypeId") or item.get("content_type_id") or item.get("id") or item.get("Id"))


def item_list_id(item: dict, fallback: str = DEFAULT_FILE_LIST_ID) -> str:
    return clean_text(item.get("listId") or item.get("list_id")) or fallback


def is_folder_like(item: dict, content_types: Optional[Dict[str, dict]] = None) -> bool:
    content_type_id = content_type_id_from_item(item)
    type_info = (content_types or {}).get(content_type_id, {})
    field_map = item.get("fieldValueMap") or {}
    item_path = item.get("itemPath") or {}
    child_count = int(item_path.get("childCount") or 0)

    if bool(type_info.get("isFile")):
        return False
    if bool(type_info.get("isFolder")):
        return True
    fs_obj_type = field_map.get("FSObjType")
    if fs_obj_type in (1, "1"):
        return True
    # В некоторых ответах isFolder не проставлен, но childCount есть.
    if child_count > 0:
        return True
    # Для дерева целевых папок лучше показывать не-файловые элементы.
    if type_info and not type_info.get("isFile"):
        return True
    return False


def permission_tokens(cell_value: object) -> List[str]:
    text = clean_text(cell_value)
    if not text:
        return []
    for sep in ("\n", "\r", ";", "/"):
        text = text.replace(sep, ",")
    return [clean_text(part) for part in text.split(",") if clean_text(part)]


def canonical_permission_actions(cell_value: object) -> List[str]:
    """Возвращает понятные действия из текста Excel/сервера.

    Примеры:
    - "Просмотр, Скачать" -> ["Просмотр", "Скачать"]
    - "Download" -> ["Скачать"]
    - "Просмотр, Скачать, Загрузить" -> ["Просмотр", "Скачать", "Загрузить"]
    """
    tokens = permission_tokens(cell_value)
    if not tokens:
        return []

    aliases = {
        "загрузить файл": "Загрузить",
        "upload": "Загрузить",
        "download": "Скачать",
        "read": "Просмотр",
        "view": "Просмотр",
        "create": "Создать",
        "delete": "Удалить",
        "remove": "Удалить",
        "update": "Изменить",
        "edit": "Изменить",
        "modify": "Изменить",
        "no access": "Без доступа",
    }
    canonical_by_norm = {normalize_text(name): name for name in PERMISSION_RANK + ["Без доступа"]}
    result: List[str] = []
    seen = set()

    for token in tokens:
        norm = normalize_text(token)
        action = canonical_by_norm.get(norm) or aliases.get(norm)
        if action and action not in seen:
            seen.add(action)
            result.append(action)

    # Стабильный порядок — как в иерархии прав.
    order = {name: idx for idx, name in enumerate(PERMISSION_RANK + ["Без доступа"])}
    return sorted(result, key=lambda name: order.get(name, 999))


def permission_primary_action(permission_text: object) -> str:
    actions = [name for name in canonical_permission_actions(permission_text) if name != "Без доступа"]
    if not actions:
        return "Без доступа" if "Без доступа" in canonical_permission_actions(permission_text) else ""
    rank = {name: idx for idx, name in enumerate(PERMISSION_RANK)}
    return max(actions, key=lambda name: rank.get(name, -1))


def implied_permission_actions(permission_text: object) -> List[str]:
    """Действия, которые обычно подразумевает наивысшее право.

    Это нужно для серверных уровней вида "Просмотр, Скачать":
    Excel может писать просто "Скачать", а серверный уровень может называться
    полным набором действий.
    """
    primary = permission_primary_action(permission_text)
    if primary not in PERMISSION_RANK:
        return canonical_permission_actions(permission_text)
    idx = PERMISSION_RANK.index(primary)
    return PERMISSION_RANK[: idx + 1]


def excel_permission_from_cell(cell_value: object, apply_no_access: bool = False) -> Optional[str]:
    """Возвращает текст права для поиска на сервере по названию, без ручных ID.

    В отличие от старого варианта, мы не теряем комбинацию прав: если в Excel
    написано "Просмотр, Скачать", то сначала ищем на сервере именно такой уровень.
    Если такого нет, ниже сработает поиск по наивысшему действию и fallback.
    """
    actions = canonical_permission_actions(cell_value)
    if not actions:
        return None
    if actions == ["Без доступа"]:
        return "Без доступа" if apply_no_access else None
    actions = [name for name in actions if name != "Без доступа"]
    if not actions:
        return None
    return ", ".join(actions)


def highest_excel_permission(cell_value: object, apply_no_access: bool = False) -> Optional[str]:
    """Оставлено для совместимости: возвращает наивысшее действие из ячейки."""
    text = excel_permission_from_cell(cell_value, apply_no_access=apply_no_access)
    if not text:
        return None
    primary = permission_primary_action(text)
    return primary or text


# -----------------------------------------------------------------------------
# Данные
# -----------------------------------------------------------------------------
@dataclass
class PermissionEntry:
    row_number: int
    path_parts: List[str]
    permissions: Dict[str, str]

    @property
    def path_text(self) -> str:
        return display_path(self.path_parts)


@dataclass
class FolderItem:
    id: str
    name: str
    parent_id: str
    list_id: str
    content_type_id: str
    type_name: str
    path_parts: List[str]
    level: int
    child_count: int = 0

    @property
    def path_text(self) -> str:
        return display_path(self.path_parts)

    @property
    def norm_path(self) -> str:
        return normalize_path_parts(self.path_parts)


@dataclass
class PrincipalResolution:
    name: str
    id: str = ""
    message: str = ""

    @property
    def ok(self) -> bool:
        return bool(self.id)


@dataclass
class PermissionLevelResolution:
    excel_permission: str
    server_level_name: str = ""
    server_level_id: str = ""
    message: str = ""

    @property
    def ok(self) -> bool:
        return bool(self.server_level_id)


@dataclass
class PlanRow:
    entry: PermissionEntry
    folder: Optional[FolderItem]
    status: str
    message: str
    resolved_principals: Dict[str, PrincipalResolution] = field(default_factory=dict)
    resolved_permission_levels: Dict[str, PermissionLevelResolution] = field(default_factory=dict)

    @property
    def can_apply(self) -> bool:
        if self.status != "OK" or not self.folder:
            return False
        for group_name, permission_name in self.entry.permissions.items():
            pr = self.resolved_principals.get(group_name)
            pl = self.resolved_permission_levels.get(permission_name)
            if not pr or not pr.ok or not pl or not pl.ok:
                return False
        return True

    @property
    def permission_count(self) -> int:
        return len(self.entry.permissions)


# -----------------------------------------------------------------------------
# Excel parser
# -----------------------------------------------------------------------------
class PermissionExcelParser:
    def __init__(self, excel_path: str, apply_no_access: bool = False):
        self.excel_path = excel_path
        self.apply_no_access = apply_no_access

    def parse(self) -> List[PermissionEntry]:
        if not os.path.exists(self.excel_path):
            raise FileNotFoundError(f"Excel-файл не найден: {self.excel_path}")

        workbook = load_workbook(self.excel_path, read_only=True, data_only=True)
        sheet = workbook.active
        header_row = self._find_header_row(sheet)

        level_columns: List[Tuple[int, int]] = []
        role_columns: List[Tuple[int, str]] = []
        for col_idx in range(1, sheet.max_column + 1):
            header = clean_text(sheet.cell(header_row, col_idx).value)
            m = re.fullmatch(r"Уровень\s*(\d+)", header, flags=re.IGNORECASE)
            if m:
                level_columns.append((int(m.group(1)), col_idx))
            elif header and not header.startswith("Unnamed"):
                role_columns.append((col_idx, header))

        level_columns.sort(key=lambda pair: pair[0])
        if not level_columns:
            raise RuntimeError("В Excel не найдены колонки вида 'Уровень 1', 'Уровень 2', ...")
        if not role_columns:
            raise RuntimeError("В Excel не найдены колонки ролей/групп после колонок уровней.")

        path_stack: List[Optional[str]] = [None] * len(level_columns)
        result: List[PermissionEntry] = []

        for row_idx in range(header_row + 1, sheet.max_row + 1):
            changed = False
            for stack_idx, (_level_num, col_idx) in enumerate(level_columns):
                value = clean_text(sheet.cell(row_idx, col_idx).value)
                if value:
                    path_stack[stack_idx] = value
                    for clear_idx in range(stack_idx + 1, len(path_stack)):
                        path_stack[clear_idx] = None
                    changed = True

            path_parts = [part for part in path_stack if part]
            if not path_parts:
                continue

            permissions: Dict[str, str] = {}
            for col_idx, role_name in role_columns:
                perm = excel_permission_from_cell(sheet.cell(row_idx, col_idx).value, self.apply_no_access)
                if perm:
                    permissions[role_name] = perm

            if permissions:
                result.append(PermissionEntry(row_number=row_idx, path_parts=list(path_parts), permissions=permissions))

        if not result:
            raise RuntimeError("В Excel не найдено ни одной строки с правами.")
        return result

    @staticmethod
    def _find_header_row(sheet) -> int:
        scan_to = min(sheet.max_row, 20)
        for row_idx in range(1, scan_to + 1):
            count = 0
            for col_idx in range(1, sheet.max_column + 1):
                value = clean_text(sheet.cell(row_idx, col_idx).value)
                if re.fullmatch(r"Уровень\s*\d+", value, flags=re.IGNORECASE):
                    count += 1
            if count >= 2:
                return row_idx
        raise RuntimeError("Не найдена строка заголовков с колонками 'Уровень N'.")


# -----------------------------------------------------------------------------
# VitroCAD client
# -----------------------------------------------------------------------------
class VitroClient:
    def __init__(self, base_url: str, token: str = "", timeout: int = 30):
        self.base_url = base_url.strip().rstrip("/")
        self.token = token
        self.timeout = timeout
        self.session = requests.Session()
        self.session.verify = VERIFY_SSL_CERTIFICATE
        # Не используем системные HTTP/HTTPS/SOCKS-прокси из Windows/переменных среды.
        # Для VitroCAD внутри сети это обычно нужно: иначе requests может попытаться
        # идти через V2Ray/Clash/другой SOCKS-прокси и упасть с ошибкой
        # "Missing dependencies for SOCKS support".
        self.session.trust_env = False

    @property
    def headers_json(self) -> Dict[str, str]:
        headers = {"Content-Type": "application/json"}
        if self.token:
            headers["Authorization"] = self.token
        return headers

    @property
    def headers_auth(self) -> Dict[str, str]:
        return {"Authorization": self.token} if self.token else {}

    def login(self, login: str, password: str) -> str:
        response = self.session.post(
            f"{self.base_url}/api/security/login",
            json={"login": login, "password": password},
            timeout=self.timeout,
        )
        self._raise_for_response(response, "Авторизация не удалась")
        data = response.json()
        token = clean_text(data.get("token"))
        if not token:
            raise RuntimeError(f"В ответе авторизации нет token. Ответ: {data}")
        self.token = token
        return token

    def get_list(self, list_id: str) -> dict:
        response = self.session.post(
            f"{self.base_url}/api/list/get/{list_id}",
            headers=self.headers_json,
            timeout=self.timeout,
        )
        self._raise_for_response(response, f"Не удалось прочитать список {list_id}")
        data = response.json()
        if not isinstance(data, dict):
            raise RuntimeError(f"Ожидался объект списка, получено: {type(data).__name__}")
        return data

    def get_children(self, parent_id: str) -> List[dict]:
        response = self.session.post(
            f"{self.base_url}/api/item/getList/{parent_id}",
            headers=self.headers_json,
            json={"sort": {"name": 0}},
            timeout=self.timeout,
        )
        self._raise_for_response(response, f"Не удалось загрузить дочерние элементы {parent_id}")
        data = response.json()
        if not isinstance(data, list):
            raise RuntimeError(f"Ожидался список дочерних элементов, получено: {type(data).__name__}")
        return data

    def get_item(self, item_id: str) -> dict:
        response = self.session.post(
            f"{self.base_url}/api/item/get/{item_id}",
            headers=self.headers_json,
            timeout=self.timeout,
        )
        self._raise_for_response(response, f"Не удалось загрузить элемент {item_id}")
        data = response.json()
        if not isinstance(data, dict):
            raise RuntimeError(f"Ожидался объект элемента, получено: {type(data).__name__}")
        return data

    def get_content_types(self, list_id: str) -> List[dict]:
        response = self.session.post(
            f"{self.base_url}/api/contentType/getByList/{list_id}",
            headers=self.headers_json,
            timeout=self.timeout,
        )
        self._raise_for_response(response, f"Не удалось загрузить типы контента списка {list_id}")
        data = response.json()
        if not isinstance(data, list):
            raise RuntimeError(f"Ожидался список типов контента, получено: {type(data).__name__}")
        return data

    def lookup_get_list(self, field_id: str, value_string: str, list_id: str) -> List[dict]:
        payload = {"field": field_id, "value_string": value_string, "list": list_id}
        response = self.session.post(
            f"{self.base_url}/api/lookup/getList",
            headers=self.headers_json,
            json=payload,
            timeout=self.timeout,
        )
        self._raise_for_response(response, f"Не удалось выполнить lookup по списку {list_id}")
        data = response.json()
        if not isinstance(data, list):
            raise RuntimeError(f"Ожидался список lookup, получено: {type(data).__name__}")
        return data

    def set_unique_permissions(self, item_id: str, copy_permission: bool = True) -> None:
        response = self.session.post(
            f"{self.base_url}/api/security/setItemUniquePermission",
            headers=self.headers_json,
            json={"id": item_id, "copy_permission": copy_permission},
            timeout=self.timeout,
        )
        self._raise_for_response(response, f"Не удалось установить уникальные права для {item_id}")

    def remove_unique_permissions(self, item_id: str) -> None:
        response = self.session.post(
            f"{self.base_url}/api/security/removeItemUniquePermission/{item_id}",
            headers=self.headers_json,
            timeout=self.timeout,
        )
        self._raise_for_response(response, f"Не удалось восстановить наследование прав для {item_id}")

    def delete_items(self, item_ids: Sequence[str]) -> None:
        if not item_ids:
            return
        payload = [{"id": item_id} for item_id in item_ids]
        response = self.session.post(
            f"{self.base_url}/api/item/delete",
            headers=self.headers_json,
            json=payload,
            timeout=self.timeout,
        )
        self._raise_for_response(response, f"Не удалось удалить записи: {', '.join(item_ids[:5])}")

    def create_permission_record(
        self,
        scope_list_id: str,
        permission_content_type_id: str,
        source_item_id: str,
        principal_id: str,
        permission_level_id: str,
        name: str,
    ) -> str:
        item_data = {
            "list_id": scope_list_id,
            "content_type_id": permission_content_type_id,
            "parent_id": scope_list_id,
            "name": name,
            "source": source_item_id,
            "principal": principal_id,
            "permission_level": permission_level_id,
        }
        response = self.session.post(
            f"{self.base_url}/api/item/update",
            headers=self.headers_auth,
            files={"itemListJson": (None, json.dumps([item_data], ensure_ascii=False), "application/json")},
            timeout=self.timeout,
        )
        self._raise_for_response(response, f"Не удалось создать запись прав для {name}")
        data = response.json()
        if not isinstance(data, list) or not data:
            raise RuntimeError(f"Неожиданный ответ item/update: {data}")
        return clean_text(data[0].get("id"))

    @staticmethod
    def _raise_for_response(response: requests.Response, context: str) -> None:
        if response.ok:
            return
        text = response.text
        if len(text) > 2000:
            text = text[:2000] + "..."
        raise RuntimeError(f"{context}. HTTP {response.status_code}: {text}")


# -----------------------------------------------------------------------------
# Resolvers and structure index
# -----------------------------------------------------------------------------
class StructureIndex:
    def __init__(self, content_types: Optional[Dict[str, dict]] = None):
        self.content_types = content_types or {}
        self.items: List[FolderItem] = []
        self.path_index: Dict[str, List[FolderItem]] = {}
        self.name_index: Dict[str, List[FolderItem]] = {}
        self.visited: set[str] = set()

    def add(self, item: FolderItem) -> None:
        self.items.append(item)
        self.path_index.setdefault(item.norm_path, []).append(item)
        self.name_index.setdefault(normalize_text(item.name), []).append(item)

    def find(self, path_parts: Sequence[str], allow_suffix: bool = False) -> Tuple[Optional[FolderItem], str]:
        norm = normalize_path_parts(path_parts)
        exact = self.path_index.get(norm, [])
        if len(exact) == 1:
            return exact[0], "Найдено по полному пути."
        if len(exact) > 1:
            msg = "Неоднозначный полный путь: " + "; ".join(item.path_text for item in exact[:8])
            return None, msg

        if allow_suffix:
            suffix = "/" + norm
            matches = [item for item in self.items if item.norm_path.endswith(suffix)]
            if len(matches) == 1:
                return matches[0], f"Найдено по хвосту пути: {matches[0].path_text}"
            if len(matches) > 1:
                return None, "Хвост пути неоднозначен: " + "; ".join(item.path_text for item in matches[:8])

        last = path_parts[-1] if path_parts else ""
        same_name = self.name_index.get(normalize_text(last), [])
        if same_name:
            return None, "Конечная папка есть, но путь отличается: " + "; ".join(item.path_text for item in same_name[:5])

        close = difflib.get_close_matches(norm, list(self.path_index.keys()), n=5, cutoff=0.72)
        if close:
            variants = []
            for key in close:
                variants.extend(item.path_text for item in self.path_index.get(key, []))
            return None, "Путь не найден. Похожие пути: " + "; ".join(variants[:5])
        return None, "Путь не найден в просканированной структуре."

    def scan(
        self,
        client: VitroClient,
        parent_id: str,
        parent_path_parts: Optional[List[str]],
        level: int,
        max_depth: int,
        max_items: int,
        log_func,
    ) -> None:
        if parent_path_parts is None:
            parent_path_parts = []
        if level > max_depth:
            log_func(f"  WARNING: достигнута максимальная глубина {max_depth}: {display_path(parent_path_parts)}\n")
            return
        if parent_id in self.visited:
            log_func(f"  WARNING: повторный parent_id пропущен: {parent_id}\n")
            return
        self.visited.add(parent_id)
        if max_items and len(self.items) >= max_items:
            log_func(f"  WARNING: достигнут лимит элементов {max_items}\n")
            return

        children = client.get_children(parent_id)
        for raw in children:
            if max_items and len(self.items) >= max_items:
                log_func(f"  WARNING: достигнут лимит элементов {max_items}\n")
                return
            name = get_item_name(raw)
            item_path = raw.get("itemPath") or {}
            child_count = int(item_path.get("childCount") or 0)
            ct_id = content_type_id_from_item(raw)
            ct_info = self.content_types.get(ct_id, {})
            type_name = clean_text(ct_info.get("name") or ct_info.get("description") or ct_id or "Неизвестный тип")
            folder = FolderItem(
                id=clean_text(raw.get("id")),
                name=name,
                parent_id=parent_id,
                list_id=item_list_id(raw),
                content_type_id=ct_id,
                type_name=type_name,
                path_parts=list(parent_path_parts) + [name],
                level=level,
                child_count=child_count,
            )
            self.add(folder)
            if is_folder_like(raw, self.content_types):
                try:
                    self.scan(
                        client,
                        folder.id,
                        folder.path_parts,
                        level + 1,
                        max_depth,
                        max_items,
                        log_func,
                    )
                except Exception as exc:
                    log_func(f"  WARNING: не удалось прочитать потомков '{folder.path_text}': {exc}\n")


class PermissionSchema:
    def __init__(self):
        self.permission_content_type_id = DEFAULT_USER_PERMISSION_CONTENT_TYPE_ID
        self.principal_field_id = ""
        self.principal_list_id = DEFAULT_PRINCIPAL_LIST_ID
        self.permission_level_field_id = ""
        self.permission_level_list_id = ""

    @staticmethod
    def from_content_types(content_types: Sequence[dict], fallback_permission_content_type_id: str) -> "PermissionSchema":
        schema = PermissionSchema()
        schema.permission_content_type_id = fallback_permission_content_type_id or DEFAULT_USER_PERMISSION_CONTENT_TYPE_ID
        selected = None
        for item in content_types:
            item_id = clean_text(item.get("id"))
            name = normalize_text(item.get("name") or item.get("description"))
            if item_id == schema.permission_content_type_id or name in {"доступ пользователя", "user permission"}:
                selected = item
                schema.permission_content_type_id = item_id or schema.permission_content_type_id
                break
        if not selected and content_types:
            selected = content_types[0]
            schema.permission_content_type_id = clean_text(selected.get("id")) or schema.permission_content_type_id

        for field in (selected or {}).get("fieldList", []) or []:
            internal = clean_text(field.get("internalName"))
            field_id = clean_text(field.get("id"))
            fmap = field.get("fieldValueMap") or {}
            if internal == "principal":
                schema.principal_field_id = field_id
                schema.principal_list_id = clean_text(fmap.get("list")) or schema.principal_list_id
            elif internal == "permission_level":
                schema.permission_level_field_id = field_id
                schema.permission_level_list_id = clean_text(fmap.get("list")) or schema.permission_level_list_id
        return schema


class DirectoryResolver:
    def __init__(self, client: VitroClient, schema: PermissionSchema, log_func):
        self.client = client
        self.schema = schema
        self.log = log_func
        self.principals_by_norm: Dict[str, dict] = {}
        self.permission_levels_by_norm: Dict[str, dict] = {}

    def load_permission_levels(self) -> None:
        if not self.schema.permission_level_field_id or not self.schema.permission_level_list_id:
            raise RuntimeError("Не удалось определить поле/список уровней доступа из схемы списка 'Разрывы прав'.")
        levels = self.client.lookup_get_list(
            self.schema.permission_level_field_id,
            "",
            self.schema.permission_level_list_id,
        )
        self.permission_levels_by_norm = {}
        for item in levels:
            name = get_item_name(item)
            if name:
                self.permission_levels_by_norm[normalize_text(name)] = item
        self.log(f"  уровней доступа с сервера: {len(self.permission_levels_by_norm)}\n")
        for name in sorted(get_item_name(item) for item in self.permission_levels_by_norm.values()):
            self.log(f"    - {name}\n")

    def resolve_permission_level(self, excel_permission: str) -> PermissionLevelResolution:
        if not self.permission_levels_by_norm:
            self.load_permission_levels()

        primary = permission_primary_action(excel_permission)
        candidates: List[str] = [excel_permission]
        if primary and normalize_text(primary) != normalize_text(excel_permission):
            candidates.append(primary)
        candidates.extend(EXCEL_PERMISSION_FALLBACKS.get(excel_permission, []))
        if primary:
            candidates.extend(EXCEL_PERMISSION_FALLBACKS.get(primary, []))

        seen = set()
        unique_candidates = []
        for cand in candidates:
            key = normalize_text(cand)
            if key and key not in seen:
                seen.add(key)
                unique_candidates.append(cand)

        # 1) Прямое совпадение по названию уровня на сервере.
        for cand in unique_candidates:
            item = self.permission_levels_by_norm.get(normalize_text(cand))
            if item:
                server_name = get_item_name(item)
                msg = "точное совпадение" if normalize_text(server_name) == normalize_text(excel_permission) else f"Excel '{excel_permission}' → VitroCAD '{server_name}'"
                return PermissionLevelResolution(excel_permission, server_name, clean_text(item.get("id")), msg)

        # 2) Совпадение по набору буквенных действий.
        # Например: Excel = "Скачать", сервер = "Просмотр, Скачать".
        excel_actions = set(canonical_permission_actions(excel_permission))
        implied_actions = set(implied_permission_actions(excel_permission))
        for expected_actions, label in (
            (excel_actions, "совпадение по набору действий"),
            (implied_actions, "совпадение по подразумеваемому набору действий"),
        ):
            if not expected_actions:
                continue
            matches = []
            for item in self.permission_levels_by_norm.values():
                server_name = get_item_name(item)
                server_actions = set(canonical_permission_actions(server_name))
                if server_actions and server_actions == expected_actions:
                    matches.append(item)
            if len(matches) == 1:
                item = matches[0]
                server_name = get_item_name(item)
                return PermissionLevelResolution(
                    excel_permission,
                    server_name,
                    clean_text(item.get("id")),
                    f"{label}: Excel '{excel_permission}' → VitroCAD '{server_name}'",
                )

        # 3) Очень близкое текстовое совпадение, но только с высоким порогом,
        # чтобы не выдать права на случайно похожий уровень.
        server_keys = list(self.permission_levels_by_norm.keys())
        close = difflib.get_close_matches(normalize_text(excel_permission), server_keys, n=1, cutoff=0.92)
        if close:
            item = self.permission_levels_by_norm[close[0]]
            server_name = get_item_name(item)
            return PermissionLevelResolution(
                excel_permission,
                server_name,
                clean_text(item.get("id")),
                f"близкое текстовое совпадение: Excel '{excel_permission}' → VitroCAD '{server_name}'",
            )

        return PermissionLevelResolution(
            excel_permission,
            message=f"Не найден уровень доступа для Excel-права '{excel_permission}'. Доступные: {', '.join(sorted(get_item_name(i) for i in self.permission_levels_by_norm.values()))}",
        )

    def resolve_principal(self, group_name: str) -> PrincipalResolution:
        norm = normalize_text(group_name)
        if norm in self.principals_by_norm:
            item = self.principals_by_norm[norm]
            return PrincipalResolution(group_name, clean_text(item.get("id")), "найдено ранее")

        # Основной способ: lookup по полю principal из content type "Доступ пользователя".
        if self.schema.principal_field_id and self.schema.principal_list_id:
            try:
                items = self.client.lookup_get_list(self.schema.principal_field_id, group_name, self.schema.principal_list_id)
                for item in items:
                    item_name = get_item_name(item)
                    self.principals_by_norm.setdefault(normalize_text(item_name), item)
                exact = self.principals_by_norm.get(norm)
                if exact:
                    return PrincipalResolution(group_name, clean_text(exact.get("id")), "найдено через lookup")
                if len(items) == 1:
                    only = items[0]
                    return PrincipalResolution(
                        group_name,
                        clean_text(only.get("id")),
                        f"точного совпадения нет; взят единственный результат lookup: {get_item_name(only)}",
                    )
            except Exception as exc:
                self.log(f"  WARNING: lookup группы '{group_name}' не сработал: {exc}\n")

        # Запасной способ: читать список principals напрямую. Может быть тяжелее, но полезно.
        try:
            items = self.client.get_children(self.schema.principal_list_id)
            for item in items:
                self.principals_by_norm.setdefault(normalize_text(get_item_name(item)), item)
            exact = self.principals_by_norm.get(norm)
            if exact:
                return PrincipalResolution(group_name, clean_text(exact.get("id")), "найдено в списке principal")
        except Exception as exc:
            self.log(f"  WARNING: список principal прочитать не удалось: {exc}\n")

        similar = difflib.get_close_matches(norm, list(self.principals_by_norm.keys()), n=5, cutoff=0.72)
        if similar:
            variants = ", ".join(get_item_name(self.principals_by_norm[key]) for key in similar)
            return PrincipalResolution(group_name, message=f"Группа не найдена точно. Похожие: {variants}")
        return PrincipalResolution(group_name, message="Группа/пользователь не найден на сервере.")


# -----------------------------------------------------------------------------
# Workers
# -----------------------------------------------------------------------------
class LoginWorker(QtCore.QObject):
    done = Signal(str)
    failed = Signal(str)

    def __init__(self, server: str, login: str, password: str):
        super().__init__()
        self.server = server
        self.login_name = login
        self.password = password

    @Slot()
    def run(self):
        try:
            client = VitroClient(self.server)
            token = client.login(self.login_name, self.password)
            self.done.emit(token)
        except Exception:
            self.failed.emit(traceback.format_exc())


class FolderTreeWorker(QtCore.QObject):
    loaded = Signal(object)
    failed = Signal(str)

    def __init__(self, server: str, token: str, parent_id: str, content_types: Dict[str, dict]):
        super().__init__()
        self.server = server
        self.token = token
        self.parent_id = parent_id
        self.content_types = content_types

    @Slot()
    def run(self):
        try:
            client = VitroClient(self.server, self.token)
            children = client.get_children(self.parent_id)
            folders = [item for item in children if is_folder_like(item, self.content_types)]
            folders.sort(key=lambda it: normalize_text(get_item_name(it)))
            self.loaded.emit(folders)
        except Exception:
            self.failed.emit(traceback.format_exc())


class PreviewWorker(QtCore.QObject):
    log = Signal(str)
    done = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        server: str,
        token: str,
        excel_path: str,
        target_id: str,
        target_name: str,
        include_target_in_path: bool,
        file_list_id: str,
        scope_list_id: str,
        permission_content_type_id: str,
        max_depth: int,
        max_items: int,
        limit: int,
        allow_suffix: bool,
        apply_no_access: bool,
    ):
        super().__init__()
        self.server = server
        self.token = token
        self.excel_path = excel_path
        self.target_id = target_id
        self.target_name = target_name
        self.include_target_in_path = include_target_in_path
        self.file_list_id = file_list_id
        self.scope_list_id = scope_list_id
        self.permission_content_type_id = permission_content_type_id
        self.max_depth = max_depth
        self.max_items = max_items
        self.limit = limit
        self.allow_suffix = allow_suffix
        self.apply_no_access = apply_no_access

    def _log(self, text: str):
        self.log.emit(text)

    @Slot()
    def run(self):
        try:
            self._log("=" * 70 + "\n")
            self._log("Построение предпросмотра выдачи прав\n")
            self._log(f"Excel: {self.excel_path}\n")
            self._log(f"Target: {self.target_name} / {self.target_id}\n")
            self._log(f"Include target in path: {self.include_target_in_path}\n")
            self._log("=" * 70 + "\n\n")

            parser = PermissionExcelParser(self.excel_path, apply_no_access=self.apply_no_access)
            entries = parser.parse()
            if self.limit > 0:
                entries = entries[: self.limit]
                self._log(f"Ограничение: первые {self.limit} строк с правами.\n")
            self._log(f"[1/5] Прочитано строк с правами: {len(entries)}\n")

            client = VitroClient(self.server, self.token)

            self._log("[2/5] Загрузка типов папочного списка\n")
            file_content_types = client.get_content_types(self.file_list_id)
            ct_map: Dict[str, dict] = {}
            for ct in file_content_types:
                ct_id = clean_text(ct.get("id"))
                if ct_id:
                    ct_map[ct_id] = ct
            self._log(f"  типов контента: {len(ct_map)}\n")

            self._log("[3/5] Загрузка схемы прав\n")
            scope_content_types = client.get_content_types(self.scope_list_id)
            schema = PermissionSchema.from_content_types(scope_content_types, self.permission_content_type_id)
            if not schema.principal_field_id:
                self._log("  WARNING: поле principal не найдено в схеме. Поиск групп будет через список principal, если доступен.\n")
            if not schema.permission_level_field_id:
                self._log("  WARNING: поле permission_level не найдено в схеме. Проверь ID списка разрывов прав.\n")
            self._log(f"  content type прав: {schema.permission_content_type_id}\n")
            self._log(f"  principal field/list: {schema.principal_field_id} / {schema.principal_list_id}\n")
            self._log(f"  permission level field/list: {schema.permission_level_field_id} / {schema.permission_level_list_id}\n")

            self._log("[4/5] Сканирование структуры\n")
            root_path = [self.target_name] if self.include_target_in_path and self.target_name and self.target_name != "Корень списка" else []
            index = StructureIndex(ct_map)
            index.scan(client, self.target_id, root_path, 0, self.max_depth, self.max_items, self._log)
            self._log(f"  найдено элементов: {len(index.items)}\n")
            self._log(f"  уникальных путей: {len(index.path_index)}\n")

            self._log("[5/5] Проверка путей, групп и уровней доступа\n")
            resolver = DirectoryResolver(client, schema, self._log)
            resolver.load_permission_levels()

            plan: List[PlanRow] = []
            role_cache: Dict[str, PrincipalResolution] = {}
            level_cache: Dict[str, PermissionLevelResolution] = {}
            for entry in entries:
                folder, folder_msg = index.find(entry.path_parts, allow_suffix=self.allow_suffix)
                status = "OK" if folder else "Папка не найдена"
                resolved_principals: Dict[str, PrincipalResolution] = {}
                resolved_levels: Dict[str, PermissionLevelResolution] = {}

                for group_name, excel_perm in entry.permissions.items():
                    if group_name not in role_cache:
                        role_cache[group_name] = resolver.resolve_principal(group_name)
                    resolved_principals[group_name] = role_cache[group_name]

                    if excel_perm not in level_cache:
                        level_cache[excel_perm] = resolver.resolve_permission_level(excel_perm)
                    resolved_levels[excel_perm] = level_cache[excel_perm]

                missing_groups = [name for name, res in resolved_principals.items() if not res.ok]
                missing_levels = [name for name, res in resolved_levels.items() if not res.ok]
                messages = [folder_msg]
                if missing_groups:
                    status = "Ошибка сопоставления"
                    messages.append("Не найдены группы: " + ", ".join(missing_groups[:10]))
                if missing_levels:
                    status = "Ошибка сопоставления"
                    messages.append("Не найдены уровни доступа: " + ", ".join(missing_levels[:10]))
                for excel_perm, res in resolved_levels.items():
                    if res.ok and normalize_text(excel_perm) != normalize_text(res.server_level_name):
                        messages.append(res.message)

                plan.append(
                    PlanRow(
                        entry=entry,
                        folder=folder,
                        status=status,
                        message=" ".join(msg for msg in messages if msg),
                        resolved_principals=resolved_principals,
                        resolved_permission_levels=resolved_levels,
                    )
                )

            ok_count = sum(1 for row in plan if row.can_apply)
            error_count = len(plan) - ok_count
            self._log(f"Предпросмотр готов. OK: {ok_count}; с ошибками: {error_count}; всего: {len(plan)}\n")
            self.done.emit({"plan": plan, "schema": schema})
        except Exception:
            self.failed.emit(traceback.format_exc())


class ApplyWorker(QtCore.QObject):
    log = Signal(str)
    done = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        server: str,
        token: str,
        plan: Sequence[PlanRow],
        scope_list_id: str,
        permission_content_type_id: str,
        dry_run: bool,
    ):
        super().__init__()
        self.server = server
        self.token = token
        self.plan = list(plan)
        self.scope_list_id = scope_list_id
        self.permission_content_type_id = permission_content_type_id
        self.dry_run = dry_run

    def _log(self, text: str):
        self.log.emit(text)

    @Slot()
    def run(self):
        try:
            client = VitroClient(self.server, self.token)
            rows = [row for row in self.plan if row.can_apply]
            stats = {
                "folders": 0,
                "deleted": 0,
                "created": 0,
                "failed": 0,
                "skipped": len(self.plan) - len(rows),
            }
            self._log("=" * 70 + "\n")
            self._log("Выдача прав по Excel\n")
            self._log(f"Режим: {'ТЕСТОВЫЙ' if self.dry_run else 'БОЕВОЙ'}\n")
            self._log(f"К обработке: {len(rows)}; пропущено из-за ошибок предпросмотра: {stats['skipped']}\n")
            self._log("=" * 70 + "\n\n")

            for idx, row in enumerate(rows, 1):
                folder = row.folder
                assert folder is not None
                stats["folders"] += 1
                self._log(f"[{idx}/{len(rows)}] Excel строка {row.entry.row_number}: {row.entry.path_text}\n")
                self._log(f"  VitroCAD: {folder.path_text}\n")
                self._log(f"  ID: {folder.id}\n")

                if self.dry_run:
                    for group_name, excel_perm in row.entry.permissions.items():
                        pr = row.resolved_principals[group_name]
                        pl = row.resolved_permission_levels[excel_perm]
                        self._log(f"  [DRY RUN] {group_name} -> {pl.server_level_name} ({excel_perm})\n")
                    continue

                try:
                    self._clear_permissions(client, folder.id, stats)
                    for group_name, excel_perm in row.entry.permissions.items():
                        pr = row.resolved_principals[group_name]
                        pl = row.resolved_permission_levels[excel_perm]
                        record_name = f"Разрывы прав - {group_name} - {datetime.now().strftime('%H:%M:%S')}"
                        client.create_permission_record(
                            self.scope_list_id,
                            self.permission_content_type_id,
                            folder.id,
                            pr.id,
                            pl.server_level_id,
                            record_name,
                        )
                        stats["created"] += 1
                        self._log(f"  + {group_name}: {pl.server_level_name} (Excel: {excel_perm})\n")
                except Exception as exc:
                    stats["failed"] += 1
                    self._log(f"  ERROR: {exc}\n")

            self._log("\nИтог\n")
            self._log(f"  обработано папок: {stats['folders']}\n")
            self._log(f"  удалено старых записей прав: {stats['deleted']}\n")
            self._log(f"  создано новых записей прав: {stats['created']}\n")
            self._log(f"  ошибок: {stats['failed']}\n")
            self._log(f"  пропущено: {stats['skipped']}\n")
            self.done.emit(stats)
        except Exception:
            self.failed.emit(traceback.format_exc())

    def _clear_permissions(self, client: VitroClient, item_id: str, stats: Dict[str, int]) -> None:
        self._log("  Очистка старых уникальных прав...\n")
        try:
            client.remove_unique_permissions(item_id)
        except Exception as exc:
            self._log(f"  WARNING: наследование не восстановлено или уже наследуется: {exc}\n")
        time.sleep(0.25)
        client.set_unique_permissions(item_id, copy_permission=True)
        time.sleep(0.25)

        for _attempt in range(5):
            permission_items = client.get_children(self.scope_list_id)
            to_delete: List[str] = []
            for item in permission_items:
                fmap = item.get("fieldValueMap") or {}
                source = field_value_id(fmap.get("source")) or clean_text(fmap.get("source"))
                if source == item_id:
                    item_id_to_delete = clean_text(item.get("id"))
                    if item_id_to_delete:
                        to_delete.append(item_id_to_delete)
            if not to_delete:
                break
            client.delete_items(to_delete)
            stats["deleted"] += len(to_delete)
            self._log(f"  удалено записей: {len(to_delete)}\n")
            time.sleep(0.35)


# -----------------------------------------------------------------------------
# UI widgets
# -----------------------------------------------------------------------------
class ModernSection(QtWidgets.QWidget):
    def __init__(self, icon: str, title: str, collapsible: bool = True, parent=None):
        super().__init__(parent)
        self._collapsible = collapsible
        self._expanded = True
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.card = QtWidgets.QFrame(self)
        self.card.setObjectName("card")
        card_layout = QtWidgets.QVBoxLayout(self.card)
        card_layout.setContentsMargins(0, 0, 0, 0)
        card_layout.setSpacing(0)
        self.header = QtWidgets.QFrame(self.card)
        self.header.setObjectName("sectionHeader")
        header_layout = QtWidgets.QHBoxLayout(self.header)
        header_layout.setContentsMargins(12, 8, 12, 6)
        header_layout.setSpacing(8)
        self.icon_label = QtWidgets.QLabel(icon)
        self.icon_label.setObjectName("sectionIcon")
        self.title_label = QtWidgets.QLabel(title)
        self.title_label.setObjectName("sectionTitle")
        self.arrow_label = QtWidgets.QLabel("⌄" if collapsible else "")
        self.arrow_label.setObjectName("sectionArrow")
        header_layout.addWidget(self.icon_label)
        header_layout.addWidget(self.title_label)
        header_layout.addStretch(1)
        header_layout.addWidget(self.arrow_label)
        self.body = QtWidgets.QFrame(self.card)
        self.body.setObjectName("sectionBody")
        self.grid = QtWidgets.QGridLayout(self.body)
        self.grid.setContentsMargins(12, 0, 12, 10)
        self.grid.setHorizontalSpacing(10)
        self.grid.setVerticalSpacing(7)
        card_layout.addWidget(self.header)
        card_layout.addWidget(self.body)
        layout.addWidget(self.card)
        if collapsible:
            self.header.setCursor(QtCore.Qt.PointingHandCursor)
            self.header.mousePressEvent = self._toggle_from_mouse

    def _toggle_from_mouse(self, event):
        self.set_expanded(not self._expanded)
        event.accept()

    def set_expanded(self, expanded: bool):
        if not self._collapsible:
            return
        self._expanded = expanded
        self.body.setVisible(expanded)
        self.arrow_label.setText("⌄" if expanded else "›")


class LogDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Log")
        self.resize(900, 560)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        self.text = QtWidgets.QTextEdit()
        self.text.setObjectName("logArea")
        self.text.setReadOnly(True)
        self.text.setFont(QtGui.QFont("Cascadia Code", 10))
        layout.addWidget(self.text, 1)

    def set_log_text(self, value: str):
        self.text.setPlainText(value)
        self.text.moveCursor(QtGui.QTextCursor.End)
        self.text.ensureCursorVisible()

    def append(self, value: str):
        self.text.moveCursor(QtGui.QTextCursor.End)
        self.text.insertPlainText(value)
        self.text.ensureCursorVisible()


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VitroCAD — выдача прав из Excel")
        self.resize(1180, 760)
        self.setMinimumSize(1040, 640)

        self.base_url = DEFAULT_SERVER
        self.auth_token = ""
        self.file_list_id = DEFAULT_FILE_LIST_ID
        self.scope_list_id = DEFAULT_SCOPE_LIST_ID
        self.permission_content_type_id = DEFAULT_USER_PERMISSION_CONTENT_TYPE_ID
        self.content_types: Dict[str, dict] = {}
        self.selected_folder_id = ""
        self.selected_folder_name = "Корень списка"
        self.current_plan: List[PlanRow] = []
        self.schema = PermissionSchema()
        self.active_threads: List[QtCore.QThread] = []
        self.active_workers: List[QtCore.QObject] = []
        self._log_lines: List[str] = []
        self.log_dialog: Optional[LogDialog] = None
        self._apply_after_preview = False

        self._setup_ui()
        self._apply_stylesheet()

    def _setup_ui(self):
        self.setFont(QtGui.QFont("Segoe UI", 9))
        central = QtWidgets.QWidget()
        central.setObjectName("page")
        self.setCentralWidget(central)
        root = QtWidgets.QVBoxLayout(central)
        root.setContentsMargins(18, 14, 18, 0)
        root.setSpacing(8)

        title = QtWidgets.QLabel("VitroCAD — выдача прав из Excel")
        title.setObjectName("mainTitle")
        subtitle = QtWidgets.QLabel("Выберите сайт и Excel. Папки будут найдены по путям из таблицы, начиная от корня списка.")
        subtitle.setObjectName("subtitle")
        root.addWidget(title)
        root.addWidget(subtitle)

        # Авторизация
        sec_auth = ModernSection("■", "Авторизация")
        auth = sec_auth.grid
        self.cmb_server = QtWidgets.QComboBox()
        self.cmb_server.setEditable(True)
        self.cmb_server.addItems([DEFAULT_SERVER, "https://192.168.12.2"])
        self.cmb_server.setCurrentText(DEFAULT_SERVER)
        self.ed_login = QtWidgets.QLineEdit()
        self.ed_login.setPlaceholderText("Введите логин")
        self.ed_password = QtWidgets.QLineEdit()
        self.ed_password.setPlaceholderText("Введите пароль")
        self.ed_password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.btn_toggle_password = QtWidgets.QToolButton()
        self.btn_toggle_password.setText("Показать")
        self.btn_toggle_password.setObjectName("eyeButton")
        self.btn_login = QtWidgets.QPushButton("Войти")
        self.btn_login.setObjectName("primaryButton")
        self.lbl_auth = QtWidgets.QLabel("Не авторизован")
        self.lbl_auth.setObjectName("badStatus")

        auth.addWidget(self._field_label("Сайт"), 0, 0)
        auth.addWidget(self.cmb_server, 0, 1, 1, 5)
        auth.addWidget(self._field_label("Логин"), 1, 0)
        auth.addWidget(self.ed_login, 1, 1, 1, 2)
        auth.addWidget(self._field_label("Пароль"), 1, 3)
        password_box = QtWidgets.QHBoxLayout()
        password_widget = QtWidgets.QWidget()
        password_box.setContentsMargins(0, 0, 0, 0)
        password_box.addWidget(self.ed_password, 1)
        password_box.addWidget(self.btn_toggle_password)
        password_widget.setLayout(password_box)
        auth.addWidget(password_widget, 1, 4, 1, 2)
        auth.addWidget(self.btn_login, 2, 1, 1, 2)
        auth.addWidget(self.lbl_auth, 2, 3, 1, 3)
        auth.setColumnStretch(1, 1)
        auth.setColumnStretch(2, 1)
        auth.setColumnStretch(4, 1)
        auth.setColumnStretch(5, 1)
        root.addWidget(sec_auth)

        # Основные настройки — без системных ID и технических флагов.
        sec_settings = ModernSection("■", "Настройка")
        settings = sec_settings.grid
        self.ed_excel = QtWidgets.QLineEdit()
        self.ed_excel.setPlaceholderText("Выберите Excel-файл с матрицей прав")
        self.btn_excel = QtWidgets.QPushButton("Обзор...")
        self.btn_excel.setObjectName("secondaryButton")
        # Старый тестовый режим убран из интерфейса: кнопка «Проверить» теперь
        # отвечает за безопасную проверку, а «Применить права» — за реальное применение.
        # Переменная оставлена только для совместимости с кодом ApplyWorker.
        self.cb_dry_run = QtWidgets.QCheckBox("Сначала только проверить, ничего не менять")
        self.cb_dry_run.setChecked(False)
        self.cb_dry_run.setVisible(False)
        self.lbl_search_scope = QtWidgets.QLabel("Можно сразу нажать «Применить права»: сначала будет автоматическая проверка, потом одно подтверждение")
        self.lbl_search_scope.setObjectName("hintLabel")

        # Технические настройки оставлены, но спрятаны в отдельной свернутой секции.
        self.ed_file_list_id = QtWidgets.QLineEdit(DEFAULT_FILE_LIST_ID)
        self.ed_scope_list_id = QtWidgets.QLineEdit(DEFAULT_SCOPE_LIST_ID)
        self.ed_permission_ct_id = QtWidgets.QLineEdit(DEFAULT_USER_PERMISSION_CONTENT_TYPE_ID)
        self.sp_max_depth = QtWidgets.QSpinBox()
        self.sp_max_depth.setRange(1, 200)
        self.sp_max_depth.setValue(50)
        self.sp_max_items = QtWidgets.QSpinBox()
        self.sp_max_items.setRange(0, 1000000)
        self.sp_max_items.setValue(0)
        self.sp_limit = QtWidgets.QSpinBox()
        self.sp_limit.setRange(0, 1000000)
        self.sp_limit.setValue(0)
        self.cb_include_target = QtWidgets.QCheckBox("Включать выбранную папку в путь")
        self.cb_include_target.setChecked(True)
        self.cb_suffix = QtWidgets.QCheckBox("Разрешить поиск по окончанию пути")
        self.cb_suffix.setChecked(False)
        self.cb_no_access = QtWidgets.QCheckBox("Создавать явное 'Без доступа'")
        self.cb_no_access.setChecked(False)

        settings.addWidget(self._field_label("Excel-файл"), 0, 0)
        settings.addWidget(self.ed_excel, 0, 1, 1, 5)
        settings.addWidget(self.btn_excel, 0, 6)
        settings.addWidget(self.lbl_search_scope, 1, 1, 1, 6)
        settings.setColumnStretch(1, 1)
        settings.setColumnStretch(2, 1)
        settings.setColumnStretch(3, 1)
        settings.setColumnStretch(4, 1)
        settings.setColumnStretch(5, 1)
        root.addWidget(sec_settings)

        sec_advanced = ModernSection("■", "Расширенные настройки")
        advanced = sec_advanced.grid
        advanced.addWidget(self._field_label("ID списка файлов"), 0, 0)
        advanced.addWidget(self.ed_file_list_id, 0, 1, 1, 2)
        advanced.addWidget(self._field_label("ID списка прав"), 0, 3)
        advanced.addWidget(self.ed_scope_list_id, 0, 4, 1, 3)
        advanced.addWidget(self._field_label("ID типа права"), 1, 0)
        advanced.addWidget(self.ed_permission_ct_id, 1, 1, 1, 2)
        advanced.addWidget(self._field_label("Глубина поиска"), 1, 3)
        advanced.addWidget(self.sp_max_depth, 1, 4)
        advanced.addWidget(self._field_label("Лимит папок"), 1, 5)
        advanced.addWidget(self.sp_max_items, 1, 6)
        advanced.addWidget(self._field_label("Лимит строк Excel"), 2, 0)
        advanced.addWidget(self.sp_limit, 2, 1)
        advanced.addWidget(self.cb_include_target, 2, 2, 1, 2)
        advanced.addWidget(self.cb_suffix, 2, 4)
        advanced.addWidget(self.cb_no_access, 2, 5, 1, 2)
        advanced.setColumnStretch(1, 1)
        advanced.setColumnStretch(2, 1)
        advanced.setColumnStretch(4, 1)
        advanced.setColumnStretch(6, 1)
        sec_advanced.set_expanded(False)
        sec_advanced.setVisible(False)
        root.addWidget(sec_advanced)

        content = QtWidgets.QHBoxLayout()
        content.setSpacing(10)
        sec_tree = ModernSection("■", "Целевая папка / проект", collapsible=False)
        tree_grid = sec_tree.grid
        self.tree = QtWidgets.QTreeWidget()
        self.tree.setHeaderLabels(["Папка", "Тип"])
        self.tree.setColumnWidth(0, 320)
        self.tree.setUniformRowHeights(True)
        tree_grid.addWidget(self.tree, 0, 0)

        sec_plan = ModernSection("■", "Предпросмотр", collapsible=False)
        plan_grid = sec_plan.grid
        self.tbl_plan = QtWidgets.QTableWidget(0, 6)
        self.tbl_plan.setHorizontalHeaderLabels(["Статус", "Строка", "Папка в Excel", "Папка в VitroCAD", "Прав", "Комментарий"])
        self.tbl_plan.verticalHeader().setVisible(False)
        self.tbl_plan.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.tbl_plan.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tbl_plan.setWordWrap(False)
        header = self.tbl_plan.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QtWidgets.QHeaderView.Stretch)
        plan_grid.addWidget(self.tbl_plan, 0, 0)
        # Обычному пользователю выбирать целевую папку не нужно:
        # Excel уже содержит полный путь, поэтому по умолчанию ищем от корня списка.
        # Дерево оставлено в коде как запасной механизм, но скрыто из основного интерфейса.
        sec_tree.setVisible(False)
        content.addWidget(sec_tree, 0)
        content.addWidget(sec_plan, 1)
        root.addLayout(content, 1)

        footer = QtWidgets.QFrame()
        footer.setObjectName("footer")
        footer_layout = QtWidgets.QHBoxLayout(footer)
        footer_layout.setContentsMargins(0, 10, 0, 10)
        self.btn_log = QtWidgets.QPushButton("Лог")
        self.btn_log.setObjectName("logBtn")
        self.btn_preview = QtWidgets.QPushButton("Проверить")
        self.btn_preview.setObjectName("secondaryActionButton")
        self.btn_apply = QtWidgets.QPushButton("Применить права")
        self.btn_apply.setObjectName("primaryButton")
        self.btn_apply.setEnabled(True)
        self.btn_apply.setToolTip("Если предпросмотр ещё не построен, программа сначала выполнит проверку автоматически.")
        footer_layout.addStretch(1)
        footer_layout.addWidget(self.btn_log)
        footer_layout.addWidget(self.btn_preview)
        footer_layout.addWidget(self.btn_apply)
        root.addWidget(footer)
        self.statusBar().showMessage("Готово")

        self.btn_toggle_password.clicked.connect(self._toggle_password)
        self.btn_login.clicked.connect(self._login)
        self.btn_excel.clicked.connect(self._pick_excel)
        self.btn_log.clicked.connect(self._show_log)
        self.btn_preview.clicked.connect(self._preview)
        self.btn_apply.clicked.connect(self._apply_permissions)
        self.tree.itemExpanded.connect(self._load_children_for_item)
        self.tree.itemClicked.connect(self._on_tree_item_selected)
        self.ed_excel.textChanged.connect(self._invalidate_plan)
        self.ed_file_list_id.textChanged.connect(self._on_list_id_changed)
        self.ed_scope_list_id.textChanged.connect(self._invalidate_plan)
        self.ed_permission_ct_id.textChanged.connect(self._invalidate_plan)

    def _field_label(self, text: str) -> QtWidgets.QLabel:
        label = QtWidgets.QLabel(text)
        label.setObjectName("fieldLabel")
        return label

    def _apply_stylesheet(self):
        self.setStyleSheet(
            """
            QMainWindow { background: #ffffff; }
            QWidget#page { background: #ffffff; color: #0f172a; }
            QLabel#mainTitle { font-size: 18px; font-weight: 700; color: #0b1220; }
            QLabel#subtitle { font-size: 12px; color: #7b879f; padding-bottom: 2px; }
            QFrame#card { background: #ffffff; border: 1px solid #dbe3ef; border-radius: 8px; }
            QFrame#sectionHeader, QFrame#sectionBody { background: transparent; border: 0; }
            QLabel#sectionIcon { color: #1277f3; font-size: 13px; font-weight: 600; min-width: 16px; }
            QLabel#sectionTitle { color: #111827; font-size: 13px; font-weight: 600; }
            QLabel#sectionArrow { color: #475569; font-size: 15px; font-weight: 600; }
            QLabel#fieldLabel { color: #111827; font-size: 13px; padding-right: 4px; }
            QLabel#hintLabel { color: #667085; font-size: 12px; }
            QLineEdit, QComboBox, QSpinBox {
                min-height: 32px; border: 1px solid #dbe3ef; border-radius: 6px;
                background: #ffffff; color: #111827; padding: 0 10px; font-size: 13px;
            }
            QLineEdit:focus, QComboBox:focus, QSpinBox:focus { border: 1px solid #93bdf8; }
            QToolButton#eyeButton { border: 0; background: transparent; color: #667085; font-size: 12px; padding: 0 6px; }
            QCheckBox { spacing: 8px; font-size: 13px; color: #1f2937; }
            QPushButton {
                min-height: 32px; border: 1px solid transparent; border-radius: 8px;
                padding: 0 18px; font-size: 13px; font-weight: 600;
                background: #eef2f7; color: #111827;
            }
            QPushButton:hover { background: #e5ebf4; }
            QPushButton:disabled { background: #e9edf3; color: #9aa6b8; border: 1px solid #e1e7f0; }
            QPushButton#primaryButton { background: #0f72ed; color: #ffffff; border: 1px solid #0f72ed; min-width: 170px; }
            QPushButton#primaryButton:hover { background: #0c66d8; border: 1px solid #0c66d8; }
            QPushButton#secondaryButton { background: #ffffff; border: 1px solid #dbe3ef; color: #1f2937; min-width: 110px; }
            QPushButton#secondaryButton:hover { background: #f8fafc; border: 1px solid #cbd7e6; }
            QPushButton#secondaryActionButton { background: #ffffff; border: 1px solid #dbe3ef; color: #1f2937; min-width: 220px; }
            QPushButton#logBtn { background: transparent; color: #667085; border: 1px solid #d0d5dd; min-width: 70px; padding: 0 14px; font-size: 12px; }
            QLabel#goodStatus { color: #15803d; font-weight: 650; font-size: 13px; }
            QLabel#badStatus { color: #ff0000; font-weight: 520; font-size: 13px; }
            QTreeWidget, QTableWidget {
                background: #ffffff; border: 1px solid #dbe3ef; border-radius: 8px;
                gridline-color: #e6edf5; color: #111827; selection-background-color: #eaf3ff; selection-color: #0b1220;
            }
            QTextEdit#logArea {
                border: 1px solid #d0d5dd; border-radius: 8px; background-color: #1e293b;
                color: #e2e8f0; padding: 10px; font-size: 12px;
            }
            QTreeWidget::item, QTableWidget::item { min-height: 24px; padding: 3px; border: 0; }
            QHeaderView::section {
                background: #ffffff; color: #111827; padding: 8px 8px; border: 0;
                border-right: 1px solid #e6edf5; border-bottom: 1px solid #e6edf5;
                font-weight: 520; font-size: 13px;
            }
            QFrame#footer { background: #ffffff; border-top: 1px solid #dbe3ef; }
            QStatusBar { background: #ffffff; color: #263244; border-top: 1px solid #dbe3ef; padding-left: 12px; font-size: 11px; min-height: 24px; }
            QStatusBar::item { border: 0; }
            QDialog { background: #ffffff; }
            """
        )

    # ------------------------------------------------------------------
    # Общие методы UI
    # ------------------------------------------------------------------
    def _toggle_password(self):
        if self.ed_password.echoMode() == QtWidgets.QLineEdit.Password:
            self.ed_password.setEchoMode(QtWidgets.QLineEdit.Normal)
            self.btn_toggle_password.setText("Скрыть")
        else:
            self.ed_password.setEchoMode(QtWidgets.QLineEdit.Password)
            self.btn_toggle_password.setText("Показать")

    def _log(self, text: str):
        chunk = str(text)
        if chunk and not chunk.endswith("\n"):
            chunk += "\n"
        self._log_lines.append(chunk)
        if self.log_dialog and self.log_dialog.isVisible():
            self.log_dialog.append(chunk)

    def _show_log(self):
        if self.log_dialog is None:
            self.log_dialog = LogDialog(self)
        self.log_dialog.set_log_text("".join(self._log_lines) if self._log_lines else "Лог пока пуст.\n")
        self.log_dialog.show()
        self.log_dialog.raise_()
        self.log_dialog.activateWindow()

    def _ensure_log_visible(self):
        if self.log_dialog is None:
            self.log_dialog = LogDialog(self)
        if not self.log_dialog.isVisible():
            self.log_dialog.set_log_text("".join(self._log_lines) if self._log_lines else "Лог пока пуст.\n")
            self.log_dialog.show()
        self.log_dialog.raise_()
        self.log_dialog.activateWindow()

    def _show_error(self, title: str, details: str):
        self._log(f"ERROR: {title}\n{details}\n")
        QtWidgets.QMessageBox.critical(self, title, details[:4000])

    def _set_busy(self, busy: bool, message: str = ""):
        widgets = [
            self.btn_login,
            self.btn_preview,
            self.btn_apply,
            self.btn_excel,
            self.cmb_server,
            self.ed_login,
            self.ed_password,
            self.ed_file_list_id,
            self.ed_scope_list_id,
            self.ed_permission_ct_id,
            self.sp_max_depth,
            self.sp_max_items,
            self.sp_limit,
            self.cb_dry_run,
            self.cb_include_target,
            self.cb_suffix,
            self.cb_no_access,
            self.tree,
        ]
        for widget in widgets:
            widget.setEnabled(not busy)
        if not busy:
            # Кнопка должна быть доступна сразу: если плана ещё нет,
            # обработчик сам построит предпросмотр и продолжит применение.
            self.btn_apply.setEnabled(True)
        if message:
            self.statusBar().showMessage(message)
        if busy:
            if QtWidgets.QApplication.overrideCursor() is None:
                QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        else:
            while QtWidgets.QApplication.overrideCursor() is not None:
                QtWidgets.QApplication.restoreOverrideCursor()

    def _run_worker(self, worker: QtCore.QObject):
        thread = QtCore.QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        # Любой рабочий сигнал завершения должен останавливать QThread.
        # Иначе поток останется жить после worker.run(), а cleanup не выполнится.
        for signal_name in ("done", "failed", "loaded"):
            signal = getattr(worker, signal_name, None)
            if signal is not None:
                try:
                    signal.connect(thread.quit)
                except Exception:
                    pass
        thread.finished.connect(lambda: self._cleanup_thread(thread, worker))
        self.active_threads.append(thread)
        self.active_workers.append(worker)
        thread.start()

    def _cleanup_thread(self, thread: QtCore.QThread, worker: QtCore.QObject):
        worker.deleteLater()
        thread.deleteLater()
        try:
            self.active_threads.remove(thread)
        except ValueError:
            pass
        try:
            self.active_workers.remove(worker)
        except ValueError:
            pass

    def _invalidate_plan(self):
        self.current_plan = []
        self.schema = PermissionSchema()
        self.tbl_plan.setRowCount(0)
        self.btn_apply.setEnabled(True)

    def _on_list_id_changed(self):
        self.file_list_id = clean_text(self.ed_file_list_id.text()) or DEFAULT_FILE_LIST_ID
        self._invalidate_plan()
        if self.auth_token:
            self._load_root_tree()

    # ------------------------------------------------------------------
    # Авторизация и дерево
    # ------------------------------------------------------------------
    def _login(self):
        server = clean_text(self.cmb_server.currentText()).rstrip("/")
        login = clean_text(self.ed_login.text())
        password = self.ed_password.text()
        if not server or not login or not password:
            QtWidgets.QMessageBox.warning(self, "Не хватает данных", "Заполни сайт/сервер, логин и пароль.")
            return
        self._set_busy(True, "Авторизация...")
        self._log(f"Авторизация: {server}\n")
        worker = LoginWorker(server, login, password)
        worker.done.connect(lambda token: self._on_login_done(server, token))
        worker.failed.connect(self._on_login_failed)
        self._run_worker(worker)

    def _on_login_done(self, server: str, token: str):
        self.base_url = server
        self.auth_token = token
        self.file_list_id = clean_text(self.ed_file_list_id.text()) or DEFAULT_FILE_LIST_ID
        self.scope_list_id = clean_text(self.ed_scope_list_id.text()) or DEFAULT_SCOPE_LIST_ID
        self.permission_content_type_id = clean_text(self.ed_permission_ct_id.text()) or DEFAULT_USER_PERMISSION_CONTENT_TYPE_ID
        self.lbl_auth.setText("Авторизован")
        self.lbl_auth.setObjectName("goodStatus")
        self.lbl_auth.style().unpolish(self.lbl_auth)
        self.lbl_auth.style().polish(self.lbl_auth)
        self._log("Авторизация успешно завершена.\n")
        try:
            client = VitroClient(self.base_url, self.auth_token)
            content_types = client.get_content_types(self.file_list_id)
            self.content_types = {clean_text(ct.get("id")): ct for ct in content_types if clean_text(ct.get("id"))}
            self._log(f"Загружены типы списка файлов: {len(self.content_types)}\n")
        except Exception as exc:
            self._log(f"WARNING: типы списка файлов не загружены: {exc}\n")
            self.content_types = {}
        self._load_root_tree()
        self._set_busy(False, "Готово")

    def _on_login_failed(self, error: str):
        self.auth_token = ""
        self.lbl_auth.setText("Ошибка авторизации")
        self.lbl_auth.setObjectName("badStatus")
        self.lbl_auth.style().unpolish(self.lbl_auth)
        self.lbl_auth.style().polish(self.lbl_auth)
        self._set_busy(False, "Ошибка")
        if "Missing dependencies for SOCKS support" in error:
            user_message = (
                "Python пытается подключиться через системный SOCKS-прокси. "
                "В этой версии программы системный прокси отключен для запросов к VitroCAD. "
                "Закрой программу и запусти обновленный файл.\n\n"
                "Если ошибка осталась, установи поддержку SOCKS командой:\n"
                "pip install PySocks"
            )
            self._log(f"ERROR: Ошибка авторизации\n{error}\n")
            QtWidgets.QMessageBox.critical(self, "Ошибка авторизации", user_message)
        else:
            self._show_error("Ошибка авторизации", error)

    def _load_root_tree(self):
        self.tree.clear()
        self.selected_folder_id = self.file_list_id
        self.selected_folder_name = "Корень списка"
        root = QtWidgets.QTreeWidgetItem(["Корень списка", "Список"])
        root.setData(0, QtCore.Qt.UserRole, self.file_list_id)
        root.setData(0, QtCore.Qt.UserRole + 1, "Корень списка")
        root.addChild(QtWidgets.QTreeWidgetItem(["...", ""]))
        self.tree.addTopLevelItem(root)
        root.setExpanded(True)
        self._load_children_for_item(root)

    def _load_children_for_item(self, item: QtWidgets.QTreeWidgetItem):
        if not self.auth_token:
            return
        if item.childCount() == 1 and item.child(0).text(0) == "...":
            item.takeChild(0)
        else:
            return
        parent_id = clean_text(item.data(0, QtCore.Qt.UserRole))
        if not parent_id:
            return
        worker = FolderTreeWorker(self.base_url, self.auth_token, parent_id, self.content_types)
        worker.loaded.connect(lambda folders, parent=item: self._on_tree_loaded(parent, folders))
        worker.failed.connect(lambda err: self._show_error("Ошибка загрузки дерева", err))
        worker.loaded.connect(lambda _: self.statusBar().showMessage("Готово"))
        worker.failed.connect(lambda _: self.statusBar().showMessage("Ошибка"))
        self.statusBar().showMessage("Загрузка дерева...")
        self._run_worker(worker)

    def _on_tree_loaded(self, parent_item: QtWidgets.QTreeWidgetItem, folders: List[dict]):
        for raw in folders:
            name = get_item_name(raw)
            ct_id = content_type_id_from_item(raw)
            ct = self.content_types.get(ct_id, {})
            type_name = clean_text(ct.get("name") or ct.get("description") or ct_id)
            item = QtWidgets.QTreeWidgetItem([name, type_name])
            item.setData(0, QtCore.Qt.UserRole, clean_text(raw.get("id")))
            item.setData(0, QtCore.Qt.UserRole + 1, name)
            item.addChild(QtWidgets.QTreeWidgetItem(["...", ""]))
            parent_item.addChild(item)
        parent_item.setExpanded(True)
        self.tree.resizeColumnToContents(1)
        self._log(f"Загружено элементов дерева: {len(folders)}\n")

    def _on_tree_item_selected(self, item: QtWidgets.QTreeWidgetItem, _column: int):
        self.selected_folder_id = clean_text(item.data(0, QtCore.Qt.UserRole))
        self.selected_folder_name = clean_text(item.data(0, QtCore.Qt.UserRole + 1)) or item.text(0)
        self._invalidate_plan()
        self._log(f"Выбрана целевая папка: {item.text(0)} / {self.selected_folder_id}\n")

    # ------------------------------------------------------------------
    # Preview / Apply
    # ------------------------------------------------------------------
    def _pick_excel(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Выберите Excel-файл", "", "Excel (*.xlsx *.xlsm)")
        if path:
            self.ed_excel.setText(path)

    def _validate_common(self) -> bool:
        if not self.auth_token:
            QtWidgets.QMessageBox.warning(self, "Нет авторизации", "Сначала войди в VitroCAD.")
            return False
        if not self.selected_folder_id:
            # По умолчанию целевая область — корень списка файлов.
            self.selected_folder_id = self.file_list_id
            self.selected_folder_name = "Корень списка"
        excel_path = clean_text(self.ed_excel.text())
        if not excel_path or not os.path.exists(excel_path):
            QtWidgets.QMessageBox.warning(self, "Excel не найден", "Выбери существующий Excel-файл.")
            return False
        return True

    def _make_preview_worker(self) -> PreviewWorker:
        self.file_list_id = clean_text(self.ed_file_list_id.text()) or DEFAULT_FILE_LIST_ID
        self.scope_list_id = clean_text(self.ed_scope_list_id.text()) or DEFAULT_SCOPE_LIST_ID
        self.permission_content_type_id = clean_text(self.ed_permission_ct_id.text()) or DEFAULT_USER_PERMISSION_CONTENT_TYPE_ID
        return PreviewWorker(
            server=self.base_url,
            token=self.auth_token,
            excel_path=clean_text(self.ed_excel.text()),
            target_id=self.selected_folder_id,
            target_name=self.selected_folder_name,
            include_target_in_path=self.cb_include_target.isChecked(),
            file_list_id=self.file_list_id,
            scope_list_id=self.scope_list_id,
            permission_content_type_id=self.permission_content_type_id,
            max_depth=int(self.sp_max_depth.value()),
            max_items=int(self.sp_max_items.value()),
            limit=int(self.sp_limit.value()),
            allow_suffix=self.cb_suffix.isChecked(),
            apply_no_access=self.cb_no_access.isChecked(),
        )

    def _start_preview(self, apply_after_preview: bool = False):
        if not self._validate_common():
            return
        self._apply_after_preview = apply_after_preview
        self._invalidate_plan()
        self._ensure_log_visible()
        status_text = "Проверка перед выдачей прав..." if apply_after_preview else "Построение предпросмотра..."
        self._set_busy(True, status_text)
        worker = self._make_preview_worker()
        worker.log.connect(self._log)
        worker.done.connect(self._on_preview_done)
        worker.failed.connect(self._on_preview_failed)
        self._run_worker(worker)

    def _preview(self):
        self._start_preview(apply_after_preview=False)

    def _on_preview_done(self, payload: dict):
        self.current_plan = payload["plan"]
        self.schema = payload["schema"]
        self._fill_plan_table(self.current_plan)
        can_apply = bool(self.current_plan) and all(row.can_apply for row in self.current_plan)
        self._set_busy(False, "Предпросмотр готов")

        if not can_apply:
            self._log("Предпросмотр содержит ошибки. Права не применены. Исправь пути/группы/уровни и запусти снова.\n")
            if self._apply_after_preview:
                self._apply_after_preview = False
                QtWidgets.QMessageBox.warning(
                    self,
                    "Права не применены",
                    "Автоматическая проверка нашла ошибки. Права не применены. Смотри таблицу предпросмотра и лог.",
                )
            return

        if self._apply_after_preview:
            self._apply_after_preview = False
            self._log("Проверка успешна. Начинаю выдачу прав без дополнительного подтверждения.\n")
            self._start_apply_worker()

    def _on_preview_failed(self, error: str):
        self.current_plan = []
        self._apply_after_preview = False
        self._fill_plan_table([])
        self._set_busy(False, "Ошибка")
        self._show_error("Ошибка предпросмотра", error)

    def _fill_plan_table(self, plan: Sequence[PlanRow]):
        self.tbl_plan.setRowCount(len(plan))
        for row_idx, row in enumerate(plan):
            folder_path = row.folder.path_text if row.folder else ""
            values = [
                row.status,
                str(row.entry.row_number),
                row.entry.path_text,
                folder_path,
                str(row.permission_count),
                row.message,
            ]
            for col_idx, value in enumerate(values):
                item = QtWidgets.QTableWidgetItem(value)
                item.setToolTip(value)
                if row.can_apply:
                    item.setBackground(QtGui.QColor("#e7f5e7"))
                elif row.status == "Папка не найдена":
                    item.setBackground(QtGui.QColor("#fff7d6"))
                else:
                    item.setBackground(QtGui.QColor("#ffe4e6"))
                self.tbl_plan.setItem(row_idx, col_idx, item)
        self.tbl_plan.resizeRowsToContents()

    def _confirm_apply(self, checked_rows: Optional[int] = None) -> bool:
        if checked_rows is None:
            rows_text = "Количество папок будет определено после автоматической проверки."
            action_text = (
                "Программа сначала проверит Excel, папки, группы и уровни доступа. "
                "Если ошибок нет, права на найденных папках будут заменены по Excel."
            )
        else:
            rows_text = f"Будет изменено папок: {checked_rows}."
            action_text = "Будет использован уже построенный предпросмотр. Права на найденных папках будут заменены по Excel."
        confirm_text = (
            f"{rows_text}\n\n"
            f"{action_text}\n\n"
            "Точно применить права?"
        )
        answer = QtWidgets.QMessageBox.question(
            self,
            "Подтвердить применение прав",
            confirm_text,
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        )
        return answer == QtWidgets.QMessageBox.Yes

    def _start_apply_worker(self):
        self._ensure_log_visible()
        self._set_busy(True, "Выдача прав...")
        worker = ApplyWorker(
            server=self.base_url,
            token=self.auth_token,
            plan=self.current_plan,
            scope_list_id=self.scope_list_id,
            permission_content_type_id=self.schema.permission_content_type_id or self.permission_content_type_id,
            dry_run=False,
        )
        worker.log.connect(self._log)
        worker.done.connect(self._on_apply_done)
        worker.failed.connect(self._on_apply_failed)
        self._run_worker(worker)

    def _apply_permissions(self):
        if not self._validate_common():
            return

        # Режим одного запуска без подтверждений:
        # 1) если предпросмотр уже есть и он без ошибок — сразу применяем его;
        # 2) если предпросмотра нет — автоматически строим его и при отсутствии ошибок
        #    сразу запускаем выдачу прав.
        # Если проверка найдёт ошибки, права не применяются.
        if self.current_plan:
            if not all(row.can_apply for row in self.current_plan):
                QtWidgets.QMessageBox.warning(
                    self,
                    "Есть ошибки",
                    "В текущем предпросмотре есть ошибки. Права не применены. Исправь Excel/пути/группы/уровни и запусти снова.",
                )
                return
            self._log("Запуск выдачи прав по уже построенному предпросмотру без дополнительного подтверждения.\n")
            self._start_apply_worker()
            return

        self._log("Предпросмотр ещё не построен. Автоматически проверяю Excel и затем применю права, если ошибок нет.\n")
        self._start_preview(apply_after_preview=True)

    def _on_apply_done(self, stats: dict):
        self._set_busy(False, "Готово")
        QtWidgets.QMessageBox.information(
            self,
            "Готово",
            f"Обработано папок: {stats['folders']}\n"
            f"Удалено старых прав: {stats['deleted']}\n"
            f"Создано новых прав: {stats['created']}\n"
            f"Ошибок: {stats['failed']}\n"
            f"Пропущено: {stats['skipped']}",
        )

    def _on_apply_failed(self, error: str):
        self._set_busy(False, "Ошибка")
        self._show_error("Ошибка выдачи прав", error)


# -----------------------------------------------------------------------------
# Entry point
# -----------------------------------------------------------------------------
def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    if hasattr(app, "exec"):
        return app.exec()
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
