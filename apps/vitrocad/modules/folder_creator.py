# -*- coding: utf-8 -*-
"""
VitroCAD Folder Creator
----------------------
Утилита для массового создания папочной структуры VitroCAD из Excel.

Ожидаемый Excel:
- несколько колонок уровней вложенности;
- отдельная колонка "Тип папки";
- пример: A:C = уровни, D = тип.

Зависимости:
    pip install requests openpyxl PySide6

Если PySide6 недоступен, можно поставить PyQt5:
    pip install requests openpyxl PyQt5
"""

from __future__ import annotations

import json
import os
import re
import sys
import traceback
from html import unescape
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import requests
import urllib3
from openpyxl import load_workbook

# Qt bindings: сначала PySide6, затем PyQt5 как запасной вариант.
try:
    from PySide6 import QtCore, QtGui, QtWidgets

    Signal = QtCore.Signal
    Slot = QtCore.Slot
except ImportError:  # pragma: no cover - зависит от окружения пользователя
    from PyQt5 import QtCore, QtGui, QtWidgets

    Signal = QtCore.pyqtSignal
    Slot = QtCore.pyqtSlot


# -----------------------------------------------------------------------------
# Константы VitroCAD по данным HAR-запросов
# -----------------------------------------------------------------------------
LIST_ID = "966e62c5-a803-49a0-a1be-e680d130c481"

# Справочник пользователей/ответственных из HAR ручного создания.
# Если на другой инсталляции VitroCAD эти id отличаются, поле responsible можно
# ввести вручную как UUID в окне обязательных атрибутов.
RESPONSIBLE_LOOKUP_FIELD_ID = "d3e8f0fa-dcda-4fbf-b143-70d92f2a1141"
PRINCIPAL_LIST_ID = "e3a94bde-0ca9-456f-b338-4465d40389ee"

CONTENT_TYPE_IDS: Dict[str, str] = {
    "Папка": "09ad2c16-6047-4dfb-8274-6c4a5f1edbe5",
    "Проект": "540ff572-4a29-4d8e-8dfb-8eda093c583a",
    "Комплект": "28a48961-35e0-43ae-ad6b-8f159cfd7202",
    "Стадия": "64a15f4c-083e-420b-b54e-fc75d44f31b7",
    "Стадия проектирования": "6733bea3-00ec-463d-adb8-0e315e73df99",
}

TYPE_NAME_BY_ID = {value: key for key, value in CONTENT_TYPE_IDS.items()}
FOLDER_LIKE_TYPE_IDS = set(CONTENT_TYPE_IDS.values())

DEFAULT_SERVER = "https://vitrocad.bim-info.ru"
SCRIPT_PATCH_NOTE = "SSL verification disabled; item/update uses multipart/form-data"

# Если VitroCAD стоит на локальном сервере/по IP и использует самоподписанный
# или внутренний сертификат, requests падает с CERTIFICATE_VERIFY_FAILED.
# False отключает проверку SSL-сертификата для всех запросов этого скрипта.
# Для публичного сервера с нормальным сертификатом можно вернуть True.
VERIFY_SSL_CERTIFICATE = False
if not VERIFY_SSL_CERTIFICATE:
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# В прямом режиме никаких ожидаемых/зашитых сопоставлений нет.
# Тип из Excel должен совпасть с типом, который реально вернул VitroCAD.
DEFAULT_TYPE_MAPPINGS: Dict[str, str] = {}


def app_dir() -> str:
    """Отдельный каталог config для пользовательских настроек."""
    if getattr(sys, "frozen", False):
        root = os.path.dirname(sys.executable)
    else:
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    config_dir = os.path.join(root, "config")
    os.makedirs(config_dir, exist_ok=True)
    return config_dir


def type_mappings_path() -> str:
    return os.path.join(app_dir(), "vitrocad_type_mappings.json")


def load_type_mappings() -> Dict[str, str]:
    """Прямой режим: пользовательские сопоставления не используются.

    Раньше рядом с программой мог сохраняться vitrocad_type_mappings.json.
    В этой версии он намеренно игнорируется, чтобы старые правила не подменяли
    типы из Excel. Проверка идёт так: значение Excel -> реальные типы VitroCAD.
    """
    return {}


def save_type_mappings(mapping: Dict[str, str]) -> None:
    path = type_mappings_path()
    data = {clean_text(k): clean_text(v) for k, v in mapping.items() if clean_text(k) and clean_text(v)}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def content_type_id_from_item(item: dict) -> str:
    return clean_text(item.get("contentTypeId") or item.get("content_type_id") or item.get("id") or item.get("Id"))


def strip_html(value: object) -> str:
    """Убирает HTML из названий/описаний типов, если VitroCAD вернул их как '<p>...</p>'."""
    text = clean_text(value)
    if not text:
        return ""
    if "<" in text and ">" in text:
        text = re.sub(r"<[^>]+>", " ", text)
    return " ".join(unescape(text).split())


def is_truthy(value: object) -> bool:
    if value is True:
        return True
    if isinstance(value, (int, float)) and value == 1:
        return True
    return clean_text(value).casefold() in {"true", "1", "yes", "да", "истина"}


def is_falsey(value: object) -> bool:
    if value is False:
        return True
    if isinstance(value, (int, float)) and value == 0:
        return True
    return clean_text(value).casefold() in {"false", "0", "no", "нет", "ложь"}


def is_content_type_folder_like(item: dict) -> bool:
    """Папочный тип для создания структуры.

    На разных инсталляциях VitroCAD флаг isFolder может приходить boolean, строкой
    или отсутствовать. Поэтому исключаем явные файловые типы и принимаем все
    не-файловые типы списка как допустимые для структуры.
    """
    if is_truthy(item.get("isFile")):
        return False
    if is_truthy(item.get("isFolder")):
        return True
    if is_falsey(item.get("isFile")):
        return True
    return False


def content_type_names_from_item(item: dict) -> List[str]:
    """Возвращает все человекочитаемые варианты имени типа: name/title/description."""
    result: List[str] = []

    def add(value: object) -> None:
        text = strip_html(value)
        if text and name_key(text) not in {name_key(v) for v in result}:
            result.append(text)

    # name обычно системное/короткое имя; description часто то, что видит пользователь.
    for key in ("name", "Name", "title", "Title", "description", "Description"):
        add(item.get(key))
    field_map = item.get("fieldValueMap") or {}
    for key in ("name", "Name", "title", "Title", "description", "Description"):
        add(field_map.get(key))
    return result


def content_type_name_from_item(item: dict) -> str:
    names = content_type_names_from_item(item)
    return names[0] if names else content_type_id_from_item(item) or "Без названия"


def register_server_content_types(content_types: Sequence[dict]) -> None:
    """Обновляет глобальные справочники имен/ID после загрузки типов с сервера."""
    for item in content_types:
        type_id = clean_text(item.get("id") or item.get("Id"))
        names = content_type_names_from_item(item)
        if not type_id or not names:
            continue
        TYPE_NAME_BY_ID[type_id] = names[0]
        for type_name in names:
            CONTENT_TYPE_IDS.setdefault(type_name, type_id)
        if is_content_type_folder_like(item):
            FOLDER_LIKE_TYPE_IDS.add(type_id)


class ContentTypeResolver:
    """Сопоставляет пользовательские типы из Excel с реальными типами VitroCAD."""

    def __init__(self, server_content_types: Sequence[dict], user_mapping: Optional[Dict[str, str]] = None):
        self.server_content_types = list(server_content_types or [])
        register_server_content_types(self.server_content_types)
        # Прямой режим: никаких Excel -> VitroCAD сопоставлений.
        self.user_mapping = {}
        self.id_to_name: Dict[str, str] = {}
        self.name_to_id: Dict[str, str] = {}
        self.folder_type_names: List[str] = []

        for item in self.server_content_types:
            type_id = clean_text(item.get("id") or item.get("Id"))
            names = content_type_names_from_item(item)
            type_name = names[0] if names else ""
            if not type_id or not type_name:
                continue
            self.id_to_name[type_id] = type_name
            # Для создания структуры берём все не-файловые типы списка: на части серверов
            # кастомные типы не всегда помечены isFolder=True, хотя создаются как папочные.
            if is_content_type_folder_like(item):
                for alias in names:
                    self.name_to_id[name_key(alias)] = type_id
                self.folder_type_names.append(type_name)

        self.server_ids = set(self.id_to_name)

    def available_folder_types_text(self) -> str:
        if not self.folder_type_names:
            return "не удалось определить папочные типы"
        return ", ".join(sorted(set(self.folder_type_names), key=name_key))

    def resolve(self, excel_type: str, project_as_folder: bool = False) -> Tuple[str, str, str]:
        excel_type = clean_text(excel_type) or "Папка"
        # Прямой режим: берём ровно то название, которое пришло из Excel.
        # Ничего не подменяем через внутренний словарь.
        target = excel_type
        note = ""

        # Можно вводить либо имя типа, либо реальный content_type_id.
        if target in self.server_ids:
            return target, self.id_to_name.get(target, target), note

        target_key = name_key(target)
        if target_key in self.name_to_id:
            type_id = self.name_to_id[target_key]
            return type_id, self.id_to_name.get(type_id, target), note

        # Запасной режим для старых HAR-значений: сработает только если ID действительно есть на сервере.
        legacy_id = CONTENT_TYPE_IDS.get(target)
        if legacy_id and legacy_id in self.server_ids:
            return legacy_id, self.id_to_name.get(legacy_id, target), note

        raise ValueError(
            f"тип из Excel '{excel_type}' не найден среди папочных типов VitroCAD для выбранного списка. "
            f"Доступные папочные типы: {self.available_folder_types_text()}. "
            "Проверь написание в Excel или убедись, что программа получает правильный справочник типов с сервера."
        )


# -----------------------------------------------------------------------------
# Данные плана
# -----------------------------------------------------------------------------
@dataclass(frozen=True)
class StructureNode:
    path: Tuple[str, ...]
    original_type: str
    effective_type: str
    source_row: int
    note: str = ""

    @property
    def name(self) -> str:
        return self.path[-1]

    @property
    def parent_path(self) -> Tuple[str, ...]:
        return self.path[:-1]

    @property
    def depth(self) -> int:
        return len(self.path)

    @property
    def content_type_id(self) -> str:
        # Для новых пользовательских типов ID определяется позже через ContentTypeResolver.
        return CONTENT_TYPE_IDS.get(self.effective_type, "")

    @property
    def type_display(self) -> str:
        if self.original_type != self.effective_type:
            return f"{self.original_type} → {self.effective_type}"
        return self.effective_type


@dataclass
class PlanRow:
    path: Tuple[str, ...]
    original_type: str
    effective_type: str
    content_type_id: str
    source_row: int
    status: str
    message: str
    existing_id: Optional[str] = None
    parent_known: bool = True

    @property
    def path_text(self) -> str:
        return " / ".join(self.path)

    @property
    def parent_path(self) -> Tuple[str, ...]:
        return self.path[:-1]

    @property
    def name(self) -> str:
        return self.path[-1]

    @property
    def depth(self) -> int:
        return len(self.path)

    @property
    def type_display(self) -> str:
        if self.original_type != self.effective_type:
            return f"{self.original_type} → {self.effective_type}"
        return self.effective_type


# -----------------------------------------------------------------------------
# Утилиты
# -----------------------------------------------------------------------------
def clean_text(value: object) -> str:
    """Аккуратно приводит значение Excel/API к строке."""
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null"}:
        return ""
    return text


def name_key(name: str) -> str:
    """Ключ для безопасного сравнения имён без учёта регистра и лишних пробелов."""
    return " ".join(clean_text(name).split()).casefold()


def get_item_name(item: dict) -> str:
    """Извлекает имя элемента VitroCAD из разных возможных мест ответа API."""
    field_map = item.get("fieldValueMap") or {}
    for key in ("name", "Name", "Название", "title", "Title"):
        value = clean_text(field_map.get(key))
        if value:
            return value
    for key in ("name", "Name", "title", "Title"):
        value = clean_text(item.get(key))
        if value:
            return value
    return "Без названия"


def item_type_name(item: dict) -> str:
    content_type_id = item.get("contentTypeId") or item.get("content_type_id") or ""
    return TYPE_NAME_BY_ID.get(content_type_id, content_type_id or "Неизвестный тип")


def is_folder_like_item(item: dict) -> bool:
    content_type_id = item.get("contentTypeId") or item.get("content_type_id")
    return content_type_id in FOLDER_LIKE_TYPE_IDS


def find_child_by_name(children: Iterable[dict], name: str) -> Optional[dict]:
    target = name_key(name)
    for child in children:
        if name_key(get_item_name(child)) == target:
            return child
    return None


# -----------------------------------------------------------------------------
# Парсер Excel
# -----------------------------------------------------------------------------
class ExcelStructureParser:
    """Читает Excel со структурой: уровни вложенности + колонка 'Тип папки'."""

    TYPE_HEADER_ALIASES = {
        "тип папки",
        "тип",
        "content type",
        "content_type",
        "contenttype",
    }

    def __init__(self, excel_path: str, project_as_folder: bool = False):
        self.excel_path = excel_path
        self.project_as_folder = project_as_folder

    def parse(self) -> List[StructureNode]:
        if not os.path.exists(self.excel_path):
            raise FileNotFoundError(f"Excel-файл не найден: {self.excel_path}")

        workbook = load_workbook(self.excel_path, read_only=True, data_only=True)
        sheet = workbook.active

        type_col_idx, header_row_idx = self._find_type_column(sheet)
        level_col_indices = list(range(1, type_col_idx))
        if not level_col_indices:
            raise ValueError("Перед колонкой 'Тип папки' не найдено колонок уровней.")

        nodes_by_path: Dict[Tuple[str, ...], StructureNode] = {}
        order: Dict[Tuple[str, ...], int] = {}

        for row in sheet.iter_rows(min_row=header_row_idx + 1):
            row_idx = row[0].row
            levels: List[str] = []
            for col_idx in level_col_indices:
                levels.append(clean_text(sheet.cell(row=row_idx, column=col_idx).value))

            path = tuple(level for level in levels if level)
            raw_type = clean_text(sheet.cell(row=row_idx, column=type_col_idx).value)

            if not path and not raw_type:
                continue
            if not path:
                raise ValueError(f"Строка {row_idx}: указан тип, но не указан путь.")
            if not raw_type:
                raw_type = "Папка"

            original_type = self._normalize_type(raw_type, row_idx)
            effective_type = original_type
            note = ""

            node = StructureNode(
                path=path,
                original_type=original_type,
                effective_type=effective_type,
                source_row=row_idx,
                note=note,
            )

            if path in nodes_by_path:
                prev = nodes_by_path[path]
                if prev.original_type != node.original_type or prev.effective_type != node.effective_type:
                    raise ValueError(
                        f"Конфликт типов для пути '{' / '.join(path)}': "
                        f"строка {prev.source_row} = {prev.type_display}, "
                        f"строка {row_idx} = {node.type_display}."
                    )
                continue

            nodes_by_path[path] = node
            order[path] = len(order)

        # Добавляем отсутствующих родителей как обычные папки.
        original_paths = list(nodes_by_path.keys())
        for full_path in original_paths:
            for depth in range(1, len(full_path)):
                parent_path = full_path[:depth]
                if parent_path not in nodes_by_path:
                    nodes_by_path[parent_path] = StructureNode(
                        path=parent_path,
                        original_type="Папка",
                        effective_type="Папка",
                        source_row=0,
                        note="Родительского узла не было в Excel; добавлен как 'Папка'.",
                    )
                    order[parent_path] = len(order)

        nodes = list(nodes_by_path.values())
        nodes.sort(key=lambda n: (n.depth, order.get(n.path, 10**9), n.path))
        if not nodes:
            raise ValueError("В Excel не найдено ни одной строки структуры.")
        return nodes

    def _find_type_column(self, sheet) -> Tuple[int, int]:
        max_scan_rows = min(sheet.max_row, 30)
        for row_idx in range(1, max_scan_rows + 1):
            for col_idx in range(1, sheet.max_column + 1):
                value = clean_text(sheet.cell(row=row_idx, column=col_idx).value)
                normalized = value.casefold()
                if normalized in self.TYPE_HEADER_ALIASES or "тип пап" in normalized:
                    return col_idx, row_idx
        raise ValueError("Не найдена колонка 'Тип папки'. Проверь заголовок Excel.")

    @staticmethod
    def _normalize_type(raw_type: str, row_idx: int) -> str:
        # Здесь больше не проверяем тип по жёсткому словарю.
        # Проверка и сопоставление выполняются позже по реальным типам сервера VitroCAD
        # и пользовательской таблице "Типы папок...".
        value = clean_text(raw_type)
        if not value:
            return "Папка"
        for known_type in CONTENT_TYPE_IDS:
            if name_key(known_type) == name_key(value):
                return known_type
        return value


# -----------------------------------------------------------------------------
# Клиент VitroCAD API
# -----------------------------------------------------------------------------
class VitroClient:
    def __init__(self, base_url: str, token: Optional[str] = None, timeout: int = 20):
        self.base_url = base_url.strip().rstrip("/")
        self.token = token
        self.timeout = timeout
        self.session = requests.Session()
        # Не используем системные/переменные окружения proxy.
        # Иначе requests может подхватить SOCKS-прокси из HTTP_PROXY/HTTPS_PROXY/ALL_PROXY
        # и упасть с ошибкой: Missing dependencies for SOCKS support, если PySocks не установлен
        # или не попал в PyInstaller-сборку. Для VitroCAD здесь нужен прямой HTTPS-запрос.
        self.session.trust_env = False
        self.session.proxies.clear()
        self.session.verify = VERIFY_SSL_CERTIFICATE

    @property
    def headers(self) -> Dict[str, str]:
        if not self.token:
            return {}
        return {"Authorization": self.token}

    def login(self, login: str, password: str) -> str:
        response = self.session.post(
            f"{self.base_url}/api/security/login",
            json={"login": login, "password": password},
            timeout=self.timeout,
        )
        self._raise_for_response(response, "Авторизация не удалась")
        data = response.json()
        token = data.get("token")
        if not token:
            raise RuntimeError(f"В ответе авторизации нет token. Ответ: {data}")
        self.token = token
        return token

    def get_children(self, parent_id: str) -> List[dict]:
        response = self.session.post(
            f"{self.base_url}/api/item/getList/{parent_id}",
            headers=self.headers,
            json={"sort": {"name": 0}},
            timeout=self.timeout,
        )
        self._raise_for_response(response, f"Не удалось загрузить дочерние элементы {parent_id}")
        data = response.json()
        if not isinstance(data, list):
            raise RuntimeError(f"Ожидался список дочерних элементов, получено: {type(data).__name__}")
        return data

    def get_item(self, item_id: str) -> dict:
        """Загружает один элемент VitroCAD по id.

        Нужен, чтобы взять обязательные значения из родительского проекта:
        например responsible для создания элемента типа "Папка проекта/Проект".
        """
        response = self.session.post(
            f"{self.base_url}/api/item/get/{item_id}",
            headers=self.headers,
            timeout=self.timeout,
        )
        self._raise_for_response(response, f"Не удалось загрузить элемент {item_id}")
        data = response.json()
        if not isinstance(data, dict):
            raise RuntimeError(f"Ожидался объект элемента, получено: {type(data).__name__}: {data}")
        return data

    def get_content_types(self, list_id: str = LIST_ID) -> List[dict]:
        """Загружает актуальные типы контента указанного списка VitroCAD с сервера."""
        list_id = clean_text(list_id) or LIST_ID
        response = self.session.post(
            f"{self.base_url}/api/contentType/getByList/{list_id}",
            headers=self.headers,
            timeout=self.timeout,
        )
        self._raise_for_response(response, "Не удалось загрузить типы контента VitroCAD")
        data = response.json()
        if not isinstance(data, list):
            raise RuntimeError(f"Ожидался список типов контента, получено: {type(data).__name__}")
        return data

    @staticmethod
    def _field_value_for_update(value: object) -> object:
        """Преобразует значения из createNewInstance в формат, который ждёт /api/item/update.

        VitroCAD часто возвращает справочники объектом {"id": "...", ...}, а при сохранении
        фронт отправляет только id. Для обычных чисел/строк значение оставляем как есть.
        """
        if isinstance(value, dict):
            value_id = clean_text(value.get("id") or value.get("Id"))
            if value_id:
                return value_id
            return value
        if isinstance(value, list):
            result = []
            for item in value:
                if isinstance(item, dict):
                    result.append(clean_text(item.get("id") or item.get("Id")) or item)
                else:
                    result.append(item)
            return result
        return value

    @staticmethod
    def _lookup_id(value: object) -> str:
        """Достаёт id из справочного значения VitroCAD.

        Значение может быть строкой-id, объектом {id: ...} или списком объектов.
        """
        if isinstance(value, str):
            return clean_text(value)
        if isinstance(value, dict):
            return clean_text(value.get("id") or value.get("Id"))
        if isinstance(value, list) and value:
            return VitroClient._lookup_id(value[0])
        return ""

    def find_default_responsible_id(self, start_item_id: str) -> str:
        """Ищет responsible в выбранной папке или ближайшем родительском проекте.

        В ручном HAR для типа "Папка проекта/Проект" VitroCAD отправляет responsible.
        createNewInstance это поле не возвращает, поэтому берём его из выбранного
        проекта/родителя. Если не найдём, сервер вернёт понятную ошибку с payload.
        """
        checked = set()

        def responsible_from_item(item: dict) -> str:
            field_map = item.get("fieldValueMap") or {}
            return self._lookup_id(field_map.get("responsible"))

        def check_item(item_id: str) -> Tuple[str, List[str]]:
            item_id = clean_text(item_id)
            if not item_id or item_id in checked:
                return "", []
            checked.add(item_id)
            item = self.get_item(item_id)
            responsible_id = responsible_from_item(item)
            path = []
            item_path = item.get("itemPath") or {}
            if isinstance(item_path, dict):
                path = [clean_text(v) for v in item_path.get("path") or [] if clean_text(v)]
            return responsible_id, path

        responsible_id, path = check_item(start_item_id)
        if responsible_id:
            return responsible_id

        # В itemPath обычно есть цепочка: site/list/project/folder/...
        # Идём назад по цепочке и ищем первый элемент с fieldValueMap.responsible.
        for ancestor_id in reversed(path):
            if ancestor_id == start_item_id:
                continue
            try:
                responsible_id, _ = check_item(ancestor_id)
                if responsible_id:
                    return responsible_id
            except Exception:
                # Не все id из path являются item-ами; site/list могут не читаться через item/get.
                continue
        return ""

    def search_responsibles(self, query: str) -> List[dict]:
        """Ищет пользователей для поля responsible через тот же lookup, что использует веб-интерфейс.

        По HAR ручного создания VitroCAD вызывает /api/lookup/getList с list=e3a94...
        и field=d3e8... . Если API на сервере отличается, пользователь всё равно может
        вставить responsible как UUID вручную.
        """
        query = clean_text(query)
        if not query:
            return []
        payload = {
            "field": RESPONSIBLE_LOOKUP_FIELD_ID,
            "value_string": query,
            "list": PRINCIPAL_LIST_ID,
        }
        response = self.session.post(
            f"{self.base_url}/api/lookup/getList",
            headers=self.headers,
            json=payload,
            timeout=self.timeout,
        )
        self._raise_for_response(response, f"Не удалось найти ответственного '{query}'")
        data = response.json()
        if not isinstance(data, list):
            raise RuntimeError(f"Ожидался список ответственных, получено: {type(data).__name__}: {data}")
        return data

    def create_new_instance(self, parent_id: str, content_type_id: str, list_id: str = LIST_ID) -> dict:
        """Получает серверный шаблон нового элемента.

        Вручную VitroCAD делает именно так: сначала createNewInstance, потом item/update.
        Это важно для типов, у которых есть обязательные/дефолтные поля: project_status и т.п.
        """
        payload = {
            "list_id": clean_text(list_id) or LIST_ID,
            "content_type_id": content_type_id,
            "parent_id": parent_id,
        }
        response = self.session.post(
            f"{self.base_url}/api/item/createNewInstance",
            headers=self.headers,
            json=payload,
            timeout=self.timeout,
        )
        self._raise_for_response(response, f"Не удалось получить шаблон создания для '{content_type_id}'")
        data = response.json()
        if not isinstance(data, dict):
            raise RuntimeError(f"Ожидался объект шаблона создания, получено: {type(data).__name__}: {data}")
        return data

    def create_item(
        self,
        parent_id: str,
        name: str,
        content_type_id: str,
        list_id: str = LIST_ID,
        default_responsible_id: str = "",
    ) -> str:
        list_id = clean_text(list_id) or LIST_ID

        # Берём серверный шаблон нового элемента. Никаких заранее вшитых правил
        # по названию типа здесь нет: что сервер вернул, то и используем.
        template = self.create_new_instance(parent_id, content_type_id, list_id)
        template_fields = template.get("fieldValueMap") or {}
        if not isinstance(template_fields, dict):
            template_fields = {}

        item_data: Dict[str, object] = {}
        for key, value in template_fields.items():
            if value is None:
                continue
            item_data[key] = self._field_value_for_update(value)

        item_data.update(
            {
                "list_id": list_id,
                "content_type_id": content_type_id,
                "parent_id": parent_id,
                "fact_work_time1": item_data.get("fact_work_time1", 0),
                "plan_work_time1": item_data.get("plan_work_time1", 0),
                "name": name,
            }
        )

        # Минимальные автозаполнения делаются только по признакам из серверного
        # шаблона createNewInstance, а не по названию типа из Excel.
        server_template_keys = {clean_text(key).casefold() for key in template_fields.keys()}
        server_project_like = "project_status" in server_template_keys or "project_status" in item_data

        if server_project_like or "code" in server_template_keys:
            item_data.setdefault("code", name)

        if default_responsible_id and not self._lookup_id(item_data.get("responsible")):
            # responsible подставляем только если серверный шаблон сам содержит такое поле
            # либо если серверный шаблон выглядит как проектный.
            if "responsible" in server_template_keys or server_project_like:
                item_data["responsible"] = default_responsible_id

        # На части инсталляций VitroCAD /api/item/update ожидает именно
        # multipart/form-data. Если отправить обычный data={...}, сервер падает
        # с ошибкой "Multipart boundary not found". Через files requests сам
        # формирует корректный Content-Type с boundary.
        item_list_json = json.dumps([item_data], ensure_ascii=False)
        response = self.session.post(
            f"{self.base_url}/api/item/update",
            headers=self.headers,
            files={
                "itemListJson": (None, item_list_json),
            },
            timeout=self.timeout,
        )

        if not response.ok:
            text = response.text
            if len(text) > 1200:
                text = text[:1200] + "..."
            template_text = json.dumps(template, ensure_ascii=False)
            if len(template_text) > 2500:
                template_text = template_text[:2500] + "..."
            raise RuntimeError(
                f"Не удалось создать '{name}'. HTTP {response.status_code}: {text}"
                f"\nОтправленный payload: {item_list_json}"
                f"\nСерверный шаблон createNewInstance: {template_text}"
            )

        data = response.json()
        if not isinstance(data, list) or not data:
            raise RuntimeError(f"Неожиданный ответ создания '{name}': {data}")
        created_id = data[0].get("id")
        if not created_id:
            raise RuntimeError(f"В ответе создания '{name}' нет id: {data[0]}")
        return created_id

    @staticmethod
    def _raise_for_response(response: requests.Response, context: str) -> None:
        if response.ok:
            return
        text = response.text
        if len(text) > 1200:
            text = text[:1200] + "..."
        raise RuntimeError(f"{context}. HTTP {response.status_code}: {text}")


# -----------------------------------------------------------------------------
# Worker: загрузка дерева папок
# -----------------------------------------------------------------------------
class FolderTreeWorker(QtCore.QObject):
    loaded = Signal(object)
    failed = Signal(str)

    def __init__(self, base_url: str, token: str, parent_id: str):
        super().__init__()
        self.base_url = base_url
        self.token = token
        self.parent_id = parent_id

    @Slot()
    def run(self):
        try:
            client = VitroClient(self.base_url, self.token)
            children = client.get_children(self.parent_id)
            folders = [child for child in children if is_folder_like_item(child)]
            folders.sort(key=lambda item: name_key(get_item_name(item)))
            self.loaded.emit(folders)
        except Exception:
            self.failed.emit(traceback.format_exc())


# -----------------------------------------------------------------------------
# Worker: предпросмотр плана
# -----------------------------------------------------------------------------
class PlanWorker(QtCore.QObject):
    log = Signal(str)
    done = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        excel_path: str,
        base_url: str,
        token: str,
        target_id: str,
        target_list_id: str,
        project_as_folder: bool,
        type_mapping: Optional[Dict[str, str]] = None,
    ):
        super().__init__()
        self.excel_path = excel_path
        self.base_url = base_url
        self.token = token
        self.target_id = target_id
        self.target_list_id = clean_text(target_list_id) or LIST_ID
        self.project_as_folder = project_as_folder
        self.type_mapping = type_mapping or load_type_mappings()

    @Slot()
    def run(self):
        try:
            self.log.emit("=" * 60 + "\n")
            self.log.emit("Построение предпросмотра структуры VitroCAD\n")
            self.log.emit(f"Excel: {self.excel_path}\n")
            self.log.emit(f"TargetId: {self.target_id}\n")
            self.log.emit(f"ListId: {self.target_list_id}\n")
            self.log.emit("=" * 60 + "\n\n")

            self.log.emit("[1/4] Чтение Excel\n")
            parser = ExcelStructureParser(self.excel_path, self.project_as_folder)
            nodes = parser.parse()
            self.log.emit(f"  узлов структуры: {len(nodes)}\n\n")

            client = VitroClient(self.base_url, self.token)

            self.log.emit("[2/4] Проверка типов контента VitroCAD\n")
            server_content_types = client.get_content_types(self.target_list_id)
            resolver = ContentTypeResolver(server_content_types, self.type_mapping)
            resolved_nodes: List[Tuple[StructureNode, str, str, str]] = []
            type_errors: List[str] = []
            for node in nodes:
                try:
                    content_type_id, effective_type, resolve_note = resolver.resolve(node.original_type, self.project_as_folder)
                    notes = " ".join(part for part in (node.note, resolve_note) if part).strip()
                    resolved_nodes.append((node, content_type_id, effective_type, notes))
                except Exception as exc:
                    type_errors.append(f"строка {node.source_row or 'авто'}; путь '{' / '.join(node.path)}': {exc}")

            if type_errors:
                raise RuntimeError(
                    "Не удалось сопоставить типы из Excel с типами VitroCAD.\n"
                    + "\n".join(type_errors[:80])
                    + "\n\nВ этой версии сопоставлений нет: типы из Excel проверяются напрямую по типам VitroCAD."
                )

            self.log.emit(f"  папочных типов на сервере: {len(resolver.folder_type_names)}\n")
            excel_types = sorted({node.original_type for node in nodes}, key=name_key)
            self.log.emit(f"  уникальных типов в Excel: {len(excel_types)}\n")
            for excel_type in excel_types:
                self.log.emit(f"    Excel type: {excel_type}\n")
            self.log.emit("\n")

            self.log.emit("[3/4] Проверка существующих элементов\n")
            plan: List[PlanRow] = []
            known_real_ids: Dict[Tuple[str, ...], str] = {tuple(): self.target_id}

            for node, content_type_id, effective_type, node_note in resolved_nodes:
                parent_path = node.parent_path
                parent_id = known_real_ids.get(parent_path)

                if not parent_id:
                    message = "Родитель будет создан выше; существование этого элемента пока нельзя проверить."
                    if node_note:
                        message += " " + node_note
                    plan.append(
                        PlanRow(
                            path=node.path,
                            original_type=node.original_type,
                            effective_type=effective_type,
                            content_type_id=content_type_id,
                            source_row=node.source_row,
                            status="Создать",
                            message=message,
                            parent_known=False,
                        )
                    )
                    continue

                children = client.get_children(parent_id)
                existing = find_child_by_name(children, node.name)
                if existing:
                    existing_type_id = content_type_id_from_item(existing)
                    existing_type = resolver.id_to_name.get(existing_type_id) or item_type_name(existing)
                    message = f"Уже существует. Тип в VitroCAD: {existing_type}."
                    if existing_type_id and existing_type_id != content_type_id:
                        message += f" В Excel ожидается: {effective_type}. Будет использован существующий элемент."
                    if node_note:
                        message += " " + node_note
                    existing_id = existing.get("id")
                    if existing_id:
                        known_real_ids[node.path] = existing_id
                    plan.append(
                        PlanRow(
                            path=node.path,
                            original_type=node.original_type,
                            effective_type=effective_type,
                            content_type_id=content_type_id,
                            source_row=node.source_row,
                            status="Уже есть",
                            message=message,
                            existing_id=existing_id,
                        )
                    )
                else:
                    message = "Будет создано."
                    if node_note:
                        message += " " + node_note
                    plan.append(
                        PlanRow(
                            path=node.path,
                            original_type=node.original_type,
                            effective_type=effective_type,
                            content_type_id=content_type_id,
                            source_row=node.source_row,
                            status="Создать",
                            message=message,
                        )
                    )

            create_count = sum(1 for row in plan if row.status == "Создать")
            exists_count = sum(1 for row in plan if row.status == "Уже есть")
            self.log.emit("[4/4] Формирование плана\n")
            self.log.emit(f"  к созданию: {create_count}\n")
            self.log.emit(f"  уже существует: {exists_count}\n")
            self.log.emit(f"  всего строк плана: {len(plan)}\n\n")

            self.done.emit(plan)
        except Exception:
            self.failed.emit(traceback.format_exc())


# -----------------------------------------------------------------------------
# Worker: создание структуры
# -----------------------------------------------------------------------------
class CreateWorker(QtCore.QObject):
    log = Signal(str)
    done = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        plan: Sequence[PlanRow],
        base_url: str,
        token: str,
        target_id: str,
        target_list_id: str,
        default_responsible_id: str = "",
    ):
        super().__init__()
        self.plan = list(plan)
        self.base_url = base_url
        self.token = token
        self.target_id = target_id
        self.target_list_id = clean_text(target_list_id) or LIST_ID
        self.default_responsible_id = clean_text(default_responsible_id)

    @Slot()
    def run(self):
        try:
            client = VitroClient(self.base_url, self.token)
            default_responsible_id = self.default_responsible_id
            if default_responsible_id:
                self.log.emit(f"  responsible задан вручную: {default_responsible_id}\n")
            else:
                try:
                    default_responsible_id = client.find_default_responsible_id(self.target_id)
                    if default_responsible_id:
                        self.log.emit(
                            f"  responsible найден у родителя: {default_responsible_id}. "
                            "Будет использован только если серверный шаблон создания содержит поле responsible или project_status.\n"
                        )
                    else:
                        self.log.emit("  responsible у родителя не найден. Будет использован только серверный шаблон createNewInstance.\n")
                except Exception as exc:
                    self.log.emit(f"  WARNING: responsible у родителя определить не удалось: {exc}\n")
            runtime_ids: Dict[Tuple[str, ...], str] = {tuple(): self.target_id}
            created_count = 0
            existing_count = 0
            warning_count = 0

            self.log.emit("=" * 60 + "\n")
            self.log.emit("Создание структуры VitroCAD\n")
            self.log.emit(f"TargetId: {self.target_id}\n")
            self.log.emit(f"ListId: {self.target_list_id}\n")
            self.log.emit(f"Элементов в плане: {len(self.plan)}\n")
            self.log.emit("=" * 60 + "\n\n")
            self.log.emit("[1/2] Подготовка\n")
            self.log.emit("  порядок создания: сначала родители, затем дочерние элементы\n\n")
            self.log.emit("[2/2] Создание и повторная проверка существующих элементов\n")

            for row in sorted(self.plan, key=lambda item: (item.depth, item.source_row or 0, item.path)):
                parent_id = runtime_ids.get(row.parent_path)
                if not parent_id:
                    raise RuntimeError(
                        f"Не найден id родителя для '{row.path_text}'. "
                        f"Вероятно, выше не создался родительский элемент."
                    )

                children = client.get_children(parent_id)
                existing = find_child_by_name(children, row.name)
                if existing:
                    existing_id = existing.get("id")
                    if not existing_id:
                        raise RuntimeError(f"У найденного элемента '{row.path_text}' нет id.")
                    runtime_ids[row.path] = existing_id
                    existing_type = item_type_name(existing)
                    existing_count += 1
                    if existing_type != row.effective_type:
                        warning_count += 1
                        self.log.emit(
                            f"  [WARNING] Уже существует, но тип отличается: {row.path_text}\n"
                            f"            VitroCAD: {existing_type}; Excel: {row.effective_type}. Используется существующий id.\n"
                        )
                    else:
                        self.log.emit(f"  [EXISTS] {row.path_text}\n")
                    continue

                self.log.emit(
                    f"  [CREATE] строка Excel: {row.source_row or 'авто'}; "
                    f"путь: {row.path_text}; тип: {row.effective_type}\n"
                )
                try:
                    new_id = client.create_item(
                        parent_id,
                        row.name,
                        row.content_type_id,
                        self.target_list_id,
                        default_responsible_id=default_responsible_id,
                    )
                except Exception as exc:
                    raise RuntimeError(
                        f"Ошибка создания элемента.\n"
                        f"Строка Excel: {row.source_row or 'авто'}\n"
                        f"Путь: {row.path_text}\n"
                        f"Тип: {row.effective_type}\n"
                        f"content_type_id: {row.content_type_id}\n\n"
                        f"{exc}"
                    ) from exc
                runtime_ids[row.path] = new_id
                created_count += 1
                self.log.emit(f"  [CREATED] {row.path_text} [{row.effective_type}]\n")

            self.log.emit("\nИтог создания\n")
            self.log.emit(f"  создано: {created_count}\n")
            self.log.emit(f"  уже было: {existing_count}\n")
            self.log.emit(f"  предупреждений: {warning_count}\n")
            self.log.emit(f"  всего обработано: {len(self.plan)}\n\n")

            self.done.emit(
                {
                    "created": created_count,
                    "existing": existing_count,
                    "warnings": warning_count,
                    "total": len(self.plan),
                }
            )
        except Exception:
            self.failed.emit(traceback.format_exc())


# -----------------------------------------------------------------------------
# UI: современный интерфейс
# -----------------------------------------------------------------------------
class ModernSection(QtWidgets.QWidget):
    """Карточка-секция с заголовком и опциональным сворачиванием."""

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

        self.icon_label = QtWidgets.QLabel(icon, self.header)
        self.icon_label.setObjectName("sectionIcon")
        self.title_label = QtWidgets.QLabel(title, self.header)
        self.title_label.setObjectName("sectionTitle")
        self.arrow_label = QtWidgets.QLabel("⌄" if collapsible else "", self.header)
        self.arrow_label.setObjectName("sectionArrow")
        self.arrow_label.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)

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

    def _toggle_from_mouse(self, event):  # noqa: D401 - Qt callback
        self.set_expanded(not self._expanded)
        event.accept()

    def set_expanded(self, expanded: bool):
        if not self._collapsible:
            return
        self._expanded = expanded
        self.body.setVisible(expanded)
        self.arrow_label.setText("⌄" if expanded else "›")


class LogDialog(QtWidgets.QDialog):
    """Окно лога в стиле PP.py: тёмная консольная область без лишней обвязки."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Log")
        self.resize(800, 500)
        self.setMinimumSize(500, 300)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(0)

        self.text = QtWidgets.QTextEdit()
        self.text.setObjectName("logArea")
        self.text.setReadOnly(True)
        self.text.setFont(QtGui.QFont("Cascadia Code", 10))
        layout.addWidget(self.text, 1)

    def append(self, text: str):
        self.text.moveCursor(QtGui.QTextCursor.End)
        self.text.insertPlainText(text)
        self.text.ensureCursorVisible()

    def clear_log(self):
        self.text.clear()

    def set_log_text(self, value: str):
        self.text.setPlainText(value)
        self.text.moveCursor(QtGui.QTextCursor.End)
        self.text.ensureCursorVisible()


class TypeMappingDialog(QtWidgets.QDialog):
    """Редактор соответствий: значение в Excel -> тип контента VitroCAD."""

    def __init__(self, mapping: Dict[str, str], server_content_types: Sequence[dict], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Типы папок")
        self.resize(760, 520)
        self.setMinimumSize(680, 420)
        self.server_content_types = list(server_content_types or [])
        self._result_mapping = dict(mapping or DEFAULT_TYPE_MAPPINGS)

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        title = QtWidgets.QLabel("Сопоставление типов из Excel с типами VitroCAD")
        title.setObjectName("dialogTitle")
        root.addWidget(title)

        hint = QtWidgets.QLabel(
            "Левая колонка — как тип написан в Excel. Правая колонка — имя папочного типа VitroCAD или его content_type_id. "
            "Если в Excel появится новый тип, добавь его здесь."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #667085; font-size: 12px;")
        root.addWidget(hint)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        self.table = QtWidgets.QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Тип в Excel", "Тип VitroCAD или content_type_id"])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked | QtWidgets.QAbstractItemView.EditKeyPressed | QtWidgets.QAbstractItemView.AnyKeyPressed)
        self.table.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        splitter.addWidget(self.table)

        self.available = QtWidgets.QTextEdit()
        self.available.setReadOnly(True)
        self.available.setMinimumHeight(100)
        self.available.setPlainText(self._available_types_text())
        splitter.addWidget(self.available)
        splitter.setSizes([320, 140])
        root.addWidget(splitter, 1)

        buttons_top = QtWidgets.QHBoxLayout()
        self.btn_add = QtWidgets.QPushButton("Добавить строку")
        self.btn_add.setObjectName("secondaryButton")
        self.btn_remove = QtWidgets.QPushButton("Удалить строку")
        self.btn_remove.setObjectName("secondaryButton")
        self.btn_default = QtWidgets.QPushButton("Заполнить по умолчанию")
        self.btn_default.setObjectName("secondaryButton")
        buttons_top.addWidget(self.btn_add)
        buttons_top.addWidget(self.btn_remove)
        buttons_top.addWidget(self.btn_default)
        buttons_top.addStretch(1)
        root.addLayout(buttons_top)

        buttons = QtWidgets.QHBoxLayout()
        buttons.addStretch(1)
        self.btn_cancel = QtWidgets.QPushButton("Отмена")
        self.btn_cancel.setObjectName("secondaryButton")
        self.btn_save = QtWidgets.QPushButton("Сохранить")
        self.btn_save.setObjectName("primaryButton")
        buttons.addWidget(self.btn_cancel)
        buttons.addWidget(self.btn_save)
        root.addLayout(buttons)

        self.btn_add.clicked.connect(self._add_empty_row)
        self.btn_remove.clicked.connect(self._remove_selected_rows)
        self.btn_default.clicked.connect(self._load_defaults)
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_save.clicked.connect(self._save_and_accept)

        self._fill_table(self._result_mapping)

    def _available_types_text(self) -> str:
        if not self.server_content_types:
            return "Доступные типы VitroCAD не загружены. Авторизуйся и открой это окно снова, чтобы увидеть список с сервера."
        lines = ["Доступные папочные типы VitroCAD:"]
        for item in sorted(self.server_content_types, key=lambda it: name_key(content_type_name_from_item(it))):
            if not is_content_type_folder_like(item):
                continue
            type_name = content_type_name_from_item(item)
            aliases = [name for name in content_type_names_from_item(item) if name_key(name) != name_key(type_name)]
            alias_text = f"    aliases: {', '.join(aliases)}" if aliases else ""
            type_id = clean_text(item.get("id") or item.get("Id"))
            lines.append(f"- {type_name}    |    {type_id}{alias_text}")
        return "\n".join(lines)

    def _fill_table(self, mapping: Dict[str, str]) -> None:
        items = sorted(mapping.items(), key=lambda pair: name_key(pair[0]))
        self.table.setRowCount(len(items))
        for row, (excel_type, vitro_type) in enumerate(items):
            self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(excel_type))
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(vitro_type))
        self.table.resizeRowsToContents()

    def _load_defaults(self):
        self._fill_table(DEFAULT_TYPE_MAPPINGS)

    def _add_empty_row(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(""))
        self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(""))
        self.table.setCurrentCell(row, 0)

    def _remove_selected_rows(self):
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        for row in rows:
            self.table.removeRow(row)

    def _collect_mapping(self) -> Dict[str, str]:
        mapping: Dict[str, str] = {}
        errors: List[str] = []
        for row in range(self.table.rowCount()):
            excel_item = self.table.item(row, 0)
            vitro_item = self.table.item(row, 1)
            excel_type = clean_text(excel_item.text() if excel_item else "")
            vitro_type = clean_text(vitro_item.text() if vitro_item else "")
            if not excel_type and not vitro_type:
                continue
            if not excel_type or not vitro_type:
                errors.append(f"строка {row + 1}: должны быть заполнены обе колонки")
                continue
            mapping[excel_type] = vitro_type
        if errors:
            raise ValueError("Ошибки в таблице типов:\n" + "\n".join(errors))
        if not mapping:
            raise ValueError("Таблица типов пустая.")
        return mapping

    def _save_and_accept(self):
        try:
            self._result_mapping = self._collect_mapping()
            save_type_mappings(self._result_mapping)
            self.accept()
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Ошибка", str(exc))

    def mapping(self) -> Dict[str, str]:
        return dict(self._result_mapping)



# Диалог обязательных атрибутов по названию типа удалён.
# Скрипт больше не решает заранее, что тип "Папка проекта/Проект" обязан иметь responsible.
# При создании используется серверный шаблон /api/item/createNewInstance.


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VitroCAD — создание структуры из Excel")
        self.resize(1050, 700)
        self.setMinimumSize(980, 620)

        self.base_url = DEFAULT_SERVER
        self.auth_token: Optional[str] = None
        self.selected_folder_id: Optional[str] = None
        self.selected_list_id: str = LIST_ID
        self.current_plan: List[PlanRow] = []
        self.active_threads: List[QtCore.QThread] = []
        self.active_workers: List[QtCore.QObject] = []
        self._log_lines: List[str] = []
        self.log_dialog: Optional[LogDialog] = None
        self.type_mapping: Dict[str, str] = load_type_mappings()

        self._setup_ui()
        self._apply_stylesheet()

    def _setup_ui(self):
        app_font = QtGui.QFont("Segoe UI", 9)
        self.setFont(app_font)

        central = QtWidgets.QWidget()
        central.setObjectName("page")
        self.setCentralWidget(central)
        root = QtWidgets.QVBoxLayout(central)
        root.setContentsMargins(18, 14, 18, 0)
        root.setSpacing(8)

        # Заголовок страницы
        title = QtWidgets.QLabel("VitroCAD — создание структуры из Excel")
        title.setObjectName("mainTitle")
        subtitle = QtWidgets.QLabel("Загрузка файла и создание структуры проекта")
        subtitle.setObjectName("subtitle")
        root.addWidget(title)
        root.addWidget(subtitle)
        root.addSpacing(2)

        # Авторизация
        sec_auth = ModernSection("■", "Авторизация")
        auth = sec_auth.grid

        self.ed_server = QtWidgets.QLineEdit(DEFAULT_SERVER)
        self.ed_server.setPlaceholderText("https://vitrocad.bim-info.ru")
        self.ed_login = QtWidgets.QLineEdit()
        self.ed_login.setPlaceholderText("Введите логин")
        self.ed_password = QtWidgets.QLineEdit()
        self.ed_password.setPlaceholderText("Введите пароль")
        self.ed_password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.btn_toggle_password = QtWidgets.QToolButton()
        self.btn_toggle_password.setText("Показать")
        self.btn_toggle_password.setObjectName("eyeButton")
        self.btn_toggle_password.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_toggle_password.setToolTip("Показать / скрыть пароль")
        self.btn_login = QtWidgets.QPushButton("Войти")
        self.btn_login.setObjectName("primaryButton")
        self.btn_login.setMinimumWidth(220)
        self.lbl_auth = QtWidgets.QLabel("Не авторизован")
        self.lbl_auth.setObjectName("badStatus")
        self.lbl_auth.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)

        server_row = self._input_with_icon("", self.ed_server)
        login_row = self._input_with_icon("", self.ed_login)
        password_row = self._input_with_icon("", self.ed_password, self.btn_toggle_password)

        auth.addWidget(self._field_label("Сервер"), 0, 0)
        auth.addWidget(server_row, 0, 1, 1, 5)
        auth.addWidget(self._field_label("Логин"), 1, 0)
        auth.addWidget(login_row, 1, 1, 1, 2)
        auth.addWidget(self._field_label("Пароль"), 1, 3)
        auth.addWidget(password_row, 1, 4, 1, 2)
        auth.addWidget(self.btn_login, 2, 1, 1, 2)
        auth.addWidget(self.lbl_auth, 2, 3, 1, 3)
        auth.setColumnStretch(1, 1)
        auth.setColumnStretch(2, 1)
        auth.setColumnStretch(4, 1)
        auth.setColumnStretch(5, 1)
        root.addWidget(sec_auth)

        # Настройка
        sec_setup = ModernSection("■", "Настройка")
        setup = sec_setup.grid
        self.ed_excel = QtWidgets.QLineEdit()
        self.ed_excel.setPlaceholderText("Выберите файл Excel")
        self.btn_pick_excel = QtWidgets.QPushButton("Обзор...")
        self.btn_pick_excel.setObjectName("secondaryButton")
        self.btn_types = QtWidgets.QPushButton("Типы с сервера...")
        self.btn_types.setObjectName("secondaryButton")
        self.btn_types.setToolTip("Показать типы, которые программа получила с сервера VitroCAD")
        self.cb_project_as_folder = QtWidgets.QCheckBox("Служебный режим: заменять 'Проект' / 'Папка проекта' на 'Папка'")
        self.cb_project_as_folder.setChecked(False)
        self.cb_project_as_folder.setVisible(False)
        self.cb_project_as_folder.setToolTip("В прямом режиме не используется.")
        setup.addWidget(self._field_label("Excel-файл"), 0, 0)
        setup.addWidget(self.ed_excel, 0, 1, 1, 4)
        setup.addWidget(self.btn_pick_excel, 0, 5)
        setup.addWidget(self.cb_project_as_folder, 1, 1, 1, 4)
        setup.addWidget(self.btn_types, 1, 5)
        setup.setColumnStretch(1, 1)
        setup.setColumnStretch(2, 1)
        setup.setColumnStretch(3, 1)
        setup.setColumnStretch(4, 1)
        root.addWidget(sec_setup)

        # Центральная зона
        content_row = QtWidgets.QHBoxLayout()
        content_row.setSpacing(10)

        sec_tree = ModernSection("■", "Целевая папка / проект", collapsible=False)
        tree_grid = sec_tree.grid
        self.tree_folders = QtWidgets.QTreeWidget()
        self.tree_folders.setHeaderLabels(["Имя", "Тип"])
        self.tree_folders.setColumnWidth(0, 280)
        self.tree_folders.setAlternatingRowColors(False)
        self.tree_folders.setUniformRowHeights(True)
        self.tree_folders.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        tree_grid.addWidget(self.tree_folders, 0, 0)

        sec_plan = ModernSection("■", "Предпросмотр плана", collapsible=False)
        plan_grid = sec_plan.grid
        self.tbl_plan = QtWidgets.QTableWidget(0, 5)
        self.tbl_plan.setHorizontalHeaderLabels(["Статус", "Строка", "Путь", "Тип", "Комментарий"])
        self.tbl_plan.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tbl_plan.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.tbl_plan.verticalHeader().setVisible(False)
        self.tbl_plan.setWordWrap(False)
        self.tbl_plan.setAlternatingRowColors(False)
        header = self.tbl_plan.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.Stretch)
        plan_grid.addWidget(self.tbl_plan, 0, 0)

        content_row.addWidget(sec_tree, 1)
        content_row.addWidget(sec_plan, 2)
        root.addLayout(content_row, 1)

        # Нижняя панель действий
        footer = QtWidgets.QFrame()
        footer.setObjectName("footer")
        footer_layout = QtWidgets.QHBoxLayout(footer)
        footer_layout.setContentsMargins(0, 10, 0, 10)
        footer_layout.setSpacing(10)
        self.btn_log = QtWidgets.QPushButton("Log")
        self.btn_log.setObjectName("logBtn")
        self.btn_preview = QtWidgets.QPushButton("Построить предпросмотр")
        self.btn_preview.setObjectName("secondaryActionButton")
        self.btn_create = QtWidgets.QPushButton("Создать структуру")
        self.btn_create.setObjectName("primaryButton")
        self.btn_create.setEnabled(False)
        footer_layout.addStretch(1)
        footer_layout.addWidget(self.btn_log)
        footer_layout.addWidget(self.btn_preview)
        footer_layout.addWidget(self.btn_create)
        root.addWidget(footer)

        # Статусбар как на скриншоте: низкая белая строка
        self.statusBar().showMessage("Готово")

        # Signals
        self.btn_login.clicked.connect(self._login)
        self.btn_toggle_password.clicked.connect(self._toggle_password_visibility)
        self.btn_pick_excel.clicked.connect(self._pick_excel)
        self.btn_types.clicked.connect(self._configure_types)
        self.btn_preview.clicked.connect(self._build_preview)
        self.btn_create.clicked.connect(self._create_structure)
        self.btn_log.clicked.connect(self._show_log_dialog)
        self.tree_folders.itemClicked.connect(self._on_folder_selected)
        self.tree_folders.itemExpanded.connect(self._load_children_for_item)
        self.ed_excel.textChanged.connect(self._invalidate_plan)
        self.cb_project_as_folder.stateChanged.connect(self._invalidate_plan)

    def _field_label(self, text: str) -> QtWidgets.QLabel:
        label = QtWidgets.QLabel(text)
        label.setObjectName("fieldLabel")
        return label

    def _input_with_icon(self, icon: str, line_edit: QtWidgets.QLineEdit, trailing_widget: Optional[QtWidgets.QWidget] = None) -> QtWidgets.QWidget:
        # В PP.py поля ввода без декоративных пиктограмм.
        # Поэтому оставляем только аккуратный контейнер и само поле,
        # чтобы не было странных символов из-за разных наборов шрифтов Windows.
        box = QtWidgets.QFrame()
        box.setObjectName("inputBox")
        layout = QtWidgets.QHBoxLayout(box)
        layout.setContentsMargins(10, 0, 8, 0)
        layout.setSpacing(8)
        line_edit.setObjectName("embeddedLineEdit")
        layout.addWidget(line_edit, 1)
        if trailing_widget is not None:
            layout.addWidget(trailing_widget)
        return box

    def _apply_stylesheet(self):
        self.setStyleSheet(
            """
            QMainWindow { background: #ffffff; }
            QWidget#page { background: #ffffff; color: #0f172a; }
            QLabel#mainTitle {
                font-size: 18px;
                line-height: 24px;
                font-weight: 700;
                color: #0b1220;
                padding: 0;
            }
            QLabel#subtitle {
                font-size: 12px;
                color: #7b879f;
                padding-bottom: 2px;
            }
            QLabel#dialogTitle {
                font-size: 14px;
                font-weight: 600;
                color: #0b1220;
            }
            QFrame#card {
                background: #ffffff;
                border: 1px solid #dbe3ef;
                border-radius: 8px;
            }
            QFrame#sectionHeader {
                background: transparent;
                border: 0;
            }
            QFrame#sectionBody {
                background: transparent;
                border: 0;
            }
            QLabel#sectionIcon {
                color: #1277f3;
                font-size: 13px;
                font-weight: 600;
                min-width: 16px;
            }
            QLabel#sectionTitle {
                color: #111827;
                font-size: 13px;
                font-weight: 600;
            }
            QLabel#sectionArrow {
                color: #475569;
                font-size: 15px;
                font-weight: 600;
            }
            QLabel#fieldLabel {
                color: #111827;
                font-size: 13px;
                padding-right: 4px;
            }
            QFrame#inputBox {
                background: #ffffff;
                border: 1px solid #dbe3ef;
                border-radius: 6px;
                min-height: 32px;
            }
            QFrame#inputBox:hover {
                border: 1px solid #c6d3e2;
            }
            QLabel#inputIcon {
                color: #71809b;
                font-size: 14px;
            }
            QLineEdit {
                min-height: 32px;
                border: 1px solid #dbe3ef;
                border-radius: 6px;
                background: #ffffff;
                color: #111827;
                padding: 0 12px;
                font-size: 13px;
            }
            QLineEdit:focus {
                border: 1px solid #93bdf8;
            }
            QLineEdit#embeddedLineEdit {
                border: 0;
                background: transparent;
                padding: 0;
                min-height: 32px;
            }
            QLineEdit#embeddedLineEdit:focus {
                border: 0;
            }
            QLineEdit::placeholder {
                color: #7b879f;
            }
            QToolButton#eyeButton {
                border: 0;
                background: transparent;
                color: #667085;
                font-size: 12px;
                padding: 0 6px;
            }
            QToolButton#eyeButton:hover {
                color: #475467;
            }
            QCheckBox {
                spacing: 8px;
                font-size: 13px;
                color: #1f2937;
                padding-top: 2px;
            }
            QPushButton {
                min-height: 32px;
                border: 1px solid transparent;
                border-radius: 8px;
                padding: 0 18px;
                font-size: 13px;
                font-weight: 600;
                background: #eef2f7;
                color: #111827;
            }
            QPushButton:hover {
                background: #e5ebf4;
            }
            QPushButton:disabled {
                background: #e9edf3;
                color: #9aa6b8;
                border: 1px solid #e1e7f0;
            }
            QPushButton#primaryButton {
                background: #0f72ed;
                color: #ffffff;
                border: 1px solid #0f72ed;
                min-width: 180px;
            }
            QPushButton#primaryButton:hover {
                background: #0c66d8;
                border: 1px solid #0c66d8;
            }
            QPushButton#secondaryButton {
                background: #ffffff;
                border: 1px solid #dbe3ef;
                color: #1f2937;
                min-width: 120px;
            }
            QPushButton#secondaryButton:hover {
                background: #f8fafc;
                border: 1px solid #cbd7e6;
            }
            QPushButton#secondaryActionButton {
                background: #ffffff;
                border: 1px solid #dbe3ef;
                color: #1f2937;
                min-width: 220px;
            }
            QPushButton#secondaryActionButton:hover {
                background: #f8fafc;
            }
            QPushButton#outlineButton {
                background: #ffffff;
                border: 1px solid #dbe3ef;
                color: #0f72ed;
                min-width: 70px;
                padding: 0 14px;
            }
            QPushButton#outlineButton:hover {
                background: #f4f8ff;
                border: 1px solid #c8d8f0;
            }
            QPushButton#logBtn {
                background-color: transparent;
                color: #667085;
                border: 1px solid #d0d5dd;
                min-width: 70px;
                padding: 0 14px;
                font-size: 12px;
            }
            QPushButton#logBtn:hover {
                background-color: #e8ecf1;
                color: #475467;
                border: 1px solid #d0d5dd;
            }
            QLabel#goodStatus {
                color: #15803d;
                font-weight: 650;
                font-size: 13px;
            }
            QLabel#badStatus {
                color: #ff0000;
                font-weight: 520;
                font-size: 13px;
            }
            QTreeWidget, QTableWidget, QPlainTextEdit {
                background: #ffffff;
                border: 1px solid #dbe3ef;
                border-radius: 8px;
                gridline-color: #e6edf5;
                color: #111827;
                selection-background-color: #eaf3ff;
                selection-color: #0b1220;
            }
            QTextEdit#logArea {
                border: 1px solid #d0d5dd;
                border-radius: 8px;
                background-color: #1e293b;
                color: #e2e8f0;
                padding: 10px;
                font-size: 12px;
                selection-background-color: #3b82f6;
                selection-color: #ffffff;
            }
            QTreeWidget::item, QTableWidget::item {
                min-height: 24px;
                padding: 3px;
                border: 0;
            }
            QTreeWidget::item:selected, QTableWidget::item:selected {
                background: #eaf3ff;
                color: #0b1220;
            }
            QHeaderView::section {
                background: #ffffff;
                color: #111827;
                padding: 8px 8px;
                border: 0;
                border-right: 1px solid #e6edf5;
                border-bottom: 1px solid #e6edf5;
                font-weight: 520;
                font-size: 13px;
            }
            QScrollBar:horizontal {
                height: 12px;
                background: #ffffff;
                margin: 2px 8px 2px 8px;
            }
            QScrollBar::handle:horizontal {
                background: #d6d9df;
                border-radius: 7px;
                min-width: 80px;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0;
            }
            QScrollBar:vertical {
                width: 12px;
                background: #ffffff;
                margin: 8px 2px 8px 2px;
            }
            QScrollBar::handle:vertical {
                background: #d6d9df;
                border-radius: 7px;
                min-height: 80px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0;
            }
            QFrame#footer {
                background: #ffffff;
                border-top: 1px solid #dbe3ef;
            }
            QStatusBar {
                background: #ffffff;
                color: #263244;
                border-top: 1px solid #dbe3ef;
                padding-left: 12px; font-size: 11px;
                min-height: 24px;
            }
            QStatusBar::item { border: 0; }
            QDialog {
                background: #ffffff;
            }
            """
        )

    # ------------------------------------------------------------------
    # UI helpers
    # ------------------------------------------------------------------
    def _load_server_content_types_for_dialog(self) -> List[dict]:
        if not self.auth_token:
            return []
        try:
            client = VitroClient(self.base_url, self.auth_token)
            content_types = client.get_content_types(getattr(self, "selected_list_id", LIST_ID))
            register_server_content_types(content_types)
            return content_types
        except Exception as exc:
            self._log(f"WARNING: не удалось загрузить типы VitroCAD для окна настройки: {exc}")
            return []

    def _configure_types(self):
        server_types = self._load_server_content_types_for_dialog()
        resolver = ContentTypeResolver(server_types, {})

        excel_types: List[str] = []
        excel_path = self.ed_excel.text().strip()
        if excel_path and os.path.exists(excel_path):
            try:
                nodes = ExcelStructureParser(excel_path, False).parse()
                excel_types = sorted({node.original_type for node in nodes if node.source_row}, key=name_key)
            except Exception as exc:
                self._log(f"WARNING: не удалось прочитать типы из Excel для просмотра: {exc}")

        available = sorted(set(resolver.folder_type_names), key=name_key)
        available_keys = {name_key(name) for name in available}
        found = [name for name in excel_types if name_key(name) in available_keys]
        missing = [name for name in excel_types if name_key(name) not in available_keys]

        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("Типы папок с сервера")
        dialog.resize(760, 520)
        layout = QtWidgets.QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        title = QtWidgets.QLabel("Проверка типов папок")
        title.setObjectName("dialogTitle")
        layout.addWidget(title)

        hint = QtWidgets.QLabel(
            "Программа не использует зашитые сопоставления. Она берёт значения из Excel "
            "и проверяет их напрямую среди папочных типов, полученных с сервера VitroCAD "
            "для выбранного списка."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #667085; font-size: 12px;")
        layout.addWidget(hint)

        text = QtWidgets.QTextEdit()
        text.setReadOnly(True)
        text.setFont(QtGui.QFont("Cascadia Code", 10))
        lines: List[str] = []
        lines.append("Типы из Excel\n")
        lines.append("-" * 60 + "\n")
        if excel_types:
            for name in excel_types:
                status = "OK" if name_key(name) in available_keys else "NOT FOUND"
                lines.append(f"[{status}] {name}\n")
        else:
            lines.append("Excel не выбран или типы не прочитаны.\n")
        lines.append("\n")
        lines.append("Папочные типы, полученные с сервера\n")
        lines.append("-" * 60 + "\n")
        if available:
            for name in available:
                lines.append(f"{name}\n")
        else:
            lines.append("Типы с сервера не получены. Проверь авторизацию и выбранную папку/проект.\n")
        lines.append("\n")
        lines.append("Итог\n")
        lines.append("-" * 60 + "\n")
        lines.append(f"найдено типов из Excel: {len(found)}\n")
        lines.append(f"не найдено типов из Excel: {len(missing)}\n")
        if missing:
            lines.append("\nНе найдены:\n")
            for name in missing:
                lines.append(f"- {name}\n")
        text.setPlainText("".join(lines))
        layout.addWidget(text, 1)

        buttons = QtWidgets.QHBoxLayout()
        buttons.addStretch(1)
        close_btn = QtWidgets.QPushButton("Закрыть")
        close_btn.setObjectName("secondaryButton")
        close_btn.clicked.connect(dialog.accept)
        buttons.addWidget(close_btn)
        layout.addLayout(buttons)
        dialog.exec()
        self._log(f"Проверены типы с сервера. Excel OK: {len(found)}; не найдено: {len(missing)}.")

    def _toggle_password_visibility(self):
        if self.ed_password.echoMode() == QtWidgets.QLineEdit.Password:
            self.ed_password.setEchoMode(QtWidgets.QLineEdit.Normal)
            self.btn_toggle_password.setText("Скрыть")
        else:
            self.ed_password.setEchoMode(QtWidgets.QLineEdit.Password)
            self.btn_toggle_password.setText("Показать")

    def _log(self, message: str):
        chunk = str(message)
        if chunk and not chunk.endswith("\n"):
            chunk += "\n"
        self._log_lines.append(chunk)
        if self.log_dialog is not None and self.log_dialog.isVisible():
            self.log_dialog.append(chunk)

    def _show_log_dialog(self):
        if self.log_dialog is None:
            self.log_dialog = LogDialog(self)
        self.log_dialog.set_log_text("".join(self._log_lines) if self._log_lines else "Лог пока пуст.\n")
        self.log_dialog.show()
        self.log_dialog.raise_()
        self.log_dialog.activateWindow()

    def _set_busy(self, busy: bool, message: str = ""):
        self.btn_login.setEnabled(not busy)
        self.btn_preview.setEnabled(not busy)
        self.btn_create.setEnabled((not busy) and bool(self.current_plan))
        self.btn_pick_excel.setEnabled(not busy)
        self.btn_types.setEnabled(not busy)
        self.ed_server.setEnabled(not busy)
        self.ed_login.setEnabled(not busy)
        self.ed_password.setEnabled(not busy)
        self.cb_project_as_folder.setEnabled(not busy)
        self.tree_folders.setEnabled(not busy)
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

    def _show_error(self, title: str, details: str):
        self._log(f"ERROR: {title}\n{details}")
        QtWidgets.QMessageBox.critical(self, title, details[:4000])

    def _invalidate_plan(self):
        self.current_plan = []
        self.tbl_plan.setRowCount(0)
        self.btn_create.setEnabled(False)

    # ------------------------------------------------------------------
    # Авторизация и дерево
    # ------------------------------------------------------------------
    def _login(self):
        server = self.ed_server.text().strip().rstrip("/")
        login = self.ed_login.text().strip()
        password = self.ed_password.text()

        if not server or not login or not password:
            QtWidgets.QMessageBox.warning(self, "Не хватает данных", "Заполни сервер, логин и пароль.")
            return

        self._set_busy(True, "Авторизация...")
        self._log("Авторизация...")
        try:
            client = VitroClient(server)
            token = client.login(login, password)
            self.base_url = server
            self.auth_token = token
            try:
                content_types = VitroClient(server, token).get_content_types(LIST_ID)
                register_server_content_types(content_types)
                folder_count = sum(1 for item in content_types if is_content_type_folder_like(item))
                self._log(f"Загружены типы VitroCAD: папочных типов: {folder_count}.")
            except Exception as type_exc:
                self._log(f"WARNING: авторизация успешна, но типы VitroCAD загрузить не удалось: {type_exc}")
            self.lbl_auth.setText("Авторизован")
            self.lbl_auth.setObjectName("goodStatus")
            self.lbl_auth.style().unpolish(self.lbl_auth)
            self.lbl_auth.style().polish(self.lbl_auth)
            self._log("Авторизация успешно завершена.")
            self._load_root_folders()
        except Exception as exc:
            self.auth_token = None
            self.lbl_auth.setText("Ошибка авторизации")
            self.lbl_auth.setObjectName("badStatus")
            self.lbl_auth.style().unpolish(self.lbl_auth)
            self.lbl_auth.style().polish(self.lbl_auth)
            self._show_error("Ошибка авторизации", str(exc))
        finally:
            self._set_busy(False, "Готово")

    def _load_root_folders(self):
        if not self.auth_token:
            return
        self.tree_folders.clear()
        root_item = QtWidgets.QTreeWidgetItem(["Корень списка", "Список"])
        root_item.setData(0, QtCore.Qt.UserRole, LIST_ID)
        root_item.setData(0, QtCore.Qt.UserRole + 1, LIST_ID)
        root_item.addChild(QtWidgets.QTreeWidgetItem(["...", ""]))
        self.tree_folders.addTopLevelItem(root_item)
        root_item.setExpanded(True)
        self._load_children_for_item(root_item)

    def _load_children_for_item(self, item: QtWidgets.QTreeWidgetItem):
        if not self.auth_token:
            return
        if item.childCount() == 1 and item.child(0).text(0) == "...":
            item.takeChild(0)
        else:
            return

        parent_id = item.data(0, QtCore.Qt.UserRole)
        if not parent_id:
            return

        worker = FolderTreeWorker(self.base_url, self.auth_token, parent_id)
        worker.loaded.connect(lambda folders, parent=item: self._on_tree_children_loaded(parent, folders))
        worker.failed.connect(lambda err: self._show_error("Ошибка загрузки дерева", err))
        worker.loaded.connect(lambda _: self.statusBar().showMessage("Готово"))
        worker.failed.connect(lambda _: self.statusBar().showMessage("Ошибка"))
        self.statusBar().showMessage("Загрузка дерева...")
        self._run_worker(worker)

    def _on_tree_children_loaded(self, parent_item: QtWidgets.QTreeWidgetItem, folders: List[dict]):
        for folder in folders:
            name = get_item_name(folder)
            type_name = item_type_name(folder)
            folder_id = folder.get("id")
            folder_list_id = clean_text(folder.get("listId") or folder.get("list_id")) or LIST_ID
            child_item = QtWidgets.QTreeWidgetItem([name, type_name])
            child_item.setData(0, QtCore.Qt.UserRole, folder_id)
            child_item.setData(0, QtCore.Qt.UserRole + 1, folder_list_id)
            child_item.addChild(QtWidgets.QTreeWidgetItem(["...", ""]))
            parent_item.addChild(child_item)
        parent_item.setExpanded(True)
        self.tree_folders.resizeColumnToContents(1)
        self._log(f"Загружено элементов дерева: {len(folders)}")

    def _on_folder_selected(self, item: QtWidgets.QTreeWidgetItem, _column: int):
        self.selected_folder_id = item.data(0, QtCore.Qt.UserRole)
        self.selected_list_id = clean_text(item.data(0, QtCore.Qt.UserRole + 1)) or LIST_ID
        self._invalidate_plan()
        if self.selected_folder_id:
            self._log(f"Выбрана целевая папка: {item.text(0)}")

    # ------------------------------------------------------------------
    # Excel и план
    # ------------------------------------------------------------------
    def _pick_excel(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Выберите Excel-файл",
            "",
            "Excel (*.xlsx *.xlsm)",
        )
        if path:
            self.ed_excel.setText(path)

    def _validate_preview_inputs(self) -> bool:
        if not self.auth_token:
            QtWidgets.QMessageBox.warning(self, "Нет авторизации", "Сначала войди в VitroCAD.")
            return False
        if not self.selected_folder_id:
            QtWidgets.QMessageBox.warning(self, "Не выбрана папка", "Выбери целевую папку или проект в дереве.")
            return False
        excel_path = self.ed_excel.text().strip()
        if not excel_path or not os.path.exists(excel_path):
            QtWidgets.QMessageBox.warning(self, "Excel не найден", "Выбери существующий Excel-файл.")
            return False
        return True

    def _build_preview(self):
        if not self._validate_preview_inputs():
            return

        self._set_busy(True, "Построение предпросмотра...")
        self._invalidate_plan()
        self._log("Построение предпросмотра...")

        worker = PlanWorker(
            excel_path=self.ed_excel.text().strip(),
            base_url=self.base_url,
            token=self.auth_token or "",
            target_id=self.selected_folder_id or "",
            target_list_id=getattr(self, "selected_list_id", LIST_ID),
            project_as_folder=self.cb_project_as_folder.isChecked(),
            type_mapping=self.type_mapping,
        )
        worker.log.connect(self._log)
        worker.done.connect(self._on_plan_ready)
        worker.failed.connect(self._on_plan_failed)
        self._run_worker(worker)

    def _on_plan_ready(self, plan: List[PlanRow]):
        self.current_plan = plan
        self._fill_plan_table(plan)
        create_count = sum(1 for row in plan if row.status == "Создать")
        exists_count = sum(1 for row in plan if row.status == "Уже есть")
        self._log(f"Предпросмотр готов: создать: {create_count}; уже есть: {exists_count}; всего: {len(plan)}.")
        self._set_busy(False, "Предпросмотр готов")
        self.btn_create.setEnabled(bool(plan))

    def _on_plan_failed(self, error: str):
        self.current_plan = []
        self._fill_plan_table([])
        self._set_busy(False, "Ошибка")
        self._show_error("Ошибка предпросмотра", error)

    def _fill_plan_table(self, plan: Sequence[PlanRow]):
        display_plan = self._sort_plan_for_display(plan)
        self.tbl_plan.setRowCount(len(display_plan))
        for row_idx, row in enumerate(display_plan):
            values = [
                row.status,
                str(row.source_row or "авто"),
                row.path_text,
                row.type_display,
                row.message,
            ]
            for col_idx, value in enumerate(values):
                item = QtWidgets.QTableWidgetItem(value)
                item.setToolTip(value)
                if row.status == "Создать":
                    item.setBackground(QtGui.QColor("#fff7d6"))
                elif row.status == "Уже есть":
                    item.setBackground(QtGui.QColor("#e7f5e7"))
                self.tbl_plan.setItem(row_idx, col_idx, item)
        self.tbl_plan.resizeRowsToContents()

    @staticmethod
    def _sort_plan_for_display(plan: Sequence[PlanRow]) -> List[PlanRow]:
        """Показывает предпросмотр в порядке строк Excel, а не в порядке создания.

        Автоматически добавленные родители ставятся перед первой дочерней строкой,
        из-за которой они понадобились. Само создание всё равно выполняется
        безопасно: сначала родители, потом дочерние элементы.
        """
        plan_list = list(plan)

        def first_child_source_row(parent: PlanRow) -> int:
            child_rows = [
                item.source_row
                for item in plan_list
                if item.source_row
                and len(item.path) > len(parent.path)
                and item.path[: len(parent.path)] == parent.path
            ]
            return min(child_rows) if child_rows else 10**9

        def display_key(item: PlanRow):
            if item.source_row:
                return (item.source_row, 1, item.depth, item.path_text)
            return (first_child_source_row(item), 0, item.depth, item.path_text)

        return sorted(plan_list, key=display_key)

    # ------------------------------------------------------------------
    # Создание
    # ------------------------------------------------------------------
    def _create_structure(self):
        if not self.current_plan:
            QtWidgets.QMessageBox.warning(self, "Нет плана", "Сначала построй предпросмотр.")
            return
        if not self.auth_token or not self.selected_folder_id:
            QtWidgets.QMessageBox.warning(self, "Нет данных", "Нет авторизации или целевой папки.")
            return

        create_count = sum(1 for row in self.current_plan if row.status == "Создать")
        answer = QtWidgets.QMessageBox.question(
            self,
            "Подтвердить создание",
            f"Будет обработано элементов: {len(self.current_plan)}.\n"
            f"По предпросмотру новых к созданию: {create_count}.\n\n"
            "Перед созданием программа ещё раз проверит существующие элементы, чтобы не создать дубли.\n"
            "Обязательные поля не определяются по вшитым названиям типов: программа берёт шаблон создания с сервера VitroCAD.\n"
            "Продолжить?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        )
        if answer != QtWidgets.QMessageBox.Yes:
            return

        self._set_busy(True, "Создание структуры...")
        self._log("Создание структуры...")
        worker = CreateWorker(
            plan=self.current_plan,
            base_url=self.base_url,
            token=self.auth_token,
            target_id=self.selected_folder_id,
            target_list_id=getattr(self, "selected_list_id", LIST_ID),
            default_responsible_id="",
        )
        worker.log.connect(self._log)
        worker.done.connect(self._on_create_done)
        worker.failed.connect(self._on_create_failed)
        self._run_worker(worker)

    def _on_create_done(self, result: dict):
        self._set_busy(False, "Готово")
        self._log(
            "Готово. "
            f"Создано: {result['created']}; "
            f"уже было: {result['existing']}; "
            f"предупреждений: {result['warnings']}; "
            f"всего обработано: {result['total']}."
        )
        QtWidgets.QMessageBox.information(
            self,
            "Готово",
            f"Создано: {result['created']}\n"
            f"Уже было: {result['existing']}\n"
            f"Предупреждений: {result['warnings']}\n"
            f"Всего обработано: {result['total']}",
        )
        self.btn_create.setEnabled(False)

    def _on_create_failed(self, error: str):
        self._set_busy(False, "Ошибка")
        self._show_error("Ошибка создания", error)


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
