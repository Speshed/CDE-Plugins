# -*- coding: utf-8 -*-
"""
Larix — импорт план-графика из Excel

Назначение:
1. Авторизуется в Larix.
2. Читает обязательные поля типа «Задача план-графика» непосредственно с сервера.
3. Сравнивает обязательные пользовательские поля Larix со столбцами Excel.
   Поля, заданные контекстом создания, определяет автоматически по ответу сервера.
4. Читает уже существующую структуру план-графика из VitroCAD.
5. Сопоставляет элементы без UUID — по полному пути из столбца «Название»:
       Проект / Стадия / Объект / Документ / Задание
6. Обновляет найденные элементы и создаёт только отсутствующие.
7. Передаёт доступные значения из Excel в поля VitroCAD.
8. При обрыве ответа переподключается, проверяет результат и безопасно продолжает.

Зависимости:
    pip install PySide6 requests pandas openpyxl

Примечание:
- Технические параметры API настроены внутри программы и не показываются пользователю.
- Пароль нигде не сохраняется.
- Для внутреннего сервера Larix используется подключение без проверки SSL-сертификата.
"""
from __future__ import annotations

import json
import math
import re
import sys
import traceback
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from shared.theme_core import themed_icon
from time import sleep
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
import requests
import urllib3
from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import QThread, Signal


APP_TITLE = "Larix — Синхронизация план-графика с Excel"
DEFAULT_BASE_URL = "https://192.168.12.2"
DEFAULT_LIST_ID = "716e8d52-90dc-4c85-bddf-582c94ab505e"
DEFAULT_CONTENT_TYPE_ID = "97af0dc5-a9da-455e-8e5a-c70da99cfcdc"
DEFAULT_UTC_OFFSET_HOURS = 3.0
VERIFY_SSL = False
CONTENT_TYPE_NAME = "Задача план-графика"
REQUEST_CONNECT_TIMEOUT = 15
REQUEST_READ_TIMEOUT = 90
RECOVERY_POLL_DELAYS = (2, 4, 8, 12)

UUID_RE = re.compile(
    r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-"
    r"[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}$"
)

# Названия столбцов нормализуются, поэтому регистр и лишние пробелы не важны.
FIELD_ALIASES: Dict[str, Sequence[str]] = {
    "content_type_id": (
        "Тип", "Тип задачи", "Тип элемента", "content_type_id",
    ),
    "name": (
        "Название", "Наименование", "Задача", "name",
    ),
    "start_date_plan": (
        "Дата начала (План)", "Дата начала план", "Начало план",
        "Плановое начало", "start_date_plan",
    ),
    "end_date_plan": (
        "Дата окончания (План)", "Дата окончания план", "Окончание план",
        "Плановое окончание", "Прогнозный срок", "Пронозный срок",
        "Срок", "end_date_plan",
    ),
    "assignedto": (
        "Исполнитель", "Ответственный исполнитель", "ID исполнителя",
        "Логин исполнителя", "assignedto",
    ),
    "task_status": (
        "Статус задачи", "Статус", "task_status",
    ),
    "importance": (
        "Важность", "Приоритет", "importance",
    ),
    "description": (
        "Описание", "description",
    ),
    "comment": (
        "Комментарий", "comment",
    ),
    "duration": (
        "Продолжительность", "duration",
    ),
    "cost_plan": (
        "Плановые трудозатраты", "Трудозатраты плановые", "Плановые",
        "cost_plan",
    ),
    "cost_max": (
        "Максимальные трудозатраты", "Максимальные", "cost_max",
    ),
    "ancestors": (
        "Предшественники", "ancestors",
    ),
    "descendants": (
        "Потомки", "descendants",
    ),
}

PATH_ALIASES = (
    "Путь", "Полный путь", "Иерархия", "Родительский путь", "path",
)


# -----------------------------------------------------------------------------
# Общие функции
# -----------------------------------------------------------------------------
def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_key(value: Any) -> str:
    text = normalize_text(value).lower().replace("ё", "е")
    return re.sub(r"[^a-zа-я0-9]+", "", text)


def is_empty(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    if isinstance(value, str):
        return not value.strip()
    return False


def is_uuid(value: Any) -> bool:
    return bool(UUID_RE.fullmatch(normalize_text(value)))


def extract_id(value: Any) -> Optional[str]:
    if isinstance(value, dict):
        candidate = value.get("id")
        return normalize_text(candidate) or None
    text = normalize_text(value)
    return text if is_uuid(text) else None


def split_path(value: Any) -> Tuple[str, ...]:
    text = normalize_text(value)
    if not text:
        return tuple()
    parts = [normalize_text(part) for part in re.split(r"\s*/\s*", text)]
    return tuple(part for part in parts if part)


def normalize_path_segment(value: Any) -> str:
    """Нормализация одного уровня пути для безопасного сопоставления.

    Игнорируются только регистр, буква ё/е и лишние пробелы. Пунктуация
    сохраняется, чтобы разные реальные названия не были ошибочно объединены.
    """
    return normalize_text(value).lower().replace("ё", "е")


def path_key(path: Sequence[str]) -> Tuple[str, ...]:
    return tuple(normalize_path_segment(part) for part in path)


def item_field_map(item: Dict[str, Any]) -> Dict[str, Any]:
    value = item.get("fieldValueMap")
    return value if isinstance(value, dict) else {}


def item_name(item: Dict[str, Any]) -> str:
    return normalize_text(item_field_map(item).get("name") or item.get("name"))


def item_id(item: Dict[str, Any]) -> str:
    return normalize_text(item.get("id") or item_field_map(item).get("id"))


def item_parent_id(item: Dict[str, Any]) -> str:
    return normalize_text(item.get("parentId") or item.get("parent_id"))


def item_content_type_id(item: Dict[str, Any]) -> str:
    return normalize_text(item.get("contentTypeId") or item.get("content_type_id"))


def excel_serial_to_datetime(value: float) -> datetime:
    # Excel использует фиктивную дату 1900-02-29; база 1899-12-30 учитывает это.
    return datetime(1899, 12, 30) + timedelta(days=float(value))


def parse_datetime_value(value: Any) -> Optional[datetime]:
    if is_empty(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, time.min)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        numeric = float(value)
        if math.isfinite(numeric) and 1 <= numeric <= 100000:
            return excel_serial_to_datetime(numeric)
    text = normalize_text(value)
    if not text:
        return None
    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    if pd.isna(parsed):
        raise ValueError(f"не удалось распознать дату: {value!r}")
    return parsed.to_pydatetime()


def to_api_datetime(value: Any, utc_offset_hours: float) -> Optional[str]:
    dt = parse_datetime_value(value)
    if dt is None:
        return None
    local_tz = timezone(timedelta(hours=float(utc_offset_hours)))
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=local_tz)
    else:
        dt = dt.astimezone(local_tz)
    utc_dt = dt.astimezone(timezone.utc)
    return utc_dt.isoformat(timespec="milliseconds").replace("+00:00", "Z")


def coerce_number(value: Any, integer: bool = False) -> Any:
    if is_empty(value):
        return None
    if isinstance(value, str):
        value = value.replace(" ", "").replace(",", ".")
    number = float(value)
    return int(number) if integer else number


def extract_json_error(response: requests.Response) -> str:
    try:
        payload = response.json()
        if isinstance(payload, dict):
            for key in ("message", "error", "detail", "title"):
                if payload.get(key):
                    return normalize_text(payload[key])
        return json.dumps(payload, ensure_ascii=False)[:600]
    except Exception:
        return normalize_text(response.text)[:600]


def find_column(columns: Iterable[Any], aliases: Iterable[str]) -> Optional[str]:
    normalized = {normalize_key(column): str(column) for column in columns}
    for alias in aliases:
        found = normalized.get(normalize_key(alias))
        if found is not None:
            return found
    return None


class UncertainWriteError(RuntimeError):
    """Результат POST /api/item/update неизвестен из-за обрыва ответа.

    Сервер мог применить изменение, даже если клиент не получил ответ.
    Поэтому такой запрос нельзя повторять вслепую.
    """

    def __init__(self, message: str):
        super().__init__(message)


@dataclass
class FieldMeta:
    internal_name: str
    display_name: str
    description: str
    field_id: str
    field_type: str
    required: bool
    readonly: bool
    lookup_list_id: Optional[str]
    default_value_guid: Optional[str]
    multi: bool
    visible: bool
    component: str

    @classmethod
    def from_api(cls, raw: Dict[str, Any]) -> "FieldMeta":
        field_type = raw.get("fieldType") if isinstance(raw.get("fieldType"), dict) else {}
        config = raw.get("fieldValueMap") if isinstance(raw.get("fieldValueMap"), dict) else {}
        return cls(
            internal_name=normalize_text(raw.get("internalName")),
            display_name=normalize_text(raw.get("name")),
            description=normalize_text(raw.get("description")),
            field_id=normalize_text(raw.get("id")),
            field_type=normalize_text(field_type.get("name")),
            required=bool(raw.get("isRequired")),
            readonly=bool(raw.get("isReadOnly")),
            lookup_list_id=normalize_text(config.get("list")) or None,
            default_value_guid=normalize_text(config.get("default_value_guid")) or None,
            multi=bool(config.get("multi")),
            visible=bool(raw.get("isVisible", True)),
            component=normalize_text(raw.get("component")),
        )


@dataclass
class ExcelData:
    path: str
    sheet_name: str
    dataframe: pd.DataFrame

    @property
    def columns(self) -> List[str]:
        return [str(column) for column in self.dataframe.columns]


@dataclass
class ScheduleNode:
    path: Tuple[str, ...]
    source_row_number: Optional[int]
    row: Optional[Dict[str, Any]]

    @property
    def name(self) -> str:
        return self.path[-1]


@dataclass
class ExistingScheduleIndex:
    """Индекс существующего дерева VitroCAD по полному пути."""

    unique: Dict[Tuple[str, ...], Dict[str, Any]]
    duplicates: Dict[Tuple[str, ...], List[Dict[str, Any]]]
    display_paths: Dict[Tuple[str, ...], Tuple[str, ...]]
    warnings: List[str]

    @classmethod
    def build(cls, items: Sequence[Dict[str, Any]], list_id: str,
              content_type_id: str) -> "ExistingScheduleIndex":
        by_id: Dict[str, Dict[str, Any]] = {}
        for raw in items:
            if not isinstance(raw, dict):
                continue
            current_id = item_id(raw)
            if current_id and current_id != list_id:
                by_id[current_id] = raw

        cache: Dict[str, Optional[Tuple[str, ...]]] = {}
        visiting: set[str] = set()
        warnings: List[str] = []

        def resolve_path(current_id: str) -> Optional[Tuple[str, ...]]:
            if current_id in cache:
                return cache[current_id]
            if current_id in visiting:
                warnings.append(f"Обнаружен цикл родителей у элемента {current_id}; элемент пропущен.")
                cache[current_id] = None
                return None

            current = by_id.get(current_id)
            if not current:
                cache[current_id] = None
                return None
            current_name = item_name(current)
            if not current_name:
                warnings.append(f"У элемента {current_id} отсутствует название; элемент пропущен.")
                cache[current_id] = None
                return None

            visiting.add(current_id)
            parent_id = item_parent_id(current)
            if not parent_id or parent_id == list_id:
                result: Optional[Tuple[str, ...]] = (current_name,)
            elif parent_id in by_id:
                parent_path = resolve_path(parent_id)
                result = parent_path + (current_name,) if parent_path else None
            else:
                warnings.append(
                    f"У элемента «{current_name}» ({current_id}) не найден родитель {parent_id}; "
                    "элемент не участвует в сопоставлении."
                )
                result = None
            visiting.discard(current_id)
            cache[current_id] = result
            return result

        grouped: Dict[Tuple[str, ...], List[Dict[str, Any]]] = {}
        display_paths: Dict[Tuple[str, ...], Tuple[str, ...]] = {}
        for current_id, current in by_id.items():
            # Обновляем только элементы того типа, который сейчас выбран сервером
            # как «Задача план-графика». Родители других типов всё равно входят
            # в расчёт пути через by_id.
            if item_content_type_id(current) != content_type_id:
                continue
            resolved = resolve_path(current_id)
            if not resolved:
                continue
            key = path_key(resolved)
            grouped.setdefault(key, []).append(current)
            display_paths.setdefault(key, resolved)

        unique = {key: values[0] for key, values in grouped.items() if len(values) == 1}
        duplicates = {key: values for key, values in grouped.items() if len(values) > 1}
        return cls(unique=unique, duplicates=duplicates,
                   display_paths=display_paths, warnings=warnings)


# -----------------------------------------------------------------------------
# Excel
# -----------------------------------------------------------------------------
class ExcelReader:
    @staticmethod
    def list_sheets(path: str) -> List[str]:
        excel = pd.ExcelFile(path)
        return [str(name) for name in excel.sheet_names]

    @staticmethod
    def read(path: str, sheet_name: str) -> ExcelData:
        df = pd.read_excel(path, sheet_name=sheet_name, dtype=object)
        df.columns = [normalize_text(column) for column in df.columns]
        df = df.dropna(how="all").reset_index(drop=True)
        return ExcelData(path=path, sheet_name=sheet_name, dataframe=df)

    @staticmethod
    def prepare_nodes(excel: ExcelData) -> Tuple[List[ScheduleNode], List[str]]:
        df = excel.dataframe
        name_col = find_column(df.columns, FIELD_ALIASES["name"])
        path_col = find_column(df.columns, PATH_ALIASES)
        if not name_col and not path_col:
            raise ValueError(
                "В Excel не найден столбец «Название». Допустимы также «Наименование» или «Путь»."
            )

        explicit: Dict[Tuple[str, ...], ScheduleNode] = {}
        order: List[Tuple[str, ...]] = []
        warnings: List[str] = []

        for index, series in df.iterrows():
            row_number = index + 2
            row = {str(key): value for key, value in series.to_dict().items()}
            name_value = row.get(name_col) if name_col else None
            path_value = row.get(path_col) if path_col else None

            if path_col:
                path = split_path(path_value)
                leaf = normalize_text(name_value)
                if leaf and (not path or path[-1] != leaf):
                    path = path + (leaf,)
            else:
                path = split_path(name_value)

            if not path:
                continue
            if path in explicit:
                warnings.append(
                    f"Строка {row_number}: путь «{' / '.join(path)}» уже встречался; повтор пропущен."
                )
                continue
            explicit[path] = ScheduleNode(path=path, source_row_number=row_number, row=row)
            order.append(path)

        if not order:
            raise ValueError("В выбранном листе нет строк с непустым названием задачи.")

        nodes: List[ScheduleNode] = []
        added: set[Tuple[str, ...]] = set()
        for path in order:
            for depth in range(1, len(path) + 1):
                prefix = path[:depth]
                if prefix in added:
                    continue
                nodes.append(
                    explicit.get(prefix)
                    or ScheduleNode(path=prefix, source_row_number=None, row=None)
                )
                added.add(prefix)
        return nodes, warnings


# -----------------------------------------------------------------------------
# Larix API
# -----------------------------------------------------------------------------
class LarixApi:
    def __init__(self, base_url: str, list_id: str, content_type_id: str,
                 verify_ssl: bool, log: Optional[Callable[[str], None]] = None):
        self.base_url = base_url.rstrip("/")
        self.list_id = normalize_text(list_id)
        self.content_type_id = normalize_text(content_type_id)
        self.verify_ssl = bool(verify_ssl)
        self.session = requests.Session()
        # API Larix возвращает отдельный токен авторизации в поле ``token``.
        # Обычные API-запросы передают его в заголовке Authorization.
        # AuthorizationToken используется отдельными загрузчиками файлов,
        # поэтому для совместимости мы устанавливаем оба заголовка.
        self.authorization_token: Optional[str] = None
        self._auth_variant_index: int = 0
        self.current_user_id: Optional[str] = None
        self.current_user_name: Optional[str] = None
        self.content_type: Optional[Dict[str, Any]] = None
        self.default_instance: Optional[Dict[str, Any]] = None
        self.fields: Dict[str, FieldMeta] = {}
        self.lookup_cache: Dict[Tuple[str, str], str] = {}
        self._log_func = log

        if not self.verify_ssl:
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    def _log(self, message: str) -> None:
        if self._log_func:
            self._log_func(message)

    def _url(self, path: str) -> str:
        return f"{self.base_url}{path}"

    def _auth_variants(self) -> List[Dict[str, str]]:
        """Варианты заголовка для разных сборок Larix/Vitro."""
        token = normalize_text(self.authorization_token)
        if not token:
            return []
        return [
            {"Authorization": token},
            {"Authorization": f"Bearer {token}"},
            {"AuthorizationToken": token},
            {"Authorization": token, "AuthorizationToken": token},
        ]

    def _apply_auth_variant(self, index: int) -> None:
        variants = self._auth_variants()
        if not variants:
            return
        index = max(0, min(index, len(variants) - 1))
        self.session.headers.pop("Authorization", None)
        self.session.headers.pop("AuthorizationToken", None)
        self.session.headers.update(variants[index])
        self._auth_variant_index = index

    def reconnect(self) -> None:
        """Создаёт новое HTTP-соединение, сохраняя действующий токен.

        Это не повторный вход с паролем: старый зависший TCP-сеанс закрывается,
        а дальнейшие запросы выполняются через новую requests.Session.
        """
        try:
            self.session.close()
        except Exception:
            pass
        self.session = requests.Session()
        self.session.headers.update({"Accept": "*/*"})
        if self.authorization_token:
            self._apply_auth_variant(self._auth_variant_index)
        self._log("HTTP-соединение с Larix создано заново.")

    def _request(self, method: str, path: str, **kwargs: Any) -> requests.Response:
        kwargs.setdefault(
            "timeout",
            (REQUEST_CONNECT_TIMEOUT, REQUEST_READ_TIMEOUT),
        )
        kwargs.setdefault("verify", self.verify_ssl)

        def send_request() -> requests.Response:
            try:
                return self.session.request(method, self._url(path), **kwargs)
            except requests.exceptions.ConnectTimeout:
                raise RuntimeError(
                    f"Не удалось подключиться к Larix за {REQUEST_CONNECT_TIMEOUT} секунд. "
                    "Проверьте доступность сервера и сети."
                ) from None
            except requests.exceptions.ReadTimeout:
                if method.upper() == "POST" and path == "/api/item/update":
                    raise UncertainWriteError(
                        f"Larix не вернул ответ на сохранение за {REQUEST_READ_TIMEOUT} секунд."
                    ) from None
                raise RuntimeError(
                    f"Larix не ответил за {REQUEST_READ_TIMEOUT} секунд. "
                    f"Запрос: {method} {path}."
                ) from None
            except requests.exceptions.ConnectionError as exc:
                if method.upper() == "POST" and path == "/api/item/update":
                    raise UncertainWriteError(
                        "Соединение оборвалось во время сохранения; результат запроса неизвестен."
                    ) from None
                raise RuntimeError(
                    f"Соединение с Larix было разорвано: {normalize_text(exc)}"
                ) from None

        response = send_request()

        # В разных версиях Larix встречаются Authorization, Bearer и
        # AuthorizationToken. При 401 перебираются варианты заголовка.
        if response.status_code == 401 and self.authorization_token:
            variants = self._auth_variants()
            start_index = self._auth_variant_index
            for index in range(len(variants)):
                if index == start_index:
                    continue
                self._apply_auth_variant(index)
                response = send_request()
                if response.status_code != 401:
                    self._log("Формат авторизации определён автоматически.")
                    break

        if response.status_code == 401:
            raise RuntimeError(
                "Сервер отклонил токен авторизации. Выйдите и войдите заново. "
                f"Запрос: {method} {path}."
            )
        if response.status_code not in (200, 201):
            raise RuntimeError(
                f"{method} {path}: HTTP {response.status_code}: {extract_json_error(response)}"
            )
        return response

    def login(self, login: str, password: str) -> Dict[str, Any]:
        response = self._request(
            "POST", "/api/security/login",
            json={"login": login, "password": password},
        )
        payload = response.json()
        if not isinstance(payload, dict):
            raise RuntimeError("Сервер вернул некорректный ответ авторизации.")

        # В ответе есть ДВА разных значения: ``id`` и ``token``.
        # Для защищённых API нужен именно ``token``. В предыдущей версии
        # ошибочно выбирался ``id``, из-за чего createNewInstance отвечал 401.
        token = normalize_text(
            payload.get("token")
            or payload.get("authorizationToken")
            or payload.get("authorization_token")
        )
        if not token:
            token = normalize_text(
                response.headers.get("Authorization")
                or response.headers.get("AuthorizationToken")
            )
        # Fallback для старых сборок Larix, где токен мог возвращаться как id.
        if not token:
            token = normalize_text(payload.get("id"))
        if not token:
            raise RuntimeError(
                "Вход выполнен, но сервер не вернул токен авторизации."
            )

        self.authorization_token = token
        self.session.headers.update({"Accept": "*/*"})
        # Начинаем с обычного Authorization. При необходимости _request
        # автоматически переключится на другой формат после первого 401.
        self._apply_auth_variant(0)

        user = payload.get("user")
        if not isinstance(user, dict):
            raise RuntimeError("Авторизация выполнена, но сервер не вернул данные пользователя.")
        user_id = normalize_text(user.get("id"))
        if not user_id:
            raise RuntimeError("В ответе авторизации отсутствует ID пользователя.")
        self.current_user_id = user_id
        self.current_user_name = normalize_text(user.get("name")) or login
        return payload

    def load_content_type(self) -> Dict[str, Any]:
        response = self._request(
            "POST", f"/api/contentType/getByList/{self.list_id}"
        )
        content_types = response.json()
        if not isinstance(content_types, list):
            raise RuntimeError("Сервер вернул некорректный список типов содержимого.")

        selected = None
        for item in content_types:
            if not isinstance(item, dict):
                continue
            if normalize_text(item.get("id")) == self.content_type_id:
                selected = item
                break
        if selected is None:
            for item in content_types:
                if not isinstance(item, dict):
                    continue
                name = normalize_text(item.get("name") or item.get("description"))
                if name.lower() == CONTENT_TYPE_NAME.lower():
                    selected = item
                    self.content_type_id = normalize_text(item.get("id"))
                    break
        if selected is None:
            names = [normalize_text(item.get("name")) for item in content_types if isinstance(item, dict)]
            raise RuntimeError(
                f"Тип «{CONTENT_TYPE_NAME}» не найден в списке. Доступны: {', '.join(filter(None, names))}"
            )

        self.content_type = selected
        self.fields = {}
        for raw in selected.get("fieldList", []):
            if not isinstance(raw, dict):
                continue
            field = FieldMeta.from_api(raw)
            if field.internal_name:
                self.fields[field.internal_name] = field
        return selected

    def required_fields(self) -> List[FieldMeta]:
        """Все обязательные редактируемые поля по метаданным VitroCAD."""
        return [
            field for field in self.fields.values()
            if field.required and not field.readonly
        ]

    def _creation_context_ids(self) -> set[str]:
        """ID, которые VitroCAD задаёт самим контекстом создания элемента.

        Список не привязан к именам полей. Значения берутся из фактического
        ответа createNewInstance: тип содержимого, список и родитель.
        """
        instance = self.default_instance or {}
        raw_values = [
            self.list_id,
            self.content_type_id,
            instance.get("listId"),
            instance.get("list_id"),
            instance.get("contentTypeId"),
            instance.get("content_type_id"),
            instance.get("parentId"),
            instance.get("parent_id"),
        ]
        return {normalize_text(value) for value in raw_values if normalize_text(value)}

    def field_instance_value(self, field: FieldMeta) -> Any:
        instance = self.default_instance or {}
        field_map = instance.get("fieldValueMap")
        if not isinstance(field_map, dict):
            return None
        return field_map.get(field.internal_name)

    def is_creation_context_field(self, field: FieldMeta) -> bool:
        """Поле уже однозначно задано VitroCAD при создании элемента.

        Например, поле «Тип» в createNewInstance содержит тот же ID, что и
        contentTypeId создаваемого элемента. Поэтому пользователь не может и
        не должен передавать его отдельным столбцом Excel. Определение
        выполняется по данным сервера, без списка имён обязательных полей.
        """
        value = self.field_instance_value(field)
        value_id = extract_id(value)
        if not value_id and not isinstance(value, (dict, list, tuple)):
            value_id = normalize_text(value) or None
        return bool(value_id and normalize_text(value_id) in self._creation_context_ids())

    def required_excel_fields(self) -> List[FieldMeta]:
        """Обязательные поля, значения которых должен предоставить Excel."""
        return [
            field for field in self.required_fields()
            if not self.is_creation_context_field(field)
        ]

    def describe_instance_value(self, field: FieldMeta) -> str:
        value = self.field_instance_value(field)
        if isinstance(value, dict):
            value_map = value.get("fieldValueMap")
            if isinstance(value_map, dict):
                for key in ("name", "title", "login", "email"):
                    text = normalize_text(value_map.get(key))
                    if text:
                        return text
            for key in ("name", "title"):
                text = normalize_text(value.get(key))
                if text:
                    return text
            return extract_id(value) or "задаётся VitroCAD"
        return normalize_text(value) or "задаётся VitroCAD"

    def create_new_instance(self, parent_id: str) -> Dict[str, Any]:
        response = self._request(
            "POST", "/api/item/createNewInstance",
            json={
                "list_id": self.list_id,
                "content_type_id": self.content_type_id,
                "parent_id": parent_id,
            },
        )
        payload = response.json()
        if not isinstance(payload, dict):
            raise RuntimeError("createNewInstance вернул некорректный ответ.")
        return payload

    def _submit_item_payload(self, item_payload: Dict[str, Any]) -> Dict[str, Any]:
        serialized = json.dumps([item_payload], ensure_ascii=False, separators=(",", ":"))
        # Формат соответствует HAR: multipart/form-data, поле itemListJson.
        files = {"itemListJson": (None, serialized, "application/json")}
        response = self._request("POST", "/api/item/update", files=files)
        payload = response.json()
        if isinstance(payload, dict):
            result = payload
        elif isinstance(payload, list) and payload and isinstance(payload[0], dict):
            result = payload[0]
        else:
            raise RuntimeError("Сервер вернул некорректный ответ после item/update.")
        if not normalize_text(result.get("id") or item_field_map(result).get("id")):
            raise RuntimeError("В ответе item/update отсутствует ID элемента.")
        return result

    def update_new_item(self, item_payload: Dict[str, Any]) -> Dict[str, Any]:
        return self._submit_item_payload(item_payload)

    def update_existing_item(self, existing: Dict[str, Any],
                             changed_values: Dict[str, Any]) -> Dict[str, Any]:
        """Обновляет существующий элемент без изменения его положения в дереве."""
        current_id = item_id(existing)
        if not current_id:
            raise RuntimeError("У существующего элемента отсутствует ID.")
        payload: Dict[str, Any] = {
            "id": current_id,
            "list_id": normalize_text(existing.get("listId") or existing.get("list_id")) or self.list_id,
            "content_type_id": item_content_type_id(existing) or self.content_type_id,
            "parent_id": item_parent_id(existing) or self.list_id,
        }
        payload.update(changed_values)
        return self._submit_item_payload(payload)

    def get_item(self, current_id: str) -> Dict[str, Any]:
        response = self._request("POST", f"/api/item/get/{current_id}")
        payload = response.json()
        if not isinstance(payload, dict):
            raise RuntimeError(f"Сервер не вернул элемент {current_id}.")
        return payload

    def find_child(self, parent_id: str, name: str) -> Optional[Dict[str, Any]]:
        """Находит единственного дочернего элемента по имени и типу."""
        response = self._request("POST", f"/api/item/getList/{parent_id}")
        payload = response.json()
        if not isinstance(payload, list):
            raise RuntimeError(f"getList/{parent_id} вернул ответ не в формате списка.")
        target = normalize_path_segment(name)
        matches = [
            item for item in payload
            if isinstance(item, dict)
            and item_content_type_id(item) == self.content_type_id
            and normalize_path_segment(item_name(item)) == target
        ]
        if len(matches) > 1:
            raise RuntimeError(
                f"После восстановления соединения найдено несколько элементов «{name}» "
                f"у одного родителя. Автоматическое продолжение небезопасно."
            )
        return matches[0] if matches else None

    def get_existing_tree(self) -> List[Dict[str, Any]]:
        """Читает текущее дерево списка.

        Основной способ — getRecursive. Для сборок, где этот метод недоступен
        для корня списка, используется безопасный обход getList/{parentId}.
        """
        try:
            response = self._request("POST", f"/api/item/getRecursive/{self.list_id}")
            payload = response.json()
            if isinstance(payload, list):
                self._log(f"Получена существующая структура: {len(payload)} элементов.")
                return [item for item in payload if isinstance(item, dict)]
            raise RuntimeError("getRecursive вернул ответ не в формате списка.")
        except Exception as recursive_error:
            self._log(
                "getRecursive для корня списка недоступен; "
                "перехожу к последовательному чтению дочерних элементов."
            )
            self._log(f"Причина переключения: {recursive_error}")

        result: List[Dict[str, Any]] = []
        seen: set[str] = set()
        queued: set[str] = {self.list_id}
        queue: List[str] = [self.list_id]
        max_items = 100000

        while queue:
            parent = queue.pop(0)
            response = self._request("POST", f"/api/item/getList/{parent}")
            payload = response.json()
            if not isinstance(payload, list):
                raise RuntimeError(f"getList/{parent} вернул ответ не в формате списка.")
            for raw in payload:
                if not isinstance(raw, dict):
                    continue
                current_id = item_id(raw)
                if not current_id or current_id in seen or current_id == self.list_id:
                    continue
                seen.add(current_id)
                result.append(raw)
                if current_id not in queued:
                    queued.add(current_id)
                    queue.append(current_id)
                if len(result) > max_items:
                    raise RuntimeError(
                        f"В списке больше {max_items} элементов. "
                        "Обход остановлен для защиты приложения."
                    )
        self._log(f"Получена существующая структура: {len(result)} элементов.")
        return result

    def resolve_lookup(self, field: FieldMeta, value: Any) -> str:
        text = normalize_text(value)
        if not text:
            raise ValueError(f"поле «{field.display_name}» пустое")
        if is_uuid(text):
            return text
        if not field.lookup_list_id:
            raise ValueError(
                f"для поля «{field.display_name}» не указан lookup-список; передайте UUID"
            )

        cache_key = (field.internal_name, normalize_key(text))
        cached = self.lookup_cache.get(cache_key)
        if cached:
            return cached

        response = self._request(
            "POST", "/api/lookup/getList",
            json={
                "field": field.field_id,
                "value_string": text,
                "list": field.lookup_list_id,
            },
        )
        items = response.json()
        if not isinstance(items, list) or not items:
            raise ValueError(
                f"значение «{text}» не найдено для поля «{field.display_name}»"
            )

        def candidate_strings(item: Dict[str, Any]) -> List[str]:
            result: List[str] = []
            field_map = item.get("fieldValueMap") if isinstance(item.get("fieldValueMap"), dict) else {}
            for key in ("name", "login", "email", "title"):
                if field_map.get(key):
                    result.append(normalize_text(field_map[key]))
                if item.get(key):
                    result.append(normalize_text(item[key]))
            return result

        target = normalize_key(text)
        exact = [
            item for item in items if isinstance(item, dict)
            and any(normalize_key(candidate) == target for candidate in candidate_strings(item))
        ]
        candidates = exact or [item for item in items if isinstance(item, dict)]
        if len(candidates) != 1:
            variants = []
            for item in candidates[:8]:
                variants.extend(candidate_strings(item)[:1])
            raise ValueError(
                f"для «{field.display_name} = {text}» найдено несколько вариантов: "
                f"{', '.join(filter(None, variants)) or 'без названий'}"
            )
        resolved = normalize_text(candidates[0].get("id"))
        if not resolved:
            raise ValueError(f"lookup «{text}» найден, но не содержит ID")
        self.lookup_cache[cache_key] = resolved
        return resolved

    def default_value(self, instance: Dict[str, Any], internal_name: str) -> Any:
        field_map = instance.get("fieldValueMap")
        if not isinstance(field_map, dict):
            field_map = {}
        value = field_map.get(internal_name)
        if isinstance(value, dict):
            return extract_id(value)
        if not is_empty(value):
            return value
        field = self.fields.get(internal_name)
        return field.default_value_guid if field else None


# -----------------------------------------------------------------------------
# Сопоставление Excel → API
# -----------------------------------------------------------------------------
class ImportMapper:
    def __init__(self, api: LarixApi, excel: ExcelData, utc_offset_hours: float):
        self.api = api
        self.excel = excel
        self.utc_offset_hours = utc_offset_hours
        self.column_map = self._build_column_map()

    def _build_column_map(self) -> Dict[str, str]:
        result: Dict[str, str] = {}
        columns = self.excel.columns
        for internal_name, field in self.api.fields.items():
            aliases = list(FIELD_ALIASES.get(internal_name, ()))
            aliases.extend([internal_name, field.display_name])
            column = find_column(columns, aliases)
            if column:
                result[internal_name] = column
        return result

    def mandatory_report(self) -> List[Dict[str, str]]:
        """Сравнивает с Excel только обязательные поля, которыми управляет пользователь.

        Поля, уже однозначно заданные самим контекстом создания VitroCAD
        (например, тип создаваемого элемента), остаются в серверной логике,
        но не показываются пользователю и не требуют столбца Excel.
        """
        report: List[Dict[str, str]] = []
        for field in self.api.required_excel_fields():
            column = self.column_map.get(field.internal_name)
            if column:
                source = f"Excel: {column}"
                status = "OK"
                details = "Обязательный столбец найден в Excel"
            else:
                source = "Не найдено в Excel"
                status = "ОШИБКА"
                details = "Добавьте отдельный столбец для этого обязательного параметра"
            report.append({
                "name": field.display_name or field.internal_name,
                "internal": field.internal_name,
                "source": source,
                "status": status,
                "details": details,
            })
        return report

    def validate_required_values(self, nodes: Sequence[ScheduleNode]) -> None:
        """Проверяет значения обязательных полей до первого запроса создания.

        Проверка столбцов выполняется в mandatory_report. Здесь дополнительно
        проверяются пустые ячейки и автоматически достроенные уровни иерархии,
        для которых в Excel нет отдельной строки с обязательными значениями.
        """
        issues: List[str] = []
        required_fields = self.api.required_excel_fields()

        for node in nodes:
            location = (
                f"строка {node.source_row_number}"
                if node.source_row_number is not None
                else "автоматически добавленный уровень"
            )
            path_text = " / ".join(node.path)

            if node.row is None:
                issues.append(
                    f"{location} «{path_text}»: в Excel нет отдельной строки, "
                    "поэтому обязательные значения прочитать невозможно"
                )
                continue

            empty_fields: List[str] = []
            for field in required_fields:
                column = self.column_map.get(field.internal_name)
                if not column:
                    # Отсутствующий столбец уже показывается в mandatory_report.
                    continue
                if is_empty(node.row.get(column)):
                    empty_fields.append(field.display_name or field.internal_name)

            if empty_fields:
                issues.append(
                    f"{location} «{path_text}»: пустые обязательные поля: "
                    + ", ".join(empty_fields)
                )

        if issues:
            shown = issues[:20]
            suffix = ""
            if len(issues) > len(shown):
                suffix = f"\n... и ещё ошибок: {len(issues) - len(shown)}"
            raise ValueError(
                "Проверка обязательных значений не пройдена:\n"
                + "\n".join(f"• {issue}" for issue in shown)
                + suffix
            )

    def _row_value(self, row: Optional[Dict[str, Any]], internal_name: str) -> Any:
        if not row:
            return None
        column = self.column_map.get(internal_name)
        return row.get(column) if column else None

    def coerce_field(self, field: FieldMeta, value: Any) -> Any:
        if is_empty(value):
            return None
        field_type = field.field_type.lower()
        if field_type == "lookup":
            resolved = self.api.resolve_lookup(field, value)
            # VitroCAD принимает множественный Lookup именно как массив UUID.
            # Даже если в Excel указан один человек, для поля с multi=true
            # необходимо отправлять ["uuid"], а не одиночную строку "uuid".
            # Иначе /api/item/update завершается HTTP 500 (например, поле ГИП).
            return [resolved] if field.multi else resolved
        if field_type == "datetime":
            return to_api_datetime(value, self.utc_offset_hours)
        if field_type in ("decimal", "double", "float"):
            return coerce_number(value, integer=False)
        if field_type in ("long", "integer", "int"):
            return coerce_number(value, integer=True)
        if field_type in ("boolean", "bool"):
            if isinstance(value, bool):
                return value
            return normalize_key(value) in {"1", "true", "да", "yes", "истина"}
        return normalize_text(value)

    def excel_values(self, node: ScheduleNode) -> Dict[str, Any]:
        """Возвращает только значения, которыми Excel должен управлять.

        Отсутствующий столбец и пустая ячейка не очищают существующее значение.
        Название берётся из последнего уровня полного пути.
        """
        values: Dict[str, Any] = {"name": node.name}
        for internal_name, column in self.column_map.items():
            field = self.api.fields.get(internal_name)
            if internal_name == "name":
                continue
            if not field or field.readonly or self.api.is_creation_context_field(field) or not node.row:
                continue
            raw_value = node.row.get(column)
            if is_empty(raw_value):
                continue
            values[internal_name] = self.coerce_field(field, raw_value)
        return values

    @staticmethod
    def _existing_scalar(value: Any) -> Any:
        if isinstance(value, dict):
            return extract_id(value) or value
        if isinstance(value, list):
            return [extract_id(item) or item for item in value]
        return value

    def values_equal(self, field: Optional[FieldMeta], current: Any, desired: Any) -> bool:
        current = self._existing_scalar(current)
        if field is None:
            return normalize_text(current) == normalize_text(desired)
        field_type = field.field_type.lower()
        try:
            if field_type == "lookup":
                if isinstance(current, list) or isinstance(desired, list):
                    left = sorted(normalize_text(x) for x in (current or []))
                    right = sorted(normalize_text(x) for x in (desired or []))
                    return left == right
                return normalize_text(current) == normalize_text(desired)
            if field_type == "datetime":
                left = pd.to_datetime(current, utc=True, errors="coerce")
                right = pd.to_datetime(desired, utc=True, errors="coerce")
                if pd.isna(left) or pd.isna(right):
                    return normalize_text(current) == normalize_text(desired)
                return abs((left - right).total_seconds()) < 1
            if field_type in ("decimal", "double", "float", "long", "integer", "int"):
                return abs(float(current) - float(desired)) < 1e-9
            if field_type in ("boolean", "bool"):
                return bool(current) == bool(desired)
        except Exception:
            pass
        return normalize_text(current) == normalize_text(desired)

    def changed_values(self, node: ScheduleNode, existing: Dict[str, Any]) -> Dict[str, Any]:
        desired = self.excel_values(node)
        current_map = item_field_map(existing)
        changed: Dict[str, Any] = {}
        for internal_name, desired_value in desired.items():
            field = self.api.fields.get(internal_name)
            current_value = current_map.get(internal_name)
            if not self.values_equal(field, current_value, desired_value):
                changed[internal_name] = desired_value
        return changed

    def make_payload(self, node: ScheduleNode, parent_id: str,
                     instance: Dict[str, Any]) -> Dict[str, Any]:
        # Контекстные значения list_id, parent_id и content_type_id задаются
        # самой операцией создания. Какие обязательные поля относятся к этому
        # контексту, определяется по фактическому ответу VitroCAD.
        payload: Dict[str, Any] = {
            "list_id": self.api.list_id,
            "content_type_id": self.api.content_type_id,
            "parent_id": parent_id,
            "name": node.name,
        }
        # createNewInstance обычно выдаёт заранее созданный UUID. Сохраняем его
        # в payload: повтор того же запроса после обрыва будет относиться к тому
        # же элементу, а не создавать новый UUID.
        instance_id = item_id(instance)
        if instance_id and instance_id != self.api.list_id:
            payload["id"] = instance_id

        # Передаются только значения, которые действительно присутствуют в Excel.
        # Автоподстановка текущего пользователя и серверных значений по умолчанию
        # намеренно отключена.
        payload.update(self.excel_values(node))

        # Защитная проверка непосредственно перед отправкой. Источником всех
        # обязательных пользовательских значений должен быть Excel.
        missing: List[str] = []
        for field in self.api.required_excel_fields():
            column = self.column_map.get(field.internal_name)
            raw_value = node.row.get(column) if node.row and column else None
            if is_empty(raw_value):
                missing.append(field.display_name or field.internal_name)

        if missing:
            location = (
                f"строка {node.source_row_number}"
                if node.source_row_number
                else "автоматически добавленный уровень"
            )
            raise ValueError(
                f"{location}, «{' / '.join(node.path)}»: не заполнены обязательные поля Excel: "
                + ", ".join(missing)
            )
        return payload


# -----------------------------------------------------------------------------
# Потоки
# -----------------------------------------------------------------------------
class LoginWorker(QThread):
    succeeded = Signal(object, object)
    failed = Signal(str)
    log = Signal(str)

    def __init__(self, base_url: str, login: str, password: str):
        super().__init__()
        self.base_url = base_url
        self.login_value = login
        self.password_value = password

    def run(self) -> None:
        try:
            api = LarixApi(
                self.base_url, DEFAULT_LIST_ID, DEFAULT_CONTENT_TYPE_ID,
                VERIFY_SSL, log=lambda text: self.log.emit(text),
            )
            self.log.emit("Авторизация...")
            api.login(self.login_value, self.password_value)
            self.log.emit("Чтение структуры полей план-графика...")
            content_type = api.load_content_type()
            self.log.emit("Чтение серверного шаблона нового элемента...")
            api.default_instance = api.create_new_instance(api.list_id)
            self.succeeded.emit(api, content_type)
        except Exception as exc:
            self.log.emit(traceback.format_exc())
            self.failed.emit(str(exc))


class ImportWorker(QThread):
    progress = Signal(int, int)
    log = Signal(str)
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, api: LarixApi, excel: ExcelData, utc_offset_hours: float):
        super().__init__()
        self.api = api
        self.excel = excel
        self.utc_offset_hours = utc_offset_hours

    def _poll_item_applied(
        self, mapper: ImportMapper, node: ScheduleNode, current_id: str,
    ) -> Optional[Dict[str, Any]]:
        """Ждёт, пока элемент появится и получит все значения из Excel."""
        if not current_id:
            return None
        last_error = ""
        for delay in (0,) + RECOVERY_POLL_DELAYS:
            if delay:
                self.log.emit(f"Проверка результата через {delay} сек.")
                sleep(delay)
            try:
                item = self.api.get_item(current_id)
                if item and not mapper.changed_values(node, item):
                    return item
                last_error = "элемент найден, но изменения ещё не применены полностью"
            except Exception as exc:
                last_error = str(exc)
        if last_error:
            self.log.emit("Результат пока не подтверждён: " + last_error)
        return None

    def _poll_created_item(
        self, mapper: ImportMapper, parent_id: str, node: ScheduleNode,
        payload: Dict[str, Any],
    ) -> Optional[Dict[str, Any]]:
        payload_id = normalize_text(payload.get("id"))
        last_error = ""
        for delay in (0,) + RECOVERY_POLL_DELAYS:
            if delay:
                self.log.emit(f"Проверка результата создания через {delay} сек.")
                sleep(delay)
            try:
                if payload_id:
                    item = self.api.get_item(payload_id)
                else:
                    item = self.api.find_child(parent_id, node.name)
                if item and not mapper.changed_values(node, item):
                    return item
                if item:
                    last_error = "элемент найден, но его поля ещё не сохранены полностью"
            except Exception as exc:
                last_error = str(exc)
        if last_error:
            self.log.emit("Создание пока не подтверждено: " + last_error)
        return None

    def _save_existing_with_recovery(
        self, mapper: ImportMapper, node: ScheduleNode, existing: Dict[str, Any],
        changes: Dict[str, Any],
    ) -> Dict[str, Any]:
        current_id = item_id(existing)
        try:
            return self.api.update_existing_item(existing, changes)
        except UncertainWriteError as exc:
            self.log.emit(f"⚠️ {exc} Проверяю фактическое состояние элемента.")

        self.api.reconnect()
        recovered = self._poll_item_applied(mapper, node, current_id)
        if recovered:
            self.log.emit("✅ Изменение найдено на сервере; продолжаю синхронизацию.")
            return recovered

        self.log.emit("Изменение не подтверждено. Безопасно повторяю обновление по тому же ID один раз.")
        try:
            return self.api.update_existing_item(existing, changes)
        except UncertainWriteError as exc:
            self.log.emit(f"⚠️ Повторное сохранение также не вернуло ответ: {exc}")
            self.api.reconnect()
            recovered = self._poll_item_applied(mapper, node, current_id)
            if recovered:
                self.log.emit("✅ Повторное изменение найдено на сервере; продолжаю.")
                return recovered
            raise RuntimeError(
                f"Не удалось подтвердить обновление элемента «{' / '.join(node.path)}». "
                "Синхронизация остановлена, чтобы не продолжать с неизвестным состоянием."
            ) from None

    def _save_new_with_recovery(
        self, mapper: ImportMapper, node: ScheduleNode, parent_id: str,
        payload: Dict[str, Any],
    ) -> Dict[str, Any]:
        try:
            return self.api.update_new_item(payload)
        except UncertainWriteError as exc:
            self.log.emit(f"⚠️ {exc} Переподключаюсь и проверяю, создан ли элемент.")

        self.api.reconnect()
        recovered = self._poll_created_item(mapper, parent_id, node, payload)
        if recovered:
            self.log.emit("✅ Созданный элемент найден после переподключения; продолжаю.")
            return recovered

        payload_id = normalize_text(payload.get("id"))
        if not payload_id:
            raise RuntimeError(
                f"Сервер не подтвердил создание «{' / '.join(node.path)}», а createNewInstance "
                "не выдал UUID. Повтор остановлен, чтобы не создать дубль."
            )

        self.log.emit("Элемент не найден. Повторяю сохранение один раз с тем же UUID.")
        try:
            return self.api.update_new_item(payload)
        except UncertainWriteError as exc:
            self.log.emit(f"⚠️ Повторное сохранение также не вернуло ответ: {exc}")
            self.api.reconnect()
            recovered = self._poll_created_item(mapper, parent_id, node, payload)
            if recovered:
                self.log.emit("✅ Элемент найден после повторного сохранения; продолжаю.")
                return recovered
            raise RuntimeError(
                f"Не удалось подтвердить создание элемента «{' / '.join(node.path)}». "
                "Синхронизация остановлена во избежание дублей."
            ) from None

    def run(self) -> None:
        try:
            nodes, warnings = ExcelReader.prepare_nodes(self.excel)
            for warning in warnings:
                self.log.emit("⚠️ " + warning)

            mapper = ImportMapper(self.api, self.excel, self.utc_offset_hours)
            report = mapper.mandatory_report()
            missing = [row["name"] for row in report if row["status"] == "ОШИБКА"]
            if missing:
                raise RuntimeError(
                    "Нельзя начать синхронизацию. В Excel отсутствуют обязательные столбцы: "
                    + ", ".join(missing)
                )

            # Проверка выполняется до первого изменения на сервере.
            mapper.validate_required_values(nodes)

            self.log.emit("Чтение существующей структуры VitroCAD...")
            existing_items = self.api.get_existing_tree()
            index = ExistingScheduleIndex.build(
                existing_items, self.api.list_id, self.api.content_type_id
            )
            for warning in index.warnings:
                self.log.emit("⚠️ " + warning)

            excel_keys = {path_key(node.path) for node in nodes}
            conflicts = {
                key: values for key, values in index.duplicates.items()
                if key in excel_keys
            }
            if conflicts:
                lines: List[str] = []
                for key, values in list(conflicts.items())[:20]:
                    display = index.display_paths.get(key, tuple(key))
                    lines.append(f"• {' / '.join(display)} — найдено элементов: {len(values)}")
                suffix = ""
                if len(conflicts) > len(lines):
                    suffix = f"\n... и ещё конфликтов: {len(conflicts) - len(lines)}"
                raise RuntimeError(
                    "Синхронизация остановлена до внесения изменений. "
                    "В VitroCAD уже есть дубликаты одинаковых полных путей:\n"
                    + "\n".join(lines) + suffix
                    + "\nУдалите лишние дубликаты в VitroCAD и повторите запуск."
                )

            # Предварительно определяем изменения существующих элементов. Это
            # позволяет обнаружить ошибки lookup до первого изменения сервера.
            prepared_changes: Dict[Tuple[str, ...], Dict[str, Any]] = {}
            create_count = 0
            update_count = 0
            unchanged_count = 0
            for node in nodes:
                key = path_key(node.path)
                existing = index.unique.get(key)
                if existing is None:
                    create_count += 1
                    continue
                changes = mapper.changed_values(node, existing)
                prepared_changes[key] = changes
                if changes:
                    update_count += 1
                else:
                    unchanged_count += 1

            self.log.emit(
                f"План синхронизации: создать {create_count}, обновить {update_count}, "
                f"без изменений {unchanged_count}."
            )
            self.log.emit("Элементы, отсутствующие в Excel, удаляться не будут.")

            # В path_to_id входят существующие уникальные пути и элементы,
            # созданные во время текущего запуска.
            path_to_id: Dict[Tuple[str, ...], str] = {
                key: item_id(value) for key, value in index.unique.items()
                if item_id(value)
            }
            created: List[Dict[str, Any]] = []
            updated: List[Dict[str, Any]] = []
            unchanged: List[Dict[str, Any]] = []
            total = len(nodes)

            for position, node in enumerate(nodes, start=1):
                key = path_key(node.path)
                parent_key = path_key(node.path[:-1])
                parent_id = self.api.list_id if not parent_key else path_to_id.get(parent_key)
                if not parent_id:
                    raise RuntimeError(
                        f"Не найден родитель для пути: {' / '.join(node.path)}"
                    )

                source = (
                    f"строка {node.source_row_number}"
                    if node.source_row_number is not None else "автоматический родитель"
                )
                self.log.emit(f"[{position}/{total}] {source}: {' / '.join(node.path)}")

                existing = index.unique.get(key)
                if existing is not None:
                    changes = prepared_changes.get(key, {})
                    current_id = item_id(existing)
                    if changes:
                        self.log.emit("⏳ Сохраняю изменения существующего элемента...")
                        item = self._save_existing_with_recovery(
                            mapper, node, existing, changes
                        )
                        current_id = item_id(item) or current_id
                        updated.append({
                            "path": " / ".join(node.path),
                            "id": current_id,
                            "row": node.source_row_number,
                            "fields": sorted(changes),
                        })
                        self.log.emit(
                            f"✅ Обновлено: {node.name}; поля: {', '.join(sorted(changes))}"
                        )
                    else:
                        unchanged.append({
                            "path": " / ".join(node.path),
                            "id": current_id,
                            "row": node.source_row_number,
                        })
                        self.log.emit(f"— Без изменений: {node.name}")
                    path_to_id[key] = current_id
                else:
                    self.log.emit("⏳ Получаю серверный шаблон нового элемента...")
                    instance = self.api.create_new_instance(parent_id)
                    payload = mapper.make_payload(node, parent_id, instance)
                    self.log.emit("⏳ Сохраняю новый элемент...")
                    item = self._save_new_with_recovery(
                        mapper, node, parent_id, payload
                    )
                    current_id = item_id(item)
                    path_to_id[key] = current_id
                    created.append({
                        "path": " / ".join(node.path),
                        "id": current_id,
                        "row": node.source_row_number,
                    })
                    self.log.emit(f"✅ Создано: {node.name} (ID: {current_id})")

                self.progress.emit(position, total)

            self.succeeded.emit({
                "created": created,
                "updated": updated,
                "unchanged": unchanged,
                "warnings": warnings + index.warnings,
                "total": total,
            })
        except Exception as exc:
            self.log.emit(traceback.format_exc())
            self.failed.emit(str(exc))


# -----------------------------------------------------------------------------
# GUI
# -----------------------------------------------------------------------------
class PasswordEdit(QtWidgets.QLineEdit):
    def __init__(self, parent: Optional[QtWidgets.QWidget] = None):
        super().__init__(parent)
        self.setEchoMode(QtWidgets.QLineEdit.Password)
        self._visible = False
        self._button = QtWidgets.QToolButton(self)
        self._button.setText("◎")
        self._button.setCursor(QtCore.Qt.PointingHandCursor)
        self._button.clicked.connect(self._toggle)
        self.setTextMargins(0, 0, 30, 0)

    def _toggle(self) -> None:
        self._visible = not self._visible
        self.setEchoMode(QtWidgets.QLineEdit.Normal if self._visible else QtWidgets.QLineEdit.Password)
        self._button.setText("◉" if self._visible else "◎")

    def resizeEvent(self, event: QtGui.QResizeEvent) -> None:
        super().resizeEvent(event)
        self._button.setGeometry(self.width() - 29, 6, 22, max(20, self.height() - 12))


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1040, 820)
        self.setMinimumSize(860, 700)

        self.api: Optional[LarixApi] = None
        self.excel_path: Optional[str] = None
        self.excel_data: Optional[ExcelData] = None
        self.login_worker: Optional[LoginWorker] = None
        self.import_worker: Optional[ImportWorker] = None
        self.dark_theme = False

        self._build_ui()
        self._apply_styles()
        self._set_authorized(False)
        self._refresh_actions()

    # ----- UI helpers ---------------------------------------------------------
    def _card(self, title: str, icon: str, tone: str) -> Tuple[QtWidgets.QFrame, QtWidgets.QVBoxLayout]:
        frame = QtWidgets.QFrame()
        frame.setObjectName("card")
        layout = QtWidgets.QVBoxLayout(frame)
        layout.setContentsMargins(16, 14, 16, 16)
        layout.setSpacing(10)
        row = QtWidgets.QHBoxLayout()
        badge = QtWidgets.QLabel()
        badge.setObjectName("badge")
        badge.setProperty("tone", tone)
        badge.setAlignment(QtCore.Qt.AlignCenter)
        badge.setFixedSize(32, 32)
        if icon == "preview.png":
            runtime_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[3]))
            badge.setPixmap(themed_icon(runtime_root / "assets", icon, self.dark_theme).pixmap(QtCore.QSize(18, 18)))
            badge.setProperty("iconAsset", icon)
            badge.setProperty("iconSize", 18)
        else:
            badge.setText(icon)
        label = QtWidgets.QLabel(title)
        label.setObjectName("cardTitle")
        row.addWidget(badge)
        row.addWidget(label)
        row.addStretch(1)
        layout.addLayout(row)
        return frame, layout

    def _field_label(self, text: str) -> QtWidgets.QLabel:
        label = QtWidgets.QLabel(text)
        label.setObjectName("fieldLabel")
        return label

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)
        outer = QtWidgets.QVBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)

        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        outer.addWidget(scroll)

        page = QtWidgets.QWidget()
        page.setObjectName("page")
        scroll.setWidget(page)
        layout = QtWidgets.QVBoxLayout(page)
        layout.setContentsMargins(24, 22, 24, 20)
        layout.setSpacing(13)

        header = QtWidgets.QHBoxLayout()
        title_box = QtWidgets.QVBoxLayout()
        title = QtWidgets.QLabel("Синхронизация план-графика с Excel")
        title.setObjectName("pageTitle")
        subtitle = QtWidgets.QLabel(
            "Войдите в Larix, загрузите Excel и синхронизируйте структуру без дублирования."
        )
        subtitle.setObjectName("pageSubtitle")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)
        header.addLayout(title_box, 1)
        self.btn_theme = QtWidgets.QPushButton("☀  Светлая тема")
        self.btn_theme.setObjectName("themeButton")
        self.btn_theme.setFixedHeight(34)
        self.btn_theme.clicked.connect(self._toggle_theme)
        header.addWidget(self.btn_theme, 0, QtCore.Qt.AlignTop)
        layout.addLayout(header)

        # Шаг 1 — авторизация. Технические параметры API скрыты внутри программы.
        auth_card, auth = self._card("1. Подключение к Larix", "▣", "purple")
        grid = QtWidgets.QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(8)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)

        self.ed_server = QtWidgets.QLineEdit(DEFAULT_BASE_URL)
        self.ed_server.setPlaceholderText("Адрес сервера Larix")
        self.ed_login = QtWidgets.QLineEdit()
        self.ed_login.setPlaceholderText("Введите логин")
        self.ed_password = PasswordEdit()
        self.ed_password.setPlaceholderText("Введите пароль")
        self.lbl_auth = QtWidgets.QLabel("Не авторизован")
        self.lbl_auth.setObjectName("statusLabel")
        self.btn_login = QtWidgets.QPushButton("↪  Войти")
        self.btn_login.setObjectName("purpleAction")
        self.btn_login.setFixedHeight(36)
        self.btn_login.setMinimumWidth(120)
        self.btn_login.clicked.connect(self._login)
        self.ed_login.returnPressed.connect(self._login)
        self.ed_password.returnPressed.connect(self._login)

        grid.addWidget(self._field_label("Адрес Larix"), 0, 0)
        grid.addWidget(self.ed_server, 0, 1, 1, 3)
        grid.addWidget(self._field_label("Логин"), 1, 0)
        grid.addWidget(self.ed_login, 1, 1)
        grid.addWidget(self._field_label("Пароль"), 1, 2)
        grid.addWidget(self.ed_password, 1, 3)
        auth.addLayout(grid)

        auth_row = QtWidgets.QHBoxLayout()
        auth_row.setContentsMargins(0, 2, 0, 0)
        auth_row.addWidget(self.lbl_auth)
        auth_row.addStretch(1)
        auth_row.addWidget(self.btn_login)
        auth.addLayout(auth_row)
        layout.addWidget(auth_card)

        # Excel + проверка
        middle = QtWidgets.QHBoxLayout()
        middle.setSpacing(13)

        excel_card, excel = self._card("2. Загрузка Excel", "▤", "green")
        self.btn_excel = QtWidgets.QPushButton("⇧  Загрузить Excel")
        self.btn_excel.setObjectName("outlineGreen")
        self.btn_excel.setFixedHeight(36)
        self.btn_excel.clicked.connect(self._pick_excel)
        self.lbl_excel = QtWidgets.QLabel("Файл не выбран")
        self.lbl_excel.setWordWrap(True)
        self.lbl_excel.setObjectName("muted")
        self.cb_sheet = QtWidgets.QComboBox()
        self.cb_sheet.setPlaceholderText("Выберите лист")
        self.cb_sheet.currentTextChanged.connect(self._sheet_changed)
        excel.addWidget(self.btn_excel)
        excel.addWidget(self.lbl_excel)
        excel.addWidget(self._field_label("Лист с план-графиком"))
        excel.addWidget(self.cb_sheet)
        excel.addStretch(1)
        middle.addWidget(excel_card, 1)

        required_card, required = self._card("3. Проверка параметров", "✓", "orange")
        self.btn_required = QtWidgets.QPushButton("Проверить обязательные поля")
        self.btn_required.setObjectName("outlineOrange")
        self.btn_required.setFixedHeight(36)
        self.btn_required.clicked.connect(self._show_required_dialog)
        self.lbl_required = QtWidgets.QLabel(
            "Программа получит список обязательных полей из Larix и проверит наличие каждого столбца в Excel."
        )
        self.lbl_required.setObjectName("muted")
        self.lbl_required.setWordWrap(True)
        self.required_table = QtWidgets.QTableWidget(0, 3)
        self.required_table.setHorizontalHeaderLabels(["Параметр", "Источник", "Статус"])
        self.required_table.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.required_table.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        self.required_table.horizontalHeader().setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        self.required_table.setMaximumHeight(160)
        self.required_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.required_table.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        required.addWidget(self.btn_required)
        required.addWidget(self.lbl_required)
        required.addWidget(self.required_table)
        middle.addWidget(required_card, 1)
        layout.addLayout(middle)

        # Предпросмотр
        preview_card, preview = self._card("Предпросмотр данных", "preview.png", "green")
        self.preview_table = QtWidgets.QTableWidget()
        self.preview_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setMinimumHeight(210)
        self.preview_table.verticalHeader().setVisible(False)
        preview.addWidget(self.preview_table)
        layout.addWidget(preview_card)

        # Запуск
        run_card, run = self._card("4. Синхронизация план-графика", "▶", "orange")
        run_info = QtWidgets.QLabel(
            "Программа сверяет полный путь из столбца «Название» с VitroCAD: найденные элементы обновляются, отсутствующие создаются. Ничего не удаляется."
        )
        run_info.setObjectName("info")
        run_info.setWordWrap(True)
        run.addWidget(run_info)
        self.progress = QtWidgets.QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        run.addWidget(self.progress)
        run_row = QtWidgets.QHBoxLayout()
        self.btn_log = QtWidgets.QPushButton("▤  Открыть журнал")
        self.btn_log.setObjectName("secondaryButton")
        self.btn_log.clicked.connect(self._show_log)
        self.btn_import = QtWidgets.QPushButton("↻  Синхронизировать")
        self.btn_import.setObjectName("orangeAction")
        self.btn_import.setFixedHeight(40)
        self.btn_import.clicked.connect(self._start_import)
        run_row.addWidget(self.btn_log)
        run_row.addStretch(1)
        run_row.addWidget(self.btn_import)
        run.addLayout(run_row)
        layout.addWidget(run_card)

        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.hide()
        self._log_dialog: Optional[QtWidgets.QDialog] = None
        self._log_view: Optional[QtWidgets.QPlainTextEdit] = None

    def _apply_styles(self) -> None:
        dark = self.dark_theme
        page_bg = "#17191F" if dark else "#F7F8FB"
        card_bg = "#22252D" if dark else "#FFFFFF"
        field_bg = "#1B1E25" if dark else "#FBFCFE"
        text = "#F4F5F8" if dark else "#172033"
        muted = "#A2A8B6" if dark else "#687187"
        border = "#383C47" if dark else "#E1E5ED"
        purple = "#6D28D9"
        green = "#22B96B"
        orange = "#FF7A00"
        disabled = "#2B2E36" if dark else "#F1F3F7"
        table_alt = "#1D2027" if dark else "#FAFBFD"

        self.setStyleSheet(f"""
            * {{ font-family: 'Segoe UI', Arial, sans-serif; font-size: 9.5pt; color: {text}; outline: none; }}
            QMainWindow, QWidget#root, QWidget#page {{ background: {page_bg}; }}
            QScrollArea {{ border: none; background: {page_bg}; }}
            QScrollArea > QWidget > QWidget {{ background: {page_bg}; }}
            QLabel#pageTitle {{ font-size: 18pt; font-weight: 700; }}
            QLabel#pageSubtitle, QLabel#muted {{ color: {muted}; }}
            QLabel#cardTitle {{ font-size: 10.5pt; font-weight: 700; }}
            QLabel#fieldLabel {{ color: {muted}; font-size: 8.5pt; }}
            QLabel#statusLabel {{ font-weight: 700; }}
            QLabel#info {{ background: {'#3A2A18' if dark else '#FFF6EA'}; border: 1px solid {'#76512C' if dark else '#FFE0B8'}; border-radius: 8px; padding: 9px; }}
            QLabel#badge {{ border-radius: 8px; font-weight: 700; }}
            QLabel#badge[tone='purple'] {{ background: {'#332153' if dark else '#F0EAFE'}; color: {purple}; }}
            QLabel#badge[tone='green'] {{ background: {'#193D2E' if dark else '#E8F8EF'}; color: {green}; }}
            QLabel#badge[tone='orange'] {{ background: {'#4A2D16' if dark else '#FFF0E4'}; color: {orange}; }}
            QFrame#card {{ background: {card_bg}; border: 1px solid {border}; border-radius: 12px; }}
            QLineEdit, QComboBox, QDoubleSpinBox, QPlainTextEdit {{
                background: {field_bg}; border: 1px solid {border}; border-radius: 8px; min-height: 34px; padding: 0 9px;
            }}
            QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus {{ border-color: {purple}; }}
            QToolButton {{ border: none; background: transparent; color: {muted}; }}
            QTableWidget {{ background: {field_bg}; alternate-background-color: {table_alt}; border: 1px solid {border}; border-radius: 8px; gridline-color: {border}; }}
            QHeaderView::section {{ background: {disabled}; color: {text}; border: none; border-right: 1px solid {border}; border-bottom: 1px solid {border}; padding: 7px; font-weight: 600; }}
            QProgressBar {{ background: {disabled}; border: 1px solid {border}; border-radius: 7px; min-height: 16px; text-align: center; }}
            QProgressBar::chunk {{ background: {green}; border-radius: 6px; }}
            QPushButton {{ border-radius: 8px; padding: 6px 14px; font-weight: 600; }}
            QPushButton#themeButton, QPushButton#secondaryButton {{ background: {card_bg}; border: 1px solid {border}; }}
            QPushButton#themeButton:hover, QPushButton#secondaryButton:hover {{ border-color: {purple}; }}
            QPushButton#purpleAction {{ background: {purple}; color: white; border: 1px solid {purple}; }}
            QPushButton#purpleAction:hover {{ background: #7C3AED; }}
            QPushButton#orangeAction {{ background: {card_bg}; color: {text}; border: 1px solid {border}; }}
            QPushButton#orangeAction:hover {{ background: {'#302619' if dark else '#FFF7ED'}; border-color: {orange}; color: {text}; }}
            QPushButton#orangeAction:pressed {{ background: {'#3A2A18' if dark else '#FFE9D2'}; border-color: {orange}; color: {text}; }}
            QPushButton#outlineGreen {{ background: transparent; color: {green}; border: 1px solid {green}; }}
            QPushButton#outlineGreen:hover {{ background: {'#1B382B' if dark else '#F1FCF6'}; }}
            QPushButton#outlineOrange {{ background: transparent; color: {orange}; border: 1px solid {orange}; }}
            QPushButton#outlineOrange:hover {{ background: {'#3A2A18' if dark else '#FFF7ED'}; }}
            QPushButton:disabled {{ background: {disabled}; color: {muted}; border: 1px solid {border}; }}
            QCheckBox {{ spacing: 7px; }}
            QToolTip {{ background: {card_bg}; color: {text}; border: 1px solid {border}; padding: 5px; }}
        """)
        self.btn_theme.setText("☾  Тёмная тема" if not dark else "☀  Светлая тема")

    # ----- state --------------------------------------------------------------
    def _toggle_theme(self) -> None:
        self.dark_theme = not self.dark_theme
        self._apply_styles()

    def _set_authorized(self, value: bool) -> None:
        authorized = bool(value and self.api)
        self.lbl_auth.setText(
            f"✓ Подключено: {self.api.current_user_name}" if authorized else "Не авторизован"
        )
        self.lbl_auth.setStyleSheet(
            "color: #22B96B; font-weight: 700;" if authorized else "color: #777777; font-weight: 700;"
        )
        self.btn_login.setText("↻  Переподключиться" if authorized else "↪  Войти")

    def _refresh_actions(self) -> None:
        authorized = self.api is not None
        has_excel = self.excel_data is not None
        busy = bool(
            (self.login_worker and self.login_worker.isRunning())
            or (self.import_worker and self.import_worker.isRunning())
        )
        self.btn_required.setEnabled(authorized and not busy)
        self.btn_import.setEnabled(authorized and has_excel and not busy)
        self.btn_login.setEnabled(not busy)
        self.btn_excel.setEnabled(not busy)

    def _log(self, text: str) -> None:
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {text}"
        self.log.appendPlainText(line)
        if self._log_view:
            self._log_view.appendPlainText(line)
            self._log_view.verticalScrollBar().setValue(self._log_view.verticalScrollBar().maximum())

    def _show_log(self) -> None:
        if self._log_dialog and self._log_dialog.isVisible():
            self._log_dialog.raise_()
            return
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("Журнал импорта")
        dialog.resize(780, 460)
        dialog.setStyleSheet(self.styleSheet())
        dialog.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        layout = QtWidgets.QVBoxLayout(dialog)
        view = QtWidgets.QPlainTextEdit()
        view.setReadOnly(True)
        view.setPlainText(self.log.toPlainText())
        layout.addWidget(view)
        close = QtWidgets.QPushButton("Закрыть")
        close.clicked.connect(dialog.close)
        layout.addWidget(close, 0, QtCore.Qt.AlignRight)
        self._log_dialog = dialog
        self._log_view = view

        def clear_refs() -> None:
            self._log_dialog = None
            self._log_view = None

        dialog.destroyed.connect(clear_refs)
        dialog.show()

    def _message(self, title: str, text: str, error: bool = False) -> None:
        box = QtWidgets.QMessageBox(self)
        box.setWindowTitle(title)
        box.setText(text)
        box.setIcon(QtWidgets.QMessageBox.Critical if error else QtWidgets.QMessageBox.Information)
        box.setStyleSheet(self.styleSheet())
        box.exec()

    # ----- login --------------------------------------------------------------
    def _login(self) -> None:
        login = self.ed_login.text().strip()
        password = self.ed_password.text()
        base_url = self.ed_server.text().strip()
        if not all((login, password, base_url)):
            self._message("Не заполнены данные", "Укажите адрес Larix, логин и пароль.", True)
            return

        self.api = None
        self._set_authorized(False)
        self.login_worker = LoginWorker(base_url, login, password)
        self.login_worker.log.connect(self._log)
        self.login_worker.succeeded.connect(self._login_ok)
        self.login_worker.failed.connect(self._login_failed)
        self.login_worker.finished.connect(self._worker_finished)
        self.login_worker.start()
        self._refresh_actions()

    def _login_ok(self, api: LarixApi, content_type: Dict[str, Any]) -> None:
        self.api = api
        self._set_authorized(True)
        self._log(
            f"✅ Авторизация успешна. Тип: {normalize_text(content_type.get('name'))}; "
            f"полей: {len(api.fields)}; обязательных на сервере: {len(api.required_fields())}; "
            f"обязательных из Excel: {len(api.required_excel_fields())}."
        )
        self._update_required_table()

    def _login_failed(self, error: str) -> None:
        self._log("❌ Ошибка авторизации/метаданных: " + error.split("\n", 1)[0])
        self._message("Ошибка подключения", error, True)

    def _worker_finished(self) -> None:
        self._refresh_actions()

    # ----- excel --------------------------------------------------------------
    def _pick_excel(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Выберите Excel с план-графиком", "",
            "Excel files (*.xlsx *.xlsm *.xls);;All files (*.*)",
        )
        if not path:
            return
        try:
            sheets = ExcelReader.list_sheets(path)
        except Exception as exc:
            self._message("Ошибка Excel", f"Не удалось прочитать книгу:\n{exc}", True)
            return
        self.excel_path = path
        self.lbl_excel.setText(Path(path).name)
        self.lbl_excel.setToolTip(path)
        self.cb_sheet.blockSignals(True)
        self.cb_sheet.clear()
        self.cb_sheet.addItems(sheets)
        self.cb_sheet.blockSignals(False)
        if sheets:
            self.cb_sheet.setCurrentIndex(0)
            self._load_selected_sheet()

    def _sheet_changed(self, _text: str) -> None:
        self._load_selected_sheet()

    def _load_selected_sheet(self) -> None:
        if not self.excel_path or not self.cb_sheet.currentText():
            return
        try:
            self.excel_data = ExcelReader.read(self.excel_path, self.cb_sheet.currentText())
            nodes, warnings = ExcelReader.prepare_nodes(self.excel_data)
            self._log(
                f"Excel загружен: {Path(self.excel_path).name}; лист: {self.cb_sheet.currentText()}; "
                f"строк: {len(self.excel_data.dataframe)}; элементов иерархии: {len(nodes)}."
            )
            for warning in warnings[:10]:
                self._log("⚠️ " + warning)
            self._fill_preview()
            self._update_required_table()
        except Exception as exc:
            self.excel_data = None
            self.preview_table.clear()
            self._message("Ошибка листа", str(exc), True)
        self._refresh_actions()

    def _fill_preview(self) -> None:
        if not self.excel_data:
            return
        df = self.excel_data.dataframe.head(12)
        columns = [str(column) for column in df.columns]
        self.preview_table.clear()
        self.preview_table.setRowCount(len(df))
        self.preview_table.setColumnCount(len(columns))
        self.preview_table.setHorizontalHeaderLabels(columns)
        for row_index, (_, row) in enumerate(df.iterrows()):
            for col_index, column in enumerate(columns):
                value = row[column]
                text = "" if is_empty(value) else normalize_text(value)
                item = QtWidgets.QTableWidgetItem(text)
                item.setToolTip(text)
                self.preview_table.setItem(row_index, col_index, item)
        self.preview_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        if columns:
            self.preview_table.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)

    # ----- mandatory fields ---------------------------------------------------
    def _mandatory_report(self) -> List[Dict[str, str]]:
        if not self.api:
            return []
        if self.excel_data:
            return ImportMapper(self.api, self.excel_data, DEFAULT_UTC_OFFSET_HOURS).mandatory_report()

        report = []
        for field in self.api.required_excel_fields():
            report.append({
                "name": field.display_name or field.internal_name,
                "internal": field.internal_name,
                "source": "Загрузите Excel",
                "status": "ОЖИДАЕТ EXCEL",
                "details": "Обязательный параметр должен иметь отдельный столбец в Excel",
            })
        return report

    def _update_required_table(self) -> None:
        report = self._mandatory_report()
        self.required_table.setRowCount(len(report))
        missing_count = 0
        for row_index, row in enumerate(report):
            values = (row["name"], row["source"], row["status"])
            for col_index, value in enumerate(values):
                item = QtWidgets.QTableWidgetItem(value)
                item.setToolTip(row.get("details", ""))
                if row["status"] == "ОШИБКА":
                    item.setForeground(QtGui.QColor("#E53935"))
                elif row["status"] == "OK":
                    item.setForeground(QtGui.QColor("#22B96B"))
                else:
                    item.setForeground(QtGui.QColor("#FF7A00"))
                self.required_table.setItem(row_index, col_index, item)
            if row["status"] == "ОШИБКА":
                missing_count += 1

        if not self.api:
            self.lbl_required.setText("Сначала выполните вход в систему.")
        elif not self.excel_data:
            self.lbl_required.setText("Метаданные прочитаны. Загрузите Excel для полного сопоставления.")
        elif missing_count:
            self.lbl_required.setText(f"Не хватает обязательных параметров: {missing_count}.")
        else:
            self.lbl_required.setText("Все обязательные пользовательские столбцы найдены в Excel.")
        self._refresh_actions()

    def _show_required_dialog(self) -> None:
        if not self.api:
            return
        report = self._mandatory_report()
        lines = ["Обязательные поля для Excel:", ""]
        for row in report:
            lines.append(
                f"• {row['name']} — {row['source']} — {row['status']}"
            )
        lines.extend([
            "",
            "Обязательность каждого поля считывается из метаданных VitroCAD (isRequired).",
            "В списке показываются только обязательные поля, которые пользователь должен передать через Excel.",
            "Системные поля, уже заданные самой операцией создания VitroCAD, в этом списке не отображаются.",
            "Текущий пользователь и обычные серверные значения по умолчанию не подставляются вместо Excel.",
        ])
        self._message("Обязательные параметры", "\n".join(lines))

    # ----- import -------------------------------------------------------------
    def _start_import(self) -> None:
        if not self.api or not self.excel_data:
            return
        report = self._mandatory_report()
        missing = [row["name"] for row in report if row["status"] == "ОШИБКА"]
        if missing:
            self._message(
                "Проверка не пройдена",
                "Добавьте в Excel обязательные столбцы: " + ", ".join(missing),
                True,
            )
            return

        try:
            nodes, _ = ExcelReader.prepare_nodes(self.excel_data)
        except Exception as exc:
            self._message("Ошибка Excel", str(exc), True)
            return

        answer = QtWidgets.QMessageBox.question(
            self,
            "Синхронизация план-графика",
            f"Будет проверено строк структуры: {len(nodes)}.\n"
            f"Лист Excel: {self.excel_data.sheet_name}\n\n"
            "Совпавшие полные пути будут обновлены, отсутствующие — созданы.\n"
            "Элементы VitroCAD, которых нет в Excel, удаляться не будут.\n\nПродолжить?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        )
        if answer != QtWidgets.QMessageBox.Yes:
            return

        self.progress.setValue(0)
        self.import_worker = ImportWorker(self.api, self.excel_data, DEFAULT_UTC_OFFSET_HOURS)
        self.import_worker.log.connect(self._log)
        self.import_worker.progress.connect(self._import_progress)
        self.import_worker.succeeded.connect(self._import_ok)
        self.import_worker.failed.connect(self._import_failed)
        self.import_worker.finished.connect(self._worker_finished)
        self.import_worker.start()
        self._refresh_actions()

    def _import_progress(self, current: int, total: int) -> None:
        percent = int(current * 100 / max(total, 1))
        self.progress.setValue(percent)
        self.progress.setFormat(f"{current} / {total} — {percent}%")

    def _import_ok(self, result: Dict[str, Any]) -> None:
        self.progress.setValue(100)
        self.progress.setFormat("Готово")
        created = len(result.get("created", []))
        updated = len(result.get("updated", []))
        unchanged = len(result.get("unchanged", []))
        self._log(
            f"✅ Синхронизация завершена. Создано: {created}; "
            f"обновлено: {updated}; без изменений: {unchanged}."
        )
        self._message(
            "Синхронизация завершена",
            f"Создано элементов: {created}.\n"
            f"Обновлено элементов: {updated}.\n"
            f"Без изменений: {unchanged}.\n"
            f"Предупреждений: {len(result.get('warnings', []))}.\n\n"
            "Элементы, отсутствующие в Excel, не удалялись. "
            "Подробности доступны в журнале.",
        )

    def _import_failed(self, error: str) -> None:
        self._log("❌ Синхронизация остановлена: " + error.split("\n", 1)[0])
        self.progress.setFormat("Ошибка")
        self._message("Ошибка синхронизации", error, True)


# -----------------------------------------------------------------------------
def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
