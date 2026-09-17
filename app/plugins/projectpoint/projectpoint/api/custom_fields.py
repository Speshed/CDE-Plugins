from __future__ import annotations

from .common import api_headers

__all__ = [
    "DATA_TYPE_MAP", "DATA_TYPE_ALIASES", "CONFIRMED_DATA_TYPES", "STILL_UNCONFIRMED",
    "COLUMN_MAP", "VALID_TYPES", "BOOL_FIELDS", "NUMERIC_FIELDS",
    "get_all_custom_fields", "create_custom_field",
]


DATA_TYPE_MAP = {
    "да / нет": 0,
    "дата": 1,
    "дата и время": 2,
    "список": 3,
    "короткий текст": 4,
    "длинный текст": 5,
    "целое число": 6,
    "дробное число": 7,
}


DATA_TYPE_ALIASES = {
    "текст": "короткий текст",
    "число": "целое число",
}


CONFIRMED_DATA_TYPES = dict(DATA_TYPE_MAP)


STILL_UNCONFIRMED = {"да / нет", "дата и время", "целое число"}


COLUMN_MAP = {
    "наименование атрибута": "Name",
    "внутренее имя атрибута": "SystemName",
    "внутреннее имя атрибута": "SystemName",
    "подсказка": "Description",
    "тип данных": "DataType",
    "значения списка": "ListValuesRaw",
}


VALID_TYPES = {
    "текст", "целое число", "список", "дата", "да / нет",
    "короткий текст", "длинный текст", "дробное число",
}


BOOL_FIELDS = [
    "UseDescription", "UseDate", "UseMark", "UseRevisionReleasePlan",
    "UseApprovedPlan", "UseReleaseTarget", "UseChangeNumber", "UseDiscipline",
    "UseProjectStage", "UseLanguage", "UseWorkingDocumentationNumber",
    "UseDeveloperDocumentNumber", "UseNameEng", "UseChangePermissionNumber",
    "UseOutgoingCoverLetterNumber", "UseIncomingCoverLetterNumber",
    "UseContractStage", "UseSheetsCount", "UseContract", "UseHasPaper",
    "UseSubobjectTitle", "UseWorkCompletionCertificateNumber",
    "UseTechnicalQueryNumber", "UseExpiredDate", "IsActive",
    "HasWorkProductionDate", "UseRevisionReleaseExpectedDate",
    "UsePositionNumber", "UseInWorkProduction", "UniqueDocumentNumber",
]


NUMERIC_FIELDS = ["DisplayParameters"]


def get_all_custom_fields(sess, access_token, base_url):
    url = f"{base_url.rstrip('/')}/ru/api/Core/CustomFieldService/GetAll"
    resp = sess.get(url, headers=api_headers(access_token))
    if resp.status_code != 200:
        raise Exception(f"GetAll CustomFields error: HTTP {resp.status_code}, {resp.text[:300]}")
    return resp.json()


def create_custom_field(sess, access_token, base_url, payload):
    url = f"{base_url.rstrip('/')}/ru/api/Core/CustomFieldService/Create"
    headers = api_headers(access_token, {"Content-Type": "application/json"})
    resp = sess.post(url, headers=headers, json=payload)
    try:
        resp_data = resp.json()
    except Exception:
        resp_data = resp.text
    return resp.status_code in (200, 201), resp.status_code, resp_data
