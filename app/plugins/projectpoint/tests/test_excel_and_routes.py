from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from projectpoint.api.routes import get_all_departments, get_all_disciplines
from projectpoint.excel.common import read_excel_table
from projectpoint.excel.templates import write_content_types_import_template


class _Response:
    def __init__(self, data, status_code=200):
        self._data = data
        self.status_code = status_code
        self.text = ""

    def json(self):
        return self._data


class _GetSession:
    def __init__(self, data):
        self.data = data
        self.urls = []
        self.verify = True

    def get(self, url, **kwargs):
        self.urls.append(url)
        return _Response(self.data)


class ExcelAndRoutesTests(unittest.TestCase):
    def test_generated_content_types_template_header_is_detected(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "types.xlsx"
            write_content_types_import_template(str(path))
            df = read_excel_table(str(path), "Типы документов", required_headers=("Code", "Title"))
            self.assertIn("Code", df.columns)
            self.assertIn("Title", df.columns)
            self.assertEqual(df.iloc[0]["Code"], "DOC-GEN")

    def test_departments_items_wrapper_is_unwrapped(self):
        sess = _GetSession({"Items": [{"Id": "1", "FullName": "Отдел"}]})
        items = get_all_departments(sess, "token", "https://host")
        self.assertEqual(items[0]["FullName"], "Отдел")
        self.assertTrue(sess.urls[0].endswith("/ru/api/Core/DepartmentsService/GetAll"))

    def test_disciplines_read_endpoint(self):
        sess = _GetSession([{"Id": "1", "Title": "АР"}])
        items = get_all_disciplines(sess, "token", "https://host")
        self.assertEqual(items[0]["Title"], "АР")
        self.assertTrue(sess.urls[0].endswith("/ru/api/Core/DisciplineService/GetAll"))


if __name__ == "__main__":
    unittest.main()
