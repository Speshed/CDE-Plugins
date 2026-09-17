from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from projectpoint.excel.common import read_excel_table
from projectpoint.excel.objects import read_objects_import_excel
from projectpoint.excel.parameters import read_excel as read_parameters_excel, validate_row
from projectpoint.excel.project_stages import read_project_stages_excel
from projectpoint.excel.roles import read_roles_import_excel
from projectpoint.excel.templates import (
    write_content_types_import_template,
    write_objects_import_template,
    write_params_import_template,
    write_project_stages_import_template,
    write_roles_import_template,
    write_routes_import_template,
)
from projectpoint.api.roles import get_builtin_role_permissions


class TemplateRoundtripTests(unittest.TestCase):
    def _path(self, directory: str, name: str) -> str:
        return str(Path(directory) / name)

    def test_parameters_template_is_readable_and_valid(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._path(tmp, "params.xlsx")
            write_params_import_template(path)
            records = read_parameters_excel(path, "Параметры")
            self.assertEqual(len(records), 2)
            self.assertEqual(records[0]["SystemName"], "ContractNumber")
            errors = []
            for idx, row in enumerate(records, 1):
                errors.extend(validate_row(row, idx))
            self.assertEqual(errors, [])

    def test_project_stages_template_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._path(tmp, "stages.xlsx")
            write_project_stages_import_template(path)
            rows = read_project_stages_excel(path, "Виды документов")
            self.assertEqual([row["Code"] for row in rows], ["РД", "ПД"])

    def test_content_types_template_header_and_stage_separator(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._path(tmp, "types.xlsx")
            write_content_types_import_template(path)
            df = read_excel_table(path, "Типы документов", required_headers=("Code", "Title"))
            self.assertEqual(df.iloc[0]["Code"], "DOC-GEN")
            stages = str(df.iloc[0]["ProjectStages"])
            self.assertIn("Рабочая документация", stages)
            self.assertIn(";", stages)

    def test_routes_template_rows_match_import_offsets(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._path(tmp, "routes.xlsx")
            write_routes_import_template(path)

            routes = pd.read_excel(path, sheet_name="Маршруты", header=None)
            conditions = pd.read_excel(path, sheet_name="Условия маршрута", header=None)
            duration = pd.read_excel(path, sheet_name="Продолжительность", header=None)

            # Current importer starts route data at Excel row 5.
            self.assertEqual(str(routes.iloc[4, 0]).strip(), "Маршрут 1")
            # Conditions and duration data both start at Excel row 6.
            self.assertEqual(str(conditions.iloc[5, 0]).strip(), "Маршрут 1")
            self.assertEqual(str(duration.iloc[5, 0]).strip(), "Маршрут 1")
            # Header rows are immediately above data rows.
            self.assertEqual(str(conditions.iloc[4, 0]).strip(), "Маршрут")
            self.assertEqual(str(duration.iloc[4, 0]).strip(), "Маршрут")

            wb = load_workbook(path, read_only=False)
            self.assertEqual(wb["Продолжительность"].freeze_panes, "A6")
            wb.close()

    def test_roles_template_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._path(tmp, "roles.xlsx")
            write_roles_import_template(get_builtin_role_permissions(), path)
            roles, permissions = read_roles_import_excel(path)
            self.assertIn("Новая роль", roles)
            self.assertGreater(len(permissions), 0)

    def test_objects_template_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._path(tmp, "objects.xlsx")
            write_objects_import_template(path)
            rows = read_objects_import_excel(path, "Объекты строительства")
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0]["Code"], "OBJ-001")
            self.assertEqual(rows[1]["ParentCode"], "OBJ-001")


if __name__ == "__main__":
    unittest.main()
