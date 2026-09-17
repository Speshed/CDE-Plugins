from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from projectpoint.api import get_builtin_role_permissions
from projectpoint.excel import read_objects_import_excel, write_objects_import_template, write_roles_import_template, read_roles_import_excel


ROOT = Path(__file__).resolve().parents[1]


class CoreSmokeTests(unittest.TestCase):
    def test_builtin_permissions_load(self):
        self.assertGreater(len(get_builtin_role_permissions()), 0)

    def test_objects_example_import(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "objects.xlsx"
            write_objects_import_template(path)
            records = read_objects_import_excel(path, "Объекты строительства")
            self.assertGreater(len(records), 0)

    def test_roles_template_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "roles.xlsx"
            write_roles_import_template(get_builtin_role_permissions(), target)
            roles, permission_rows = read_roles_import_excel(target)
            self.assertEqual(len(roles), 1)
            self.assertGreater(len(permission_rows), 0)


if __name__ == "__main__":
    unittest.main()
