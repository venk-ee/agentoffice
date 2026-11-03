import tempfile
import unittest
from pathlib import Path

import openpyxl

from agentoffice.tools.excel.workspace import ExcelWorkspace


class TestExcelWorkspace(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.temp_path = Path(self.temp_dir.name)
        self.workspace = ExcelWorkspace(workspace_dir=self.temp_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_resolve_path(self):
        # Test absolute path
        abs_path = self.temp_path / "test.xlsx"
        resolved = self.workspace.resolve_path(abs_path, must_exist=False)
        self.assertEqual(resolved, abs_path.resolve())

        # Test relative path
        rel_path = "test.xlsx"
        resolved = self.workspace.resolve_path(rel_path, must_exist=False)
        self.assertEqual(resolved, (self.temp_path / rel_path).resolve())

        # Test path outside workspace
        outside_path = Path("/tmp/test.xlsx")
        with self.assertRaises(PermissionError):
            self.workspace.resolve_path(outside_path, must_exist=False)

    def test_open_workbook(self):
        # Test opening a non-existent file
        with self.assertRaises(FileNotFoundError):
            with self.workspace.open_workbook("nonexistent.xlsx", must_exist=True):
                pass

        test_file = self.temp_path / "test.xlsx"
        wb = openpyxl.Workbook()
        wb.active["A1"] = "Hello"
        wb.save(test_file)
        wb.close()

        with self.workspace.open_workbook(test_file, must_exist=True) as open_wb:
            self.assertEqual(open_wb.active["A1"].value, "Hello")
