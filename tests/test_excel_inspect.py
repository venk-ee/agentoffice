import tempfile
import unittest
from pathlib import Path

import openpyxl
from openpyxl.worksheet.table import Table

from agentoffice.tools.excel.inspect import InspectWorkbookTool, inspect_workbook
from agentoffice.tools.excel.workspace import ExcelWorkspace


class TestExcelInspect(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.workspace_path = Path(self.temp_dir.name)
        self.workspace = ExcelWorkspace(workspace_dir=self.workspace_path)

        # Create a test workbook: test_data.xlsx
        self.wb_path = self.workspace_path / "test_data.xlsx"
        wb = openpyxl.Workbook()

        # Sheet 1: Summary with a formula
        ws1 = wb.active
        ws1.title = "Summary"
        ws1["A1"] = "Revenue"
        ws1["B1"] = 1000
        ws1["A2"] = "Tax"
        ws1["B2"] = "=B1*0.2"

        # Sheet 2: Orders with a native table
        ws2 = wb.create_sheet(title="Orders")
        ws2["A1"] = "Item"
        ws2["B1"] = "Price"
        ws2["A2"] = "Laptop"
        ws2["B2"] = 1200

        tab = Table(displayName="OrdersTable", ref="A1:B2")
        ws2.add_table(tab)

        wb.save(self.wb_path)
        wb.close()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_inspect_workbook_finds_sheets_formulas_and_tables(self):
        result = inspect_workbook(self.workspace, "test_data.xlsx")
        self.assertIn("sheet_names", result)
        self.assertIn("formulas", result)
        self.assertIn("tables", result)
        self.assertEqual(result["sheet_names"], ["Summary", "Orders"])
        self.assertEqual(result["formulas"][0]["formula"], "=B1*0.2")
        self.assertEqual(result["tables"][0]["name"], "OrdersTable")

    def test_inspect_tool_execution(self):
        tool = InspectWorkbookTool(workspace=self.workspace)
        result = tool.execute("test_data.xlsx")
        self.assertEqual(result["active_sheet"], "Summary")
        self.assertEqual(result["sheet_names"], ["Summary", "Orders"])
        self.assertEqual(result["tables"][0]["name"], "OrdersTable")

    def test_inspect_tool_error_handling(self):
        tool = InspectWorkbookTool(workspace=self.workspace)
        result = tool.execute("nonexistent.xlsx")
        self.assertIn("error", result)
        self.assertIn("FileNotFoundError", result["error"])
