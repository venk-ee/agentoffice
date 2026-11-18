import tempfile
import unittest
from pathlib import Path

import openpyxl

from agentoffice.tools.excel.read import ReadRangesTool, read_ranges
from agentoffice.tools.excel.workspace import ExcelWorkspace


class TestExcelRead(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.workspace_path = Path(self.temp_dir.name)
        self.workspace = ExcelWorkspace(workspace_dir=self.workspace_path)

        # Create multi-sheet test workbook
        self.wb_path = self.workspace_path / "test_data.xlsx"
        wb = openpyxl.Workbook()

        # Sheet 1: Summary
        ws1 = wb.active
        ws1.title = "Summary"
        ws1["A1"] = "Revenue"
        ws1["B1"] = 1000
        ws1["A2"] = "Tax"
        ws1["B2"] = "=B1*0.2"

        # Sheet 2: Orders
        ws2 = wb.create_sheet(title="Orders")
        ws2["A1"] = "Laptop"
        ws2["B1"] = 5

        wb.save(self.wb_path)
        wb.close()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_read_ranges_single_sheet(self):
        result = read_ranges(self.workspace, "test_data.xlsx", ["A1:B2"])
        self.assertEqual(result["file_name"], "test_data.xlsx")
        self.assertEqual(len(result["ranges"]), 1)
        self.assertEqual(result["ranges"][0]["range"], "A1:B2")
        self.assertEqual(
            result["ranges"][0]["values"], [["Revenue", 1000], ["Tax", "=B1*0.2"]]
        )

    def test_read_ranges_quoted_sheet_name(self):
        result = read_ranges(self.workspace, "test_data.xlsx", ["'Summary'!A1:B2"])
        self.assertEqual(len(result["ranges"]), 1)
        self.assertEqual(
            result["ranges"][0]["values"], [["Revenue", 1000], ["Tax", "=B1*0.2"]]
        )

    def test_read_ranges_multi_sheet(self):
        result = read_ranges(
            self.workspace, "test_data.xlsx", ["Summary!A1:B2", "Orders!A1:B1"]
        )
        self.assertEqual(len(result["ranges"]), 2)
        self.assertEqual(result["ranges"][0]["range"], "Summary!A1:B2")
        self.assertEqual(result["ranges"][1]["range"], "Orders!A1:B1")

    def test_read_ranges_tool_execution(self):
        tool = ReadRangesTool(workspace=self.workspace)
        result = tool.execute("test_data.xlsx", ["A1:B2"])
        self.assertEqual(result["file_name"], "test_data.xlsx")
        self.assertEqual(len(result["ranges"]), 1)
        self.assertEqual(result["ranges"][0]["range"], "A1:B2")
        self.assertEqual(
            result["ranges"][0]["values"], [["Revenue", 1000], ["Tax", "=B1*0.2"]]
        )

    def test_read_ranges_tool_error_handling(self):
        tool = ReadRangesTool(workspace=self.workspace)
        result = tool.execute("nonexistent.xlsx", ["A1:B2"])
        self.assertIn("error", result)
        self.assertIn("FileNotFoundError", result["error"])

    def test_read_ranges_tool_max_cells_limit(self):
        tool = ReadRangesTool(workspace=self.workspace)
        result = tool.execute("test_data.xlsx", ["A1:B100000"])
        self.assertIn("error", result)
        self.assertIn("ValueError", result["error"])
        self.assertIn("Cell limit exceeded", result["error"])
        self.assertIn("inspect_workbook", result["error"])

    def test_read_ranges_giant_range_rejected_without_allocating_cells(self):
        import time

        start_time = time.perf_counter()
        with self.assertRaises(ValueError) as ctx:
            read_ranges(self.workspace, "test_data.xlsx", ["A1:XFD1048576"])
        elapsed = time.perf_counter() - start_time

        # Cell limit check should fail in less than 50ms without allocating 17B cells
        self.assertLess(elapsed, 0.05)
        self.assertIn("Cell limit exceeded", str(ctx.exception))
        self.assertIn("inspect_workbook", str(ctx.exception))

    def test_read_ranges_unbounded_columns_and_rows(self):
        # A:B bounded by max_row (which is 2 in Summary sheet); openpyxl column slice yields columns
        result_cols = read_ranges(self.workspace, "test_data.xlsx", ["Summary!A:B"])
        self.assertEqual(len(result_cols["ranges"]), 1)
        self.assertEqual(
            result_cols["ranges"][0]["values"],
            [["Revenue", "Tax"], [1000, "=B1*0.2"]],
        )

        # 1:2 bounded by max_column (which is 2 in Summary sheet); row slice yields rows
        result_rows = read_ranges(self.workspace, "test_data.xlsx", ["Summary!1:2"])
        self.assertEqual(len(result_rows["ranges"]), 1)
        self.assertEqual(
            result_rows["ranges"][0]["values"],
            [["Revenue", 1000], ["Tax", "=B1*0.2"]],
        )

    def test_read_ranges_cumulative_cell_limit_exceeded_on_dimension_change(self):
        # A100 is 1 cell, A:Z on a 2x2 sheet initially looks like 52 cells.
        # But reading A100 causes max_row to expand to 100, making A:Z require 2,600 cells.
        with self.assertRaises(ValueError) as ctx:
            read_ranges(
                self.workspace,
                "test_data.xlsx",
                ["Summary!A100", "Summary!A:Z"],
                max_cells=200,
            )
        self.assertIn("Cell limit exceeded", str(ctx.exception))

    def test_read_ranges_escaped_apostrophe_sheet_name(self):
        wb = openpyxl.load_workbook(self.wb_path)
        ws = wb.create_sheet(title="O'Brien")
        ws["A1"] = "Found"
        wb.save(self.wb_path)
        wb.close()

        result = read_ranges(self.workspace, "test_data.xlsx", ["'O''Brien'!A1"])
        self.assertEqual(result["ranges"][0]["values"], [["Found"]])
