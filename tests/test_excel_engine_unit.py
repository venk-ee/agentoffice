from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo

from agentoffice.tools.office.excel import (
    ExcelEngine,
    InspectWorkbookTool,
    ReadRangeTool,
)


class TestExcelEngine(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.temp_path = Path(self.temp_dir.name)
        self.engine = ExcelEngine()

        # Create a multi-sheet workbook with tables, formulas, and named ranges
        self.wb_path = self.temp_path / "test_multisheet.xlsx"
        wb = openpyxl.Workbook()

        # Sheet 1: Summary
        ws1 = wb.active
        ws1.title = "Summary"
        ws1["A1"] = "Metric"
        ws1["B1"] = "Value"
        ws1["A2"] = "Revenue"
        ws1["B2"] = 1000
        ws1["A3"] = "Expenses"
        ws1["B3"] = 400
        ws1["A4"] = "Net"
        ws1["B4"] = "=B2-B3"

        # Sheet 2: Breakdown
        ws2 = wb.create_sheet(title="Breakdown")
        ws2["A1"] = "Dept"
        ws2["B1"] = "Cost"
        ws2["A2"] = "R&D"
        ws2["B2"] = 250
        ws2["A3"] = "Sales"
        ws2["B3"] = 150
        ws2["A4"] = "Total"
        ws2["B4"] = "=SUM(B2:B3)"

        # Add a table to Sheet 2
        tab = Table(displayName="DeptCostTable", ref="A1:B3")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        ws2.add_table(tab)

        # Defined name 1: single cell
        wb.defined_names.add(
            openpyxl.workbook.defined_name.DefinedName(
                name="SingleSlot", attr_text="'Summary'!$B$2"
            )
        )

        # Defined name 2: multi-cell range
        wb.defined_names.add(
            openpyxl.workbook.defined_name.DefinedName(
                name="DeptRange", attr_text="'Breakdown'!$A$2:$B$3"
            )
        )

        # Save workbook
        wb.save(self.wb_path)
        wb.close()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_multi_sheet_formulas_and_tables(self):
        result = self.engine.inspect(self.wb_path)

        # Check sheets
        self.assertEqual(result["sheets"], ["Summary", "Breakdown"])

        # Check formulas from BOTH sheets
        formulas = result["formulas"]
        self.assertEqual(len(formulas), 2)
        formula_sheets = {f["sheet"] for f in formulas}
        self.assertEqual(formula_sheets, {"Summary", "Breakdown"})

        summary_formula = next(f for f in formulas if f["sheet"] == "Summary")
        self.assertEqual(summary_formula["cell"], "B4")
        self.assertEqual(summary_formula["formula"], "=B2-B3")

        breakdown_formula = next(f for f in formulas if f["sheet"] == "Breakdown")
        self.assertEqual(breakdown_formula["cell"], "B4")
        self.assertEqual(breakdown_formula["formula"], "=SUM(B2:B3)")

        # Check native tables from Sheet 2
        tables = result["native_tables"]
        self.assertEqual(len(tables), 1)
        self.assertEqual(tables[0]["sheet"], "Breakdown")
        self.assertEqual(tables[0]["name"], "DeptCostTable")
        self.assertEqual(tables[0]["range"], "A1:B3")



    def test_safe_named_ranges_single_and_multicell(self):
        result = self.engine.inspect(self.wb_path)
        slots = {s["name"]: s for s in result["named_slots"]}

        # Single cell slot
        self.assertIn("SingleSlot", slots)
        self.assertFalse(slots["SingleSlot"]["is_range"])
        self.assertEqual(slots["SingleSlot"]["value"], 1000)
        self.assertEqual(slots["SingleSlot"]["sheet"], "Summary")

        # Multi-cell range slot
        self.assertIn("DeptRange", slots)
        self.assertTrue(slots["DeptRange"]["is_range"])
        self.assertEqual(slots["DeptRange"]["sheet"], "Breakdown")
        self.assertEqual(
            slots["DeptRange"]["value"],
            [["R&D", 250], ["Sales", 150]],
        )

    def test_safe_named_range_missing_sheet(self):
        # Create a workbook with an orphaned / invalid defined name
        orphan_wb_path = self.temp_path / "orphan.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Main"
        wb.defined_names.add(
            openpyxl.workbook.defined_name.DefinedName(
                name="DeadSlot",
                attr_text="'NonExistentSheet'!$A$1",
            )
        )
        wb.save(orphan_wb_path)
        wb.close()

        # Should not crash, simply ignore missing sheet
        result = self.engine.inspect(orphan_wb_path)
        self.assertEqual(result["named_slots"], [])

    def test_data_only_support(self):
        # With data_only=False, formulas are returned
        res_formulas = self.engine.read_range(
            self.wb_path, "Summary", "A4", "B4", data_only=False
        )
        self.assertEqual(res_formulas["values"], [["Net", "=B2-B3"]])

        # With data_only=True, it reads the cached value (may be None if not evaluated by Excel)
        res_data = self.engine.read_range(
            self.wb_path, "Summary", "A4", "B4", data_only=True
        )
        self.assertEqual(res_data["values"][0][0], "Net")
        # Since openpyxl created the file without Excel opening it, cached value is None
        self.assertIsNone(res_data["values"][0][1])

    def test_close_workbook_called(self):
        with patch.object(openpyxl.Workbook, "close", autospec=True) as mock_close:
            self.engine.inspect(self.wb_path)
            self.assertTrue(mock_close.called)

        with patch.object(openpyxl.Workbook, "close", autospec=True) as mock_close:
            self.engine.read_range(self.wb_path, "Summary", "A1", "B2")
            self.assertTrue(mock_close.called)

        with patch.object(openpyxl.Workbook, "close", autospec=True) as mock_close:
            self.engine.edit_cell(self.wb_path, "Summary", "A1", "UpdatedMetric")
            self.assertTrue(mock_close.called)

    def test_error_handling(self):
        with self.assertRaises(FileNotFoundError):
            self.engine.inspect(self.temp_path / "non_existent.xlsx")

        with self.assertRaises(FileNotFoundError):
            self.engine.read_range(self.temp_path / "non_existent.xlsx", "Sheet1", "A1", "A2")

        with self.assertRaises(KeyError):
            self.engine.read_range(self.wb_path, "NoSuchSheet", "A1", "A2")

    def test_tools_integration(self):
        inspect_tool = InspectWorkbookTool(engine=self.engine)
        inspect_res = inspect_tool.execute(self.wb_path)
        self.assertIn("sheets", inspect_res)

        read_tool = ReadRangeTool(engine=self.engine)
        read_res = read_tool.execute(
            file_path=self.wb_path,
            sheet_name="Summary",
            start_cell="A1",
            end_cell="B2",
            data_only=False,
        )
        self.assertEqual(read_res["range"], "A1:B2")
        self.assertEqual(len(read_res["values"]), 2)


if __name__ == "__main__":
    unittest.main()
