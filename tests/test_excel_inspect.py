import tempfile
import unittest
from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table

from agentoffice.tools.excel.inspect import InspectWorkbookTool, inspect_workbook
from agentoffice.tools.excel.slots import classify_slot, collect_named_slots
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

        # Named slots: settable single-cell names plus non-settable shapes
        wb.defined_names.add(DefinedName("TaxRate", attr_text="Summary!$C$1"))
        ws2.defined_names.add(DefinedName("LocalNote", attr_text="Orders!$C$1"))
        wb.defined_names.add(DefinedName("Quarterly", attr_text="Summary!$A$1:$B$2"))
        wb.defined_names.add(DefinedName("Commission", attr_text="0.05"))
        wb.defined_names.add(
            DefinedName("Computed", attr_text="=SUM(Summary!$B$1:$B$2)")
        )

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

    def test_inspect_reports_named_slots_across_scopes(self):
        result = inspect_workbook(self.workspace, "test_data.xlsx")
        slots = {s["name"]: s for s in result["named_slots"]}
        self.assertEqual(slots["TaxRate"]["scope"], "workbook")
        self.assertEqual(slots["TaxRate"]["ref"], "Summary!$C$1")
        self.assertEqual(slots["LocalNote"]["scope"], "Orders")

    def test_inspect_classifies_slot_settability(self):
        result = inspect_workbook(self.workspace, "test_data.xlsx")
        slots = {s["name"]: s for s in result["named_slots"]}
        self.assertTrue(slots["TaxRate"]["settable"])
        self.assertEqual(slots["TaxRate"]["kind"], "cell")
        self.assertEqual(slots["Quarterly"]["kind"], "range")
        self.assertFalse(slots["Quarterly"]["settable"])
        self.assertEqual(slots["Commission"]["kind"], "constant")
        self.assertFalse(slots["Commission"]["settable"])
        self.assertEqual(slots["Computed"]["kind"], "formula")
        self.assertFalse(slots["Computed"]["settable"])

    def test_classify_slot_multi_destination(self):
        defn = DefinedName("TwoCells", attr_text="Summary!$A$1,Summary!$B$2")
        self.assertEqual(classify_slot(defn), "multi")

    def test_classify_slot_handles_formula_expressions_and_broken_refs(self):
        expr_defn = DefinedName("Expr", attr_text="Summary!$A$1+$B$1")
        self.assertEqual(classify_slot(expr_defn), "formula")

        ref_defn = DefinedName("Broken", attr_text="#REF!")
        self.assertEqual(classify_slot(ref_defn), "constant")

        bad_defn = DefinedName("BadCoord", attr_text="Summary!NOT_A_VALID_RANGE")
        self.assertEqual(classify_slot(bad_defn), "range")

    def test_collect_named_slots_filters_internal_and_hidden(self):
        wb = openpyxl.Workbook()
        wb.defined_names.add(DefinedName("TaxRate", attr_text="Summary!$B$2"))
        wb.defined_names.add(
            DefinedName("_xlnm.Print_Area", attr_text="Summary!$A$1:$B$2")
        )
        wb.defined_names.add(
            DefinedName("Secret", attr_text="Summary!$B$5", hidden=True)
        )
        names = [name for name, _, _ in collect_named_slots(wb)]
        self.assertEqual(names, ["TaxRate"])

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
