import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo

from agentoffice.tools.office.excel import (
    ApplyTemplateUpdateTool,
    InspectWorkbookTool,
    ReadRangeTool,
)
from agentoffice.tools.office.excel_engine import ExcelEngine


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
        tab.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9", showRowStripes=True
        )
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

        # Defined name 3: points to a formula cell
        wb.defined_names.add(
            openpyxl.workbook.defined_name.DefinedName(
                name="FormulaSlot", attr_text="'Summary'!$B$4"
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

        out_path = self.temp_path / "closed_test_output.xlsx"
        with patch.object(openpyxl.Workbook, "close", autospec=True) as mock_close:
            self.engine.apply_template_update(
                self.wb_path, out_path, slots={"SingleSlot": 999}
            )
            self.assertTrue(mock_close.called)

    def test_apply_template_update_success(self):
        out_path = self.temp_path / "report_out.xlsx"
        res = self.engine.apply_template_update(
            template_path=self.wb_path,
            output_path=out_path,
            slots={"SingleSlot": 5000},
        )
        self.assertEqual(res["status"], "success")
        self.assertEqual(len(res["updated_slots"]), 1)
        self.assertEqual(res["updated_slots"][0]["name"], "SingleSlot")
        self.assertEqual(res["updated_slots"][0]["value"], 5000)

        # Verify output file has updated value
        wb_out = openpyxl.load_workbook(out_path, data_only=False)
        self.assertEqual(wb_out["Summary"]["B2"].value, 5000)
        wb_out.close()

        # Verify ORIGINAL template was NOT modified
        wb_orig = openpyxl.load_workbook(self.wb_path, data_only=False)
        self.assertEqual(wb_orig["Summary"]["B2"].value, 1000)
        wb_orig.close()

    def test_apply_template_update_formula_protection(self):
        out_path = self.temp_path / "protected_out.xlsx"
        # Overwriting a formula cell must raise PermissionError and abort transaction
        with self.assertRaises(PermissionError):
            self.engine.apply_template_update(
                template_path=self.wb_path,
                output_path=out_path,
                slots={"FormulaSlot": 999999},
            )
        # Verify transaction was completely aborted and no output file was created
        self.assertFalse(out_path.exists())

    def test_apply_template_update_unknown_slot_fails_transaction(self):
        out_path = self.temp_path / "unknown_slot_out.xlsx"
        # Unknown slots must raise ValueError and abort transaction
        with self.assertRaises(ValueError) as ctx:
            self.engine.apply_template_update(
                template_path=self.wb_path,
                output_path=out_path,
                slots={"GhostSlot": 1234},
            )
        self.assertIn("Unknown slot 'GhostSlot'", str(ctx.exception))
        self.assertFalse(out_path.exists())

    def test_apply_template_update_overwriting_template_forbidden(self):
        # Setting output_path == template_path must fail to protect the template
        with self.assertRaises(ValueError):
            self.engine.apply_template_update(
                template_path=self.wb_path,
                output_path=self.wb_path,
                slots={"SingleSlot": 100},
            )

    def test_apply_template_update_table_validation(self):
        out_path = self.temp_path / "table_val_out.xlsx"

        # 1. Unknown table
        with self.assertRaises(ValueError):
            self.engine.apply_template_update(
                self.wb_path, out_path, tables={"NoSuchTable": [{"A": 1}]}
            )

        # 2. Unexpected columns in dict row
        with self.assertRaises(ValueError) as ctx:
            self.engine.apply_template_update(
                self.wb_path,
                out_path,
                tables={
                    "DeptCostTable": [{"Dept": "IT", "Cost": 100, "ExtraColumn": 999}]
                },
            )
        self.assertIn("unexpected column", str(ctx.exception))

        # 3. Column count mismatch in list row
        with self.assertRaises(ValueError) as ctx:
            self.engine.apply_template_update(
                self.wb_path,
                out_path,
                tables={"DeptCostTable": [["IT", 100, "ExtraValue"]]},
            )
        self.assertIn("expects 2 columns", str(ctx.exception))

        # 4. Invalid row type
        with self.assertRaises(TypeError):
            self.engine.apply_template_update(
                self.wb_path, out_path, tables={"DeptCostTable": ["not a dict or list"]}
            )

    def test_approved_template_enforcement(self):
        approved_dir = self.temp_path / "approved_templates"
        approved_dir.mkdir()
        engine_strict = ExcelEngine(approved_templates_dir=approved_dir)

        # self.wb_path is in self.temp_path, NOT in approved_dir
        with self.assertRaises(PermissionError):
            engine_strict.apply_template_update(
                template_path=self.wb_path,
                output_path=self.temp_path / "out.xlsx",
                slots={"SingleSlot": 10},
            )

    def test_atomic_save_cleans_up_temp_on_failure(self):
        out_path = self.temp_path / "failure_out.xlsx"
        # Mock load_workbook to fail during reopen verification
        original_load = openpyxl.load_workbook
        calls = [0]

        def failing_load(*args, **kwargs):
            calls[0] += 1
            if calls[0] == 2:  # 1st is src_path load, 2nd is verify_wb
                raise OSError("Simulated corruption during verification")
            return original_load(*args, **kwargs)

        with patch("openpyxl.load_workbook", side_effect=failing_load):
            with self.assertRaises(IOError):
                self.engine.apply_template_update(
                    self.wb_path, out_path, slots={"SingleSlot": 10}
                )

        # Verify no .tmp_*.xlsx files left behind in temp_path
        tmp_files = list(self.temp_path.glob(".tmp_*.xlsx"))
        self.assertEqual(tmp_files, [])
        self.assertFalse(out_path.exists())

    def test_apply_template_update_table_append(self):
        out_path = self.temp_path / "table_out.xlsx"
        res = self.engine.apply_template_update(
            template_path=self.wb_path,
            output_path=out_path,
            tables={"DeptCostTable": [{"Dept": "Marketing", "Cost": 300}]},
        )
        self.assertEqual(len(res["updated_tables"]), 1)
        self.assertEqual(res["updated_tables"][0]["rows_appended"], 1)
        self.assertEqual(res["updated_tables"][0]["new_range"], "A1:B4")

        # Verify row exists in new file
        wb_out = openpyxl.load_workbook(out_path, data_only=False)
        ws2 = wb_out["Breakdown"]
        self.assertEqual(ws2["A4"].value, "Marketing")
        self.assertEqual(ws2["B4"].value, 300)
        wb_out.close()

    def test_error_handling(self):
        with self.assertRaises(FileNotFoundError):
            self.engine.inspect(self.temp_path / "non_existent.xlsx")

        with self.assertRaises(FileNotFoundError):
            self.engine.read_range(
                self.temp_path / "non_existent.xlsx", "Sheet1", "A1", "A2"
            )

        with self.assertRaises(KeyError):
            self.engine.read_range(self.wb_path, "NoSuchSheet", "A1", "A2")

        with self.assertRaises(FileNotFoundError):
            self.engine.apply_template_update(
                self.temp_path / "non_existent.xlsx",
                self.temp_path / "out.xlsx",
                slots={},
            )

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

        update_tool = ApplyTemplateUpdateTool(engine=self.engine)
        out_path = self.temp_path / "tool_out.xlsx"
        update_res = update_tool.execute(
            template_path=self.wb_path,
            output_path=out_path,
            slots={"SingleSlot": 7777},
        )
        self.assertEqual(update_res["status"], "success")
        self.assertTrue(out_path.exists())

    def test_apply_template_update_multicell_range_rejected(self):
        out_path = self.temp_path / "range_rejected.xlsx"
        # Updating a multi-cell range like DeptRange ('Breakdown'!$A$2:$B$3) with a scalar must fail
        with self.assertRaises(ValueError) as ctx:
            self.engine.apply_template_update(
                template_path=self.wb_path,
                output_path=out_path,
                slots={"DeptRange": "NotA2DMatrix"},
            )
        self.assertIn("multi-cell range", str(ctx.exception))
        self.assertFalse(out_path.exists())

    def test_inspect_safe_with_constant_defined_name(self):
        const_wb_path = self.temp_path / "const_wb.xlsx"
        wb = openpyxl.Workbook()
        # Add a constant defined name (not pointing to a cell destination)
        wb.defined_names.add(
            openpyxl.workbook.defined_name.DefinedName(name="TaxRate", attr_text="0.15")
        )
        wb.save(const_wb_path)
        wb.close()

        # inspect must handle it safely without raising exceptions
        res = self.engine.inspect(const_wb_path)
        self.assertIn("sheets", res)

    def test_template_id_resolution_in_tool(self):
        # Create an approved template directory with a named template
        approved_dir = self.temp_path / "approved"
        approved_dir.mkdir()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Test"
        wb.defined_names.add(
            openpyxl.workbook.defined_name.DefinedName(
                name="CompanySlot", attr_text="'Sheet'!$A$1"
            )
        )
        template_file = approved_dir / "burn_report.xlsx"
        wb.save(template_file)
        wb.close()

        engine = ExcelEngine(approved_templates_dir=approved_dir)
        tool = ApplyTemplateUpdateTool(engine=engine)
        out_file = self.temp_path / "resolved_out.xlsx"
        res = tool.execute(
            template_id="burn_report",
            output_filename=str(out_file),
            slots={"CompanySlot": "ResolvedCorp"},
        )
        self.assertEqual(res["status"], "success")
        self.assertTrue(out_file.exists())


if __name__ == "__main__":
    unittest.main()
