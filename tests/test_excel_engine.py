import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo

from agentoffice.tools.excel import (
    ExcelEngine,
    InspectWorkbookTool,
    ReadRangeTool,
    UpdateWorkbookTool,
)


class TestExcelEngine(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.temp_path = Path(self.temp_dir.name)
        self.engine = ExcelEngine(workspace_dir=self.temp_path)

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

    def test_engine_requires_valid_directory(self):
        # Empty/None raises ValueError
        with self.assertRaises(ValueError):
            ExcelEngine(workspace_dir="")

        # Non-existent path raises FileNotFoundError
        with self.assertRaises(FileNotFoundError):
            ExcelEngine(workspace_dir=self.temp_path / "does_not_exist")

        # File instead of directory raises NotADirectoryError
        with self.assertRaises(NotADirectoryError):
            ExcelEngine(workspace_dir=self.wb_path)

    def test_multi_sheet_formulas_and_tables(self):
        result = self.engine.inspect(self.wb_path.name)

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
        result = self.engine.inspect(self.wb_path.name)
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

    def test_safe_named_range_unresolved_names_reported(self):
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

        result = self.engine.inspect("orphan.xlsx")
        self.assertEqual(result["named_slots"], [])
        self.assertIn("DeadSlot", result["unresolved_names"])

    def test_data_only_support(self):
        res_formulas = self.engine.read_range(
            self.wb_path.name, "Summary", "A4", "B4", data_only=False
        )
        self.assertEqual(res_formulas["values"], [["Net", "=B2-B3"]])

        res_data = self.engine.read_range(
            self.wb_path.name, "Summary", "A4", "B4", data_only=True
        )
        self.assertEqual(res_data["values"][0][0], "Net")
        self.assertIsNone(res_data["values"][0][1])

    def test_close_workbook_called(self):
        with patch.object(openpyxl.Workbook, "close", autospec=True) as mock_close:
            self.engine.inspect(self.wb_path.name)
            self.assertTrue(mock_close.called)

        with patch.object(openpyxl.Workbook, "close", autospec=True) as mock_close:
            self.engine.read_range(self.wb_path.name, "Summary", "A1", "B2")
            self.assertTrue(mock_close.called)

        with patch.object(openpyxl.Workbook, "close", autospec=True) as mock_close:
            self.engine.update_workbook(
                self.wb_path.name,
                "closed_test_output.xlsx",
                slots={"SingleSlot": 999},
            )
            self.assertTrue(mock_close.called)

    def test_update_workbook_cells_success(self):
        out_name = "cells_out.xlsx"
        res = self.engine.update_workbook(
            file_name=self.wb_path.name,
            output_filename=out_name,
            cells=[
                {"sheet": "Summary", "cell": "A1", "value": "UpdatedMetric"},
                {"sheet": "Breakdown", "cell": "B2", "value": 300},
            ],
        )
        self.assertEqual(res["status"], "success")
        self.assertEqual(len(res["updated_cells"]), 2)

        # Verify on-disk
        wb_out = openpyxl.load_workbook(self.temp_path / out_name, data_only=False)
        self.assertEqual(wb_out["Summary"]["A1"].value, "UpdatedMetric")
        self.assertEqual(wb_out["Breakdown"]["B2"].value, 300)
        wb_out.close()

    def test_conflicting_updates_rejected(self):
        out_name = "conflict_out.xlsx"

        # 1. Duplicate within cells
        with self.assertRaises(ValueError) as ctx:
            self.engine.update_workbook(
                file_name=self.wb_path.name,
                output_filename=out_name,
                cells=[
                    {"sheet": "Summary", "cell": "A1", "value": "First"},
                    {"sheet": "Summary", "cell": "A1", "value": "Second"},
                ],
            )
        self.assertIn("targeted multiple times", str(ctx.exception))

        # 2. Collision between cells and slots (SingleSlot points to Summary!B2)
        with self.assertRaises(ValueError) as ctx:
            self.engine.update_workbook(
                file_name=self.wb_path.name,
                output_filename=out_name,
                cells=[{"sheet": "Summary", "cell": "B2", "value": 999}],
                slots={"SingleSlot": 888},
            )
        self.assertIn("already targeted in this batch update", str(ctx.exception))

    def test_cells_formula_protection(self):
        out_name = "cells_protect_out.xlsx"
        with self.assertRaises(PermissionError) as ctx:
            self.engine.update_workbook(
                file_name=self.wb_path.name,
                output_filename=out_name,
                cells=[{"sheet": "Summary", "cell": "B4", "value": 9999}],
            )
        self.assertIn("contains a formula", str(ctx.exception))
        self.assertFalse((self.temp_path / out_name).exists())

    def test_formula_injection_prevention(self):
        out_name = "injection_out.xlsx"
        with self.assertRaises(PermissionError) as ctx:
            self.engine.update_workbook(
                file_name=self.wb_path.name,
                output_filename=out_name,
                cells=[
                    {
                        "sheet": "Summary",
                        "cell": "A1",
                        "value": "=cmd|' /C calc'!A0",
                    }
                ],
            )
        self.assertIn("Formula injection prohibited", str(ctx.exception))
        self.assertFalse((self.temp_path / out_name).exists())

    def test_merged_cell_interior_rejected(self):
        merged_wb_path = self.temp_path / "merged_wb.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "Header"
        ws.merge_cells("A1:B2")
        wb.save(merged_wb_path)
        wb.close()

        with self.assertRaises(ValueError) as ctx:
            self.engine.update_workbook(
                file_name="merged_wb.xlsx",
                output_filename="merged_out.xlsx",
                cells=[{"sheet": "Sheet1", "cell": "B2", "value": "BadChild"}],
            )
        self.assertIn("merged range", str(ctx.exception))

    def test_protected_sheet_rejection(self):
        prot_wb_path = self.temp_path / "prot_wb.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "LockedSheet"
        ws["A1"] = "Initial"
        ws.protection.sheet = True
        wb.save(prot_wb_path)
        wb.close()

        with self.assertRaises(PermissionError) as ctx:
            self.engine.update_workbook(
                file_name="prot_wb.xlsx",
                output_filename="prot_out.xlsx",
                cells=[{"sheet": "LockedSheet", "cell": "A1", "value": "Hacked"}],
            )
        self.assertIn("is protected", str(ctx.exception))

    def test_output_overwrite_policy(self):
        existing_out = self.temp_path / "already_exists.xlsx"
        existing_out.touch()

        # By default overwrite=False -> FileExistsError
        with self.assertRaises(FileExistsError):
            self.engine.update_workbook(
                file_name=self.wb_path.name,
                output_filename="already_exists.xlsx",
                slots={"SingleSlot": 500},
                overwrite=False,
            )

        # With overwrite=True -> succeeds
        res = self.engine.update_workbook(
            file_name=self.wb_path.name,
            output_filename="already_exists.xlsx",
            slots={"SingleSlot": 500},
            overwrite=True,
        )
        self.assertEqual(res["status"], "success")

    def test_table_append_collision_protection(self):
        # In self.wb_path, DeptCostTable is A1:B3. Let's put a note in row 4
        wb = openpyxl.load_workbook(self.wb_path)
        ws = wb["Breakdown"]
        ws["A4"] = "Total Depts"  # Collision!
        wb.save(self.wb_path)
        wb.close()

        with self.assertRaises(PermissionError) as ctx:
            self.engine.update_workbook(
                file_name=self.wb_path.name,
                output_filename="collision_out.xlsx",
                tables={"DeptCostTable": [{"Dept": "HR", "Cost": 100}]},
            )
        self.assertIn("Table append collision", str(ctx.exception))

    def test_read_range_max_cell_limit(self):
        with self.assertRaises(ValueError) as ctx:
            self.engine.read_range(
                file_name=self.wb_path.name,
                sheet_name="Summary",
                start_cell="A1",
                end_cell="Z500",  # 26 * 500 = 13,000 cells
            )
        self.assertIn("exceeding limit", str(ctx.exception))

    def test_path_traversal_and_symlink_rejection(self):
        with self.assertRaises(PermissionError):
            self.engine.inspect("../../etc/passwd")

        # Symlink pointing outside workspace
        outside_file = self.temp_path.parent / "outside_secret.txt"
        outside_file.touch()
        link_file = self.temp_path / "link.xlsx"
        os.symlink(outside_file, link_file)

        with self.assertRaises(PermissionError):
            self.engine.inspect("link.xlsx")

    def test_xlsm_macro_handling(self):
        xlsm_path = self.temp_path / "macro_test.xlsm"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "MacroInitial"
        wb.save(xlsm_path)
        wb.close()

        with patch("openpyxl.load_workbook", wraps=openpyxl.load_workbook) as mock_load:
            res = self.engine.update_workbook(
                file_name="macro_test.xlsm",
                output_filename="macro_out.xlsm",
                cells=[{"sheet": ws.title, "cell": "A1", "value": "MacroUpdated"}],
            )
            self.assertEqual(res["status"], "success")
            # Verify keep_vba=True was passed to all load calls for the .xlsm file
            for call in mock_load.call_args_list:
                self.assertTrue(call.kwargs.get("keep_vba"))


    def test_table_append_success(self):
        table_wb_path = self.temp_path / "table_clean.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "DataSheet"
        ws["A1"] = "Dept"
        ws["B1"] = "Cost"
        ws["A2"] = "R&D"
        ws["B2"] = 250
        tab = Table(displayName="CleanDeptTable", ref="A1:B2")
        ws.add_table(tab)
        wb.save(table_wb_path)
        wb.close()

        out_name = "table_append_out.xlsx"
        res = self.engine.update_workbook(
            file_name="table_clean.xlsx",
            output_filename=out_name,
            tables={"CleanDeptTable": [{"Dept": "Marketing", "Cost": 300}]},
        )
        self.assertEqual(len(res["updated_tables"]), 1)
        self.assertEqual(res["updated_tables"][0]["rows_appended"], 1)
        self.assertEqual(res["updated_tables"][0]["new_range"], "A1:B3")

        # Verify on-disk
        wb_out = openpyxl.load_workbook(self.temp_path / out_name, data_only=False)
        self.assertEqual(wb_out["DataSheet"]["A3"].value, "Marketing")
        self.assertEqual(wb_out["DataSheet"]["B3"].value, 300)
        wb_out.close()

    def test_table_validation_errors(self):
        out_name = "table_errors_out.xlsx"

        # 1. Unknown table
        with self.assertRaises(ValueError):
            self.engine.update_workbook(
                file_name=self.wb_path.name,
                output_filename=out_name,
                tables={"NoSuchTable": [{"A": 1}]},
            )

        # 2. Unexpected columns in dict row
        with self.assertRaises(ValueError) as ctx:
            self.engine.update_workbook(
                file_name=self.wb_path.name,
                output_filename=out_name,
                tables={
                    "DeptCostTable": [{"Dept": "IT", "Cost": 100, "ExtraColumn": 999}]
                },
            )
        self.assertIn("unexpected column", str(ctx.exception))

        # 3. Column count mismatch in list row
        with self.assertRaises(ValueError) as ctx:
            self.engine.update_workbook(
                file_name=self.wb_path.name,
                output_filename=out_name,
                tables={"DeptCostTable": [["IT", 100, "ExtraValue"]]},
            )
        self.assertIn("expects 2 columns", str(ctx.exception))

    def test_post_save_table_verification_mismatch(self):
        clean_wb = self.temp_path / "tamper_test.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "DataSheet"
        ws["A1"] = "Dept"
        ws["B1"] = "Cost"
        tab = Table(displayName="TamperTable", ref="A1:B1")
        ws.add_table(tab)
        wb.save(clean_wb)
        wb.close()

        out_name = "table_corrupt_out.xlsx"
        original_load = openpyxl.load_workbook
        calls = [0]

        def tampering_load(*args, **kwargs):
            calls[0] += 1
            wb_loaded = original_load(*args, **kwargs)
            if calls[0] == 2:  # verify_wb
                ws_loaded = wb_loaded["DataSheet"]
                if "TamperTable" in ws_loaded.tables:
                    ws_loaded.tables["TamperTable"].ref = "A1:B100"
            return wb_loaded

        with patch("openpyxl.load_workbook", side_effect=tampering_load):
            with self.assertRaises(OSError) as ctx:
                self.engine.update_workbook(
                    file_name="tamper_test.xlsx",
                    output_filename=out_name,
                    tables={"TamperTable": [{"Dept": "Support", "Cost": 120}]},
                )
            self.assertIn("range mismatch on disk", str(ctx.exception))
        self.assertFalse((self.temp_path / out_name).exists())

    def test_tools_integration(self):
        inspect_tool = InspectWorkbookTool(engine=self.engine)
        inspect_res = inspect_tool.execute(self.wb_path.name)
        self.assertIn("sheets", inspect_res)

        read_tool = ReadRangeTool(engine=self.engine)
        read_res = read_tool.execute(
            file_name=self.wb_path.name,
            sheet_name="Summary",
            start_cell="A1",
            end_cell="B2",
            data_only=False,
        )
        self.assertEqual(read_res["range"], "A1:B2")
        self.assertEqual(len(read_res["values"]), 2)

        update_tool = UpdateWorkbookTool(engine=self.engine)
        out_name = "tool_out.xlsx"
        update_res = update_tool.execute(
            file_name=self.wb_path.name,
            output_filename=out_name,
            cells=[{"sheet": "Summary", "cell": "A1", "value": "NewMetric"}],
        )
        self.assertEqual(update_res["status"], "success")
        self.assertTrue((self.temp_path / out_name).exists())

    def test_tools_return_errors_to_model(self):
        # Tools should catch exceptions and return structured {"status": "error", "error": ...}
        # so the model receives the error in its context to reflect and self-heal.
        update_tool = UpdateWorkbookTool(engine=self.engine)
        err_res = update_tool.execute(
            file_name=self.wb_path.name,
            output_filename="err_out.xlsx",
            cells=[{"sheet": "Summary", "cell": "B4", "value": 9999}],  # B4 has formula
        )
        self.assertEqual(err_res["status"], "error")
        self.assertIn("PermissionError", err_res["error"])
        self.assertIn("formula", err_res["error"])

        inspect_tool = InspectWorkbookTool(engine=self.engine)
        err_inspect = inspect_tool.execute("missing_file.xlsx")
        self.assertEqual(err_inspect["status"], "error")
        self.assertIn("FileNotFoundError", err_inspect["error"])

    def test_tools_require_engine_with_workspace(self):
        with self.assertRaises(ValueError):
            InspectWorkbookTool(engine="not_an_engine")
        with self.assertRaises(ValueError):
            ReadRangeTool(engine="not_an_engine")
        with self.assertRaises(ValueError):
            UpdateWorkbookTool(engine="not_an_engine")


if __name__ == "__main__":
    unittest.main()
