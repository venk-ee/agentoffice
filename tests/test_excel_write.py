import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
from openpyxl.worksheet.table import Table

from agentoffice.tools.excel.workspace import ExcelWorkspace
from agentoffice.tools.excel.write import (
    AppendTableHandler,
    SetCellHandler,
    UpdateWorkbookTool,
    update_workbook,
)


class TestExcelWrite(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.workspace_path = Path(self.temp_dir.name)
        self.workspace = ExcelWorkspace(workspace_dir=self.workspace_path)

        # Create base test workbook
        self.wb_path = self.workspace_path / "test_data.xlsx"
        wb = openpyxl.Workbook()

        # Sheet 1: Summary with value and formula
        ws1 = wb.active
        ws1.title = "Summary"
        ws1["A1"] = "Revenue"
        ws1["B1"] = 1000
        ws1["A2"] = "Tax"
        ws1["B2"] = "=B1*0.2"

        # Sheet 2: Orders with native table
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

    def test_cell_update_success_and_disk_verification(self):
        out_name = "test_output.xlsx"
        result = update_workbook(
            self.workspace,
            "test_data.xlsx",
            operations=[
                {"action": "set_cell", "sheet": "Summary", "cell": "B1", "value": 2500}
            ],
            output_filename=out_name,
        )

        self.assertEqual(result["status"], "success")
        self.assertEqual(result["operations_applied"], 1)

        # Verify on-disk output file
        with self.workspace.open_workbook(out_name) as wb_out:
            self.assertEqual(wb_out["Summary"]["B1"].value, 2500)

        # Verify original file remains untouched
        with self.workspace.open_workbook("test_data.xlsx") as wb_orig:
            self.assertEqual(wb_orig["Summary"]["B1"].value, 1000)

    def test_inplace_file_update(self):
        result = update_workbook(
            self.workspace,
            "test_data.xlsx",
            operations=[
                {"action": "set_cell", "sheet": "Summary", "cell": "B1", "value": 3000}
            ],
        )

        self.assertEqual(result["status"], "success")
        with self.workspace.open_workbook("test_data.xlsx") as wb:
            self.assertEqual(wb["Summary"]["B1"].value, 3000)

    def test_formula_overwrite_protection(self):
        # B2 has '=B1*0.2', overwriting must raise PermissionError
        with self.assertRaises(PermissionError):
            update_workbook(
                self.workspace,
                "test_data.xlsx",
                operations=[
                    {
                        "action": "set_cell",
                        "sheet": "Summary",
                        "cell": "B2",
                        "value": 999,
                    }
                ],
            )

        # Ensure original formula wasn't changed
        with self.workspace.open_workbook("test_data.xlsx") as wb:
            self.assertEqual(wb["Summary"]["B2"].value, "=B1*0.2")

    def test_formula_injection_protection(self):
        # Setting a cell to a formula injection payload must raise PermissionError
        with self.assertRaises(PermissionError):
            update_workbook(
                self.workspace,
                "test_data.xlsx",
                operations=[
                    {
                        "action": "set_cell",
                        "sheet": "Summary",
                        "cell": "B1",
                        "value": "=SUM(A1:A10)",
                    }
                ],
            )

    def test_merged_cell_overwrite_protection(self):
        # Create a merged cell in a new test workbook
        merged_file = "merged_data.xlsx"
        merged_path = self.workspace_path / merged_file
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "Title"
        ws.merge_cells("A1:B2")
        wb.save(merged_path)
        wb.close()

        # Writing to child cell B2 should be prohibited
        with self.assertRaises(ValueError) as ctx:
            update_workbook(
                self.workspace,
                merged_file,
                operations=[
                    {
                        "action": "set_cell",
                        "sheet": "Sheet1",
                        "cell": "B2",
                        "value": "Child",
                    }
                ],
            )
        self.assertIn("merged range", str(ctx.exception))

    def test_batch_conflict_rejection(self):
        # Targeting the same cell multiple times in the same batch
        with self.assertRaises(ValueError) as ctx:
            update_workbook(
                self.workspace,
                "test_data.xlsx",
                operations=[
                    {
                        "action": "set_cell",
                        "sheet": "Summary",
                        "cell": "B1",
                        "value": 10,
                    },
                    {
                        "action": "set_cell",
                        "sheet": "Summary",
                        "cell": "B1",
                        "value": 20,
                    },
                ],
            )
        self.assertIn("Conflicting update", str(ctx.exception))

    def test_table_append_and_range_expansion(self):
        result = update_workbook(
            self.workspace,
            "test_data.xlsx",
            operations=[
                {
                    "action": "append_table",
                    "table": "OrdersTable",
                    "rows": [["Mouse", 50], ["Keyboard", 80]],
                }
            ],
            output_filename="orders_expanded.xlsx",
        )

        self.assertEqual(result["status"], "success")

        with self.workspace.open_workbook("orders_expanded.xlsx") as wb:
            ws = wb["Orders"]
            self.assertEqual(ws.tables["OrdersTable"].ref, "A1:B4")
            self.assertEqual(ws["A3"].value, "Mouse")
            self.assertEqual(ws["B3"].value, 50)
            self.assertEqual(ws["A4"].value, "Keyboard")
            self.assertEqual(ws["B4"].value, 80)

    def test_table_landing_zone_collision(self):
        # Put data in row 3 of Orders sheet where table would expand
        with self.workspace.open_workbook("test_data.xlsx") as wb:
            wb["Orders"]["A3"] = "Existing Note"
            wb.save(self.wb_path)

        with self.assertRaises(PermissionError) as ctx:
            update_workbook(
                self.workspace,
                "test_data.xlsx",
                operations=[
                    {
                        "action": "append_table",
                        "table": "OrdersTable",
                        "rows": [["Monitor", 300]],
                    }
                ],
            )
        self.assertIn("already contains data", str(ctx.exception))

    def test_output_file_overwrite_policy(self):
        out_name = "already_exists.xlsx"
        (self.workspace_path / out_name).touch()

        # overwrite=False raises FileExistsError
        with self.assertRaises(FileExistsError):
            update_workbook(
                self.workspace,
                "test_data.xlsx",
                operations=[
                    {
                        "action": "set_cell",
                        "sheet": "Summary",
                        "cell": "B1",
                        "value": 900,
                    }
                ],
                output_filename=out_name,
                overwrite=False,
            )

        # overwrite=True succeeds
        res = update_workbook(
            self.workspace,
            "test_data.xlsx",
            operations=[
                {"action": "set_cell", "sheet": "Summary", "cell": "B1", "value": 900}
            ],
            output_filename=out_name,
            overwrite=True,
        )
        self.assertEqual(res["status"], "success")

    def test_disk_verification_failure_handling(self):
        # Simulate a disk verification failure
        with patch.object(
            SetCellHandler, "verify", side_effect=OSError("Verification check failed")
        ):
            with self.assertRaises(OSError):
                update_workbook(
                    self.workspace,
                    "test_data.xlsx",
                    operations=[
                        {
                            "action": "set_cell",
                            "sheet": "Summary",
                            "cell": "B1",
                            "value": 9999,
                        }
                    ],
                )

        # Confirm original file was not mutated
        with self.workspace.open_workbook("test_data.xlsx") as wb:
            self.assertEqual(wb["Summary"]["B1"].value, 1000)

        # Confirm temporary files were cleaned up
        tmp_files = list(self.workspace_path.glob(".tmp_*"))
        self.assertEqual(tmp_files, [])

    def test_update_workbook_tool_execution(self):
        tool = UpdateWorkbookTool(workspace=self.workspace)
        result = tool.execute(
            file_name="test_data.xlsx",
            operations=[
                {"action": "set_cell", "sheet": "Summary", "cell": "B1", "value": 5000}
            ],
        )
        self.assertEqual(result["status"], "success")
        self.assertEqual(result["operations_applied"], 1)

    def test_update_workbook_tool_error_handling(self):
        tool = UpdateWorkbookTool(workspace=self.workspace)
        # Missing file_name
        res = tool.execute(file_name="")
        self.assertEqual(res["status"], "error")
        self.assertIn("file_name", res["error"])

        # Missing operations
        res = tool.execute(file_name="test_data.xlsx", operations=[])
        self.assertEqual(res["status"], "error")
        self.assertIn("operations", res["error"])

        # Operation exception caught gracefully
        res = tool.execute(
            file_name="test_data.xlsx",
            operations=[
                {"action": "set_cell", "sheet": "Summary", "cell": "B2", "value": 123}
            ],
        )
        self.assertEqual(res["status"], "error")
        self.assertIn("PermissionError", res["error"])

    def test_table_append_dict_rows_map_by_header(self):
        # Regression: dict rows must write VALUES mapped by column name,
        # never the dict keys, and key order must not matter.
        result = update_workbook(
            self.workspace,
            "test_data.xlsx",
            operations=[
                {
                    "action": "append_table",
                    "table": "OrdersTable",
                    "rows": [
                        {"Item": "Mouse", "Price": 50},
                        {"Price": 80, "Item": "Keyboard"},
                    ],
                }
            ],
            output_filename="orders_dict_rows.xlsx",
        )

        self.assertEqual(result["status"], "success")

        with self.workspace.open_workbook("orders_dict_rows.xlsx") as wb:
            ws = wb["Orders"]
            self.assertEqual(ws["A3"].value, "Mouse")
            self.assertEqual(ws["B3"].value, 50)
            self.assertEqual(ws["A4"].value, "Keyboard")
            self.assertEqual(ws["B4"].value, 80)

    def test_table_append_dict_row_unknown_column_rejected(self):
        with self.assertRaises(ValueError) as ctx:
            update_workbook(
                self.workspace,
                "test_data.xlsx",
                operations=[
                    {
                        "action": "append_table",
                        "table": "OrdersTable",
                        "rows": [{"Item": "Mouse", "Cost": 50}],
                    }
                ],
            )
        self.assertIn("no column", str(ctx.exception))

        # Rejected batch must leave the source workbook untouched
        with self.workspace.open_workbook("test_data.xlsx") as wb:
            self.assertIsNone(wb["Orders"]["A3"].value)

    def test_table_append_rejects_invalid_row_shapes(self):
        # Regression: a bare string with len == column count previously passed
        # the width check and wrote its characters as cell values.
        with self.assertRaises(ValueError) as ctx:
            update_workbook(
                self.workspace,
                "test_data.xlsx",
                operations=[
                    {"action": "append_table", "table": "OrdersTable", "rows": ["ab"]}
                ],
            )
        self.assertIn("rows must be lists", str(ctx.exception))

        # Non-scalar values (nested containers) are rejected up front
        with self.assertRaises(ValueError) as ctx2:
            update_workbook(
                self.workspace,
                "test_data.xlsx",
                operations=[
                    {
                        "action": "append_table",
                        "table": "OrdersTable",
                        "rows": [{"Item": ["Mouse"], "Price": 50}],
                    }
                ],
            )
        self.assertIn("must be scalars", str(ctx2.exception))

    def test_append_value_verification_catches_disk_mismatch(self):
        # Tamper AFTER the real apply so the genuine verify logic has to
        # catch the corruption on the saved temp file.
        real_apply = AppendTableHandler.apply

        def corrupted_apply(wb, validated):
            real_apply(wb, validated)
            wb["Orders"]["B3"].value = 999999  # simulate silent corruption

        with patch.object(AppendTableHandler, "apply", side_effect=corrupted_apply):
            with self.assertRaises(OSError) as ctx:
                update_workbook(
                    self.workspace,
                    "test_data.xlsx",
                    operations=[
                        {
                            "action": "append_table",
                            "table": "OrdersTable",
                            "rows": [["Mouse", 50]],
                        }
                    ],
                    output_filename="orders_verify_fail.xlsx",
                )
        self.assertIn("mismatch on disk", str(ctx.exception))

        # Atomic swap never happened: no output, source untouched, tmp cleaned
        self.assertFalse((self.workspace_path / "orders_verify_fail.xlsx").exists())
        with self.workspace.open_workbook("test_data.xlsx") as wb:
            self.assertIsNone(wb["Orders"]["A3"].value)
        self.assertEqual(list(self.workspace_path.glob(".tmp_*")), [])

    def test_formulas_survive_successful_update(self):
        # Round-trip guard: an accepted update must leave unrelated formulas
        result = update_workbook(
            self.workspace,
            "test_data.xlsx",
            operations=[
                {"action": "set_cell", "sheet": "Summary", "cell": "B1", "value": 2500},
                {
                    "action": "append_table",
                    "table": "OrdersTable",
                    "rows": [{"Item": "Mouse", "Price": 50}],
                },
            ],
            output_filename="roundtrip.xlsx",
        )
        self.assertEqual(result["status"], "success")

        with self.workspace.open_workbook("roundtrip.xlsx") as wb:
            self.assertEqual(wb["Summary"]["B2"].value, "=B1*0.2")
            self.assertEqual(wb["Orders"]["A3"].value, "Mouse")
