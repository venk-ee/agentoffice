import unittest
from unittest.mock import MagicMock

from openpyxl.cell.cell import Cell

from agentoffice.tools.excel.guards import (
    format_limit_exceeded_error,
    guard_not_formula_injection,
    is_formula,
)


class TestExcelGuards(unittest.TestCase):
    def test_guard_formula_injection_blocks_trigger_prefixes(self):
        malicious_inputs = [
            "=SUM(A1:A10)",
            "+cmd|' /C calc'!A0",
            "-1+1",
            "@HYPERLINK('http://attacker.com')",
            "\t=1+1",
            "\r-2+2",
            "\n@calc",
        ]
        for val in malicious_inputs:
            with self.subTest(val=val):
                with self.assertRaises(PermissionError) as ctx:
                    guard_not_formula_injection(val, "Test cell")
                self.assertIn("Formula injection prohibited", str(ctx.exception))

    def test_guard_formula_injection_allows_numeric_strings_and_scalars(self):
        safe_inputs = [
            "-50000",
            "+12.5",
            "-.15",
            "+0",
            "-0.0",
            -42,
            3.14159,
            0,
            True,
            False,
            None,
            "Normal text",
            "Revenue (Q3)",
        ]
        for val in safe_inputs:
            with self.subTest(val=val):
                # Should not raise
                guard_not_formula_injection(val, "Test cell")

    def test_is_formula_detection(self):
        # Explicit formula data type
        mock_f_cell = MagicMock(spec=Cell)
        mock_f_cell.data_type = "f"
        mock_f_cell.value = "100"
        self.assertTrue(is_formula(mock_f_cell))

        # Formula string representations
        for expr in ["=SUM(A1)", "+A1+B1", "-1+1", "@FUNC()"]:
            cell = MagicMock(spec=Cell)
            cell.data_type = "s"
            cell.value = expr
            self.assertTrue(is_formula(cell))

        # Safe numbers and text
        for safe in ["-500", "+3.14", "Normal text", 100, None]:
            cell = MagicMock(spec=Cell)
            cell.data_type = "n" if isinstance(safe, (int, float)) else "s"
            cell.value = safe
            self.assertFalse(is_formula(cell))

    def test_format_limit_exceeded_error_steering(self):
        msg = format_limit_exceeded_error(requested_cells=5000, limit=1000)
        self.assertIn("requested 5000 cells", msg)
        self.assertIn("maximum allowed per request is 1000", msg)
        self.assertIn("inspect_workbook", msg)
        self.assertNotIn("preview_sheet", msg)
        self.assertNotIn("paginate_table", msg)
