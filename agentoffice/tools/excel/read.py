from pathlib import Path
from typing import Any

from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.worksheet import Worksheet

from agentoffice.core.tool import BaseTool
from agentoffice.tools.excel.guards import format_limit_exceeded_error
from agentoffice.tools.excel.workspace import ExcelWorkspace

DEFAULT_MAX_CELLS: int = 1000


def _resolve_sheet_and_coords(
    workbook: Any, range_str: str, file_name: str | Path
) -> tuple[Worksheet, str]:
    """Resolve worksheet and coordinate string from 'Sheet!A1' or 'A1'."""
    if "!" in range_str:
        sheet_name, coords = range_str.rsplit("!", 1)
        clean_sheet = sheet_name.strip("'").replace("''", "'")
        if clean_sheet not in workbook.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found in workbook '{file_name}'")
        return workbook[clean_sheet], coords
    return workbook.active, range_str


def _count_range_cells(ws: Worksheet, coords: str) -> int:
    """Calculate total cells in coordinates."""
    min_col, min_row, max_col, max_row = range_boundaries(coords)
    cols = (max_col or ws.max_column or 1) - (min_col or 1) + 1
    rows = (max_row or ws.max_row or 1) - (min_row or 1) + 1
    return max(0, cols) * max(0, rows)


def _extract_values_from_target(target: Any) -> list[list[Any]]:
    """Extract 2D list of cell values handling Cell, 2D tuple, and 1D tuple."""
    if not isinstance(target, tuple):
        return [[target.value]]
    if not target:
        return []
    if isinstance(target[0], tuple):
        return [[cell.value for cell in row] for row in target]
    if (
        len(target) > 1
        and hasattr(target[0], "column")
        and hasattr(target[1], "column")
    ):
        if target[0].column == target[1].column:
            return [[cell.value] for cell in target]
    return [[cell.value for cell in target]]


def read_ranges(
    workspace: ExcelWorkspace,
    file_name: str | Path,
    ranges: list[str],
    data_only: bool = False,
    max_cells: int = DEFAULT_MAX_CELLS,
) -> dict[str, Any]:
    """Reads one or more A1 ranges from a workbook in a single
    atomic turn."""
    with workspace.open_workbook(file_name, data_only=data_only) as workbook:
        initial_total = sum(
            _count_range_cells(*_resolve_sheet_and_coords(workbook, r, file_name))
            for r in ranges
        )
        if initial_total > max_cells:
            raise ValueError(format_limit_exceeded_error(initial_total, max_cells))

        accumulated_cells = 0
        value_ranges = []
        for range_str in ranges:
            ws, coords = _resolve_sheet_and_coords(workbook, range_str, file_name)
            count = _count_range_cells(ws, coords)
            if accumulated_cells + count > max_cells:
                raise ValueError(
                    format_limit_exceeded_error(accumulated_cells + count, max_cells)
                )
            accumulated_cells += count
            values = _extract_values_from_target(ws[coords])
            value_ranges.append({"range": range_str, "values": values})

        return {"file_name": str(file_name), "ranges": value_ranges}


class ReadRangesTool(BaseTool):
    def __init__(self, workspace: ExcelWorkspace):
        self.workspace = workspace
        super().__init__(
            name="read_ranges",
            description=(
                "Read one or more A1 coordinate ranges from an "
                "Excel workbook atomically. "
                "Examples: ['Summary!A1:B5', 'Orders!C1:D10'] or "
                "['A1:B10']."
            ),
            parameters={
                "type": "object",
                "properties": {
                    "file_name": {
                        "type": "string",
                        "description": "Name of the Excel workbook "
                        "in the workspace (e.g. 'data.xlsx')",
                    },
                    "ranges": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of A1 ranges to read "
                        "(e.g. ['Summary!A1:B5', 'Orders!A1:C10'])",
                    },
                    "data_only": {
                        "type": "boolean",
                        "description": "If True, returns cached "
                        "calculated numbers instead of formula "
                        "strings. Default is False.",
                    },
                },
                "required": ["file_name", "ranges"],
            },
        )

    def execute(
        self, file_name: str | Path, ranges: list[str], data_only: bool = False
    ) -> dict[str, Any]:
        if not file_name:
            return {
                "status": "error",
                "error": "ValueError: Parameter 'file_name' is required.",
            }
        if not ranges:
            return {
                "status": "error",
                "error": "ValueError: Parameter 'ranges' must be a non-empty list.",
            }
        try:
            return read_ranges(self.workspace, file_name, ranges, data_only=data_only)
        except Exception as e:
            return {"status": "error", "error": f"{type(e).__name__}: {str(e)}"}
