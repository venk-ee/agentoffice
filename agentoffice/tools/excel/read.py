from pathlib import Path
from typing import Any

from agentoffice.core.tool import BaseTool
from agentoffice.tools.excel.workspace import ExcelWorkspace

DEFAULT_MAX_CELLS: int = 1000


def read_ranges(
    workspace: ExcelWorkspace,
    file_name: str | Path,
    ranges: list[str],
    data_only: bool = False,
    max_cells: int = DEFAULT_MAX_CELLS,
) -> dict[str, Any]:
    """Reads one or more A1 ranges from a workbook in a single
    atomic turn."""
    value_ranges = []
    total_cells = 0
    with workspace.open_workbook(file_name, data_only=data_only) as workbook:
        for range_str in ranges:
            if "!" in range_str:
                sheet_name, coords = range_str.split("!", 1)
                if sheet_name not in workbook.sheetnames:
                    raise KeyError(
                        f"Sheet '{sheet_name}' not found in workbook '{file_name}'"
                    )
                ws = workbook[sheet_name]
            else:
                ws = workbook.active
                coords = range_str

            target = ws[coords]
            if isinstance(target, tuple):
                values = [[cell.value for cell in row] for row in target]
                total_cells += sum(len(row) for row in values)
            else:
                values = [[target.value]]
                total_cells += 1
            value_ranges.append({"range": range_str, "values": values})

        if total_cells > max_cells:
            raise ValueError(
                f"Requested ranges contain {total_cells} cells, exceeding "
                f"the safety ceiling of {max_cells}."
            )

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
