from pathlib import Path
from typing import Any

from agentoffice.core.tool import BaseTool
from agentoffice.tools.excel.guards import is_formula
from agentoffice.tools.excel.slots import classify_slot, collect_named_slots
from agentoffice.tools.excel.workspace import ExcelWorkspace


def inspect_workbook(
    workspace: ExcelWorkspace, file_name: str | Path
) -> dict[str, Any]:
    """Inspect workbook structure, sheet dimensions, formula locations, native tables, and named slots.

    Args:
        workspace: Workspace environment managing the file.
        file_name: Name or relative path of the Excel workbook in the workspace.

    Returns:
        A dictionary containing sheet_names, active_sheet, sheet_details,
        formulas, tables, and discoverable named_slots.
    """
    with workspace.open_workbook(file_name) as wb:
        sheet_details = {
            ws.title: {
                "dimensions": ws.dimensions,
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "is_protected": ws.protection.sheet,
            }
            for ws in wb.worksheets
        }
        formula = [
            {"sheet": ws.title, "cell": cell.coordinate, "formula": cell.value}
            for ws in wb.worksheets
            for row in ws.iter_rows()
            for cell in row
            if is_formula(cell)
        ]
        tables = [
            {"sheet": ws.title, "name": t.name, "range": t.ref}
            for ws in wb.worksheets
            for t in ws.tables.values()
        ]
        named_slots = [
            {
                "name": name,
                "scope": scope,
                "ref": defn.attr_text or defn.value,
                "kind": (kind := classify_slot(defn)),
                "settable": kind == "cell",
            }
            for name, scope, defn in collect_named_slots(wb)
        ]
        return {
            "sheet_names": list(sheet_details.keys()),
            "active_sheet": wb.active.title,
            "sheet_details": sheet_details,
            "formulas": formula,
            "tables": tables,
            "named_slots": named_slots,
        }


class InspectWorkbookTool(BaseTool):
    """Tool adapter exposing workbook structural inspection to the agent."""

    def __init__(self, workspace: ExcelWorkspace):
        self.workspace = workspace
        super().__init__(
            name="inspect_workbook",
            description=(
                "Inspect workbook structure and return sheet names, active sheet, "
                "sheet details, formulas, tables, and named_slots (defined names "
                "with a 'kind' and 'settable' flag; only settable=true, kind='cell' "
                "names can be updated via the set_slot action of update_workbook)."
            ),
            parameters={
                "type": "object",
                "properties": {
                    "file_name": {
                        "type": "string",
                        "description": "Name of the Excel workbook file in the workspace directory to inspect (e.g. 'data.xlsx').",
                    }
                },
                "required": ["file_name"],
            },
        )

    def execute(self, file_name: str) -> dict[str, Any]:
        if not file_name:
            return {
                "status": "error",
                "error": "ValueError: Parameter 'file_name' is required.",
            }
        try:
            return inspect_workbook(self.workspace, file_name)
        except Exception as e:
            return {"status": "error", "error": f"{type(e).__name__}: {str(e)}"}
