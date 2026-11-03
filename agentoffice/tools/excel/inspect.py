from pathlib import Path
from typing import Any
import openpyxl

from agentoffice.tools.excel.workspace import ExcelWorkspace
from agentoffice.tools.excel.guards import is_formula

from agentoffice.core.tool import BaseTool

def inspect_workbook(workspace: ExcelWorkspace, file_name: str | Path) -> dict[str, Any]:

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
        formula=[
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
        return {
            "sheet_names": list(sheet_details.keys()),
            "active_sheet": wb.active.title,
            "sheet_details": sheet_details,
            "formulas": formula,
            "tables": tables,
        }


class  InspectWorkbookTool(BaseTool):
    def __init__(self, workspace: ExcelWorkspace):
        self.workspace = workspace
        super().__init__(
            name="inspect_workbook",
            description=(
                "inspect workbook and return sheet names, active sheet, "
                "sheet details, formulas, and tables"
            ),
            parameters={
                "type":"object",
                "properties":{
                    "file_name":{
                        "type":"string",
                        "description":"name of the excel workbook file in the workspace directory to inspect (e.g., 'data.xlsx')"
                    }
                },
                "required":["file_name"]
            }
        )
        
    def execute(self, file_name: str) -> dict[str, Any]:
        if not file_name:
            return {"status": "error", "error": "ValueError: Parameter 'file_name' is required."}
        try:
            return inspect_workbook(self.workspace, file_name)
        except Exception as e:
            return {"status": "error", "error": f"{type(e).__name__}: {str(e)}"}