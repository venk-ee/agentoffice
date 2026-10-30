from pathlib import Path
from typing import Any

import openpyxl

from agentoffice.tools.base_tools import BaseTool


class ExcelEngine:
    def get_named_slots(self, wb) -> list[dict[str, Any]]:
        named_slots = []
        for name, defined_name in wb.defined_names.items():
            try:
                destinations = list(defined_name.destinations)
            except Exception:
                continue

            for sheet_title, cell_ref in destinations:
                if sheet_title not in wb:
                    continue
                clean_ref = cell_ref.replace("$", "")
                try:
                    target = wb[sheet_title][clean_ref]
                except Exception:
                    continue

                if isinstance(target, tuple):
                    values = [[cell.value for cell in row] for row in target]
                    named_slots.append(
                        {
                            "name": name,
                            "sheet": sheet_title,
                            "cell": cell_ref,
                            "value": values,
                            "is_range": True,
                        }
                    )
                else:
                    named_slots.append(
                        {
                            "name": name,
                            "sheet": sheet_title,
                            "cell": cell_ref,
                            "value": target.value,
                            "is_range": False,
                        }
                    )
        return named_slots

    def get_native_tables(self, wb) -> list[dict[str, Any]]:
        tables = []
        for ws in wb.worksheets:
            for tbl in ws.tables.values():
                tables.append({"sheet": ws.title, "name": tbl.name, "range": tbl.ref})
        return tables

    def get_formulas_cells(self, wb) -> list[dict[str, Any]]:
        formulas = []
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type == "f" or (
                        isinstance(cell.value, str) and cell.value.startswith("=")
                    ):
                        formulas.append(
                            {
                                "sheet": ws.title,
                                "cell": cell.coordinate,
                                "formula": cell.value,
                            }
                        )
        return formulas

    def read_range(
        self,
        file_path: str | Path,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        data_only: bool = False,
    ) -> dict[str, Any]:
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        wb = openpyxl.load_workbook(file_path, data_only=data_only)
        try:
            if sheet_name not in wb:
                raise KeyError(f"Sheet '{sheet_name}' not found in workbook")
            ws = wb[sheet_name]
            return {
                "sheet": sheet_name,
                "range": f"{start_cell}:{end_cell}",
                "values": [
                    [cell.value for cell in row] for row in ws[start_cell:end_cell]
                ],
            }
        finally:
            wb.close()

    def inspect(self, file_path: str | Path) -> dict[str, Any]:
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        wb = openpyxl.load_workbook(file_path, data_only=False)
        try:
            active_sheet = wb.active
            active_title = active_sheet.title if active_sheet else None
            dimensions = active_sheet.dimensions if active_sheet else None
            max_row = active_sheet.max_row if active_sheet else None
            max_column = active_sheet.max_column if active_sheet else None

            sheet_details = {}
            for ws in wb.worksheets:
                sheet_details[ws.title] = {
                    "dimensions": ws.dimensions,
                    "max_row": ws.max_row,
                    "max_column": ws.max_column,
                }

            return {
                "sheets": wb.sheetnames,
                "active_sheet": active_title,
                "dimensions": dimensions,
                "max_row": max_row,
                "max_column": max_column,
                "sheet_details": sheet_details,
                "named_slots": self.get_named_slots(wb),
                "formulas": self.get_formulas_cells(wb),
                "native_tables": self.get_native_tables(wb),
            }
        finally:
            wb.close()

    def edit_cell(
        self, file_path: str | Path, sheet_name: str, cell: str, value: Any
    ) -> None:
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        wb = openpyxl.load_workbook(file_path, data_only=False)
        try:
            if sheet_name not in wb:
                raise KeyError(f"Sheet '{sheet_name}' not found in workbook")
            ws = wb[sheet_name]
            ws[cell] = value
            wb.save(file_path)
        finally:
            wb.close()


class InspectWorkbookTool(BaseTool):
    def __init__(self, engine: ExcelEngine | None = None):
        super().__init__(
            name="inspect_workbook",
            description="Inspect an Excel workbook and return metadata about sheets, named ranges, formulas, and tables.",
            parameters={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Path to the Excel file to inspect",
                    }
                },
                "required": ["file_path"],
            },
        )
        self.engine = engine or ExcelEngine()

    def execute(self, file_path: str | Path) -> dict[str, Any]:
        return self.engine.inspect(file_path)


class ReadRangeTool(BaseTool):
    def __init__(self, engine: ExcelEngine | None = None):
        super().__init__(
            name="read_range",
            description="Read a range of cells from an Excel sheet.",
            parameters={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Path to the Excel file to read from",
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the sheet to read from",
                    },
                    "start_cell": {
                        "type": "string",
                        "description": "Start cell of the range to read (e.g., 'A1')",
                    },
                    "end_cell": {
                        "type": "string",
                        "description": "End cell of the range to read (e.g., 'B10')",
                    },
                    "data_only": {
                        "type": "boolean",
                        "description": (
                            "If True, returns cached evaluated values saved by Excel instead of formula strings. "
                            "Note: openpyxl does not compute formulas itself; uncached formulas may return None. Default is False."
                        ),
                    },
                },
                "required": ["file_path", "sheet_name", "start_cell", "end_cell"],
            },
        )
        self.engine = engine or ExcelEngine()

    def execute(
        self,
        file_path: str | Path,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        data_only: bool = False,
    ) -> dict[str, Any]:
        return self.engine.read_range(
            file_path, sheet_name, start_cell, end_cell, data_only=data_only
        )
