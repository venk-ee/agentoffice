from pathlib import Path
from typing import Any

import openpyxl


class ExcelEngine:
    def get_named_slots(self, wb) -> list[dict[str, Any]]:
        named_slots = []
        for name, defined_name in wb.defined_names.items():
            for sheet_title, cell_ref in defined_name.destinations:
                val = wb[sheet_title][cell_ref.replace("$", "")].value
                named_slots.append(
                    {"name": name, "sheet": sheet_title, "cell": cell_ref, "value": val}
                )
        return named_slots

    def get_native_tabels(self, wb_tables) -> list[dict[str, Any]]:
        tables = []
        for tbl in wb_tables.values():
            tables.append({"name": tbl.name, "range": tbl.ref})
        return tables

    def get_formulas_cells(self, wb) -> list[dict[str, Any]]:
        formulas = []
        for row in wb.active.iter_rows():
            for cell in row:
                if str(cell.value).startswith("="):
                    formulas.append({"cell": cell.coordinate, "formula": cell.value})
        return formulas

    def read_range(
        self, file_path: str | Path, sheet_name: str, start_cell: str, end_cell: str
    ) -> list[list[Any]]:
        wb = openpyxl.load_workbook(file_path, data_only=False)
        ws = wb[sheet_name]
        return [[cell.value for cell in row] for row in ws[start_cell:end_cell]]

    def inspect(self, file_path: str | Path) -> dict[str, Any]:
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        wb = openpyxl.load_workbook(file_path, data_only=False)
        # overvewiew ie meatadata of the excel of then sheet
        return {
            "sheets": wb.sheetnames,
            "active_sheet": wb.active.title,
            "dimensions": wb.active.dimensions,
            "max_row": wb.active.max_row,
            "max_column": wb.active.max_column,
            "named_slots": self.get_named_slots(wb),
            "formulas": self.get_formulas_cells(wb),
            "native_tables": self.get_native_tabels(wb.active.tables),
        }
