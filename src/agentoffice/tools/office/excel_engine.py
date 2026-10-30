import os
import tempfile
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries


class ExcelEngine:
    def __init__(self, approved_templates_dir: str | Path | None = None):
        self.approved_templates_dir = (
            Path(approved_templates_dir).resolve() if approved_templates_dir else None
        )

    def _validate_template_path(self, path: str | Path) -> Path:
        src = Path(path).resolve()
        if not src.exists():
            raise FileNotFoundError(f"Template not found: {src}")
        if self.approved_templates_dir and not src.is_relative_to(
            self.approved_templates_dir
        ):
            raise PermissionError(
                f"Template '{src}' is outside approved directory: {self.approved_templates_dir}"
            )
        return src

    @staticmethod
    def _is_formula(cell) -> bool:
        return getattr(cell, "data_type", None) == "f" or (
            isinstance(getattr(cell, "value", None), str) and cell.value.startswith("=")
        )

    @staticmethod
    def _iter_named_destinations(wb, dn):
        try:
            destinations = list(getattr(dn, "destinations", []))
        except Exception:
            return
        for s, ref in destinations:
            if s in wb:
                yield s, ref.replace("$", "")

    @staticmethod
    def _read_cell_or_range(target) -> tuple[Any, bool]:
        if isinstance(target, tuple):
            return [[c.value for c in row] for row in target], True
        return target.value, False

    def inspect(self, file_path: str | Path) -> dict[str, Any]:
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        wb = openpyxl.load_workbook(path, data_only=False)
        try:
            named_slots = []
            for name, dn in wb.defined_names.items():
                for s, clean_ref in self._iter_named_destinations(wb, dn):
                    try:
                        target = wb[s][clean_ref]
                        val, is_range = self._read_cell_or_range(target)
                        named_slots.append(
                            {
                                "name": name,
                                "sheet": s,
                                "cell": clean_ref,
                                "value": val,
                                "is_range": is_range,
                            }
                        )
                    except Exception:
                        continue

            formulas = [
                {"sheet": ws.title, "cell": c.coordinate, "formula": c.value}
                for ws in wb.worksheets
                for row in ws.iter_rows()
                for c in row
                if self._is_formula(c)
            ]

            tables = [
                {"sheet": ws.title, "name": t.name, "range": t.ref}
                for ws in wb.worksheets
                for t in ws.tables.values()
            ]

            return {
                "sheets": wb.sheetnames,
                "active_sheet": wb.active.title if wb.active else None,
                "dimensions": wb.active.dimensions if wb.active else None,
                "max_row": wb.active.max_row if wb.active else None,
                "max_column": wb.active.max_column if wb.active else None,
                "sheet_details": {
                    ws.title: {
                        "dimensions": ws.dimensions,
                        "max_row": ws.max_row,
                        "max_column": ws.max_column,
                    }
                    for ws in wb.worksheets
                },
                "named_slots": named_slots,
                "formulas": formulas,
                "native_tables": tables,
            }
        finally:
            wb.close()

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

        wb = openpyxl.load_workbook(path, data_only=data_only)
        try:
            if sheet_name not in wb:
                raise KeyError(f"Sheet '{sheet_name}' not found in workbook")
            ws = wb[sheet_name]
            return {
                "sheet": sheet_name,
                "range": f"{start_cell}:{end_cell}",
                "values": [[c.value for c in row] for row in ws[start_cell:end_cell]],
            }
        finally:
            wb.close()

    def apply_template_update(
        self,
        template_path: str | Path,
        output_path: str | Path,
        slots: dict[str, Any] | None = None,
        tables: dict[str, list[dict[str, Any] | list[Any]]] | None = None,
    ) -> dict[str, Any]:
        src = self._validate_template_path(template_path)
        dst = Path(output_path).resolve()
        if src == dst:
            raise ValueError("Output path cannot be identical to template path.")

        wb = openpyxl.load_workbook(src, data_only=False)
        tmp_path = None
        try:
            updated_slots = []
            for name, val in (slots or {}).items():
                if name not in wb.defined_names:
                    raise ValueError(f"Unknown slot '{name}'.")

                destinations = list(
                    self._iter_named_destinations(wb, wb.defined_names[name])
                )
                if not destinations:
                    raise ValueError(
                        f"Slot '{name}' has no resolvable cell destination."
                    )

                for s, clean_ref in destinations:
                    target = wb[s][clean_ref]
                    if isinstance(target, tuple):
                        raise ValueError(
                            f"Slot '{name}' references multi-cell range '{clean_ref}', which cannot be filled with a single value."
                        )

                    if self._is_formula(target):
                        raise PermissionError(
                            f"Slot '{name}' targets formula cell {clean_ref}. Overwrite prohibited."
                        )
                    target.value = val
                    updated_slots.append(
                        {"name": name, "sheet": s, "cell": clean_ref, "value": val}
                    )

            updated_tables = []
            for t_name, rows in (tables or {}).items():
                if not isinstance(rows, list):
                    raise TypeError(f"Table rows for '{t_name}' must be a list")
                ws = next((s for s in wb.worksheets if t_name in s.tables), None)
                if not ws:
                    raise ValueError(f"Table '{t_name}' not found in workbook.")

                tbl = ws.tables[t_name]
                min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
                headers = [
                    ws.cell(row=min_row, column=c).value
                    for c in range(min_col, max_col + 1)
                ]

                for row in rows:
                    if isinstance(row, dict):
                        unknown = set(row.keys()) - set(headers)
                        if unknown:
                            raise ValueError(
                                f"Table '{t_name}' row contains unexpected column(s): {unknown}"
                            )
                        row_vals = [row.get(h) for h in headers]
                    elif isinstance(row, (list, tuple)):
                        if len(row) != len(headers):
                            raise ValueError(
                                f"Table '{t_name}' expects {len(headers)} columns, got {len(row)}"
                            )
                        row_vals = list(row)
                    else:
                        raise TypeError(f"Invalid row type: {type(row).__name__}")

                    max_row += 1
                    for c_idx, val in enumerate(row_vals, start=min_col):
                        ws.cell(row=max_row, column=c_idx, value=val)

                tbl.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
                updated_tables.append(
                    {"table": t_name, "rows_appended": len(rows), "new_range": tbl.ref}
                )

            dst.parent.mkdir(parents=True, exist_ok=True)
            with tempfile.NamedTemporaryFile(
                dir=dst.parent, prefix=".tmp_", suffix=".xlsx", delete=False
            ) as tmp:
                tmp_path = Path(tmp.name)

            wb.save(tmp_path)

            verify_wb = openpyxl.load_workbook(tmp_path, data_only=False)
            try:
                if not verify_wb.sheetnames:
                    raise OSError("Verification failed: saved workbook has no sheets")
                for slot_info in updated_slots:
                    saved_val = verify_wb[slot_info["sheet"]][slot_info["cell"]].value
                    if saved_val != slot_info["value"]:
                        raise OSError(
                            f"Verification failed: slot '{slot_info['name']}' value mismatch on disk"
                        )
            finally:
                verify_wb.close()

            os.replace(tmp_path, dst)
            tmp_path = None

            return {
                "status": "success",
                "output_path": str(dst),
                "updated_slots": updated_slots,
                "updated_tables": updated_tables,
            }
        finally:
            wb.close()
            if tmp_path and tmp_path.exists():
                try:
                    tmp_path.unlink()
                except OSError:
                    pass
