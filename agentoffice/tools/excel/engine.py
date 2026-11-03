import os
import tempfile
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import coordinate_to_tuple, get_column_letter, range_boundaries

MAX_READ_CELLS = 10_000


class ExcelEngine:
    """Core domain engine for safe Excel (.xlsx, .xlsm) inspection and atomic updates.

    Note on Formulas:
        openpyxl preserves existing formula expressions, layout, and styling,
        but does not execute a formula recalculation engine. Cached evaluation values
        are computed and refreshed when the workbook is opened in Microsoft Excel,
        LibreOffice, or an external calculation service.
    """

    def __init__(self, workspace_dir: str | Path):
        if not workspace_dir:
            raise ValueError(
                "ExcelEngine requires a 'workspace_dir' to enforce secure filesystem boundaries."
            )
        self.workspace_dir = Path(workspace_dir).resolve()
        if not self.workspace_dir.exists():
            raise FileNotFoundError(
                f"Workspace directory not found: {self.workspace_dir}"
            )
        if not self.workspace_dir.is_dir():
            raise NotADirectoryError(
                f"Workspace path is not a directory: {self.workspace_dir}"
            )

    def _resolve_in_workspace(self, path: str | Path, must_exist: bool = True) -> Path:
        raw_path = Path(path)
        candidate = (
            (self.workspace_dir / raw_path) if not raw_path.is_absolute() else raw_path
        )
        resolved = candidate.resolve()

        if not resolved.is_relative_to(self.workspace_dir):
            raise PermissionError(
                f"Path '{path}' resolves outside workspace boundary: {self.workspace_dir}"
            )
        if must_exist and not resolved.exists():
            raise FileNotFoundError(f"File not found in workspace: {resolved.name}")
        return resolved

    @staticmethod
    def _is_formula(cell) -> bool:
        if getattr(cell, "data_type", None) == "f":
            return True
        val = getattr(cell, "value", None)
        return isinstance(val, str) and val.strip().startswith("=")

    @staticmethod
    def _iter_named_destinations(wb, dn):
        try:
            destinations = list(getattr(dn, "destinations", []))
        except Exception:
            return
        for s, ref in destinations:
            if s in wb:
                yield s, str(ref).replace("$", "")

    @staticmethod
    def _read_cell_or_range(target) -> tuple[Any, bool]:
        if isinstance(target, tuple):
            return [[c.value for c in row] for row in target], True
        return target.value, False

    def inspect(self, file_name: str | Path) -> dict[str, Any]:
        path = self._resolve_in_workspace(file_name, must_exist=True)
        keep_vba = str(path).lower().endswith(".xlsm")
        wb = openpyxl.load_workbook(path, data_only=False, keep_vba=keep_vba)
        try:
            named_slots = []
            unresolved_names = []
            for name, dn in wb.defined_names.items():
                destinations = list(self._iter_named_destinations(wb, dn))
                if not destinations:
                    unresolved_names.append(name)
                    continue
                for s, clean_ref in destinations:
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
                        unresolved_names.append(name)

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
                "file_name": path.name,
                "sheets": wb.sheetnames,
                "active_sheet": wb.active.title if wb.active else None,
                "dimensions": wb.active.dimensions if wb.active else None,
                "sheet_details": {
                    ws.title: {
                        "dimensions": ws.dimensions,
                        "max_row": ws.max_row,
                        "max_column": ws.max_column,
                        "is_protected": ws.protection.sheet,
                    }
                    for ws in wb.worksheets
                },
                "named_slots": named_slots,
                "unresolved_names": unresolved_names,
                "formulas": formulas,
                "native_tables": tables,
            }
        finally:
            wb.close()

    def read_range(
        self,
        file_name: str | Path,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        data_only: bool = False,
    ) -> dict[str, Any]:
        path = self._resolve_in_workspace(file_name, must_exist=True)
        clean_start = str(start_cell).strip().replace("$", "")
        clean_end = str(end_cell).strip().replace("$", "")

        try:
            min_col, min_row, max_col, max_row = range_boundaries(
                f"{clean_start}:{clean_end}"
            )
        except Exception as e:
            raise ValueError(
                f"Malformed range coordinates '{start_cell}:{end_cell}': {e}"
            )

        cell_count = (max_col - min_col + 1) * (max_row - min_row + 1)
        if cell_count > MAX_READ_CELLS:
            raise ValueError(
                f"Requested range '{start_cell}:{end_cell}' contains {cell_count} cells, exceeding limit of {MAX_READ_CELLS}."
            )

        keep_vba = str(path).lower().endswith(".xlsm")
        wb = openpyxl.load_workbook(path, data_only=data_only, keep_vba=keep_vba)
        try:
            if sheet_name not in wb:
                raise KeyError(f"Sheet '{sheet_name}' not found in workbook")
            ws = wb[sheet_name]
            return {
                "file_name": path.name,
                "sheet": sheet_name,
                "range": f"{clean_start}:{clean_end}",
                "values": [[c.value for c in row] for row in ws[clean_start:clean_end]],
            }
        finally:
            wb.close()

    def update_workbook(
        self,
        file_name: str | Path,
        output_filename: str | Path,
        cells: list[dict[str, Any]] | None = None,
        slots: dict[str, Any] | None = None,
        tables: dict[str, list[dict[str, Any] | list[Any]]] | None = None,
        overwrite: bool = False,
    ) -> dict[str, Any]:
        src = self._resolve_in_workspace(file_name, must_exist=True)
        dst = self._resolve_in_workspace(output_filename, must_exist=False)

        if src == dst:
            raise ValueError("Output path cannot be identical to source path.")
        if dst.exists() and not overwrite:
            raise FileExistsError(
                f"Output file '{dst.name}' already exists. Pass overwrite=True to overwrite."
            )

        if not cells and not slots and not tables:
            raise ValueError(
                "At least one update payload ('cells', 'slots', or 'tables') must be provided."
            )

        keep_vba = str(src).lower().endswith(".xlsm")
        wb = openpyxl.load_workbook(src, data_only=False, keep_vba=keep_vba)
        tmp_path = None
        try:
            # Set to detect conflicting/duplicate updates in the same transaction
            targeted_destinations: set[tuple[str, str]] = set()

            # --- Preflight Step 1: Validate explicit cell updates ---
            validated_cells = []
            for idx, entry in enumerate(cells or []):
                if not isinstance(entry, dict):
                    raise TypeError(
                        f"Cell entry at index {idx} must be a dict, got {type(entry).__name__}"
                    )
                if "cell" not in entry or "value" not in entry:
                    raise ValueError(
                        f"Cell entry at index {idx} must contain 'cell' and 'value' keys"
                    )

                val = entry["value"]
                if isinstance(val, str) and val.strip().startswith("="):
                    raise PermissionError(
                        f"Formula injection prohibited: value '{val}' for cell '{entry['cell']}' begins with '='"
                    )

                sheet_name = entry.get("sheet")
                if not sheet_name:
                    if not wb.active:
                        raise ValueError(
                            "Workbook has no active sheet; 'sheet' must be specified in cell entry"
                        )
                    sheet_name = wb.active.title

                if sheet_name not in wb:
                    raise KeyError(f"Sheet '{sheet_name}' not found in workbook")

                ws = wb[sheet_name]
                if ws.protection.sheet:
                    raise PermissionError(
                        f"Sheet '{sheet_name}' is protected. Updates prohibited."
                    )

                cell_coord = str(entry["cell"]).strip().replace("$", "")
                try:
                    coordinate_to_tuple(cell_coord)
                except Exception as e:
                    raise ValueError(
                        f"Malformed cell coordinate '{entry['cell']}': {e}"
                    )

                dest_key = (sheet_name, cell_coord)
                if dest_key in targeted_destinations:
                    raise ValueError(
                        f"Conflicting update: cell '{sheet_name}!{cell_coord}' is targeted multiple times in the same batch update."
                    )
                targeted_destinations.add(dest_key)

                target = ws[cell_coord]
                if isinstance(target, MergedCell):
                    raise ValueError(
                        f"Cell '{cell_coord}' in sheet '{sheet_name}' is part of a merged range and not the top-left cell. Overwrite prohibited."
                    )
                if self._is_formula(target):
                    raise PermissionError(
                        f"Cell '{cell_coord}' in sheet '{sheet_name}' contains a formula ({target.value}). Overwrite prohibited."
                    )

                validated_cells.append(
                    {"sheet": sheet_name, "cell": cell_coord, "value": val}
                )

            validated_slots = []
            for name, val in (slots or {}).items():
                if isinstance(val, str) and val.strip().startswith("="):
                    raise PermissionError(
                        f"Formula injection prohibited: value for slot '{name}' begins with '='"
                    )
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
                    dest_key = (s, clean_ref)
                    if dest_key in targeted_destinations:
                        raise ValueError(
                            f"Conflicting update: slot '{name}' targets cell '{s}!{clean_ref}', which is already targeted in this batch update."
                        )
                    targeted_destinations.add(dest_key)

                    ws = wb[s]
                    if ws.protection.sheet:
                        raise PermissionError(
                            f"Sheet '{s}' containing slot '{name}' is protected."
                        )
                    target = ws[clean_ref]
                    if isinstance(target, tuple):
                        raise ValueError(
                            f"Slot '{name}' references multi-cell range '{clean_ref}', which cannot be filled with a single value."
                        )
                    if isinstance(target, MergedCell):
                        raise ValueError(
                            f"Slot '{name}' references merged child cell '{clean_ref}'."
                        )
                    if self._is_formula(target):
                        raise PermissionError(
                            f"Slot '{name}' targets formula cell {clean_ref}. Overwrite prohibited."
                        )
                    validated_slots.append(
                        {"name": name, "sheet": s, "cell": clean_ref, "value": val}
                    )

            validated_tables = []
            for t_name, rows in (tables or {}).items():
                if not isinstance(rows, list):
                    raise TypeError(f"Table rows for '{t_name}' must be a list")

                ws = next((s for s in wb.worksheets if t_name in s.tables), None)
                if not ws:
                    raise ValueError(f"Table '{t_name}' not found in workbook.")
                if ws.protection.sheet:
                    raise PermissionError(
                        f"Sheet '{ws.title}' containing table '{t_name}' is protected."
                    )

                tbl = ws.tables[t_name]
                min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
                headers = [
                    ws.cell(row=min_row, column=c).value
                    for c in range(min_col, max_col + 1)
                ]

                row_data_to_append = []
                current_max_row = max_row
                for r_idx, row in enumerate(rows):
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
                        raise TypeError(
                            f"Invalid row type for table '{t_name}': {type(row).__name__}"
                        )

                    for col_idx, v in enumerate(row_vals):
                        if isinstance(v, str) and v.strip().startswith("="):
                            raise PermissionError(
                                f"Formula injection prohibited: row value for column '{headers[col_idx]}' begins with '='"
                            )

                    current_max_row += 1
                    for c_idx in range(min_col, max_col + 1):
                        cell_coord = f"{get_column_letter(c_idx)}{current_max_row}"
                        dest_key = (ws.title, cell_coord)
                        if dest_key in targeted_destinations:
                            raise ValueError(
                                f"Conflicting update: table '{t_name}' append occupies cell '{ws.title}!{cell_coord}', which is already targeted in this batch update."
                            )
                        targeted_destinations.add(dest_key)

                        existing_cell = ws.cell(row=current_max_row, column=c_idx)
                        if existing_cell.value is not None:
                            raise PermissionError(
                                f"Table append collision: cell ({ws.title}!R{current_max_row}C{c_idx}) "
                                f"already contains data or formula ('{existing_cell.value}'). Overwrite prohibited."
                            )

                    row_data_to_append.append(row_vals)

                new_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{current_max_row}"
                validated_tables.append(
                    {
                        "table_name": t_name,
                        "ws": ws,
                        "tbl": tbl,
                        "min_col": min_col,
                        "start_row": max_row + 1,
                        "rows": row_data_to_append,
                        "new_ref": new_ref,
                    }
                )

            updated_cells = []
            for c_info in validated_cells:
                wb[c_info["sheet"]][c_info["cell"]].value = c_info["value"]
                updated_cells.append(c_info)

            updated_slots = []
            for s_info in validated_slots:
                wb[s_info["sheet"]][s_info["cell"]].value = s_info["value"]
                updated_slots.append(s_info)

            updated_tables = []
            appended_cells = []
            for t_info in validated_tables:
                ws = t_info["ws"]
                tbl = t_info["tbl"]
                min_col = t_info["min_col"]
                curr_row = t_info["start_row"]
                for row_vals in t_info["rows"]:
                    for c_offset, val in enumerate(row_vals):
                        col_num = min_col + c_offset
                        ws.cell(row=curr_row, column=col_num, value=val)
                        appended_cells.append((ws.title, curr_row, col_num, val))
                    curr_row += 1
                tbl.ref = t_info["new_ref"]
                updated_tables.append(
                    {
                        "table": t_info["table_name"],
                        "sheet": ws.title,
                        "rows_appended": len(t_info["rows"]),
                        "new_range": tbl.ref,
                    }
                )

            dst.parent.mkdir(parents=True, exist_ok=True)
            suffix = dst.suffix or ".xlsx"
            with tempfile.NamedTemporaryFile(
                dir=dst.parent, prefix=".tmp_", suffix=suffix, delete=False
            ) as tmp:
                tmp_path = Path(tmp.name)

            wb.save(tmp_path)

            verify_wb = openpyxl.load_workbook(
                tmp_path, data_only=False, keep_vba=keep_vba
            )
            try:
                if not verify_wb.sheetnames:
                    raise OSError("Verification failed: saved workbook has no sheets")

                for cell_info in updated_cells:
                    saved_val = verify_wb[cell_info["sheet"]][cell_info["cell"]].value
                    if saved_val != cell_info["value"]:
                        raise OSError(
                            f"Verification failed: cell '{cell_info['sheet']}!{cell_info['cell']}' value mismatch on disk"
                        )

                for slot_info in updated_slots:
                    saved_val = verify_wb[slot_info["sheet"]][slot_info["cell"]].value
                    if saved_val != slot_info["value"]:
                        raise OSError(
                            f"Verification failed: slot '{slot_info['name']}' value mismatch on disk"
                        )

                for tbl_info in updated_tables:
                    t_name = tbl_info["table"]
                    sheet_name = tbl_info["sheet"]
                    expected_ref = tbl_info["new_range"]
                    if sheet_name not in verify_wb:
                        raise OSError(
                            f"Verification failed: sheet '{sheet_name}' for table '{t_name}' not found on disk"
                        )
                    ws_verify = verify_wb[sheet_name]
                    if t_name not in ws_verify.tables:
                        raise OSError(
                            f"Verification failed: table '{t_name}' not found on disk"
                        )
                    saved_tbl = ws_verify.tables[t_name]
                    if saved_tbl.ref != expected_ref:
                        raise OSError(
                            f"Verification failed: table '{t_name}' range mismatch on disk (expected {expected_ref}, got {saved_tbl.ref})"
                        )

                for sheet_name, r, c, expected_val in appended_cells:
                    saved_val = verify_wb[sheet_name].cell(row=r, column=c).value
                    if saved_val != expected_val:
                        raise OSError(
                            f"Verification failed: table cell ({sheet_name}!R{r}C{c}) value mismatch on disk"
                        )
            finally:
                verify_wb.close()

            os.replace(tmp_path, dst)
            tmp_path = None

            return {
                "status": "success",
                "output_file": dst.name,
                "output_path": str(dst),
                "updated_cells": updated_cells,
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
