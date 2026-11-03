import os
from pathlib import Path
import tempfile
from typing import Any

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.cell import get_column_letter, range_boundaries

from agentoffice.core.tool import BaseTool
from agentoffice.tools.excel.guards import (
    clean_coord,
    guard_not_formula_injection,
    guard_sheet_unprotected,
    is_formula,
)
from agentoffice.tools.excel.workspace import ExcelWorkspace


class ConflictTracker:
    """Tracks targeted cell coordinates during batch updates to
    detect internal collisions."""

    def __init__(self) -> None:
        self._targets: set[tuple[str, str]] = set()

    def track(self, sheet: str, cell_coord: str, error_message: str) -> None:
        dest_key = (sheet, cell_coord)
        if dest_key in self._targets:
            raise ValueError(error_message)
        self._targets.add(dest_key)


class SetCellHandler:
    @staticmethod
    def validate(
        wb: openpyxl.Workbook,
        operation: dict[str, Any],
        tracker: ConflictTracker,
    ) -> dict[str, Any]:

        if "cell" not in operation or "value" not in operation:
            raise ValueError("Operation must contain 'cell' and 'value' keys")

        value = operation["value"]
        guard_not_formula_injection(value, f"Value at cell {operation['cell']}")

        sheet_name = operation.get("sheet") or (wb.active.title if wb.active else None)
        if not sheet_name or sheet_name not in wb.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found in workbook")

        ws = wb[sheet_name]
        guard_sheet_unprotected(ws, f"Sheet '{sheet_name}' is protected")

        cell_coord = clean_coord(operation["cell"])
        tracker.track(
            sheet_name,
            cell_coord,
            f"Conflicting update: cell '{sheet_name}!{cell_coord}' targeted multiple times in same batch.",
        )

        target = ws[cell_coord]
        if isinstance(target, MergedCell):
            raise ValueError(
                f"Cell '{cell_coord}' is part of a merged range. Overwrite prohibited."
            )
        if is_formula(target):
            raise PermissionError(
                f"Cell '{cell_coord}' contains formula ({target.value}). Overwrite prohibited."
            )

        return {"sheet": sheet_name, "cell": cell_coord, "value": value}

    @staticmethod
    def apply(wb: openpyxl.Workbook, validated: dict[str, Any]) -> None:
        wb[validated["sheet"]][validated["cell"]].value = validated["value"]

    @staticmethod
    def verify(verify_wb: openpyxl.Workbook, validated: dict[str, Any]) -> None:
        saved_val = verify_wb[validated["sheet"]][validated["cell"]].value
        if saved_val != validated["value"]:
            raise OSError(
                f"Disk verification failed for cell '{validated['sheet']}!{validated['cell']}'"
            )


class AppendTableHandler:
    @staticmethod
    def validate(
        wb: openpyxl.Workbook,
        operation: dict[str, Any],
        tracker: ConflictTracker,
    ) -> dict[str, Any]:

        t_name = operation.get("table")
        rows = operation.get("rows", [])

        # 1. Find the table's sheet
        ws = next((s for s in wb.worksheets if t_name in s.tables), None)
        if not ws:
            raise ValueError(f"Table '{t_name}' not found.")

        guard_sheet_unprotected(
            ws, f"Sheet '{ws.title}' containing table '{t_name}' is protected."
        )

        # 2. Parse table boundaries
        tbl = ws.tables[t_name]
        min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
        expected_cols = max_col - min_col + 1

        # 3. Validate row shapes & formula injection
        for row in rows:
            if len(row) != expected_cols:
                raise ValueError(
                    f"Table '{t_name}' expects {expected_cols} columns, got {len(row)}."
                )
            for val in row:
                guard_not_formula_injection(val, f"table '{t_name}' row value")

        # 4. Register landing zone with tracker & check cells are empty
        start_row = max_row + 1
        AppendTableHandler._check_landing_zone(
            ws, min_col, max_col, start_row, len(rows), tracker
        )

        # 5. Compute new table ref
        new_max_row = max_row + len(rows)
        new_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{new_max_row}"

        return {
            "table_name": t_name,
            "sheet": ws.title,
            "min_col": min_col,
            "start_row": start_row,
            "rows": rows,
            "new_ref": new_ref,
        }

    @staticmethod
    def _check_landing_zone(
        ws,
        min_col: int,
        max_col: int,
        start_row: int,
        num_rows: int,
        tracker: ConflictTracker,
    ) -> None:
        """Check that every cell in the landing zone is empty and
        register each with the conflict tracker."""
        for row in ws.iter_rows(
            min_row=start_row,
            max_row=start_row + num_rows - 1,
            min_col=min_col,
            max_col=max_col,
        ):
            for cell in row:
                tracker.track(
                    ws.title,
                    cell.coordinate,
                    f"Conflict: cell '{ws.title}!{cell.coordinate}' already targeted.",
                )
                if cell.value is not None:
                    raise PermissionError(
                        f"Collision: cell '{ws.title}!{cell.coordinate}' already contains data."
                    )

    @staticmethod
    def apply(wb: openpyxl.Workbook, validated: dict[str, Any]) -> None:
        ws = wb[validated["sheet"]]
        min_col = validated["min_col"]
        curr_row = validated["start_row"]

        for row_vals in validated["rows"]:
            for c_offset, val in enumerate(row_vals):
                ws.cell(row=curr_row, column=min_col + c_offset, value=val)
            curr_row += 1

        ws.tables[validated["table_name"]].ref = validated["new_ref"]

    @staticmethod
    def verify(verify_wb: openpyxl.Workbook, validated: dict[str, Any]) -> None:
        sheet_name = validated["sheet"]
        t_name = validated["table_name"]
        ws_verify = verify_wb[sheet_name]

        if t_name not in ws_verify.tables:
            raise OSError(f"Verification failed: table '{t_name}' not found on disk")
        if ws_verify.tables[t_name].ref != validated["new_ref"]:
            raise OSError(
                f"Verification failed: table '{t_name}' range mismatch on disk"
            )


OPERATION_HANDLERS = {
    "set_cell": SetCellHandler,
    "append_table": AppendTableHandler,
    "append_rows": AppendTableHandler,
}


def update_workbook(
    workspace: ExcelWorkspace,
    file_name: str | Path,
    operations: list[dict[str, Any]] | None = None,
    output_filename: str | Path | None = None,
    overwrite: bool = False,
    mutations: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Applies an atomic batch of ordered operations to an Excel workbook."""
    src_path = workspace.resolve_path(file_name, must_exist=True)

    # Determine destination path
    if output_filename and str(output_filename) != str(file_name):
        dest_path = workspace.resolve_path(output_filename, must_exist=False)
        if dest_path.exists() and not overwrite:
            raise FileExistsError(
                f"Destination file '{output_filename}' already exists and overwrite=False."
            )
        is_inplace = False
    else:
        dest_path = src_path
        is_inplace = True

    ops = operations if operations is not None else (mutations or [])
    if not ops:
        raise ValueError("No operations provided.")

    keep_vba = src_path.suffix.lower() in (".xlsm", ".xltm")
    wb = openpyxl.load_workbook(src_path, keep_vba=keep_vba, data_only=False)
    tmp_path: Path | None = None

    try:
        # 1. Preflight validation
        tracker = ConflictTracker()
        validated_ops: list[tuple[Any, dict[str, Any]]] = []

        for op in ops:
            action = op.get("action")
            if not action or action not in OPERATION_HANDLERS:
                raise ValueError(
                    f"Unsupported operation action: '{action}'. Supported actions: {list(OPERATION_HANDLERS.keys())}"
                )
            handler = OPERATION_HANDLERS[action]
            validated = handler.validate(wb, op, tracker)
            validated_ops.append((handler, validated))

        # 2. In-memory application
        for handler, validated in validated_ops:
            handler.apply(wb, validated)

        # 3. Atomic temporary write
        tmp_fd, tmp_path_str = tempfile.mkstemp(
            dir=dest_path.parent, prefix=".tmp_", suffix=dest_path.suffix
        )
        os.close(tmp_fd)
        tmp_path = Path(tmp_path_str)

        wb.save(tmp_path)
    finally:
        wb.close()

    # 4. Post-save verification & 5. Atomic swap
    try:
        verify_wb = openpyxl.load_workbook(
            tmp_path, keep_vba=keep_vba, data_only=False
        )
        try:
            for handler, validated in validated_ops:
                handler.verify(verify_wb, validated)
        finally:
            verify_wb.close()

        os.replace(tmp_path, dest_path)
    except Exception:
        if tmp_path and tmp_path.exists():
            tmp_path.unlink()
        raise

    return {
        "status": "success",
        "file_name": str(file_name),
        "output_file": str(dest_path.name if is_inplace else output_filename),
        "operations_applied": len(validated_ops),
    }


class UpdateWorkbookTool(BaseTool):
    def __init__(self, workspace: ExcelWorkspace):
        self.workspace = workspace
        super().__init__(
            name="update_workbook",
            description=(
                "Apply an atomic batch of ordered operations to an Excel workbook. "
                "Supported actions: 'set_cell' (keys: sheet, cell, value), "
                "'append_table' (keys: table, rows). "
                "Strictly preserves formulas. If output_filename is omitted, "
                "updates the workbook in-place safely."
            ),
            parameters={
                "type": "object",
                "properties": {
                    "file_name": {
                        "type": "string",
                        "description": "Name of the Excel workbook in the workspace (e.g. 'data.xlsx')",
                    },
                    "output_filename": {
                        "type": "string",
                        "description": "Optional output filename. If omitted or equal to file_name, updates in-place safely.",
                    },
                    "operations": {
                        "type": "array",
                        "items": {"type": "object"},
                        "description": "List of operations to apply.",
                    },
                    "overwrite": {
                        "type": "boolean",
                        "description": "Whether to overwrite destination file if it already exists.",
                        "default": False,
                    },
                },
                "required": ["file_name", "operations"],
            },
        )

    def execute(
        self,
        file_name: str,
        operations: list[dict[str, Any]] | None = None,
        output_filename: str | None = None,
        overwrite: bool = False,
        **kwargs: Any,
    ) -> dict[str, Any]:
        if not file_name:
            return {
                "status": "error",
                "error": "ValueError: Parameter 'file_name' is required.",
            }
        ops = operations or kwargs.get("mutations")
        if not ops:
            return {
                "status": "error",
                "error": "ValueError: Parameter 'operations' must be a non-empty list.",
            }
        try:
            return update_workbook(
                self.workspace,
                file_name=file_name,
                operations=ops,
                output_filename=output_filename,
                overwrite=overwrite,
            )
        except Exception as e:
            return {"status": "error", "error": f"{type(e).__name__}: {str(e)}"}
