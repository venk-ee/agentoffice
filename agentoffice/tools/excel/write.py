import os
import tempfile
from difflib import get_close_matches
from pathlib import Path
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
from agentoffice.tools.excel.slots import collect_named_slots
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
    REQUIRED_KEYS = frozenset({"cell", "value"})
    OPTIONAL_KEYS = frozenset({"sheet"})

    @staticmethod
    def validate(
        wb: openpyxl.Workbook,
        operation: dict[str, Any],
        tracker: ConflictTracker,
    ) -> dict[str, Any]:

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


class SetSlotHandler:
    REQUIRED_KEYS = frozenset({"name", "value"})
    OPTIONAL_KEYS = frozenset()

    @staticmethod
    def validate(
        wb: openpyxl.Workbook,
        operation: dict[str, Any],
        tracker: ConflictTracker,
    ) -> dict[str, Any]:
        name = operation["name"]
        value = operation["value"]

        # Index slots across all scopes
        by_name: dict[str, list[tuple[str, Any]]] = {}
        for slot_name, scope, defn in collect_named_slots(wb):
            by_name.setdefault(slot_name, []).append((scope, defn))

        matches = by_name.get(str(name))
        if not matches:
            available = sorted(by_name, key=str)
            close = get_close_matches(str(name), available, n=1)
            suggestion = f" Did you mean '{close[0]}'?" if close else ""
            raise KeyError(
                f"Named slot '{name}' not found.{suggestion} "
                f"Available slots: {available}"
            )
        if len(matches) > 1:
            scopes = sorted(scope for scope, _ in matches)
            raise ValueError(
                f"Named slot '{name}' is ambiguous (defined at scopes {scopes})."
            )
        defn = matches[0][1]

        try:
            dests = list(defn.destinations)
        except Exception:
            # Non-grid expressions crash openpyxl's tokenizer
            dests = []

        if not dests:
            ref_desc = defn.attr_text or defn.value or ""
            raise ValueError(
                f"Named slot '{name}' does not resolve to a worksheet grid cell "
                f"(defined as constant/formula: '{ref_desc}'). Slot updates only apply to grid-backed cells."
            )
        if len(dests) > 1:
            raise ValueError(f"Named slot '{name}' points to multiple destinations.")

        sheet_title, cell_coord = dests[0]
        sheet_name = (sheet_title.strip("'") if sheet_title else None) or (
            wb.active.title if wb.active else None
        )
        if not sheet_name:
            raise ValueError(
                f"Named slot '{name}' destination worksheet could not be determined."
            )

        clean_c = clean_coord(cell_coord)
        try:
            min_col, min_row, max_col, max_row = range_boundaries(clean_c)
        except (ValueError, TypeError):
            raise ValueError(f"Named slot '{name}' coordinate '{clean_c}' is invalid.")
        is_single_cell = (
            None not in (min_col, min_row, max_col, max_row)
            and min_col == max_col
            and min_row == max_row
        )
        if not is_single_cell:
            raise ValueError(
                f"Named slot '{name}' spans multiple cells ({clean_c}); "
                f"scalar slot updates require a single cell."
            )

        # Delegate to SetCellHandler for protection, collision tracking, formula checks
        return SetCellHandler.validate(
            wb,
            {"sheet": sheet_name, "cell": clean_c, "value": value},
            tracker,
        )

    apply = SetCellHandler.apply
    verify = SetCellHandler.verify


class AppendTableHandler:
    REQUIRED_KEYS = frozenset({"table", "rows"})
    OPTIONAL_KEYS = frozenset()

    @staticmethod
    def validate(
        wb: openpyxl.Workbook,
        operation: dict[str, Any],
        tracker: ConflictTracker,
    ) -> dict[str, Any]:

        t_name = operation["table"]
        rows = operation["rows"]
        if not isinstance(rows, list):
            raise ValueError(
                f"append_table 'rows' must be a list, got {type(rows).__name__}."
            )
        if not rows:
            raise ValueError(
                "append_table 'rows' must be a non-empty list of row values."
            )

        # 1. Find the table's sheet
        ws = next((s for s in wb.worksheets if t_name in s.tables), None)
        if not ws:
            raise ValueError(f"Table '{t_name}' not found.")

        guard_sheet_unprotected(
            ws, f"Sheet '{ws.title}' containing table '{t_name}' is protected."
        )

        # 2. Parse table boundaries & read the header row
        tbl = ws.tables[t_name]
        min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
        expected_cols = max_col - min_col + 1
        headers = [
            ws.cell(row=min_row, column=c).value for c in range(min_col, max_col + 1)
        ]
        header_str_set = {str(h) for h in headers if h is not None}

        # 3. Validate row shapes, normalize to header order, & check values
        normalized_rows = []
        for row in rows:
            if isinstance(row, dict):
                if not row:
                    raise ValueError(
                        f"Table '{t_name}' row dictionary cannot be empty."
                    )
                str_row = {str(k): v for k, v in row.items()}
                unknown = [k for k in row if str(k) not in header_str_set]
                if unknown:
                    raise ValueError(
                        f"Table '{t_name}' has no column(s) {sorted(unknown, key=str)}. "
                        f"Expected columns: {headers}"
                    )
                row_vals = [
                    row[h] if h in row else str_row.get(str(h)) for h in headers
                ]
            elif isinstance(row, (list, tuple)):
                if len(row) != expected_cols:
                    raise ValueError(
                        f"Table '{t_name}' expects {expected_cols} columns, got {len(row)}."
                    )
                row_vals = list(row)
            else:
                raise ValueError(
                    f"Table '{t_name}' rows must be lists of values or dicts "
                    f"keyed by column names {headers}."
                )
            for val in row_vals:
                if not (val is None or isinstance(val, (str, int, float, bool))):
                    raise ValueError(
                        f"Table '{t_name}' row values must be scalars "
                        f"(string, number, boolean, or null), got {type(val).__name__}."
                    )
                guard_not_formula_injection(val, f"table '{t_name}' row value")
            normalized_rows.append(row_vals)

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
            "rows": normalized_rows,
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

        # Verify every appended value landed on disk
        for r, row_vals in enumerate(validated["rows"]):
            for c, expected in enumerate(row_vals):
                coord = (
                    f"{get_column_letter(validated['min_col'] + c)}"
                    f"{validated['start_row'] + r}"
                )
                saved_val = ws_verify[coord].value
                if saved_val != expected:
                    raise OSError(
                        f"Verification failed: table '{t_name}' cell "
                        f"'{sheet_name}!{coord}' value mismatch on disk "
                        f"(expected {expected!r}, got {saved_val!r})"
                    )


OPERATION_HANDLERS = {
    "set_cell": SetCellHandler,
    "append_table": AppendTableHandler,
    "set_slot": SetSlotHandler,
}


def _resolve_dest_path(
    workspace: ExcelWorkspace,
    file_name: str | Path,
    output_filename: str | Path | None,
    overwrite: bool,
) -> tuple[Path, Path, str]:
    """Resolves source and destination paths, validating the overwrite policy."""
    src_path = workspace.resolve_path(file_name, must_exist=True)

    if output_filename and str(output_filename) != str(file_name):
        dest_path = workspace.resolve_path(output_filename, must_exist=False)
        if dest_path.exists() and not overwrite:
            raise FileExistsError(
                f"Destination file '{output_filename}' already exists and overwrite=False."
            )
        output_name = str(output_filename)
    else:
        dest_path = src_path
        output_name = src_path.name

    return src_path, dest_path, output_name


def _validate_operation_envelope(
    op: dict[str, Any],
    handler: Any,
    action: str,
) -> None:
    """Validates that required keys are present and no unknown keys were passed."""
    keys = op.keys() - {"action"}
    unknown = keys - handler.REQUIRED_KEYS - handler.OPTIONAL_KEYS
    if unknown:
        allowed = sorted(handler.REQUIRED_KEYS | handler.OPTIONAL_KEYS)
        hints = []
        for bad in sorted(unknown, key=str):
            str_bad = str(bad)
            close = get_close_matches(str_bad, allowed, n=1)
            hints.append(
                f"'{str_bad}'" + (f" (did you mean '{close[0]}'?)" if close else "")
            )
        raise ValueError(
            f"Unknown key(s) {', '.join(hints)} for '{action}'. Allowed keys: {allowed}"
        )

    missing = handler.REQUIRED_KEYS - keys
    if missing:
        raise ValueError(
            f"'{action}' requires key(s) {sorted(handler.REQUIRED_KEYS)}, "
            f"missing: {sorted(missing)}"
        )


def _atomic_commit(
    wb: openpyxl.Workbook,
    dest_path: Path,
    validated_ops: list[tuple[Any, dict[str, Any]]],
    keep_vba: bool,
) -> None:
    """Saves workbook to a temp file, verifies on disk, and atomically replaces dest_path."""
    tmp_fd, tmp_path_str = tempfile.mkstemp(
        dir=dest_path.parent, prefix=".tmp_", suffix=dest_path.suffix
    )
    os.close(tmp_fd)
    tmp_path = Path(tmp_path_str)

    try:
        wb.save(tmp_path)

        verify_wb = openpyxl.load_workbook(tmp_path, keep_vba=keep_vba, data_only=False)
        try:
            for handler, validated in validated_ops:
                handler.verify(verify_wb, validated)
        finally:
            verify_wb.close()

        os.replace(tmp_path, dest_path)
    except Exception:
        if tmp_path.exists():
            tmp_path.unlink()
        raise


def update_workbook(
    workspace: ExcelWorkspace,
    file_name: str | Path,
    operations: list[dict[str, Any]] | None = None,
    output_filename: str | Path | None = None,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Applies an atomic batch of ordered operations to an Excel workbook."""
    src_path, dest_path, out_name = _resolve_dest_path(
        workspace, file_name, output_filename, overwrite
    )

    ops = operations or []
    if not ops:
        raise ValueError("No operations provided.")

    keep_vba = src_path.suffix.lower() in (".xlsm", ".xltm")
    wb = openpyxl.load_workbook(src_path, keep_vba=keep_vba, data_only=False)

    try:
        # 1. validation
        tracker = ConflictTracker()
        validated_ops: list[tuple[Any, dict[str, Any]]] = []

        for op in ops:
            action = op.get("action")
            if not action or action not in OPERATION_HANDLERS:
                raise ValueError(
                    f"Unsupported operation action: '{action}'. Supported actions: {list(OPERATION_HANDLERS.keys())}"
                )
            handler = OPERATION_HANDLERS[action]
            _validate_operation_envelope(op, handler, action)
            validated_ops.append((handler, handler.validate(wb, op, tracker)))

        # 2. In-memory application
        for handler, validated in validated_ops:
            handler.apply(wb, validated)

        # 3. Atomic temporary write
        _atomic_commit(wb, dest_path, validated_ops, keep_vba)
    finally:
        wb.close()

    return {
        "status": "success",
        "file_name": str(file_name),
        "output_file": out_name,
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
                "'set_slot' (keys: name, value; updates a single-cell defined "
                "name — discover available names via inspect_workbook's "
                "named_slots), "
                "'append_table' (keys: table, rows; each row is a dict keyed by the "
                "table's column names, e.g. {'Item': 'Mouse', 'Price': 50}, preferred, "
                "or a list of scalar values in column order). "
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
                        "items": {"type": "object", "required": ["action"]},
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
    ) -> dict[str, Any]:
        if not file_name:
            return {
                "status": "error",
                "error": "ValueError: Parameter 'file_name' is required.",
            }
        ops = operations
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
