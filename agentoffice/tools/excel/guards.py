from openpyxl.cell.cell import Cell


def clean_coord(coord: str) -> str:
    """Strip whitespace and remove absolute reference '$' markers."""
    return str(coord).strip().replace("$", "")


def is_formula(cell: Cell) -> bool:
    """Check if a cell contains a formula (starts with '=')."""
    if cell.data_type == "f":
        return True
    return isinstance(cell.value, str) and cell.value.strip().startswith("=")


def guard_not_formula_injection(value: object, target_desc: str) -> None:
    """Guard against formula injection in a cell."""
    if isinstance(value, str) and value.strip().startswith("="):
        raise PermissionError(
            f"Formula injection prohibited: {target_desc} begins with '='"
        )


def guard_sheet_unprotected(sheet, error_message: str = "Sheet is protected") -> None:
    """Guard against writing to a protected sheet."""
    if sheet.protection.sheet:
        raise PermissionError(error_message)


def format_limit_exceeded_error(requested_cells: int, limit: int) -> str:
    """Format error message for format limit exceeded."""
    return f"Format limit exceeded: requested {requested_cells} cells, consider using smaller targeted coordinates, please consider using preview_sheet for headers and totals,paginate_table for scrolling rows."  # noqa: E501
