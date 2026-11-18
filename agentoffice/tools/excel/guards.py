from openpyxl.cell.cell import Cell

FORMULA_TRIGGERS = ("=", "+", "-", "@")


def clean_coord(coord: str) -> str:
    """Strip whitespace and remove absolute reference '$' markers."""
    return str(coord).strip().replace("$", "")


def get_formula_trigger(value: object) -> str | None:
    """Return the trigger character if value is an unescaped formula-like string, else None."""
    if not isinstance(value, str):
        return None
    cleaned = value.strip()
    if cleaned.startswith(FORMULA_TRIGGERS):
        try:
            float(cleaned)
            return None
        except ValueError:
            return cleaned[0]
    return None


def is_formula(cell: Cell) -> bool:
    """Check if a cell contains a formula (data_type == 'f' or formula trigger string)."""
    if cell.data_type == "f":
        return True
    return get_formula_trigger(cell.value) is not None


def guard_not_formula_injection(value: object, target_desc: str) -> None:
    """Guard against formula injection in a cell value (CWE-1236)."""
    trigger = get_formula_trigger(value)
    if trigger is not None:
        raise PermissionError(
            f"Formula injection prohibited: {target_desc} begins with a formula trigger character ({trigger!r})"
        )


def guard_sheet_unprotected(sheet, error_message: str = "Sheet is protected") -> None:
    """Guard against writing to a protected sheet."""
    if sheet.protection.sheet:
        raise PermissionError(error_message)


def format_limit_exceeded_error(requested_cells: int, limit: int) -> str:
    """Format error message for format/cell limit exceeded with steering suggestions."""
    return (
        f"Cell limit exceeded: requested {requested_cells} cells, "
        f"maximum allowed per request is {limit}. "
        f"Please narrow your range (e.g. read specific sub-ranges or columns), "
        f"or use 'inspect_workbook' to examine sheet dimensions and table structures."
    )
