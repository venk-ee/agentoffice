"""Named-slot domain helpers shared by the inspect and write tools."""

import re

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import range_boundaries
from openpyxl.workbook.defined_name import DefinedName

from agentoffice.tools.excel.guards import clean_coord

CELL_OR_RANGE_RE = re.compile(
    r"^=?"
    r"(?:'(?:[^']|'')*'|[^'!\s]+)!"  # Sheet
    r"(?:"
    r"\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?"  # Cell or 2-cell range
    r"|"
    r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"  # Col range
    r"|"
    r"\$?[1-9]\d*:\$?[1-9]\d*"  # Row range
    r")$"
    r"|"
    r"^=?"
    r"(?:"
    r"\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?"
    r"|"
    r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
    r"|"
    r"\$?[1-9]\d*:\$?[1-9]\d*"
    r")$"
)


def is_direct_grid_reference(defn: DefinedName) -> bool:
    """Return True if defn is a direct single grid reference, not an expression or formula."""
    raw = (defn.attr_text or defn.value or "").strip()
    if not raw:
        return False
    try:
        tok = Tokenizer("=" + raw.lstrip("="))
        if len(tok.items) != 1 or tok.items[0].subtype != "RANGE":
            return False
        return bool(CELL_OR_RANGE_RE.match(raw))
    except Exception:
        return False


def collect_named_slots(workbook) -> list[tuple[str, str, DefinedName]]:
    """Collect user-defined named slots from every scope.

    Returns (name, scope, defn) tuples covering both workbook-global
    names (scope='workbook') and sheet-scoped names (scope=sheet
    title). Excel-internal names ('_xlnm.*', e.g. Print_Area) and
    hidden names are excluded so tool output stays LLM-readable.
    """
    slots: list[tuple[str, str, DefinedName]] = []
    for name, defn in workbook.defined_names.items():
        if name.startswith("_xlnm.") or defn.hidden:
            continue
        slots.append((name, "workbook", defn))
    for ws in workbook.worksheets:
        for name, defn in ws.defined_names.items():
            if name.startswith("_xlnm.") or defn.hidden:
                continue
            slots.append((name, ws.title, defn))
    return slots


def classify_slot(defn: DefinedName) -> str:
    """Classify a defined name by what set_slot could do with it.

    Returns one of:
      'cell'     - single-cell grid reference (set_slot-able)
      'range'    - multi-cell or whole-column/row reference
      'constant' - literal value, not a grid reference
      'formula'  - formula definition, not a grid reference
      'multi'    - resolves to more than one destination
    """
    try:
        dests = list(defn.destinations)
    except Exception:
        return "formula"

    if len(dests) > 1:
        return "multi"
    if not dests:
        attr = (defn.attr_text or defn.value or "").strip()
        return "formula" if attr.startswith("=") else "constant"

    raw = (defn.attr_text or defn.value or "").strip()
    try:
        tok = Tokenizer("=" + raw.lstrip("="))
        if len(tok.items) > 1:
            return "formula"
    except Exception:
        return "formula"

    if not is_direct_grid_reference(defn):
        return "formula" if raw.startswith("=") else "range"

    _, coord = dests[0]
    try:
        min_col, min_row, max_col, max_row = range_boundaries(clean_coord(coord))
    except (ValueError, TypeError):
        return "range"

    is_single_cell = (
        None not in (min_col, min_row, max_col, max_row)
        and min_col == max_col
        and min_row == max_row
    )
    return "cell" if is_single_cell else "range"
