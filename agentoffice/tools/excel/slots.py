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
    raw = (defn.attr_text or defn.value or "").strip()
    if not raw:
        return "constant"

    try:
        tok = Tokenizer("=" + raw.lstrip("="))
    except Exception:
        return "formula"

    items = tok.items
    if not items:
        return "constant"

    # Single-token operand
    if len(items) == 1:
        item = items[0]
        if item.subtype == "RANGE" and CELL_OR_RANGE_RE.match(raw):
            try:
                dests = list(defn.destinations)
                if dests:
                    min_col, min_row, max_col, max_row = range_boundaries(
                        clean_coord(dests[0][1])
                    )
                    if (
                        None not in (min_col, min_row, max_col, max_row)
                        and min_col == max_col
                        and min_row == max_row
                    ):
                        return "cell"
            except Exception:
                pass
            return "range"

        if item.subtype == "RANGE":
            return "formula" if raw.startswith("=") else "range"
        return "constant"

    # Multi-token: check if it is a pure comma-separated union of direct grid references
    # e.g. [RANGE, ',', RANGE] like Summary!$A$1,Summary!$B$2
    is_union = True
    range_count = 0
    for i, it in enumerate(items):
        if i % 2 == 0:
            if it.subtype != "RANGE" or not CELL_OR_RANGE_RE.match(it.value):
                is_union = False
                break
            range_count += 1
        else:
            if it.value != ",":
                is_union = False
                break

    if is_union and range_count > 1:
        return "multi"

    return "formula"
