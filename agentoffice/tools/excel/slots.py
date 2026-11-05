"""Named-slot domain helpers shared by the inspect and write tools."""

from openpyxl.utils.cell import range_boundaries
from openpyxl.workbook.defined_name import DefinedName

from agentoffice.tools.excel.guards import clean_coord


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
