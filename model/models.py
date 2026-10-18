from enum import StrEnum
from typing import Any
from pydantic import BaseModel, Field


class ColumnType(StrEnum):
    """Supported data types for columns in generic sheets."""
    STRING = "string"
    NUMBER = "number"
    CURRENCY = "currency"
    PERCENTAGE = "percentage"
    DATE = "date"


class SummaryFunction(StrEnum):
    """Supported summary/aggregation functions for table bottoms."""
    SUM = "SUM"
    AVERAGE = "AVERAGE"
    COUNT = "COUNT"
    MIN = "MIN"
    MAX = "MAX"


# =====================================================================
# Mode 1: Inspection & Template-Patching Models
# =====================================================================

class ExcelSlot(BaseModel):
    """Represents a discrete placeholder/named-range in an existing template."""
    name: str
    coordinate: str  # e.g. "B3" or "'Sheet1'!$B$3"
    current_value: Any = None
    sheet_name: str = "Sheet1"
    is_protected_formula: bool = False


class ExcelTableSchema(BaseModel):
    """Represents an existing table structure detected in a template."""
    name: str
    sheet_name: str
    header_row: int
    start_col: int
    end_col: int
    columns: list[str] = Field(default_factory=list)


class TemplateInspectionSchema(BaseModel):
    """Schema returned to the agent describing an inspected template."""
    template_path: str
    sheet_names: list[str] = Field(default_factory=list)
    slots: dict[str, ExcelSlot] = Field(default_factory=dict)
    tables: dict[str, ExcelTableSchema] = Field(default_factory=dict)
    protected_formula_cells: list[str] = Field(default_factory=list)


class TemplatePatchRequest(BaseModel):
    """Payload sent by the agent to patch an existing template atomically."""
    slots: dict[str, Any] = Field(
        default_factory=dict,
        description="Key-value mapping of slot names to new values, e.g. {'CompanyName': 'Acme'}"
    )
    table_appends: dict[str, list[dict[str, Any]]] = Field(
        default_factory=dict,
        description="Mapping of table name to rows to append, e.g. {'MonthlyExpenses': [{'Month': 'Oct', 'Cost': 50000}]}"
    )


# =====================================================================
# Mode 2: Generic Scaffold & Declarative Formula Models
# =====================================================================

class ColumnSpec(BaseModel):
    """Definition for a single column in a generic spreadsheet."""
    name: str
    type: ColumnType = ColumnType.STRING
    formula: str | None = Field(
        default=None,
        description="Declarative formula using bracketed column tokens, e.g. '=[Base Salary]*[Bonus Pct]'"
    )


class SummaryRowSpec(BaseModel):
    """Definition for a summary calculation row at the bottom of a table."""
    label: str = "Total"
    column: str  # Column name to aggregate
    function: SummaryFunction = SummaryFunction.SUM


class GenericSheetSpec(BaseModel):
    """Complete specification to generate a new corporate Excel sheet from scratch."""
    title: str = "Corporate Report"
    sheet_name: str = "Sheet1"
    metadata_slots: dict[str, Any] = Field(
        default_factory=dict,
        description="Header-level key-values, e.g. {'Department': 'Engineering', 'Date': '2026-10-18'}"
    )
    columns: list[ColumnSpec]
    rows: list[dict[str, Any]] = Field(
        default_factory=list,
        description="List of data rows matching column names"
    )
    summary_rows: list[SummaryRowSpec] = Field(default_factory=list)
