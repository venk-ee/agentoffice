from pathlib import Path
from typing import Any

from agentoffice.tools.base_tools import BaseTool
from agentoffice.tools.office.excel_engine import ExcelEngine


class InspectWorkbookTool(BaseTool):
    def __init__(self, engine: ExcelEngine | None = None):
        super().__init__(
            name="inspect_workbook",
            description="Inspect an Excel workbook and return metadata about sheets, named ranges, formulas, and tables.",
            parameters={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Path to the Excel file to inspect",
                    }
                },
                "required": ["file_path"],
            },
        )
        self.engine = engine or ExcelEngine()

    def execute(self, file_path: str | Path) -> dict[str, Any]:
        return self.engine.inspect(file_path)


class ReadRangeTool(BaseTool):
    def __init__(self, engine: ExcelEngine | None = None):
        super().__init__(
            name="read_range",
            description="Read a range of cells from an Excel sheet.",
            parameters={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Path to the Excel file to read from",
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the sheet to read from",
                    },
                    "start_cell": {
                        "type": "string",
                        "description": "Start cell of the range to read (e.g., 'A1')",
                    },
                    "end_cell": {
                        "type": "string",
                        "description": "End cell of the range to read (e.g., 'B10')",
                    },
                    "data_only": {
                        "type": "boolean",
                        "description": (
                            "If True, returns cached evaluated values saved by Excel instead of formula strings. "
                            "Note: openpyxl does not compute formulas itself; uncached formulas may return None. Default is False."
                        ),
                    },
                },
                "required": ["file_path", "sheet_name", "start_cell", "end_cell"],
            },
        )
        self.engine = engine or ExcelEngine()

    def execute(
        self,
        file_path: str | Path,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        data_only: bool = False,
    ) -> dict[str, Any]:
        return self.engine.read_range(
            file_path, sheet_name, start_cell, end_cell, data_only=data_only
        )


class ApplyTemplateUpdateTool(BaseTool):
    def __init__(
        self,
        engine: ExcelEngine,
        output_dir: str | Path | None = None,
    ):
        if not isinstance(engine, ExcelEngine) or not engine.approved_templates_dir:
            raise ValueError(
                "ApplyTemplateUpdateTool requires an ExcelEngine configured with 'approved_templates_dir'."
            )
        super().__init__(
            name="apply_template_update",
            description=(
                "Apply an atomic batch update to an approved Excel template by filling named slots and tables. "
                "Preserves all existing formulas, branding, and layout, and saves to a new output file."
            ),
            parameters={
                "type": "object",
                "properties": {
                    "template_id": {
                        "type": "string",
                        "description": "Identifier of the approved template (e.g. 'startup_burn_report').",
                    },
                    "output_filename": {
                        "type": "string",
                        "description": "Output filename for the generated report (e.g. 'q3_report.xlsx').",
                    },
                    "slots": {
                        "type": "object",
                        "description": "Key-value dictionary mapping approved named slot names to their new values (e.g. {'CompanyName': 'Nexus Labs', 'StartingCash': 4500000}).",
                    },
                    "tables": {
                        "type": "object",
                        "description": "Optional mapping of native table names to lists of row dictionaries or row lists to append.",
                    },
                },
                "required": ["template_id", "output_filename", "slots"],
            },
        )
        self.engine = engine
        self.output_dir = Path(output_dir).resolve() if output_dir else None

    def execute(
        self,
        template_id: str,
        output_filename: str,
        slots: dict[str, Any],
        tables: dict[str, list[dict[str, Any] | list[Any]]] | None = None,
    ) -> dict[str, Any]:
        filename = (
            template_id if template_id.endswith(".xlsx") else f"{template_id}.xlsx"
        )
        template_path = self.engine.approved_templates_dir / filename
        if not template_path.exists():
            raise FileNotFoundError(
                f"Approved template '{template_id}' not found in {self.engine.approved_templates_dir}"
            )

        if self.output_dir:
            dest_path = (self.output_dir / output_filename).resolve()
            if not dest_path.is_relative_to(self.output_dir):
                raise PermissionError(
                    f"Output filename '{output_filename}' resolves outside approved output directory: {self.output_dir}"
                )
        else:
            dest_path = Path(output_filename).resolve()

        return self.engine.apply_template_update(
            template_path=template_path,
            output_path=dest_path,
            slots=slots,
            tables=tables,
        )
