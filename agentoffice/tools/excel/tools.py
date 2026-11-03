from pathlib import Path
from typing import Any

from agentoffice.core.tool import BaseTool
from agentoffice.tools.excel.engine import ExcelEngine


class InspectWorkbookTool(BaseTool):
    def __init__(self, engine: Any):
        if not isinstance(engine, ExcelEngine):
            raise ValueError(
                "InspectWorkbookTool requires an ExcelEngine instance with a configured workspace_dir."
            )
        super().__init__(
            name="inspect_workbook",
            description=(
                "Inspect an Excel workbook in the workspace and return metadata about sheets, "
                "named ranges, formulas, and tables."
            ),
            parameters={
                "type": "object",
                "properties": {
                    "file_name": {
                        "type": "string",
                        "description": "Name of the Excel workbook in the workspace to inspect (e.g. 'data.xlsx' or 'invoice.xlsx')",
                    }
                },
                "required": ["file_name"],
            },
        )
        self.engine = engine

    def execute(self, file_name: str | Path) -> dict[str, Any]:
        if not file_name:
            return {
                "status": "error",
                "error": "ValueError: Parameter 'file_name' is required.",
            }
        try:
            return self.engine.inspect(file_name)
        except Exception as e:
            return {
                "status": "error",
                "error": f"{type(e).__name__}: {str(e)}",
            }


class ReadRangeTool(BaseTool):
    def __init__(self, engine: Any):
        if not isinstance(engine, ExcelEngine):
            raise ValueError(
                "ReadRangeTool requires an ExcelEngine instance with a configured workspace_dir."
            )
        super().__init__(
            name="read_range",
            description=(
                "Read a range of cells from a sheet in an Excel workbook within the workspace. "
                "Note: Returns formula strings or cached evaluated values. Formulas are not evaluated dynamically."
            ),
            parameters={
                "type": "object",
                "properties": {
                    "file_name": {
                        "type": "string",
                        "description": "Name of the Excel workbook in the workspace to read from (e.g. 'data.xlsx')",
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the worksheet to read from",
                    },
                    "start_cell": {
                        "type": "string",
                        "description": "Start cell of the range to read (e.g. 'A1')",
                    },
                    "end_cell": {
                        "type": "string",
                        "description": "End cell of the range to read (e.g. 'B10')",
                    },
                    "data_only": {
                        "type": "boolean",
                        "description": (
                            "If True, returns cached evaluated values saved by Excel instead of formula strings. "
                            "Default is False."
                        ),
                    },
                },
                "required": ["file_name", "sheet_name", "start_cell", "end_cell"],
            },
        )
        self.engine = engine

    def execute(
        self,
        file_name: str | Path,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        data_only: bool = False,
    ) -> dict[str, Any]:
        if not file_name:
            return {
                "status": "error",
                "error": "ValueError: Parameter 'file_name' is required.",
            }
        try:
            return self.engine.read_range(
                file_name=file_name,
                sheet_name=sheet_name,
                start_cell=start_cell,
                end_cell=end_cell,
                data_only=data_only,
            )
        except Exception as e:
            return {
                "status": "error",
                "error": f"{type(e).__name__}: {str(e)}",
            }


class UpdateWorkbookTool(BaseTool):
    def __init__(
        self,
        engine: Any,
        output_dir: str | Path | None = None,
    ):
        if not isinstance(engine, ExcelEngine):
            raise ValueError(
                "UpdateWorkbookTool requires an ExcelEngine instance with a configured workspace_dir."
            )
        super().__init__(
            name="update_workbook",
            description=(
                "Apply an atomic batch update to an Excel workbook in the workspace. "
                "Supports updating explicit cells, named slots, and appending to native tables, "
                "while strictly preserving existing formulas and saving to a new output file."
            ),
            parameters={
                "type": "object",
                "properties": {
                    "file_name": {
                        "type": "string",
                        "description": "Filename of the source Excel workbook in the workspace (e.g. 'invoice.xlsx').",
                    },
                    "output_filename": {
                        "type": "string",
                        "description": "Output filename for the updated workbook (e.g. 'invoice_filled.xlsx').",
                    },
                    "cells": {
                        "type": "array",
                        "description": (
                            "List of explicit cell updates. Each item must be a dictionary with 'cell' and 'value', "
                            "and an optional 'sheet' (defaults to active sheet). E.g. [{'sheet': 'Invoice', 'cell': 'B2', 'value': 'Acme Corp'}]."
                        ),
                        "items": {
                            "type": "object",
                            "properties": {
                                "sheet": {"type": "string"},
                                "cell": {"type": "string"},
                                "value": {},
                            },
                            "required": ["cell", "value"],
                        },
                    },
                    "slots": {
                        "type": "object",
                        "description": "Key-value dictionary mapping defined names to their new values (e.g. {'TaxRate': 0.18}).",
                    },
                    "tables": {
                        "type": "object",
                        "description": "Mapping of native table names to lists of row dictionaries or row lists to append.",
                    },
                    "overwrite": {
                        "type": "boolean",
                        "description": "Whether to overwrite the output file if it already exists. Default is False.",
                    },
                },
                "required": ["file_name", "output_filename"],
            },
        )
        self.engine = engine
        self.output_dir = Path(output_dir).resolve() if output_dir else None

    def execute(
        self,
        file_name: str,
        output_filename: str,
        cells: list[dict[str, Any]] | None = None,
        slots: dict[str, Any] | None = None,
        tables: dict[str, list[dict[str, Any] | list[Any]]] | None = None,
        overwrite: bool = False,
    ) -> dict[str, Any]:
        if not file_name:
            return {
                "status": "error",
                "error": "ValueError: Parameter 'file_name' must be provided.",
            }
        if not output_filename:
            return {
                "status": "error",
                "error": "ValueError: Parameter 'output_filename' must be provided.",
            }

        try:
            if self.output_dir:
                dest_path = (self.output_dir / output_filename).resolve()
                if not dest_path.is_relative_to(self.output_dir):
                    return {
                        "status": "error",
                        "error": (
                            f"PermissionError: Output filename '{output_filename}' resolves outside "
                            f"approved output directory: {self.output_dir}"
                        ),
                    }
                out_target = dest_path
            else:
                out_target = output_filename

            return self.engine.update_workbook(
                file_name=file_name,
                output_filename=out_target,
                cells=cells,
                slots=slots,
                tables=tables,
                overwrite=overwrite,
            )
        except Exception as e:
            return {
                "status": "error",
                "error": f"{type(e).__name__}: {str(e)}",
            }


__all__ = [
    "ExcelEngine",
    "InspectWorkbookTool",
    "ReadRangeTool",
    "UpdateWorkbookTool",
]
