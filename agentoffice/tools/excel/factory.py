from agentoffice.tools.excel.inspect import InspectWorkbookTool
from agentoffice.tools.excel.read import ReadRangesTool
from agentoffice.tools.excel.workspace import ExcelWorkspace
from agentoffice.tools.excel.write import UpdateWorkbookTool
from agentoffice.core.tool import BaseTool
from pathlib import Path


def create_excel_tools(workspace_dir: str | Path) -> list[BaseTool]:
    """Create and return all Excel tools bound to the given workspace directory."""
    ws = ExcelWorkspace(workspace_dir)
    return [
        InspectWorkbookTool(ws),
        ReadRangesTool(ws),
        UpdateWorkbookTool(ws),
    ]