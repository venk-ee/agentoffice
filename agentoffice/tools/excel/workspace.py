from pathlib import Path
from contextlib import contextmanager
from typing import Generator
import openpyxl


class ExcelWorkspace:
    def __init__(self, workspace_dir: str | Path):
        self.workspace_dir = Path(workspace_dir)
        self.workspace_dir.mkdir(parents=True, exist_ok=True)
        
    def resolve_path(self, path: str | Path, must_exist: bool = True) -> Path:
        raw_path = Path(path)
        candidate = raw_path if raw_path.is_absolute() else (self.workspace_dir / raw_path)
        target_path = candidate.resolve()
        if not target_path.is_relative_to(self.workspace_dir):
            raise PermissionError(
                f"Path {path} is outside the workspace directory {self.workspace_dir}"
            )
        if must_exist and not target_path.exists():
            raise FileNotFoundError(f"Path {path} does not exist")
        return target_path

    @contextmanager
    def open_workbook(
        self, filename: str | Path, must_exist: bool = True, data_only: bool = False
    ) -> Generator[openpyxl.Workbook, None, None]:
        wb_path = self.resolve_path(filename, must_exist=must_exist)
        keep_vba = wb_path.suffix.lower() in (".xlsm", ".xltm")
        wb = openpyxl.load_workbook(wb_path, keep_vba=keep_vba, data_only=data_only)
        try:
            yield wb
        finally:
            wb.close()
