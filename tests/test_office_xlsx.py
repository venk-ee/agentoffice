from pathlib import Path
import openpyxl

file_path = Path(__file__).resolve().parent.parent / "templates" / "tests" / "startup_burn_report.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)
named_ranges = list(wb.defined_names.keys())
print(named_ranges)



def fill_templet(template_path: Path, output_path: Path, data: dict):
    wb = openpyxl.load_workbook(template_path, data_only=False)
    for name, value in data.items():
        if name in wb.defined_names:
            defined_name = wb.defined_names[name]
            for sheet_name, cell_ref in defined_name.destinations:
                cell = wb[sheet_name][cell_ref.replace("$", "")]
                if str(cell.value).startswith("="):
                    print(f"Skipping formula cell: {cell_ref}")
                    continue
                cell.value = value
    wb.save(output_path)


output_file = Path(__file__).resolve().parent / "report_output.xlsx"
fill_templet(
    template_path=file_path,
    output_path=output_file,
    data={"CompanyName": "AgentNexus Inc", "StartingCash": 2750000}
)
print("Finished writing to:", output_file)