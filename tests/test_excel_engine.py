from pathlib import Path

from agentoffice.tools.office.excel import ExcelEngine

engine = ExcelEngine()
template_path = Path("testing_templates/tests/startup_burn_report.xlsx")

info = engine.inspect(template_path)
print("--- Inspection Result ---")
print("Sheets:", info["sheets"])
print("Dimensions:", info["dimensions"])
print("Named Slots:", info["named_slots"])
print("Formulas Found:", len(info["formulas"]))

rows = engine.read_range(template_path, "Burn & Runway", "A9", "E12")
print("\n--- Read Range (A9:E12) ---")
for row in rows:
    print(row)
