import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

def create_startup_burn_template(filename: str = "templates/startup_burn_report.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Burn & Runway"

    # Styling definitions
    brand_blue = "1F4E79"
    soft_gray = "F2F2F2"
    border_gray = "D9D9D9"
    
    title_font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    section_font = Font(name="Calibri", size=11, bold=True, color=brand_blue)
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True)
    normal_font = Font(name="Calibri", size=11)
    
    title_fill = PatternFill(start_color=brand_blue, end_color=brand_blue, fill_type="solid")
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    summary_fill = PatternFill(start_color=soft_gray, end_color=soft_gray, fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color=border_gray),
        right=Side(style='thin', color=border_gray),
        top=Side(style='thin', color=border_gray),
        bottom=Side(style='thin', color=border_gray)
    )

    # 1. Title Banner (Rows 1-2)
    ws.merge_cells("A1:F1")
    title_cell = ws["A1"]
    title_cell.value = "STARTUP EXECUTIVE BURN & RUNWAY REPORT"
    title_cell.font = title_font
    title_cell.fill = title_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    # 2. Company Metadata (Rows 3-5)
    ws["A3"] = "Company Name:"
    ws["A3"].font = bold_font
    ws["B3"] = "Acme AI Corp"  # Placeholder / Default
    ws["B3"].font = normal_font

    ws["A4"] = "Reporting Quarter:"
    ws["A4"].font = bold_font
    ws["B4"] = "Q3 2026"
    ws["B4"].font = normal_font

    # EDGE CASE 4: Field with label but user forgot to make a named range!
    ws["A5"] = "Lead Partner:"
    ws["A5"].font = bold_font
    ws["B5"] = "Unassigned"
    ws["B5"].font = normal_font

    # 3. High-Level Metrics Box (Columns D to F, Rows 3-7)
    ws["D3"] = "Starting Cash:"
    ws["D3"].font = bold_font
    ws["E3"] = 1500000.00  # $1.5M
    ws["E3"].font = bold_font
    ws["E3"].number_format = "$#,##0.00"

    ws["D4"] = "Total Qtr Revenue:"
    ws["D4"].font = bold_font
    # EDGE CASE 1: Formula cell (auto sums monthly revenue)
    ws["E4"] = "=SUM(C10:C12)"
    ws["E4"].font = bold_font
    ws["E4"].number_format = "$#,##0.00"

    ws["D5"] = "Total Qtr Burn:"
    ws["D5"].font = bold_font
    # Formula cell (auto sums monthly expenses)
    ws["E5"] = "=SUM(D10:D12)"
    ws["E5"].font = bold_font
    ws["E5"].number_format = "$#,##0.00"

    ws["D6"] = "Net Monthly Burn:"
    ws["D6"].font = bold_font
    # Formula: (Expenses - Revenue) / 3
    ws["E6"] = "=(E5-E4)/3"
    ws["E6"].font = bold_font
    ws["E6"].number_format = "$#,##0.00"

    ws["D7"] = "Runway (Months):"
    ws["D7"].font = bold_font
    # Formula: Starting Cash / Net Monthly Burn
    ws["E7"] = "=ROUND(E3/E6, 1)"
    ws["E7"].font = Font(name="Calibri", size=11, bold=True, color="C00000")  # Alert color
    ws["E7"].number_format = "0.0"

    for r in range(3, 8):
        ws[f"D{r}"].fill = summary_fill
        ws[f"E{r}"].fill = summary_fill
        ws[f"D{r}"].border = thin_border
        ws[f"E{r}"].border = thin_border

    # 4. Monthly Financials Breakdown Table (Rows 9-13)
    ws["A9"] = "Month"
    ws["B9"] = "Headcount Cost"
    ws["C9"] = "Revenue"
    ws["D9"] = "Total Expenses"
    ws["E9"] = "Notes / Milestones"

    for col in ["A", "B", "C", "D", "E"]:
        cell = ws[f"{col}9"]
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Sample pre-existing 3 rows (Q3: July, Aug, Sept)
    sample_data = [
        ("July 2026", 45000, 20000, 65000, "Hired 2 Engineers"),
        ("August 2026", 50000, 35000, 72000, "AWS Credits exhausted"),
        ("September 2026", 52000, 48000, 78000, "Series A prep initiated"),
    ]

    for idx, row in enumerate(sample_data, start=10):
        ws[f"A{idx}"] = row[0]
        ws[f"B{idx}"] = row[1]
        ws[f"B{idx}"].number_format = "$#,##0"
        ws[f"C{idx}"] = row[2]
        ws[f"C{idx}"].number_format = "$#,##0"
        ws[f"D{idx}"] = row[3]
        ws[f"D{idx}"].number_format = "$#,##0"
        ws[f"E{idx}"] = row[4]
        for col in ["A", "B", "C", "D", "E"]:
            ws[f"{col}{idx}"].border = thin_border

    # Adjust Column Widths
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 30
    ws.column_dimensions["F"].width = 5

    # 5. Define Official Named Ranges (The Approved Slots for Agents)
    # Target syntax: 'SheetName'!$Col$Row
    wb.defined_names.add(DefinedName("CompanyName", attr_text="'Burn & Runway'!$B$3"))
    wb.defined_names.add(DefinedName("ReportingQuarter", attr_text="'Burn & Runway'!$B$4"))
    wb.defined_names.add(DefinedName("StartingCash", attr_text="'Burn & Runway'!$E$3"))

    wb.save(filename)
    print(f"Generated realistic template at: {filename}")

if __name__ == "__main__":
    create_startup_burn_template()
