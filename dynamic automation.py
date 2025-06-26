from openpyxl import Workbook, load_workbook
from copy import copy
from datetime import datetime
import os 

#-----------COPY FUNCTIONS--------------
def copy_sheet(source_ws, target_ws):
    copy_cells(source_ws, target_ws)
    copy_sheet_attributes(source_ws, target_ws)

def copy_sheet_attributes(source_ws, target_ws):
    target_ws.sheet_format = copy(source_ws.sheet_format)
    target_ws.sheet_properties = copy(source_ws.sheet_properties)
    target_ws.merged_cells = copy(source_ws.merged_cells)
    target_ws.page_margins = copy(source_ws.page_margins)
    target_ws.freeze_panes = copy(source_ws.freeze_panes)

    for rn in range(len(source_ws.row_dimensions)):
        target_ws.row_dimensions[rn] = copy(source_ws.row_dimensions[rn])

    if source_ws.sheet_format.defaultColWidth is None:
        print('Unable to copy default column wide')
    else:
        target_ws.sheet_format.defaultColWidth = copy(source_ws.sheet_format.defaultColWidth)

    for key, value in source_ws.column_dimensions.items():
        target_ws.column_dimensions[key].min = copy(value.min)
        target_ws.column_dimensions[key].max = copy(value.max)
        target_ws.column_dimensions[key].width = copy(value.width)
        target_ws.column_dimensions[key].hidden = copy(value.hidden)

def copy_cells(source_ws, target_ws):
    for (row, col), source_cell in source_ws._cells.items():
        target_cell = target_ws.cell(column=col, row=row)
        target_cell._value = source_cell._value
        target_cell.data_type = source_cell.data_type

        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)

        if source_cell.hyperlink:
            target_cell._hyperlink = copy(source_cell.hyperlink)

        if source_cell.comment:
            target_cell.comment = copy(source_cell.comment)

#-----------END OF COPY FUNCTIONS--------------
def find_column_by_month(ws, month_label, header_row=3):
    """
    Find the column index for a given month label in the header row.
    month_label: e.g. "May-25" or "2025-05-01"
    header_row: the row number where months are listed (3 in your screenshot)
    """
    # Try to parse month_label as "May-25"
    try:
        target_date = datetime.strptime(month_label, "%b-%y")
    except ValueError:
        try:
            target_date = datetime.strptime(month_label, "%Y-%m-%d")
        except Exception:
            target_date = None

    for cell in ws[header_row]:
        if cell.value:
            # If cell is a date
            if isinstance(cell.value, datetime):
                if target_date and cell.value.year == target_date.year and cell.value.month == target_date.month:
                    return cell.column
            # If cell is a string (fallback)
            elif month_label.lower() in str(cell.value).lower():
                return cell.column
    return None

def find_section_total_row(ws, section_label, total_label="TOTAL", label_col=1):
    """Find the first 'TOTAL' row after a section label."""
    section_row = None
    for row in ws.iter_rows(min_col=label_col, max_col=label_col):
        cell = row[0]
        if cell.value and section_label.lower() in str(cell.value).lower():
            section_row = cell.row
            break
    if section_row:
        for row in ws.iter_rows(min_row=section_row+1, min_col=label_col, max_col=label_col):
            cell = row[0]
            if cell.value and total_label.lower() in str(cell.value).lower():
                return cell.row
    return None

def find_column_by_header(ws, header_label, header_row=1):
    """Find the column index for a given header in the dashboard."""
    for cell in ws[header_row]:
        if cell.value and header_label.lower() in str(cell.value).lower():
            return cell.column
    return None


def find_first_row_by_label(ws, label="TOTAL", label_col=1):
    """Find the first row index where the cell value is exactly 'TOTAL' (all caps, no extra spaces)."""
    for row in ws.iter_rows(min_col=label_col, max_col=label_col):
        cell = row[0]
        if cell.value and str(cell.value).strip() == label:
            return cell.row
    return None
# --- Main Script ---
# Financial Source Workbook
financial_wb = load_workbook(r'C:\Users\Dsouzaj\Downloads\Automation excels\May\ATAC NRC Financial Report May 2025 (1).xlsx', data_only=True)
financial_ws = financial_wb['USD Monthly Totals']

# Dashboard Template Workbook
source_wb = load_workbook(r'C:\Users\Dsouzaj\Downloads\Automation excels\April\ATAC Project Dashbaord_May30, 25 with updated projections (1).xlsx')
sheet_names = ['Presentation Working Sheet', 'NRC ATAC All Phases', 'CAD to USD Savings']

# Target Workbook
target_wb = Workbook()
if 'Sheet' in target_wb.sheetnames:
    std = target_wb['Sheet']
    target_wb.remove(std)

# Loop through each sheet and copy
for name in sheet_names:
    source_ws = source_wb[name]
    target_ws = target_wb.create_sheet(title=name)
    copy_sheet(source_ws, target_ws)


presentation_ws = target_wb['Presentation Working Sheet']

# --- Custom cell logic for "Presentation Working Sheet" ALL PHASES TABLE-----

# --- Custom cell logic for "Presentation Working Sheet" ALL PHASES TABLE-----

target_month = "May-25"
section_label = "ATAC Total Billable"

# Find the column for May-25 in the financial report
month_col = find_column_by_month(financial_ws, target_month, header_row=3)
print("month_col:", month_col)

# Find the TOTAL row under the section in the financial report
total_row = find_first_row_by_label(financial_ws, label="TOTAL", label_col=1)
print("total_row:", total_row)

if month_col and total_row:
    section_total = financial_ws.cell(row=total_row, column=month_col).value
    print(f"{section_label} TOTAL for {target_month}:", section_total)
    
    # Find the "Current Invoice" column for 2025 in the dashboard (header row 1)
    dash_col = find_column_by_header(presentation_ws, "Current Invoice", header_row=1)
    # Find the row for "2025" in the dashboard (assuming label is in column 1)
    dash_row = None
    for row in presentation_ws.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        if cell.value and "2025" in str(cell.value):
            dash_row = cell.row
            break

    if dash_col and dash_row:
        presentation_ws.cell(row=dash_row, column=dash_col).value = section_total
        print(f"Pasted value {section_total} to dashboard at {presentation_ws.cell(row=dash_row, column=dash_col).coordinate}")
    else:
        print("Could not find the 'Current Invoice' column or 2025 row in the dashboard.")
else:
    print("Could not find the required section total in the financial report.")



target_wb.save('ATAC Project Dashboard Trial May.xlsx')


