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
    
# Find columns by header (header is in row 5 in your screenshot)
prev_inv_col = find_column_by_header(presentation_ws, "Previously Invoiced", header_row=5)
curr_inv_col = find_column_by_header(presentation_ws, "Current Invoice", header_row=5)

# Find the row for 2025 (year is in column B, which is col=2)
year_row = None
for row in presentation_ws.iter_rows(min_row=6, max_row=20, min_col=2, max_col=2):
    cell = row[0]
    if cell.value and str(cell.value).strip() == "2025":
        year_row = cell.row
        break

if prev_inv_col and curr_inv_col and year_row:
    # Step 1: Get the original values
    prev_val = presentation_ws.cell(row=year_row, column=prev_inv_col).value or 0
    curr_val = presentation_ws.cell(row=year_row, column=curr_inv_col).value or 0
    total = prev_val + curr_val
    print(f"Sum of Previously Invoiced and Current Invoice for 2025: {total}")

    # (Optional) Step 2: Update Previously Invoiced (G9) with the sum
    presentation_ws.cell(row=year_row, column=prev_inv_col).value = total

    # Step 3: Replace Current Invoice (H9) with the value from May-25
    presentation_ws.cell(row=year_row, column=curr_inv_col).value = section_total
    print(f"Replaced Current Invoice for 2025 with {section_total}")
else:
    print("Could not find required columns or row for 2025.")

# --- Second Table: Revised table - Labour + Travel ---

# 1. Find columns by header for the second table (adjust header_row and min_row as needed)
labour_table_header_row = 17  # Example: change if your second table starts elsewhere
labour_table_first_data_row = 18  # Example: change as needed

prev_inv_col2 = find_column_by_header(presentation_ws, "Previously Invoiced", header_row=labour_table_header_row)
curr_inv_col2 = find_column_by_header(presentation_ws, "Current Invoice", header_row=labour_table_header_row)

# 2. Find the row for 2025 in the second table (assuming year is in column B)
year_row2 = None
for row in presentation_ws.iter_rows(min_row=labour_table_first_data_row, max_row=labour_table_first_data_row+10, min_col=2, max_col=2):
    cell = row[0]
    if cell.value and str(cell.value).strip() == "2025":
        year_row2 = cell.row
        break

if prev_inv_col2 and curr_inv_col2 and year_row2:
    prev_val2 = presentation_ws.cell(row=year_row2, column=prev_inv_col2).value or 0
    curr_val2 = presentation_ws.cell(row=year_row2, column=curr_inv_col2).value or 0
    total2 = prev_val2 + curr_val2
    # Update Previously Invoiced with the sum
    presentation_ws.cell(row=year_row2, column=prev_inv_col2).value = total2
    print(f"Updated Labour+Travel Previously Invoiced for 2025: {total2}")
else:
    print("Could not find required columns or row for 2025 in Labour+Travel table.")

# --- Get Labour + Travel from Financial Report ---

# Find row numbers for Labour and Travel in the financial report
def find_row_by_label(ws, label, label_col=1):
    for row in ws.iter_rows(min_col=label_col, max_col=label_col):
        cell = row[0]
        if cell.value and label.lower() in str(cell.value).lower():
            return cell.row
    return None

labour_row = find_row_by_label(financial_ws, "labour", label_col=1)
travel_row = find_row_by_label(financial_ws, "travel", label_col=1)

if month_col and labour_row and travel_row:
    labour_val = financial_ws.cell(row=labour_row, column=month_col).value or 0
    travel_val = financial_ws.cell(row=travel_row, column=month_col).value or 0
    sum_labour_travel = (labour_val or 0) + (travel_val or 0)
    print(f"Labour + Travel for {target_month}: {sum_labour_travel}")

    # Update Current Invoice for 2025 in the dashboard's second table
    if curr_inv_col2 and year_row2:
        presentation_ws.cell(row=year_row2, column=curr_inv_col2).value = sum_labour_travel
        print(f"Updated Labour+Travel Current Invoice for 2025 with {sum_labour_travel}")
else:
    print("Could not find Labour/Travel rows or month column in financial report.")

target_wb.save('ATAC Project Dashboard Trial May.xlsx')


