from openpyxl import Workbook, load_workbook
from copy import copy
import os 

#Fincial Source Workbook  - ATAC Fincial Report 
financial_wb = load_workbook(r'C:\Users\Dsouzaj\Documents\NRC Python Files\ATAC NRC Financial Report Mar 2025.xlsx' )
financial_ws = financial_wb['USD Monthly Totals']

#Target Workbook - New excel files with template copied over 
target_wb = Workbook()
target_ws = target_wb.active 
target_ws.title = 'Presentation working Sheet'
#target_wb.save('ATAC Project Dashboard March 2025')

#Dashboard Template Workbook - Copy these worksheets to target workbook 
source_wb = load_workbook(r'C:\Users\Dsouzaj\Documents\NRC Python Files\ATAC Project Dashbaord.xlsx')
source_ws = source_wb['Presentation Working Sheet']

#Copy Functions 
def copy_sheet(source_ws, target_ws):
    copy_cells(source_ws, target_ws)  # copy all the cel values and styles
    copy_sheet_attributes(source_ws, target_ws)

def copy_sheet_attributes(source_ws,target_ws):
   target_ws.sheet_format = copy(source_ws.sheet_format)
   target_ws.sheet_properties = copy(source_ws.sheet_properties)
   target_ws.merged_cells = copy(source_ws.merged_cells)
   target_ws.page_margins = copy(source_ws.page_margins)
   target_ws.freeze_panes = copy(source_ws.freeze_panes)

    # set row dimensions
    # So you cannot copy the row_dimensions attribute. Does not work (because of meta data in the attribute I think). So we copy every row's row_dimensions. That seems to work.
   for rn in range(len(source_ws.row_dimensions)):
       target_ws.row_dimensions[rn] = copy(source_ws.row_dimensions[rn])
       
       if source_ws.sheet_format.defaultColWidth is None:
        print('Unable to copy default column wide')
        
       else: 
          target_ws.sheet_format.defaultColWidth = copy(source_ws.sheet_format.defaultColWidth)

    # set specific column width and hidden property
    # we cannot copy the entire column_dimensions attribute so we copy selected attributes
for key, value in source_ws.column_dimensions.items():
       target_ws.column_dimensions[key].min = copy(source_ws.column_dimensions[key].min)   # Excel actually groups multiple columns under 1 key. Use the min max attribute to also group the columns in the targetSheet
       target_ws.column_dimensions[key].max = copy(source_ws.column_dimensions[key].max)  # https://stackoverflow.com/questions/36417278/openpyxl-can-not-read-consecutive-hidden-columns discussed the issue. Note that this is also the case for the width, not onl;y the hidden property
       target_ws.column_dimensions[key].width = copy(source_ws.column_dimensions[key].width) # set width for every column
       target_ws.column_dimensions[key].hidden = copy(source_ws.column_dimensions[key].hidden)


def copy_cells(source_ws,target_ws):
    for (row, col), source_cell in source_ws._cells.items():
        target_cell =target_ws.cell(column=col, row=row)

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





copy_sheet(source_ws,target_ws)

if 'Sheet' in target_wb.sheetnames:  # remove default sheet
    target_wb.remove(target_wb['Sheet'])

target_wb.save('ATAC Project Dashboard APR 2025.xlsx')

