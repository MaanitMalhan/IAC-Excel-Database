import openpyxl
from copy import copy
from datetime import datetime, timedelta

current_date = datetime.now()
yesterday = current_date - timedelta(days=1)    
date = yesterday.strftime('%Y%m%d')

# Load the source workbook
source_workbook = openpyxl.load_workbook("/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database_20240107.xlsx")
source_sheet = source_workbook['ASSESS']  #name of your source sheet

# Load the destination workbook
workbook = openpyxl.Workbook()
workbook.save('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')

destination_workbook = openpyxl.load_workbook('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')
new_sheet = destination_workbook.create_sheet(title="ASSESS")

destination_sheet = destination_workbook['ASSESS']

def copy_first_row(source_sheet, target_workbook, target_sheet):
    for source_cell in source_sheet[1]:
        target_sheet.cell(row=1, column=source_cell.column, value=source_cell.value)

copy_first_row(source_sheet, destination_workbook, destination_sheet)
def copy_rows_with_value(source_sheet, target_workbook, target_sheet, target_value):
    for row in source_sheet.iter_rows():
        for cell in row:
            if cell.value == target_value:
                # Create a new row in the target sheet
                target_row = target_sheet.max_row + 1
                print("made it to line 30")
                # Copy each cell from the source row to the target row
                for source_cell in row:
                    target_sheet.cell(row=target_row, column=source_cell.column, value=source_cell.value)


# Specify the value you want to find and copy
target_value = 'UC'

# Copy rows with the target value from the source to the target workbook
copy_rows_with_value(source_sheet, destination_workbook, destination_sheet, target_value)

# Save the changes to the destination workbook
destination_workbook.save('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')
