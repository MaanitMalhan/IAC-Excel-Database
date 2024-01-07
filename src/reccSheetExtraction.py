import openpyxl
from copy import copy
from datetime import datetime, timedelta

current_date = datetime.now()
yesterday = current_date - timedelta(days=1)    
date = yesterday.strftime('%Y%m%d')

# Load the source workbook
source_workbook = openpyxl.load_workbook("/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database_20240107.xlsx")
source_sheet = source_workbook['RECC5']  #name of your source sheet

# Load the destination workbook

destination_workbook = openpyxl.load_workbook('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')

new_sheet = destination_workbook.create_sheet(title="RECC")

destination_sheet = destination_workbook['RECC']

def copy_first_row(source_sheet, target_workbook, target_sheet):
    for source_cell in source_sheet[1]:
        target_sheet.cell(row=1, column=source_cell.column, value=source_cell.value)

copy_first_row(source_sheet, destination_workbook, destination_sheet)

def copy_rows_with_value(source_sheet, target_workbook, target_sheet, target_value, column_index):
    for row in source_sheet.iter_rows():
        if row[column_index - 1].value == target_value:  # Adjust index to start from 0
            # Create a new row in the target sheet
            target_row = target_sheet.max_row + 1

            # Copy each cell from the source row to the target row
            for source_cell in row:
                target_sheet.cell(row=target_row, column=source_cell.column, value=source_cell.value)

# Target column index 
target_column_index = 2  

# Copy rows with the target value from the source to the target workbook
for i in range(1, 10):
    target_value = 'UC230' + str(i)
    print(target_value)
    copy_rows_with_value(source_sheet, destination_workbook, destination_sheet, target_value, target_column_index)

for i in range(10, 17):
    target_value = 'UC23' + str(i)
    print(target_value)
    copy_rows_with_value(source_sheet, destination_workbook, destination_sheet, target_value, target_column_index)

# Save the changes to the destination workbook
destination_workbook.save('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')
