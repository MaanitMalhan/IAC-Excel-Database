import openpyxl
from datetime import datetime, timedelta


def copy_term(date):
# Load the source workbook
    source_workbook = openpyxl.load_workbook(f"/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database_{date}.xlsx")
    source_sheet = source_workbook['Terms']  

# Load the destination workbook
    destination_workbook = openpyxl.load_workbook('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')
    new_sheet = destination_workbook.create_sheet(title="Terms")
    destination_sheet = destination_workbook['Terms']  

    for row in source_sheet.iter_rows(min_row=1, max_row=source_sheet.max_row, min_col=1, max_col=source_sheet.max_column):
        for cell in row:
            destination_sheet[cell.coordinate].value = cell.value
   
    img = openpyxl.drawing.image.Image('IAC_PREF_MARCH_18_2024.png')
    img.anchor = 'K1'
    destination_sheet.add_image(img)

# Save the changes to the destination workbook
    destination_workbook.save('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')
