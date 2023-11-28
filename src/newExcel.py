import openpyxl
import os

#workbook = openpyxl.load_workbook('/Users/maanitmalhan/Documents/IAC Center/excel-data-iac/IAC_Database.zip')

new_workbook = openpyxl.Workbook()
new_worksheet = new_workbook.active
assessment = new_workbook.create_sheet('Assessment', 0)

new_workbook.save('/Users/maanitmalhan/Documents/IAC Center/excel-data-iac/SNE_IAC_Database.xlsx')

print(new_workbook.sheetnames)