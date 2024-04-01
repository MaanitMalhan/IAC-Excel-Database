import openpyxl
from calculcation import calculations, cost_savings
destination_workbook = openpyxl.load_workbook('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')
print(cost_savings(destination_workbook))
calculations(destination_workbook)
destination_workbook.save('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')

print("Calculations done successfully!")
