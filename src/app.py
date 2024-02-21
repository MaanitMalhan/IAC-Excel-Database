from download import download_file
from zipExtract import extract_file
from convertFormat import convert_xls_to_xlsx
from assessSheetExtraction import copy_rows_with_values, copy_first_rows
from reccSheetExtraction import copy_rows_with_value, copy_first_row
from termCopy import copy_term
from datetime import datetime, timedelta
from arcCodes import arc_code_sheet
import openpyxl


#create SNE workbook
workbook = openpyxl.Workbook()
workbook.save('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')
print("File Created!")

#Download ZIP from server 
file_url = 'https://iac.university/storage/IAC_Database.zip'
destination_path = '/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database.zip'

download_file(file_url, destination_path)
print(f"File downloaded successfully to {destination_path}")

#Extract file from ZIP
current_date = datetime.now()
yesterday = current_date - timedelta(days=1)    
date = current_date.strftime('%Y%m%d')

zip_file_path = '/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database.zip'
file_to_extract = f'IAC_Database_{date}.xls'
extraction_path = '/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files'

extract_file(zip_file_path, file_to_extract, extraction_path)

print(f"File '{file_to_extract}' extracted successfully to '{extraction_path}'.")

#Convert Format from XLS to XLSX
input_path = f"/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database_{date}.xls"
output_path = f"/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database_{date}.xlsx"
    
convert_xls_to_xlsx(input_path, output_path)

#Extract Assesment data
source_workbook = openpyxl.load_workbook(f"/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database_{date}.xlsx")
source_sheet = source_workbook['ASSESS']  #name of your source sheet

destination_workbook = openpyxl.load_workbook('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')
default_sheet = destination_workbook.active
default_sheet.title = "ASSESS"  
destination_sheet = destination_workbook['ASSESS']
target_value = 'UC'

copy_first_rows(source_sheet, destination_workbook, destination_sheet)
copy_rows_with_values(source_sheet, destination_workbook, destination_sheet, target_value)
destination_workbook.save('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')
print("Assessment data extracted successfully!")

#Extract Recommendation data
source_workbook = openpyxl.load_workbook(f"/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database_{date}.xlsx")
source_sheet = source_workbook['RECC5']  #name of your source sheet

destination_workbook = openpyxl.load_workbook('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')
new_sheet = destination_workbook.create_sheet(title="RECC")
destination_sheet = destination_workbook['RECC']
target_column_index = 2  

copy_first_row(source_sheet, destination_workbook, destination_sheet)

for i in range(1, 10):
    target_value = 'UC230' + str(i)
    copy_rows_with_value(source_sheet, destination_workbook, destination_sheet, target_value, target_column_index)

for i in range(10, 100):
    target_value = 'UC23' + str(i)
    copy_rows_with_value(source_sheet, destination_workbook, destination_sheet, target_value, target_column_index)

destination_workbook.save('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')
print("Recommendation data extracted successfully!")

#Term copy 
copy_term(date)
print("Term data copied successfully!")
#ARC codes
arc_workbook = '/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx'
arc_code_sheet(arc_workbook)
print("ARC Codes imported successfully!")
print("File Prepared!")

exit()