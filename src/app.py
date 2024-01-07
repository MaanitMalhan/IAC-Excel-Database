import download
import zipExtract
import convertFormat
import assessSheetExtraction
import reccSheetExtraction
import termCopy
from datetime import datetime, timedelta
import openpyxl

current_date = datetime.now()
yesterday = current_date - timedelta(days=1)    
date = yesterday.strftime('%Y%m%d')

workbook = openpyxl.Workbook()
workbook.save('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/SNE_IAC_Database.xlsx')

file_url = 'https://iac.university/storage/IAC_Database.zip'
destination_path = '/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files'

zip_file_path = '/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database.zip'
file_to_extract = f'IAC_Database_{date}.xls'
extraction_path = '/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files'


input_path = f"/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database_{date}.xls"
output_path = f"/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database_{date}.xlsx"

download.download_file(file_url, destination_path)
print("File downloaded successfully!")

zipExtract.extract_file(zip_file_path, file_to_extract, extraction_path)
print("File extracted successfully!")

convertFormat.convert_xls_to_xlsx(input_path, output_path)
print("File converted successfully!")

