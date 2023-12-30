import zipExtract
import download
from datetime import datetime, timedelta

current_date = datetime.now()
#yesterday = current_date - timedelta(days=1)    
date = today = current_date.strftime('%Y%m%d')

file_url = 'https://iac.university/storage/IAC_Database.zip'
destination_path = '/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files'

zip_file_path = '/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database.zip'
file_to_extract = f'IAC_Database_{date}.xls'
extraction_path = '/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files'



download.download_file(file_url, destination_path)
print("File downloaded successfully!")

zipExtract.extract_file(zip_file_path, file_to_extract, extraction_path)
print("File extracted successfully!")