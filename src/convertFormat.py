import xlrd
from openpyxl import Workbook
import zipfile, os, time
from datetime import datetime, timedelta
import pandas as pd

# Get the current date and time
current_date = datetime.now()

# Calculate the date from yesterday
yesterday = current_date - timedelta(days=1)
date = yesterday.strftime('%Y%m%d')



def convert_xls_to_xlsx(input_path, output_path):
    # Read the .xls file into a pandas ExcelFile object
    xls_file = pd.ExcelFile(input_path)

    # Create a new .xlsx file and writer object
    with pd.ExcelWriter(output_path, engine='openpyxl') as xlsx_writer:
        # Iterate through each sheet in the .xls file and write to the .xlsx file
        for sheet_name in xls_file.sheet_names:
            xls_sheet = xls_file.parse(sheet_name)
            xls_sheet.to_excel(xlsx_writer, sheet_name=sheet_name, index=False)

    print(f"Conversion complete. Output file saved at: {output_path}")


    input_path = f"/Users/maanitmalhan/Documents/IAC Center/excel-data-iac/IAC_Database_20231224.xls"#NAME FOR TESTING
    output_path = f"/Users/maanitmalhan/Documents/IAC Center/excel-data-iac/IAC_Database_{date}.xlxs"
    
    convert_xls_to_xlsx(input_path, output_path)
