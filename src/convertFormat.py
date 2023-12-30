
from openpyxl import Workbook
import os, time
from datetime import datetime, timedelta
import pandas as pd

# Get the current date and time
current_date = datetime.now()

# Calculate the date from yesterday
date = current_date.strftime('%Y%m%d')


def convert_xls_to_xlsx(input_path, output_path):
    xls_file_path = input_path
    xlsx_file_path = output_path
    xls = pd.ExcelFile(xls_file_path)


    with pd.ExcelWriter(xlsx_file_path, engine='xlsxwriter') as writer:
    # Iterate over each sheet in the .xls file
        for sheet_name in xls.sheet_names:
        # Read the sheet from .xls
            df = pd.read_excel(xls_file_path, sheet_name=sheet_name)

        # Write the sheet to .xlsx
            df.to_excel(writer, sheet_name=sheet_name, index=False)


    print(f"Conversion complete. Output file saved at: {output_path}")


input_path = f"/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database_{date}.xls"
output_path = f"/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database_{date}.xlsx"
    
convert_xls_to_xlsx(input_path, output_path)
print("File converted successfully!")
