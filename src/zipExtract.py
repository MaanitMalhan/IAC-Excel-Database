import zipfile, os, time
from datetime import datetime, timedelta

# Get the current date and time
current_date = datetime.now()

# Calculate the date from yesterday
yesterday = current_date - timedelta(days=1)
date = yesterday.strftime('%Y%m%d')

# Extracts a file from a zip file to a specified path
def extract_file(zip_file_path, file_to_extract, extraction_path):
    try:
        # Open the zip file
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            # Extract the specified file to the extraction path
            zip_ref.extract(file_to_extract, extraction_path)

        print(f"File '{file_to_extract}' extracted successfully to '{extraction_path}'.")

    except zipfile.BadZipFile as e:
        print(f"Error: {zip_file_path} is not a valid zip file.")
    except FileNotFoundError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Example usage:
zip_file_path = '/Users/maanitmalhan/Documents/IAC Center/excel-data-iac/IAC_Database.zip'
file_to_extract = f'IAC_Database_{date}.xls'
extraction_path = '/Users/maanitmalhan/Documents/IAC Center/excel-data-iac'

# Make sure the extraction directory exists
os.makedirs(extraction_path, exist_ok=True)

extract_file(zip_file_path, file_to_extract, extraction_path)
