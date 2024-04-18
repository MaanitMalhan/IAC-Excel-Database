import zipfile, os, time
from datetime import datetime, timedelta


# Extracts a file from a zip file to a specified path
def extract_file(zip_file_path, file_to_extract, extraction_path):
    try:
        # Open the zip file
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            # Extract the specified file to the extraction path
            zip_ref.extract(file_to_extract, extraction_path)

        #print(f"File '{file_to_extract}' extracted successfully to '{extraction_path}'.")

    except zipfile.BadZipFile as e:
        print(f"Error: {zip_file_path} is not a valid zip file.")
    except FileNotFoundError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

