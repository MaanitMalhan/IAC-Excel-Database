import requests # pip install requests

def download_file(url, destination):
    try:
        # Make a GET request to the URL
        response = requests.get(url, stream=True)

        # Check if the request was successful (status code 200)
        if response.status_code == 200:
            # Open the destination file in binary mode and write the content
            with open(destination, 'wb') as file:
                for chunk in response.iter_content(chunk_size=128):
                    file.write(chunk)
            print(f"File downloaded successfully to {destination}")
        else:
            print(f"Failed to download file. Status code: {response.status_code}")

    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Example usage:
file_url = 'https://iac.university/storage/IAC_Database.zip'
destination_path = '/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/IAC_Database.zip'

download_file(file_url, destination_path)
print("File downloaded successfully!")