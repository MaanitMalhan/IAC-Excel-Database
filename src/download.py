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
        else:
            print(f"Failed to download file. Status code: {response.status_code}")

    except Exception as e:
        print(f"An error occurred: {str(e)}")
