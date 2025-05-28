import os
import requests
from bs4 import BeautifulSoup

# Base URL for the archive pages
base_url = "https://www.in.gov/igc/publications/archived-monthly-revenue-reports/"
years = ["2019", "2020", "2021", "2022"]

# Directory to save downloaded Excel files
output_dir = r"C:\Users\Julien.Whitter\Downloads\Temporary Files\State Sheets\IN\Archive"

# Ensure the output directory exists
os.makedirs(output_dir, exist_ok=True)

# Loop through each year and download Excel files
for year in years:
    year_url = f"{base_url}archived-monthly-revenue-reports-{year}/"
    print(f"Processing archive for {year}...")

    try:
        # Fetch the HTML content of the archive page
        response = requests.get(year_url)
        response.raise_for_status()  # Raise an error for HTTP issues

        # Parse the HTML content
        soup = BeautifulSoup(response.text, 'html.parser')

        # Find all links to Excel files
        links = soup.find_all('a', href=True)
        for link in links:
            href = link['href']
            if href.endswith(".xlsx"):  # Check if the link points to an Excel file
                # Full URL to the Excel file
                file_url = f"https://www.in.gov{href}"
                file_name = os.path.basename(href)

                # Download the Excel file
                file_path = os.path.join(output_dir, file_name)
                print(f"Downloading {file_name} from {file_url}...")
                file_response = requests.get(file_url)
                with open(file_path, 'wb') as file:
                    file.write(file_response.content)
    except Exception as e:
        print(f"Error processing {year_url}: {e}")

print("Download completed!")
