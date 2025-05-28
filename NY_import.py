import os
import requests
from bs4 import BeautifulSoup

# URL of the page
URL = "https://gaming.ny.gov/revenue-reports"

# Directory to save the Excel files
SAVE_DIR = "sports_wagering_reports"

# Ensure the save directory exists
os.makedirs(SAVE_DIR, exist_ok=True)

def fetch_reports(url):
    # Fetch the HTML content
    response = requests.get(url)
    if response.status_code != 200:
        print("Failed to fetch the page.")
        return

    soup = BeautifulSoup(response.content, "html.parser")
    
    # Locate the "Sports Wagering" section
    sports_wagering_section = soup.find("section", id="toc_591")
    if not sports_wagering_section:
        print("Sports Wagering section not found.")
        return

    # Find all Excel report links in the section
    excel_links = sports_wagering_section.find_all("a", href=True)
    excel_files = [
        (link.text.strip(), link["href"])
        for link in excel_links
        if "excel" in link["href"].lower() and "monthly" in link.text.lower()
    ]
    
    # Download each Excel file
    for name, link in excel_files:
        file_name = f"{name.replace(' ', '_')}.xlsx"
        file_path = os.path.join(SAVE_DIR, file_name)
        download_file(link, file_path)

def download_file(url, file_path):
    # Full URL resolution
    base_url = "https://gaming.ny.gov"
    full_url = url if url.startswith("http") else f"{base_url}{url}"

    print(f"Downloading: {full_url}")
    response = requests.get(full_url, stream=True)
    if response.status_code == 200:
        with open(file_path, "wb") as file:
            for chunk in response.iter_content(chunk_size=1024):
                file.write(chunk)
        print(f"Saved: {file_path}")
    else:
        print(f"Failed to download: {full_url}")

if __name__ == "__main__":
    fetch_reports(URL)
