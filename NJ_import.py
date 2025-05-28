import os
import re
import requests
from bs4 import BeautifulSoup
from datetime import datetime

# Define the URL and the output directory
url = "https://www.njoag.gov/about/divisions-and-offices/division-of-gaming-enforcement-home/financial-and-statistical-information/monthly-press-releases-and-statistical-summaries/"
output_dir = r"C:\Users\Julien.Whitter\Downloads\Temporary Files\State Sheets\NJ\OSB\Handle"

# Ensure the output directory exists
os.makedirs(output_dir, exist_ok=True)

# Define the date range
start_date = datetime(2018, 6, 1)  # June 2018
end_date = datetime(2024, 10, 31)  # October 2024

# Month mapping to handle different cases
month_mapping = {
    'JANUARY': '01', 'FEBRUARY': '02', 'MARCH': '03', 'APRIL': '04',
    'MAY': '05', 'JUNE': '06', 'JULY': '07', 'AUGUST': '08',
    'SEPTEMBER': '09', 'OCTOBER': '10', 'NOVEMBER': '11', 'DECEMBER': '12'
}

# Fetch the webpage content
response = requests.get(url)
response.raise_for_status()  # Ensure we notice bad responses

# Parse the webpage content
soup = BeautifulSoup(response.text, 'html.parser')

# Find all links on the page
links = soup.find_all('a', href=True)

# Process each link
for link in links:
    href = link['href']
    text = link.get_text(strip=True).upper()

    # Check if the link text is a month name
    if text in month_mapping:
        # Find the year by checking the previous siblings
        year_tag = link.find_previous(string=re.compile(r'^\d{4}$'))
        if year_tag:
            year = year_tag.strip()
            month = month_mapping[text]
            file_date = datetime(int(year), int(month), 1)

            # Check if the file date is within the specified range
            if start_date <= file_date <= end_date:
                # Construct the full URL if the href is relative
                if not href.startswith('http'):
                    href = requests.compat.urljoin(url, href)

                # Define the output file path
                output_file = os.path.join(output_dir, f"{year}_{month}_{text}.pdf")

                # Download and save the PDF
                try:
                    pdf_response = requests.get(href)
                    pdf_response.raise_for_status()
                    with open(output_file, 'wb') as f:
                        f.write(pdf_response.content)
                    print(f"Downloaded: {output_file}")
                except requests.RequestException as e:
                    print(f"Failed to download {href}: {e}")

print("Download process completed.")
