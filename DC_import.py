import requests
from bs4 import BeautifulSoup
import pandas as pd

# Base URL format
base_url = "https://dclottery.com/olg/financials/{month}-{year}-unaudited"

# Define the date range
months = [
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december"
]
years = range(2020, 2025)

# Output Excel file
output_file = "all_table_data.xlsx"

# Create an Excel writer
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for year in years:
        for month in months:
            # Skip months outside the range
            if year == 2020 and months.index(month) < 4:  # Skip Jan-Apr 2020
                continue
            if year == 2024 and months.index(month) > 9:  # Skip Nov-Dec 2024
                break
            
            # Construct the URL
            url = base_url.format(month=month, year=year)
            
            try:
                # Fetch the webpage content
                response = requests.get(url)
                response.raise_for_status()
                
                # Parse the HTML
                soup = BeautifulSoup(response.content, "html.parser")
                
                # Locate the table
                table_div = soup.find("div", class_="node__content")
                if not table_div:
                    print(f"No table found for {month} {year}. Skipping...")
                    continue
                
                table = table_div.find("table")
                if not table:
                    print(f"No table element found for {month} {year}. Skipping...")
                    continue
                
                # Parse table rows
                rows = table.find_all("tr")
                headers = [header.get_text(strip=True) for header in rows[0].find_all("td")]
                data = [
                    [cell.get_text(strip=True) for cell in row.find_all("td")]
                    for row in rows[1:]
                ]
                
                # Create a DataFrame
                df = pd.DataFrame(data, columns=headers)
                
                # Save to a sheet in the Excel file
                sheet_name = f"{month.capitalize()}_{year}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Data for {month} {year} added to the Excel file.")
            
            except Exception as e:
                print(f"Failed to fetch or process data for {month} {year}: {e}")

print(f"All data saved to {output_file}.")
