import os
import re
import pdfplumber
import pandas as pd

# Define the input and output paths
input_dir = r"C:\Users\Julien.Whitter\Downloads\Temporary Files\State Sheets\OR\Files"
output_dir = r"C:\Users\Julien.Whitter\Downloads\Temporary Files\State Sheets\OR\Converted"

# Ensure the output directory exists
os.makedirs(output_dir, exist_ok=True)

# Function to extract tables from PDF and save as Excel
def convert_pdf_to_excel(pdf_path, xlsx_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_tables = []
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    # Convert table to a pandas DataFrame
                    df = pd.DataFrame(table[1:], columns=table[0])  # First row as header
                    all_tables.append(df)
            
            # Combine all DataFrames into one Excel file with multiple sheets
            if all_tables:
                with pd.ExcelWriter(xlsx_path) as writer:
                    for i, table_df in enumerate(all_tables, start=1):
                        table_df.to_excel(writer, index=False, sheet_name=f"Table_{i}")
                print(f"Converted {pdf_path} to {xlsx_path}")
            else:
                print(f"No tables found in {pdf_path}")
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")

# Process each PDF in the input directory
for filename in os.listdir(input_dir):
    if filename.endswith('.pdf'):
        # Extract month and year from the file name
        match = re.search(r'(\w+)\s(\d{4})', filename)
        if match:
            month, year = match.groups()
            xlsx_name = f"{month}_{year}.xlsx"
            pdf_path = os.path.join(input_dir, filename)
            xlsx_path = os.path.join(output_dir, xlsx_name)
            
            # Convert PDF to Excel
            convert_pdf_to_excel(pdf_path, xlsx_path)

print("PDF to XLSX conversion completed!")
