import os
import fitz  # PyMuPDF
import pandas as pd

# Define the input and output paths
input_dir = r"C:\Users\Julien.Whitter\Downloads\Temporary Files\State Sheets\NJ\OSB\Handle"
output_dir = r"C:\Users\Julien.Whitter\Downloads\Temporary Files\State Sheets\NJ\OSB\Converted"

# Ensure the output directory exists
os.makedirs(output_dir, exist_ok=True)

# Function to extract the last page's table as text and convert to an Excel file
def extract_table_from_last_page(pdf_path, xls_path):
    try:
        # Open the PDF
        pdf_document = fitz.open(pdf_path)
        # Get the last page
        last_page = pdf_document[-1]
        # Extract text from the last page
        text = last_page.get_text("text")

        # Split the text into lines
        lines = text.splitlines()

        # Assume the table starts after a specific marker (adjust based on your data)
        table_start_index = 0
        for i, line in enumerate(lines):
            if "Casino Licensee Total" in line:  # Adjust the marker based on the PDF
                table_start_index = i
                break

        # Extract table rows starting from the marker
        table_lines = lines[table_start_index:]

        # Parse each line into columns by splitting on whitespace or delimiters (adjust as needed)
        table_data = [line.split() for line in table_lines if line.strip()]

        # Create a DataFrame from the parsed table data
        df = pd.DataFrame(table_data)

        # Save the DataFrame as an Excel file
        df.to_excel(xls_path, index=False, header=False)
        print(f"Converted last page of {pdf_path} to {xls_path}")

    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")

# Process each PDF in the input directory
for filename in os.listdir(input_dir):
    if filename.endswith('.pdf'):
        # Define the output XLS file path
        xls_filename = filename.replace('.pdf', '.xlsx')
        pdf_path = os.path.join(input_dir, filename)
        xls_path = os.path.join(output_dir, xls_filename)
        
        # Extract the table from the last page
        extract_table_from_last_page(pdf_path, xls_path)

print("PDF to XLS conversion completed!")
