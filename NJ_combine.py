import os
import pandas as pd

# Define the input and output paths
input_dir = r"C:\Users\Julien.Whitter\Downloads\Temporary Files\State Sheets\NJ\OSB\Converted"
output_file = r"C:\Users\Julien.Whitter\Downloads\Temporary Files\State Sheets\NJ\OSB\Combined.xlsx"

# Create an Excel writer object for the combined file
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for filename in os.listdir(input_dir):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(input_dir, filename)
            
            # Read the Excel file into a DataFrame
            df = pd.read_excel(file_path, sheet_name=0)  # Read the first sheet
            
            # Use the filename (without extension) as the sheet name
            sheet_name = os.path.splitext(filename)[0]
            
            # Write the DataFrame to the writer object
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            print(f"Added {filename} as sheet '{sheet_name}'")

print(f"All files have been combined into {output_file}")
