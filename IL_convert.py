import pandas as pd
import os

# Define input and output paths
input_folder = r'C:\Users\Julien.Whitter\Downloads\Temporary Files\State Sheets\IL\Downloads'
output_file = 'combined_workbook.xlsx'

# Initialize Excel writer
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for file in os.listdir(input_folder):
        if file.endswith('.csv'):
            try:
                # Skip metadata and read the header
                df = pd.read_csv(
                    os.path.join(input_folder, file),
                    skiprows=4,  # Skip metadata rows
                    header=0     # Use the correct header row
                )
                
                # Drop unnamed or empty columns
                df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

                # Check if DataFrame is empty
                if df.empty:
                    print(f"Skipping empty or malformed file: {file}")
                    continue
                
                # Generate a sheet name (limit to 31 characters for Excel compatibility)
                sheet_name = os.path.splitext(file)[0][:31]
                
                # Write DataFrame to Excel sheet
                df.to_excel(writer, index=False, sheet_name=sheet_name)
                
            except Exception as e:
                print(f"Error processing file {file}: {e}")
    
    print(f"Combined workbook created: {output_file}")
