import os
import pandas as pd

# Define the directory containing the Excel workbooks
directory = r"C:\Users\Julien.Whitter\Downloads\Temporary Files\State Sheets\MD\Data"

# Output file path for the combined workbook
output_file = os.path.join(directory, "Combined_Workbook.xlsx")

# Create a writer object for the combined workbook
writer = pd.ExcelWriter(output_file, engine='openpyxl')

# Iterate through all files in the directory
for file in os.listdir(directory):
    # Check if the file is an Excel workbook and avoid the output file
    if file.endswith((".xlsx", ".xls")) and file != "Combined_Workbook.xlsx":
        file_path = os.path.join(directory, file)
        
        try:
            # Extract the "Month-Year" from the file name
            base_name = os.path.splitext(file)[0]
            month_year = "-".join(base_name.split("-")[:2])  # Assumes consistent naming like "April-2023-Sports-Wagering-Data"
            
            # Load the workbook
            xls = pd.ExcelFile(file_path)
            
            # Iterate through each sheet in the workbook
            for i, sheet_name in enumerate(xls.sheet_names, start=1):
                # Read the sheet into a DataFrame
                df = pd.read_excel(xls, sheet_name=sheet_name)
                
                # Create a unique sheet name using the extracted month-year and sheet index
                combined_sheet_name = f"{month_year}-{i}"
                
                # Write the DataFrame to the combined workbook
                df.to_excel(writer, sheet_name=combined_sheet_name, index=False)
        except Exception as e:
            print(f"Error processing file {file}: {e}")

# Save and close the combined workbook
writer.close()
print(f"Combined workbook saved to {output_file}")
