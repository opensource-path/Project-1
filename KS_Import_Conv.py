import os
import requests
from urllib.parse import urljoin
from bs4 import BeautifulSoup
from PyPDF2 import PdfReader
import pandas as pd

def download_pdfs(url, output_directory):
    # Send a GET request to the URL
    response = requests.get(url)
    
    # Check if the request was successful
    if response.status_code == 200:
        # Parse HTML content
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Find all links to PDFs on the page
        pdf_links = soup.find_all('a', href=True)
        
        # Iterate through each PDF link
        for link in pdf_links:
            if link['href'].endswith('.pdf'):
                pdf_url = urljoin(url, link['href'])
                pdf_filename = link['href'].split('/')[-1]
                pdf_path = os.path.join(output_directory, pdf_filename)
                
                # Download the PDF
                with open(pdf_path, 'wb') as file:
                    file.write(requests.get(pdf_url).content)
                print(f"Downloaded: {pdf_filename}")

def convert_pdfs_to_excels(input_directory, output_directory):
    # Iterate through each file in the input directory
    for filename in os.listdir(input_directory):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(input_directory, filename)
            
            # Initialize a PDF reader object
            pdf_reader = PdfReader(pdf_path)
            
            # Initialize an empty list to store text from each page
            pages_text = []
            
            # Iterate through each page of the PDF
            for page_num in range(len(pdf_reader.pages)):
                # Extract text from the current page
                page_text = pdf_reader.pages[page_num].extract_text()
                pages_text.append(page_text)
            
            # Concatenate text from all pages
            full_text = '\n'.join(pages_text)
            
            # Create a pandas DataFrame from the concatenated text
            df = pd.DataFrame({'Content': [full_text]})
            
            # Define the output Excel file path
            output_excel_path = os.path.join(output_directory, os.path.splitext(filename)[0] + '.xlsx')
            
            # Write the DataFrame to an Excel file
            df.to_excel(output_excel_path, index=False)
            print(f"Conversion successful. Excel file saved at: {output_excel_path}")

# Specify the URL containing the PDFs
url = "https://www.kslottery.com/publications/sports-monthly-detail/"
# Specify the directory where PDFs will be downloaded
output_directory = r'C:\Users\Julien.Whitter\Documents\Gambling\Imports\KS\Import'

# Create the output directory if it doesn't exist
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Download all the PDFs from the URL
download_pdfs(url, output_directory)

# Specify the input directory containing the downloaded PDFs
input_directory = output_directory
# Specify the output directory where Excel files will be saved
output_excel_directory = r'C:\Users\Julien.Whitter\Documents\Gambling\Imports\KS\Import\Converted_Excels'

# Create the output Excel directory if it doesn't exist
if not os.path.exists(output_excel_directory):
    os.makedirs(output_excel_directory)

# Convert all downloaded PDFs to Excel files
convert_pdfs_to_excels(input_directory, output_excel_directory)
