import os
import pdfplumber
import pandas as pd

# Directory containing the PDF files
pdf_directory = '/users/mikekuo/downloads/contracts'

# List to store extracted data from all PDF files
all_data = []

# Loop through PDF files in the directory
for pdf_file in os.listdir(pdf_directory):
    if pdf_file.endswith('.pdf'):
        pdf_path = os.path.join(pdf_directory, pdf_file)
        with pdfplumber.open(pdf_path) as pdf:
            # Extract the 4th line of the first page
            if len(pdf.pages) >= 1:  # Ensure the PDF has at least 1 page
                first_page = pdf.pages[0]  # Extract the first page (index 0)
                lines_first_page = first_page.extract_text().split('\n')
                if len(lines_first_page) >= 4:  # Ensure the first page has at least 4 lines
                    fourth_line_first_page = lines_first_page[3]  # Extract the fourth line

            # Extract the 6th line of the third page
            if len(pdf.pages) >= 3:  # Ensure the PDF has at least 3 pages
                third_page = pdf.pages[2]  # Extract the third page (index 2)
                lines_third_page = third_page.extract_text().split('\n')
                if len(lines_third_page) >= 6:  # Ensure the third page has at least 6 lines
                    sixth_line_third_page = lines_third_page[5]  # Extract the sixth line

            data = {'Fourth Line First Page': fourth_line_first_page,
                    'Sixth Line Third Page': sixth_line_third_page}
            all_data.append(data)

# Create DataFrame from the collected data
df = pd.DataFrame(all_data)

# Path to the Excel file to save the extracted data
excel_file = '/users/mikekuo/downloads/theexcel.xlsx'

# Write the DataFrame to an Excel file
df.to_excel(excel_file, index=False)

print("Extraction completed.")
