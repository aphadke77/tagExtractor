import os
import pandas as pd
import PyPDF2
from openpyxl import Workbook
from openpyxl.styles import Font

# Load the reference list from a CSV file
reference_list = pd.read_csv('reference_list.csv')

# Create an Excel workbook to store the results
workbook = Workbook()
worksheet = workbook.active
worksheet.title = 'Tag Results'
worksheet['A1'] = 'Tag'
worksheet['B1'] = 'Found'
worksheet['C1'] = 'Page Number'
bold_font = Font(bold=True)
worksheet['A1'].font = bold_font
worksheet['B1'].font = bold_font
worksheet['C1'].font = bold_font

# Specify the folder containing the PDF files
pdf_folder = 'pdf_files/'

# Loop through each PDF file in the folder
for root, dirs, files in os.walk(pdf_folder):
    for file in files:
        if file.endswith('.pdf'):
            pdf_file_path = os.path.join(root, file)

            # Open the PDF file
            pdf_file = open(pdf_file_path, 'rb')
            pdf_reader = PyPDF2.PdfFileReader(pdf_file)

            # Loop through each page of the PDF
            for page_num in range(pdf_reader.getNumPages()):
                page = pdf_reader.getPage(page_num)
                page_text = page.extractText()

                # Check if each tag in the reference list is present on the page
                for index, row in reference_list.iterrows():
                    tag = row['Tag']
                    if tag in page_text:
                        worksheet.append([tag, 'Yes', page_num + 1])
                    else:
                        worksheet.append([tag, 'No', ''])

# Save the results to an Excel file
workbook.save('tag_results.xlsx')
