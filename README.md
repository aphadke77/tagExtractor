# tagExtractor

Create a Python script, which will read data from a CSV file containing a tag list then use this as a reference and process PDF files in a folder, 
identify if any of the items in the reference list exist on the PDF document, 
identify the page on which the list item is found then Store it, 
the page number the tag from the list into an Excel file with the flag Weather tag from the list found yes or no

The script still starts by importing the necessary libraries such as os, pandas, PyPDF2, and openpyxl.
It loads the reference list from a CSV file named 'reference_list.csv' and stores it in a Pandas DataFrame.
An Excel workbook is created to store the results, and a worksheet is added to the workbook with headers 'File Name', 'TAG_NAME', 'Found', and 'Page Number'.
The script specifies the folder containing the PDF files with the pdf_folder variable. You've set this to a specific path on your local system. Ensure that this path is accurate and contains the PDF files you want to process.
It then enters a loop that traverses through the directory and subdirectories of the pdf_folder.
For each file found with a '.pdf' extension, the script opens the PDF file using the PyPDF2 library.
It iterates through each page of the PDF using a for loop that loops through the pages using the pdf_reader.pages.
For each page, it extracts the text content using page.extract_text().
It searches for each tag in the reference list within the extracted text, and for each tag found, it appends a row to the Excel worksheet with information about the file name, tag name, whether the tag was found ('Yes' or 'No'), and the page number where it was found.
After processing all PDF files, the script saves the results to an Excel file named 'tag_results.xlsx'.
Finally, it prints a message indicating where the results have been saved.
Please note that the script assumes that the provided pdf_folder contains the PDF files you want to process. Make sure to adjust the pdf_folder variable to match the folder path on your system where your PDF files are located.




