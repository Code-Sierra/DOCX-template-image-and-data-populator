Certainly! Below is a sample README file that you can use for your code:

---

# Document Generation and Conversion Utility

This utility generates QR codes for content from an Excel file, embeds these QR codes into Word document templates, and converts the Word documents to PDF format.

## Features

- Load data from an Excel file
- Generate QR codes for specific content from the Excel file
- Embed the generated QR codes into a Word document template
- Convert the Word documents to PDF format
- Save the QR code images, Word documents, and PDF files in specified output folders

## Prerequisites

Ensure that you have the following software and libraries installed:

- Python 3
- Microsoft Word (for the conversion from DOCX to PDF)
- pandas
- qrcode
- docx2pdf
- docx
- docxtpl

## Usage

1. Run the script.
2. Select the Excel file containing the data.
3. Enter the sheet name from the selected Excel file.
4. Select the Word document template file (.docx).
5. Select the output folder where the generated files will be stored.
6. Enter the QR image size in inches.
7. Enter the start and end index for slicing the QR content.

The utility will then generate the QR code images, Word documents, and PDF files in the specified output folder.

## Output Folders

The utility will create three directories in the selected output folder:

1. `QRCODE IMAGES`: Contains the generated QR code images
2. `DOCX`: Contains the generated Word documents with the embedded QR codes
3. `PDF`: Contains the PDF files converted from the Word documents

## Note

On macOS, you may need to grant full disk access to your terminal application to avoid repeated access permission prompts during the DOCX to PDF conversion.

## Author

[Sriram Santhosh Rajkumar]

---
