import os
from tkinter import Tk
from tkinter import filedialog
import pandas as pd
import qrcode
from docx2pdf import convert
from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage

def select_file(prompt):
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=prompt)
    return file_path


def select_folder(prompt):
    root = Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title=prompt)
    return folder_path

# Load the Excel sheet into a DataFrame
prompt1="Select Excel File"
print("Select an Excel File")
excel_file = select_file(prompt1)
print("Selected Excel File Path: ", excel_file)
sheet_name=input("Enter Sheet Name: ")
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Load the Word document template
prompt2="Select a .docx Template Word File"
print("Select .docx Template Word File: ")
template_file = select_file(prompt2)
print("Selected .docx Template Word File: ", template_file)
template = DocxTemplate(template_file)


# Select the output folder location
prompt3="Select Output Folder"
print("Select Output Folder: ")
output_folder = select_folder(prompt3)
print("Selected Output Folder: ", output_folder)

qr_size=float(input("Enter QR image size in inches: "))
qr_slice_start, qr_slice_end=map(int,input("Enter start and end index: ").split())


# Create a directory for storing the QR code images
qr_dir = os.path.join(output_folder, "QRCODE IMAGES")
if not os.path.exists(qr_dir):
    os.makedirs(qr_dir)

# Create a directory for storing the PDF documents in the output folder
pdf_dir = os.path.join(output_folder, "PDF")
if not os.path.exists(pdf_dir):
    os.makedirs(pdf_dir)

# Create a directory for storing the DOCX documents in the output folder
docx_dir = os.path.join(output_folder, "DOCX")
if not os.path.exists(docx_dir):
    os.makedirs(docx_dir)

# Initialize counts
countpdf=0
countdocx=0
countqr=0

# Loop over each row of data in the DataFrame
for i, row in df.iterrows():
    # Generate a QR code for the content in this row
    qr_content = str(row["QRNUM"])
    qr = qrcode.QRCode(version=1, box_size=40, border=0)
    qr.add_data(qr_content)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    # if qr_slice_end is greater than slice length
    qr_slice_end = min(qr_slice_end, len(qr_content))
    qr_sliced=qr_content[qr_slice_start:qr_slice_end]

    # Save the QR code image to a file
    qr_filename = os.path.join(qr_dir, f"{qr_content}.png")
    img.save(qr_filename)
    countqr+=1
    print("SAVED QR CODE IMAGE", countqr, ":", qr_content, "IN 'QRCODE IMAGES' FOLDER")

    # Load the Word document template for each iteration
    template = DocxTemplate(template_file)

    # Fill in the placeholders in the Word template
    context = {
        "qr_code": InlineImage(template, qr_filename, width=Inches(qr_size), height=Inches(qr_size)),
        "qr_num": qr_sliced
    }
    template.render(context)

   # Save the resulting Word document with a filename based on the content
    doc_filename = os.path.join(docx_dir, f"{qr_content}.docx")
    template.save(doc_filename)
    countdocx+=1
    print("SAVED .docx FILE", countdocx, ":", qr_content, " IN SELECTED OUTPUT FOLDER")
    print("Converting DOCX to PDF...")
    try:
        pdf_filename = os.path.join(pdf_dir, f"{qr_content}.pdf")
        convert(doc_filename, pdf_filename)
        countpdf += 1
        print("SAVED .pdf FILE", countpdf, ":", qr_content, " IN PDF FILES FOLDER")
    except Exception as e:
        print("Error converting DOCX to PDF:", str(e))
    print()
    print()
print("Sucessfully Completed!!!!!")
print("Total Number of DOCX files Generated:", countdocx)
print("Total Number of PDF files Generated:", countpdf)