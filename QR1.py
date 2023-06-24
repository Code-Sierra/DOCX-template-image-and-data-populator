import os
import pandas as pd
import qrcode
from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage

# Load the Excel sheet into a DataFrame
excel_file = "QR1excel.xlsx"
sheet_name = "Sheet1"  
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Create a directory for storing the QR code images
qr_dir = "qr_codes"
if not os.path.exists(qr_dir):
    os.makedirs(qr_dir)

# Load the Word document template
template_file = "6x8temp.docx"
template = DocxTemplate(template_file)

# Loop over each row of data in the DataFrame
for i, row in df.iterrows():
    # Generate a QR code for the content in this row
    qr_content = str(row["QRNUM"])
    qr = qrcode.QRCode(version=1, box_size=40, border=0)
    qr.add_data(qr_content)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    # Save the QR code image to a file
    qr_filename = os.path.join(qr_dir, f"{qr_content}.png")
    img.save(qr_filename)

    # Load the Word document template for each iteration
    template = DocxTemplate(template_file)

    # Fill in the placeholders in the Word template
    context = {
        "qr_code": InlineImage(template, qr_filename, width=Inches(4), height=Inches(4)),
        "qr_num": qr_content
    }
    template.render(context)

    # Save the resulting Word document with a filename based on the content
    doc_filename = os.path.join("output", f"{qr_content}.docx")
    template.save(doc_filename)
