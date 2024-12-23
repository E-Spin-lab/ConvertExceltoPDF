import argparse
import glob
import os
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from pyhanko.sign.fields import SigFieldSpec, append_signature_field, VisibleSigSettings
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from PyPDF2 import PdfReader, PdfWriter

## Declare Flags
parser = argparse.ArgumentParser(description='Batch convert xlsx files to pdf')
parser.add_argument('--directory', action="store", dest='directory', default=0)
parser.add_argument('--password ', action="store", dest='password', default=0)
parser.add_argument('--output', action="store", dest='output', default='pdf_outputs')
args = parser.parse_args()
DIRECTORY = args.directory
PASSWORD = args.password
OUTPUTFOLDER = args.output

## Check Flags
if DIRECTORY == 0:
    print("please include the full path to the directory that contains the .xlsx files \n Example python .\ConvertExcelPDF.py --directory \"C:/Users/Your.Name/Downloads/Example Reports/\" --password password \n (note file paths with spaces must be enclosed by quotation marks)" , flush=True)
    quit()
else:
    Directory = DIRECTORY.replace("\\", "/")
    if Directory[-1] != "/":
        Directory = Directory + "/*.xlsx"
    else:
        Directory = Directory + "*.xlsx"
if PASSWORD == 0:
    print("please include an encryption password \n Example python .\ConvertExcelPDF.py --directory C:/Users/Your.Name/Downloads/Example Reports// --password password \n (note file paths with spaces must be enclosed by quotation marks)", flush=True)
    quit()

## Create output folder
OutputDirectory = Directory.split("*.xlsx")[0] + OUTPUTFOLDER + "/"
if not os.path.exists(OutputDirectory):
    os.makedirs(OutputDirectory)


## Loop through all ".xlsx files in specified directory"
ExcelList = glob.glob(Directory)
for file_path in ExcelList:
    file_name = file_path.split("\\")[-1].split(".xlsx")[0]
    pdf_output = OutputDirectory + file_name  + ".pdf"
    pdf_output2 = OutputDirectory + file_name  + "encrypt.pdf"
    print("##################################\n##################################\n##################################\n")
    print(f"Converting {file_name} to pdf")
    
    ## Convert Excel to pandas df
    df = pd.read_excel(file_path)
    df = df.fillna('')

    ## Clean up strings
    df = df.replace("\r", " ", regex=True)
    df = df.replace("\n", " ", regex=True)
    df = df.replace("  ", " ", regex=True)
    df.rename(columns=lambda x: x.replace('\n', ''), inplace=True)
    df.rename(columns=lambda x: x.replace('_', '\n'), inplace=True)

    ## Convert df to reportlab format
    data = [df.columns.tolist()] + df.values.tolist()
    table = Table(data)

    ## Stylize the Table
    style = TableStyle([
        ('FONTSIZE', (0,0), (-1,-1), 12),
        ('BACKGROUND', (0, 0), (-1, 0), colors.ReportLabBlue),
        ('ROWBACKGROUNDS', (0,1),(-1,-1), (colors.whitesmoke, colors.lightgrey)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
        ('TOPPADDING', (0, 0), (-1, -1), 2.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey)
    ])

    table.setStyle(style)

    ## Determine table dimentions for pdf
    t = Table(data, style=style)
    table_w, table_h = t.wrap(0, 0)
    BottomSpace = 150
    VerticalMargin = 10
    
    ## Save pdf
    pagesize = (table_w + VerticalMargin + VerticalMargin, table_h + BottomSpace)
    doc = SimpleDocTemplate(filename=pdf_output, 
        pagesize=pagesize, 
        rightMargin=VerticalMargin, 
        leftMargin=VerticalMargin, 
        topMargin=0, 
        bottomMargin=0)
    elements = []
    elements.append(table)
    doc.build(elements)

    ## Add signature field to pdf
    with open(pdf_output, 'rb+') as temp_pdf:
        w = IncrementalPdfFileWriter(temp_pdf, strict = False)
        append_signature_field(w, 
            SigFieldSpec(sig_field_name="Signature",
            on_page = 0,
            box = (VerticalMargin, BottomSpace - 25, 600, BottomSpace - 100),
            visible_sig_settings = VisibleSigSettings(rotate_with_page=True)
            )
        )
        w.write_in_place()
    temp_pdf.close()

    ## Change PDF Permissions 
    out = PdfWriter()
    file = PdfReader(pdf_output)
    num = len(file.pages)
    for idx in range(num):
        page = file.pages[idx] 
        out.add_page(page)
    ## Permissions flags based on Table 22
    ## https://developer.adobe.com/document-services/docs/assets/35e4369068f86065372c18787171a17e/PDF_ISO_32000-1.pdf
    ## Allowed: Printing, Content Copying for Accessibility, Filling of form fields, Signing, and Creating of Template Pages
    out.encrypt(user_password='', owner_pwd=PASSWORD, permissions_flag=0b001111000100)
    with open(pdf_output, "wb") as f:
        out.write(f)
    f.close()

    print("##################################\n##################################\n##################################\n")
print("##################################\n##################################\n##################################\n")
print('Complete!')