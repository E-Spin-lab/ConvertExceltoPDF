# Convert Migration Verification Reports to pdf
This repository contains ConvertExcelPDF, a python script for converting excel documents to pdfs. Pdfs are created with a signature field and only certain permissions enabled.

## PDF permissions
#### PDF will be created with the following permissions enabled:

-Printing

-Content Copying for accessibility

-Filling of form fields

-Signing

-Creation of Template Pages

#### PDF will be created with the following permissions disabled:

-Changing the Document

-Document Assembly

-Content Copying

-Page Extraction

-Commenting

## Requirements
The python script utilizes the following python libraries: argparse, glob, os, pandas, openpyxl, pyhanko, reportlab, and PyPDF2.

To create a virtual environment with all required libraries run the following code (recommended)

        conda env create -f requirements.yml
To acivate the nely created environment run

        conda activate ExcelPDF
 

## Usage
Example 

        python .\ConvertExcelPDF_03.py --directory "C:/Users/Your.Name/ Downloads/Example Reports/" --password password --output pdf_output
### Options
--directory: (Required) File path to the excel files. Must be enclosed in parenthesis if the file path contains spaces.

--password: (Required) Password that is needed to change the pdf permissions. 

--output: (Optional) Directory within filepath where pdfs will be saved. Default string is pdf_output
