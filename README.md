# BarBendindSchedule_pdf_to_excel

What is this tool?
This is an automated tool designed to take Bar Bending Schedule (BBS) documents in PDF format and instantly convert them into fully functional Excel spreadsheets.

If your team normally spends hours manually typing customer order details, dimensions, and rebar requirements from a PDF into an Excel system, this tool does that entire job for you in seconds.

Key Benefits
Eliminates Manual Data Entry: No more typing out numbers row by row. The tool "reads" the PDF and organizes the data into the correct Excel columns automatically.

Zero Human Error: By automating the transfer, it ensures that the exact measurements and quantities from the original PDF make it into the final spreadsheet without typos.

Automatic Calculations & Drawings: Once the data is in Excel, the tool automatically calculates the final steel weights and even generates the visual rebar bending shapes directly inside the spreadsheet.

Faster Processing: What used to take hours of tedious paperwork can now be completed instantly, speeding up the workflow for production, support, and sales teams.

How It Works (In 3 Simple Steps)
Input: You provide the original Bar Bending Schedule PDF file.

Process: The tool scans the document, identifies the tables, and securely extracts all the necessary information.

Output: It instantly generates a ready-to-use Excel file containing all the structured data, weight calculations, and rebar diagrams.

An automated data processing pipeline that extracts tabular rebar data from Bar Bending Schedule (BBS) PDF documents and converts it into formatted, calculation-ready Excel spreadsheets. This tool streamlines steel manufacturing and trading workflows by eliminating manual data entry.

## Features

* **Advanced PDF Table Extraction:** Leverages `camelot` and `pdfplumber` to accurately parse complex rebar tables, bending dimensions, and specifications directly from BBS PDFs.
* **Data Structuring:** Cleans, organizes, and maps the extracted rebar data using `pandas` for seamless spreadsheet integration.
* **Excel VBA Automation:** Uses `win32com` to interface with Excel directly, automatically triggering custom VBA macros to generate visual rebar shapes and calculate accurate rebar weights across hundreds of rows.
* **Document Handling:** Utilizes `pypdf` for initial document processing, splitting, or merging as needed before data extraction.

## Built With

* **Python 3.x**
* [camelot-py](https://camelot-py.readthedocs.io/) - For precise tabular data extraction from PDFs.
* [pdfplumber](https://github.com/jsvine/pdfplumber) - For deep inspection and text/table extraction from PDFs.
* [pandas](https://pandas.pydata.org/) - For data manipulation and analysis.
* [pypdf](https://pypi.org/project/pypdf/) - For PDF file manipulation.
* [pywin32 (win32com)](https://pypi.org/project/pywin32/) - For COM interoperability to execute Excel VBA macros.

## Prerequisites

Before running the scripts, ensure you have the following installed:
* Python 3.x
* Microsoft Excel (required for `win32com` macro execution)
* Ghostscript (required by `camelot` for PDF parsing)

## Installation

1. Clone the repository:
   ```bash
   git clone [https://github.com/YH-chunkhai/BarBendindSchedule_pdf_to_excel.git](https://github.com/YH-chunkhai/BarBendindSchedule_pdf_to_excel.git)
   cd BarBendindSchedule_pdf_to_excel
