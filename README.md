# Software Inventory Exporter

This project contains two Python scripts that read all installed programmes on a Windows computer and export them to an overview file.

## Features
- Reads software information from the **Windows Registry**.
- Exports name, version, language, bits (32/64), publisher, product code and an additional column ‘type of programme’.
- Supports two formats:
  - **ODS** (LibreOffice Calc)
  - **XLSX** (Microsoft Excel, also compatible with LibreOffice and Google Sheets)

## Files
- `export_libreoffice.py`
Exports the list to a **LibreOffice Calc (.ods)** file.
Uses the [`odfpy`](https://pypi.org/project/odfpy/) library.

- `export_excel.py`  
  Exports the list to an **Excel (.xlsx)** file.  
  Uses the [`openpyxl`](https://pypi.org/project/openpyxl/) library.

Both scripts create a file titled **‘Programmes on my laptop’** and a table with columns:

1. Name of the software  
2. Version  
3. Language  
4. Bits (32/64)  
5. Publisher  
6. Product code  
7. Type of programme (default: *Unknown*)  

The table starts in row 4, with a header row and your name at the top in row 2.

## Requirements
- **Windows** (script reads from the Windows Registry)
- **Python 3.8+**
- Install libraries with pip:

```bash
pip install odfpy openpyxl
