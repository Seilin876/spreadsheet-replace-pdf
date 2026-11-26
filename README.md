# Batch Replace Text in Excel and CSV Files and Export to PDF

Run this tool from Python source on Windows. It uses the built-in PDF export feature in desktop Microsoft Excel, which must be installed on the computer. The tool's prompts are in English.

## Set up a new Windows computer

### 1. Verify that desktop Excel can export to PDF (Optional)

Open Microsoft Excel, create a blank workbook, and enter some text. Choose **File → Save As** and save it as a PDF. If you only have Excel for the web, ask your IT team to install the desktop version. [Microsoft Excel PDF export documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.exportasfixedformat)

### 2. Install Python 3.12 (64-bit)

Use your company's software center or the installation method provided by IT. If your company allows the official Python installer, select **Add python.exe to PATH** if the option appears. [Python downloads for Windows](https://www.python.org/downloads/windows/)

After installation, open a new PowerShell window and check the version.

```powershell
py -3.12 --version
```

If PowerShell cannot find `py`, try `python --version`. If it reports `Python 3.12.x`, use `python -m venv .venv` for the first command in step 4.

### 3. Extract the source files

On the GitHub repository page, choose **Code → Download ZIP**. Extract the ZIP in File Explorer, then open the folder containing both `replace_and_export.py` and `requirements.txt`. Type `powershell` in the File Explorer address bar and press Enter to open PowerShell in that folder. [GitHub repository](https://github.com/Seilin876/spreadsheet-replace-pdf)

### 4. Create a virtual environment and install the dependencies

Run these commands in order. The last command prints `OK` if Python can load `openpyxl`, Pillow, and `pywin32`. If you already have a virtual environment, rerun the last two commands after updating the source files.

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -c "import openpyxl, win32com.client; from PIL import Image; print('OK')"
```

If your company network blocks package downloads, ask IT for an approved Python package source.

## Run the tool

Open PowerShell in the folder containing the program and run this command each time you use it.

```powershell
.\.venv\Scripts\python.exe .\replace_and_export.py
```

Respond to the prompts in order.

1. **Folder path** — Enter the path to the folder containing the files to process.
2. **Text to find** — Enter the text to search for. Matching is case-sensitive.
3. **Replacement text** — Enter the new text. Press Enter with the field empty to delete matching text.

Test with copies of your files the first time. The tool saves changes to the original files and exports PDFs to the same folder. For example, `report.xlsx` produces `report.xlsx.pdf`. An existing PDF with the same name is overwritten.

- The tool processes `.xlsx`, `.xlsm`, `.xls`, and `.csv` files directly in the selected folder. It replaces text in cells across all Excel worksheets and preserves numbers and formulas.
- Excel uses the workbook's print settings to export the PDF. When text changes in a legacy `.xls` file, format conversion may affect complex layouts or features. Edited cells with mixed fonts or colors use the cell's default formatting.
- CSV files support UTF-8, UTF-16 and UTF-32 with a byte order mark, CP950, and GB18030. The tool preserves the original encoding and rewrites the field quoting. It reports a file as failed if the replacement text cannot be saved in the original encoding.
- The tool exports a PDF even when there are no matches. If PDF export fails, it leaves the original file intact and continues with the remaining files.

## Changelog

### 2026-11-26

- Added Pillow to preserve supported embedded images when saving Excel workbooks after text replacement. Verified the fix with the workbook that originally lost its images when run from source on Python 3.12.
- Excel text replacement now stops before saving if Pillow is missing and displays the dependency installation command.
