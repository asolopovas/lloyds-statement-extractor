import os
import sys
from PyPDF2 import PdfReader
from win32com import client
import re

yr = "2024"

if len(sys.argv) != 2:
    print("Usage: extract-statement.py <PDF filename>")
    sys.exit(1)

pdf_path = sys.argv[1]

if not os.path.exists(pdf_path):
    print(f"File not found: {pdf_path}")
    sys.exit(1)

spreadSheet = client.Dispatch("Excel.Application")
spreadSheet.Workbooks.Add()
spreadSheet.ActiveSheet.Cells(1, 1).Value = "Date"
spreadSheet.ActiveSheet.Cells(1, 2).Value = "Description"
spreadSheet.ActiveSheet.Cells(1, 3).Value = "Amount"

stData = ""

try:
    reader = PdfReader(pdf_path)
    for page in reader.pages:
        stData += page.extract_text()
except Exception as e:
    print(f"Error in reading file: {pdf_path}\n{e}")
    spreadSheet.Quit()
    sys.exit(1)

stData = stData.replace("Customer Services: 0345 606 2174", "")
statementData = re.findall(r"\d{2}\s\w*\s\d{2}\s.*", stData)

# Process extracted data into Excel
for i, row in enumerate(statementData):
    date_match = re.findall(r"\d{2}\s[a-zA-Z]{3,9}", row)
    desc_match = re.findall(r"\d{2}\s\w*\s\d{2}\s\w*\s(.*(?=\s+\d+\.*))", row)
    amount_match = re.findall(r"(\d+,\d{1,3}\.\d{1,2}|\d+\.\d{1,2})", row)

    date = f"{date_match[0]} {yr}" if date_match else ""
    desc = desc_match[0] if desc_match else ""
    amount = amount_match[0] if amount_match else ""

    cr = re.search(r"CR$", row)
    if cr:
        amount = "-" + amount

    spreadSheet.Cells(i + 2, 1).Value = date
    spreadSheet.Cells(i + 2, 2).Value = desc
    spreadSheet.Cells(i + 2, 3).Value = amount

excel_filename = f"statement - {yr}.xlsx"
if os.path.exists(excel_filename):
    os.remove(excel_filename)

try:
    spreadSheet.ActiveWorkbook.SaveAs(os.path.join(os.getcwd(), excel_filename))
    spreadSheet.ActiveWorkbook.Close()
finally:
    spreadSheet.Quit()
