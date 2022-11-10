from hashlib import new
import os
import PyPDF2
from win32com import client
import re
import json

yr = "2020"
pwd = os.getcwd()

spreadSheet = client.Dispatch("Excel.Application")
spreadSheet.Workbooks.Add()
spreadSheet.ActiveSheet.Cells(1, 1).Value = "Date"
spreadSheet.ActiveSheet.Cells(1, 2).Value = "Description"
spreadSheet.ActiveSheet.Cells(1, 3).Value = "Amount"

stData = ""

for pdf in os.listdir(pwd):
    if pdf.endswith(".xlsx"):
        continue
    pdfFile = open(os.path.join(pwd, pdf), "rb")
    try:
        pdfReader = PyPDF2.PdfFileReader(pdfFile)
        for pageNum in range(pdfReader.numPages):
            pageObj = pdfReader.getPage(pageNum)
            stData += pageObj.extractText()
    except:
        print ("Error in reading file: " + pdf)
        exit(0)
stData = stData.replace("Customer Services: 0345 606 2174", "")
statementData = re.findall(r"\d{2}\s\w*\s\d{2}\s.*", stData)


# map list items values
for i in range(len(statementData)):
    row = statementData[i]
    # save each row to a file
    date = re.findall(r"\d{2}\s[a-zA-Z]{3,9}", row)[0] + " " + yr
    desc = re.findall(r"\d{2}\s\w*\s\d{2}\s\w*\s(.*(?=\s+\d+\.*))", row)
    amount = re.findall(r"(\d+,\d{1,3}\.\d{1,2}|\d+\.\d{1,2})", row)[0]

    cr = re.search(r"CR$", row)
    if cr != None:
        amount = "-" + amount

    spreadSheet.Cells(i+2, 1).Value = date
    if desc:
        spreadSheet.Cells(i+2, 2).Value = desc[0]

    if amount:
        spreadSheet.Cells(i+2, 3).Value = amount


# save excel sheet
if os.path.exists(os.path.join(pwd, "statement - "+yr+".xlsx")):
    os.remove(os.path.join(pwd, "statement - "+yr+".xlsx"))

try:
    spreadSheet.ActiveWorkbook.SaveAs(
        os.path.join(pwd, "statement - "+yr+".xlsx"))
    spreadSheet.ActiveWorkbook.Close()
    spreadSheet.Quit()
except:
    spreadSheet.Quit()
