from hashlib import new
import os
import PyPDF2
from win32com import client
import re

yr = "2020"
pwd = os.getcwd()
excel = client.Dispatch("Excel.Application")
pdfFolder = os.path.join(pwd, yr)

stData = ""
for pdf in os.listdir(pdfFolder):
    pdfFile = open(os.path.join(pdfFolder, pdf), "rb")
    pdfReader = PyPDF2.PdfFileReader(pdfFile)
    for pageNum in range(pdfReader.numPages):
        pageObj = pdfReader.getPage(pageNum)
        stData += pageObj.extractText()
stData = stData.replace("Customer Services: 0345 606 2174", "")
statementData = re.findall(r"\d{2}\s\w*\s\d{2}\s.*", stData)

newExcel = client.Dispatch("Excel.Application")
newExcel.Workbooks.Add()
newExcel.Visible = True
newExcel.ActiveSheet.Cells(1, 1).Value = "Date"
newExcel.ActiveSheet.Cells(1, 2).Value = "Description"
newExcel.ActiveSheet.Cells(1, 3).Value = "Amount"

# map list items values
for i in range(len(statementData)):
    row = statementData[i]
    date = re.findall(r"\d{2}\s\w*", row)[0] + " " + yr
    desc = re.findall(r"\d{2}\s\w*\s\d{2}\s\w*\s(.*(?=\s+\d+\.*))", row)[0]
    amount = re.findall(r"(\d+,\d{1,3}\.\d{1,2}|\d+\.\d{1,2})", row)[0]

    cr = re.search(r"CR$", row)
    if cr != None:
        amount = "-" + amount

    newExcel.Cells(i+2, 1).Value = date
    newExcel.Cells(i+2, 2).Value = desc
    newExcel.Cells(i+2, 3).Value = amount

# save excel sheet
if os.path.exists(os.path.join(pwd, "statement - "+yr+".xlsx")):
    os.remove(os.path.join(pwd, "statement - "+yr+".xlsx"))
newExcel.ActiveWorkbook.SaveAs(os.path.join(pwd, "statement - "+yr+".xlsx"))

