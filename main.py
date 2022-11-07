from hashlib import new
import os
import PyPDF2
from win32com import client
import re

yr = "2021"
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


# 17 DECEMBER 20 DECEMBER WWW.JUST-EAT.CO.UK 02067 362061 12.93
# print(data)
# print("date: ", str(date), "desc: ", str(desc), "amount: ", str(amount))
# sheets = excel.Workbooks.Open(os.path.join(pwd, "jan state.xlsx"))
# def findRowThatContainsString(text, string):
#     for row in text.splitlines():
#         if string in row:
#           print(row[row.find(string):])
# def pdfsContainString(pdf_path, string):
#     pdfFileObj = open(pdf_path, 'rb')
#     pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
#     for i in range(0, pdfReader.numPages):
#         pageObj = pdfReader.getPage(i)
#         if string in pageObj.extractText():
#             return pageObj.extractText()
#     return False
# def check_folder(folder_path, string):
#     for file in os.listdir(folder_path):
#         if check_pdf(os.path.join(folder_path, file), string):
#             return True
#         return False

# # get the first sheet
# sheet = sheets.Worksheets(1)
# # loop through the rows
# for row in range(2, sheet.UsedRange.Rows.Count + 1):
#     # get the value of the cell
#     date=sheet.Cells(row, 1).Value
#     # convert date to string
#     description =sheet.Cells(row, 2).Value
#     amount = sheet.Cells(row, 3).Value
#     string = pdfsContainString(os.path.join(pwd, "pdf"), description)
#     # row = findRowThatContainsString(string, description)
#     # if not string:
#     #     print ("note found - date: " + str(date) + "; description: " + str(description) + "; amount: " + str(amount))
