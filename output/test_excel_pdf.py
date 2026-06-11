import sys, pythoncom, win32com.client
from pathlib import Path

xlsx_path = Path(r"C:\Users\jumsu\Desktop\견적서\1\price-compare-quote-main\output\tmp_pdf_work\quote.xlsx")
pdf_path = Path(r"C:\Users\jumsu\Desktop\견적서\1\price-compare-quote-main\output\tmp_pdf_work\quote_test.pdf")

print("xlsx exists:", xlsx_path.exists())
pythoncom.CoInitialize()
excel = None
workbook = None
try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    print("Excel dispatched")
    workbook = excel.Workbooks.Open(str(xlsx_path))
    print("Workbook opened")
    for worksheet in workbook.Worksheets:
        worksheet.PageSetup.Zoom = False
        worksheet.PageSetup.FitToPagesWide = 1
        worksheet.PageSetup.FitToPagesTall = 1
    workbook.ExportAsFixedFormat(0, str(pdf_path))
    print("PDF saved, exists:", pdf_path.exists())
finally:
    if workbook is not None:
        workbook.Close(False)
    if excel is not None:
        excel.Quit()
    pythoncom.CoUninitialize()
