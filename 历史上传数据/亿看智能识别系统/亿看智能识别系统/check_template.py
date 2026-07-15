import openpyxl
import os

try:
    if os.path.exists("Template.xlsx"):
        wb = openpyxl.load_workbook("Template.xlsx")
        ws = wb.active
        headers = [str(cell.value).strip() for cell in ws[1] if cell.value]
        print(f"Template.xlsx headers: {headers}")
    else:
        print("Template.xlsx not found.")
except Exception as e:
    print(f"Error reading Template.xlsx: {e}")