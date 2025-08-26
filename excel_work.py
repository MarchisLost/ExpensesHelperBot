from openpyxl import load_workbook

import requests
from io import BytesIO
from openpyxl import load_workbook

def read_excel_from_drive(file_id: str, sheet: str, cell: str):
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    response = requests.get(url)
    response.raise_for_status()

    file = BytesIO(response.content)
    wb = load_workbook(file)
    ws = wb[sheet]

    value = ws[cell].value
    wb.close()
    return value

# Usage: Pass the Google Drive file ID (the long string in the link)
val = read_excel_from_drive("YOUR_FILE_ID", "Sheet1", "A1")
print("Value in A1:", val)

def read_cells(sheet_name: str, cell_x: str, cell_y: str):
    wb = load_workbook(FILE_PATH)
    sheet = wb[sheet_name]

    value_x = sheet[cell_x].value
    value_y = sheet[cell_y].value

    wb.close()
    return value_x, value_y

def write_cell(sheet_name: str, cell: str, value):
    wb = load_workbook(FILE_PATH)
    sheet = wb[sheet_name]

    sheet[cell] = value  # write value into the cell

    wb.save(FILE_PATH)
    wb.close()
