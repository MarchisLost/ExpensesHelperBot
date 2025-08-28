#import requests
#from io import BytesIO
from openpyxl import load_workbook
from dotenv import load_dotenv

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

def read_cells(file_path:str, sheet_name: str, cell_x: str, cell_y: str):
    wb = load_workbook(file_path)
    sheet = wb[sheet_name]
    
    # Get the values to calculate from the correct months
    value_x = sheet[cell_x].value
    value_y = sheet[cell_y].value

    # Calculate all the expenses

    # Call the write function to write the final value

    wb.close()
    return value_x, value_y

def write_cell(file_path:str, sheet_name: str, cell: str, value):
    # Get the value to write and write
    wb = load_workbook("file_path")
    sheet = wb["]

    sheet[cell] = value  # write value into the cell

    wb.save(FILE_PATH)
    wb.close()

def main():
    load_dotenv()
    file_path = os.getenv("FILE_PATH")
    sheet_name = "Despesas David 25"

    # Get the values from first person

    # Get the values from the second person

    # Calculate who spent more and how much owes the other person

    # Write how much it owes and who owes who 
    v_x, v_y = read_cells(file_path, sheet_name, "D16", "E16")	
    print(v_x)
    print(v_y)

if __name__ == "__main__":
    main()
