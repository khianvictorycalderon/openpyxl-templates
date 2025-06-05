import openpyxl

def opxl_read(file_path, sheet_name):
    try:
        wb = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        raise FileNotFoundError(f"The file '{file_path}' does not exist.")
    except Exception as e:
        raise RuntimeError(f"Failed to open the file '{file_path}': {e}")

    if sheet_name not in wb.sheetnames:
        raise ValueError(f"The sheet '{sheet_name}' does not exist in the workbook.")

    ws = wb[sheet_name]

    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        data.append(row)
    return data

# Sample Usage

db = "sample.xlsx"
sheet = "Sample_Sheet"

read_data = opxl_read(db, sheet)

for row in read_data:
    print(row)