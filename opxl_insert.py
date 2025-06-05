import os
import openpyxl

def opxl_insert(file_path, sheet_name, data):
    # Check if the file exists
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Excel file '{file_path}' does not exist.")
    
    # Load workbook
    workbook = openpyxl.load_workbook(file_path)
    
    # Check if the sheet exists
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' does not exist in the workbook.")
    
    # Select sheet
    sheet = workbook[sheet_name]
    
    # Append data as a new row
    sheet.append(data)
    
    # Save workbook
    workbook.save(file_path)
    print(f"Data inserted successfully into '{sheet_name}'.")

# Sample usage:

db = "sample.xlsx"
sheet = "Sample_Sheet"
data = ["John", "Doe", "Male", 17, "Brooklyn Street"]

opxl_insert(db, sheet, data)
