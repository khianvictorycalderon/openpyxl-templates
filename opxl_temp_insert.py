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
    
    # Determine if data is a single row or multiple rows
    # If the first element is a list or tuple, treat data as multiple rows
    if isinstance(data, (list, tuple)) and len(data) > 0 and isinstance(data[0], (list, tuple)):
        for row in data:
            sheet.append(row)
    else:
        sheet.append(data)
    
    # Save workbook
    workbook.save(file_path)
    print(f"Data inserted successfully into '{sheet_name}'.")
    

# Sample usage:

db = "sample.xlsx"
sheet = "Sample_Sheet"

data = ["John", "Doe", "Male", 17, "Brooklyn Street"]

data2 = [
    ["Maria", "Currey", "Female", 18, "Brooklyn Street"],
    ["Jane", "Doe", "Female", 20, "Brooklyn Street"],
    ["Michael", "Smith", "Male", 25, "5th Avenue"],
    ["Emily", "Johnson", "Female", 22, "Main Street"],
    ["David", "Williams", "Male", 30, "Oak Lane"],
    ["Sarah", "Brown", "Female", 27, "Pine Road"],
    ["James", "Jones", "Male", 35, "Maple Drive"],
    ["Linda", "Garcia", "Female", 28, "Cedar Street"],
    ["Robert", "Martinez", "Male", 40, "Elm Avenue"],
    ["Patricia", "Rodriguez", "Female", 26, "Birch Boulevard"],
    ["Charles", "Wilson", "Male", 33, "Chestnut Street"],
    ["Barbara", "Lee", "Female", 24, "Spruce Court"],
    ["Joseph", "Walker", "Male", 29, "Willow Way"],
    ["Susan", "Hall", "Female", 31, "Poplar Street"],
    ["Thomas", "Allen", "Male", 38, "Sycamore Road"],
    ["Jessica", "Young", "Female", 21, "Magnolia Lane"],
    ["Daniel", "Hernandez", "Male", 36, "Aspen Drive"],
    ["Karen", "King", "Female", 23, "Fir Street"],
    ["Matthew", "Wright", "Male", 34, "Hawthorn Avenue"],
    ["Nancy", "Lopez", "Female", 19, "Dogwood Circle"]
]

opxl_insert(db, sheet, data)
opxl_insert(db, sheet, data2)