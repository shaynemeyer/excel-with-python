# adding_data.py 
from openpyxl import Workbook

def create_spreadsheet(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Hello"
    sheet["A2"] = "from"
    sheet["A3"] = "OpenPyXL"
    workbook.save(path)

if __name__ == '__main__':
    create_spreadsheet("hello.xlsx")