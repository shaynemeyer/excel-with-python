# Creating and Editing Spreadsheets

## Creating a Spreadsheet

```python
# creating_spreadsheet.py
from openpyxl import Workbook

def create_spreadsheet(path):
    workbook = Workbook()
    workbook.save(path)

if __name__ == '__main__':
    create_spreadsheet("hello.xlsx")
```

```python
# adding_data_rows.py

from openpyxl import Workbook


def create_workbook(path):
    workbook = Workbook()
    sheet = workbook.active
    data = [[1, 2, 3],
            ["a", "b", "c"],
            [44, 55, 66]]
    for row in data:
        sheet.append(row)
    workbook.save(path)


if __name__ == "__main__":
    create_workbook("write_rows.xlsx")
```

---

## Adding and Removing Sheets

Adding a worksheet to a workbook happens automatically when you create a new Workbook. The Worksheet will be named “Sheet” by default. If you want, you can set the name of the sheet yourself.

```python
# creating_sheet_title.py
from openpyxl import Workbook

def create_sheets(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Hello"
    sheet2 = workbook.create_sheet(title="World")
    workbook.save(path)

if __name__ == "__main__":
    create_sheets("hello_sheets.xlsx")
```

```python
# delete_sheets.py

import openpyxl

def create_worksheets(path):
    workbook = openpyxl.Workbook()
    workbook.create_sheet()
    print(workbook.sheetnames)
    # Insert a worksheet
    workbook.create_sheet(index=1, title="Second sheet")
    print(workbook.sheetnames)
    del workbook["Second sheet"]
    print(workbook.sheetnames)
    workbook.save(path)

if __name__ == "__main__":
    create_worksheets("del_sheets.xlsx")
```

---

## Inserting and Deleting Rows and Columns

```python
# insert_demo.py

from openpyxl import Workbook


def inserting_cols_rows(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Hello"
    sheet["A2"] = "from"
    sheet["A3"] = "OpenPyXL"
    # insert a column before A
    sheet.insert_cols(idx=1)
    # insert 2 rows starting on the second row
    sheet.insert_rows(idx=2, amount=2)
    workbook.save(path)


if __name__ == "__main__":
    inserting_cols_rows("inserting.xlsx")
```

```python
# delete_demo.py

from openpyxl import Workbook

def deleting_cols_rows(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Hello"
    sheet["B1"] = "from"
    sheet["C1"] = "OpenPyXL"
    sheet["A2"] = "row 2"
    sheet["A3"] = "row 3"
    sheet["A4"] = "row 4"
    # Delete column A
    sheet.delete_cols(idx=1)
    # delete 2 rows starting on the second row
    sheet.delete_rows(idx=2, amount=2)
    workbook.save(path)


if __name__ == "__main__":
    deleting_cols_rows("deleting.xlsx")
```

---

## Editing Cell Data

```python
# editing_demo.py

from openpyxl import load_workbook

def edit(path, data):
    workbook = load_workbook(filename=path)
    sheet = workbook.active
    for cell in data:
        current_value = sheet[cell].value
        sheet[cell] = data[cell]
        print(f'Changing {cell} from {current_value} to {data[cell]}')
    workbook.save(path)

if __name__ == "__main__":
    data = {"B1": "Hi", "B5": "Python"}
    edit("inserting.xlsx", data)
```

---

## Creating Merged Cells

A merged cell is where two or more cells get merged into one. To set a MergedCell’s value, you have to use the top-left-most cell. For example, if you merge “A2:E2”, you would set the value of cell “A2” for the merged cells.

```python
# merged_cells.py

from openpyxl import Workbook
from openpyxl.styles import Alignment


def create_merged_cells(path, value):
    workbook = Workbook()
    sheet = workbook.active
    sheet.merge_cells("A2:E2")
    top_left_cell = sheet["A2"]
    top_left_cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet["A2"] = value
    workbook.save(path)

if __name__ == "__main__":
    create_merged_cells("merged.xlsx", "Hello World")
```

---

## Folding Rows and Columns

Microsoft Excel supports the folding of rows and columns. The term “folding” is also called “hiding” or creating an “outline”. The rows or columns that get folded can be unfolded (or expanded) to make them visible again. You can use this functionality to make a spreadsheet briefer. For example, you might want to only show the sub-totals or the results of equations rather than all of the data at once.

```python
# folding.py

import openpyxl

def folding(path, rows=None, cols=None, hidden=True):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    if rows:
        begin_row, end_row = rows
        sheet.row_dimensions.group(begin_row, end_row, hidden=hidden)

    if cols:
        begin_col, end_col = cols
        sheet.column_dimensions.group(begin_col, end_col, hidden=hidden)

    workbook.save(path)

if __name__ == "__main__":
    folding("folded.xlsx", rows=(1, 5), cols=("C", "F"))
```
