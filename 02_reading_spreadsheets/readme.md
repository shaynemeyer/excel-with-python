# Reading Spreadsheets

## Open a Spreadsheet

```python
# open_workbook.py
from openpyxl import load_workbook

def open_workbook(path):
    workbook = load_workbook(filename=path)
    print(f"Worksheet names: {workbook.sheetnames}")
    sheet = workbook.active
    print(sheet)
    print(f"The title of the Worksheet is: {sheet.title}")

if __name__ == "__main__":
    open_workbook("books.xlsx")
```

---

## Read Specific Cells

```python
# reading_specific_cells.py

from openpyxl import load_workbook

def get_cell_info(path):
    workbook = load_workbook(filename=path)
    sheet = workbook.active
    print(sheet)
    print(f"The title of the worksheet is: {sheet.title}")
    print(f'The value of A2 is {sheet["A2"].value}')
    print(f'The value of A3 is {sheet["A3"].value}')
    cell = sheet['B3']
    print(f'The variable "cell" is {cell.value}')

if __name__ == '__main__':
    get_cell_info('books.xlsx')
```

---

## Read Cells From Specific Row

```python
# reading_row_cells.py

from openpyxl import load_workbook

def iterating_row(path, sheet_name, row):
    workbook = load_workbook(filename=path)
    if sheet_name not in workbook.sheetnames:
        print(f'{sheet_name} not found. Quitting.')
        return
    sheet = workbook[sheet_name]
    for cell in sheet[row]:
        print(f'{cell.column_letter}{cell.row} = {cell.value}')

if __name__ == '__main__':
    iterating_row("books.xlsx", sheet_name="Sheet 1 - Books", row=2)
```

---

## Read Cells From Specific Column

```python
# reading_column_cells.py

from openpyxl import load_workbook

def iterating_column(path, sheet_name, col):
    workbook = load_workbook(filename=path)
    if sheet_name not in workbook.sheetnames:
        print(f"'{sheet_name} not found. Quitting.")
        return
    sheet = workbook[sheet_name]
    for cell in sheet[col]:
        print(f"{cell.column_letter}{cell.row} = {cell.value}")

if __name__ == '__main__':
    iterating_column("books.xlsx", sheet_name="Sheet 1 - Books", col="A")
```

## Read Cells from Multiple Rows or Columns

```python
# iterating_over_cells_in_rows.py

from openpyxl import load_workbook

def iterating_over_values(path, sheet_name):
    workbook = load_workbook(filename=path)
    if sheet_name not in workbook.sheetnames:
        print(f"'{sheet_name}' not found. Quitting.")
        return
    sheet = workbook[sheet_name]
    for value in sheet.iter_rows(min_row=1, max_row=3, min_col=1, max_col=3,values_only=True):
        print(value)

if __name__ == '__main__':
    iterating_over_values("books.xlsx", sheet_name="Sheet 1 - Books")
```

---

## Read Cells from a Range

Excel lets you specify a range of cells using the following format: (col)(row):(col)(row). In other words, you can say that you want to start in column A, row 1, using A1. If you wanted to specify a range, you would use something like this: A1:B6. That tells Excel that you are selecting the cells starting at A1 and going to B6.

```python
# read_cells_from_range.py

import openpyxl
from openpyxl import load_workbook

def iterating_over_values(path, sheet_name, cell_range):
    workbook = load_workbook(filename=path)
    if sheet_name not in workbook.sheetnames:
        print(f"'{sheet_name} not found. Quitting.'")
        return

    sheet = workbook[sheet_name]
    for column in sheet[cell_range]:
        for cell in column:
            if isinstance(cell, openpyxl.cell.cell.MergedCell):
                # skip this cell
                continue
            print(f"{cell.column_letter}{cell.row} = {cell.value}")

if __name__ == '__main__':
    iterating_over_values("books.xlsx", sheet_name="Sheet 1 - Books", cell_range="A1:B6")
```

---

## Read All Cells in All Sheets

Microsoft Excel isn’t as simple to read as a CSV file, or a regular text file. That is because Excel needs to store each cell’s data, which includes its location, formatting, and value, and that value could be a number, a date, an image, a link, etc. Consequently, reading an Excel file is a lot more work! openpyxl does all that hard work for us, though.

The natural way to iterate through an Excel file is to read the sheets from left to right, and within each sheet, you would read it row by row, from top to bottom.

```python
# read_all_data.py

import openpyxl
from openpyxl import load_workbook

def read_all_data(path):
    workbook = load_workbook(filename=path)
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        print(f"Title = {sheet.title}")
        for row in sheet.rows:
            for cell in row:
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    # Skip this cell
                    continue

if  __name__ == '__main__':
    read_all_data("books.xlsx")
```

```python
# read_all_data_values.py

import openpyxl
from openpyxl import load_workbook

def read_all_data(path):
    workbook = load_workbook(filename=path)
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        print(f"Title = {sheet.title}")
        for value in sheet.iter_rows(values_only=True):
            print(value)


if __name__ == "__main__":
    read_all_data("books.xlsx")
```
