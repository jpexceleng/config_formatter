import openpyxl
# RESOURCES:
# https://openpyxl.readthedocs.io/en/stable/
# https://openpyxl.readthedocs.io/en/stable/tutorial.html

""" COMMENTS
- Openpyxl doesn't support old .xls format; must be .xlsx or .xlsm.
- Make sure workbook is not currently open when calling load.workbook().
"""

# set filepath to excel file.
path = 'res\\LST-200219-SLP_CONFIG-TM01-Mix.xlsx'

# create workbook object and open workbook.
wb_obj = openpyxl.load_workbook(path)

# get workbook active sheet object from active attribute
active_sheet_obj = wb_obj.active

# cell object is created by using sheet object's cell() method;
# NOTE: cell references is base 1 (i.e. first row/column starts at 1)
cell_obj = active_sheet_obj.cell(row = 1, column = 1)

# print value of cell object
print(cell_obj.value)