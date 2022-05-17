import openpyxl

""" COMMENTS
- objective is to read all worksheet names from existing workbook.
-  
"""

# open existing workbook.
path = 'res\LST-200219-SLP_CONFIG-TM01-Mix.xlsx'
wb = openpyxl.load_workbook(path)

# get all worksheet names in workbook
for ws in wb:
    print(ws.title)

# set cell properties
ws = wb.worksheets('P_ANALOG_INPUT')
