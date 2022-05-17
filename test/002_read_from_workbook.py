import openpyxl

""" COMMENTS
- objective is to read all worksheet names from existing workbook.
-  
"""

# open existing workbook
filename = 'LST-200219-SLP_CONFIG-TM01-Mix'
file_ext = '.xlsx'
filepath = 'res\\' + filename + file_ext
wb = openpyxl.load_workbook(filepath)

# # print worksheet names in workbook
# for ws in wb:
#     print(ws.title)

# select worksheet by name
ws = wb['P_ANALOG_INPUT']

try:
    # read from cell in worksheet
    c = ws['B10']
    #c = ws.cell(row=10, column=2)
    print(c.value)

    # write to cell in worksheet
    ws['B8'] = 'TEST'
    #ws.cell(row=8, column=2, value=4)
except:
    print('Error')

# save changes to workbook
wb.save(filename + '_edited' + file_ext)