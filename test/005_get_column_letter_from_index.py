import openpyxl.utils.cell

col_index = 5
col_letter = openpyxl.utils.cell.get_column_letter(col_index)

print(col_letter)