import xlwings
# from appscript import k


wb = xlwings.books.active
ws = wb.sheets.active

ws[f'B2'].font.bold = True
ws[f'B2'].font.italic = True
ws[f'B2'].api.Font.Underline = 2  # True == 2 - single, 3 - double

print(ws[f'B2'].font.bold)
print(ws[f'B2'].font.italic)
print(ws[f'B2'].api.Font.Underline)
