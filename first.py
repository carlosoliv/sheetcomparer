import xlrd

loc = "/home/carlos/Documents/coding/rota.xlsx"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

print (sheet.cell_value (0,1))

for i in range (2,70):
    if sheet.cell_value (i,1) == "":
        break
    print (sheet.cell_value (i,1))
