import openpyxl

# читаем excel-файл
wb = openpyxl.load_workbook('Сетевые_Порох3_вода.xlsx')

# печатаем список листов
sheets = wb.sheetnames
for sheet in sheets:
    print(sheet)

# получаем активный лист
sheet = wb.active

# печатаем значение ячейки A1
print(sheet['A1'].value)
# печатаем значение ячейки B1
print(sheet['B1'].value)