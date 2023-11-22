import openpyxl
from config import file_name_excel

wb = openpyxl.load_workbook(file_name_excel)

sheet = wb['Лист1']


d = sheet['D6'].value
d = d*2

if d <= 5:
    epsilon = 1
elif 10<= d <=25:
    epsilon = 0.85
elif d>= 50:
    epsilon = 1
elif 5<= d <= 10:
    epsilon = 1 - 0.03 * (d - 5)
elif 25 <= d <= 50:
    epsilon = 0.85 + 0.006 * (d - 25)

sheet['F47'].value = epsilon
wb.save(file_name_excel)