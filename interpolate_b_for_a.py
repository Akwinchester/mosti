import numpy as np

from utils import get_values_cell

a = [0, 1, 2, 4, 8, 16, 32, 100, float('inf')]
b = [31, 10.4, 7.7, 5.9, 4.7, 4, 3.6, 3.3, 3.14]

x = np.array(a)
y = np.array(b)


def interpolate(x, y, xq):
    idx = np.searchsorted(x, xq)
    x0 = x[idx - 1]
    y0 = y[idx - 1]

    x1 = x[idx]
    y1 = y[idx]

    if xq == 0.1 or xq < 0.1:
        return 31
    if xq == 100:
        return 3.3
    if xq > 100:
        return 3.14

    return y0 + (xq - x0) * (y1 - y0) / (x1 - x0)


xq = 0.1578
# пример значения для интерполяции
print(interpolate(x, y, xq))

import openpyxl
from config import file_name_excel

wb = openpyxl.load_workbook(file_name_excel)

sheet = wb['Лист1']
# Labda_1_4 = sheet['F30'].value
# R_y = sheet['D19'].value
values = get_values_cell(file_name_excel, {'alpha':'F32'}, calculate=True)
x = interpolate(x, y, values['alpha'])
print(x)
sheet['F37'] = x


# Сохранение изменений в файл
wb.save(file_name_excel)