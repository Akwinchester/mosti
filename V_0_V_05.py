import numpy as np
from  config import file_name_excel

table = [[1, 49.03, 49.03], [1.5, 39.15, 34.25], [2, 30.55, 26.73],
         [3, 24.16, 21.14], [4, 21.69, 18.99], [5, 20.37, 17.82],
         [6, 19.5, 17.06], [7, 18.84, 16.48], [8, 18.32, 16.02],
         [9, 17.87, 15.63], [10, 17.47, 15.28], [12, 16.78, 14.68],
         [14, 16.19, 14.16], [16, 15.66, 13.71], [18, 15.19, 13.3],
         [20, 14.76, 12.92], [25, 13.85, 12.12], [30, 13.1, 11.46],
         [35, 12.5, 10.94]]

x = np.array([row[0] for row in table])
y1 = np.array([row[1] for row in table])
y2 = np.array([row[2] for row in table])


def interpolate(x_data, y1_data, y2_data, x, variant='0'):
    if variant not in ['0', '0.5']:
        raise ValueError('Variant must be 0 or 0.5')

    if variant == '0':
        y_data = y1_data
    else:
        y_data = y2_data

    idx = np.argsort(x_data)
    x_data = x_data[idx]
    y_data = y_data[idx]

    if x < x_data[0] or x > x_data[-1]:
        return None

    i = np.searchsorted(x_data, x)
    x0 = x_data[i - 1]
    y0 = y_data[i - 1]

    x1 = x_data[i]
    y1 = y_data[i]

    return y0 + (x - x0) * (y1 - y0) / (x1 - x0)


import openpyxl

wb = openpyxl.load_workbook(file_name_excel)

sheet = wb['Лист1']


# Запись значения в ячейку A1
sheet['F11'] = interpolate(x, y1, y2, sheet['D6'].value)

sheet['F12'] = interpolate(x, y1, y2, sheet['D6'].value, variant='0.5')

# Сохранение изменений в файл
wb.save(file_name_excel)