import numpy as np

y_keys = [0.01, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]

ae1_table = {
    0: [1.243, 1.248, 1.253, 1.258, 1.264, 1.269, 1.274, 1.279, 1.283, 1.267, 1.243],
    0.1: [1.187, 1.191, 1.195, 1.199, 1.202, 1.206, 1.209, 1.212, 1.214, 1.16, None],
    0.2: [1.152, 1.155, 1.158, 1.162, 1.165, 1.168, 1.17, 1.172, 1.15, None, None],
    0.3: [1.128, 1.131, 1.133, 1.136, 1.139, 1.142, 1.144, 1.145, 1.097, None, None],
    0.4: [1.11, 1.113, 1.115, 1.118, 1.12, 1.123, 1.125, 1.126, 1.069, None, None],
    0.5: [1.097, 1.099, 1.102, 1.104, 1.106, 1.109, 1.11, 1.106, 1.061, None, None],
    0.6: [1.087, 1.089, 1.091, 1.093, 1.095, 1.097, 1.099, 1.079, None, None, None],
    0.7: [1.078, 1.08, 1.082, 1.084, 1.086, 1.088, 1.09, 1.055, None, None, None],
    0.8: [1.071, 1.073, 1.075, 1.077, 1.079, 1.081, 1.082, 1.044, None, None, None],
    0.9: [1.065, 1.067, 1.069, 1.071, 1.073, 1.074, 1.076, 1.036, None, None, None],
    1.0: [1.06, 1.062, 1.064, 1.066, 1.067, 1.069, 1.071, 1.031, None, None, None],
    2.0: [1.035, 1.036, 1.037, 1.038, 1.039, 1.04, 1.019, None, None, None, None],
    3.0: [1.024, 1.025, 1.026, 1.027, 1.028, 1.029, 1.017, None, None, None, None],
    4.0: [1.019, 1.019, 1.02, 1.021, 1.021, 1.022, 1.015, None, None, None, None],
    5.0: [1.015, 1.015, 1.016, 1.017, 1.018, 1.018, None, None, None, None, None]
}


def interpolate(x, x0, x1, y0, y1):
    return y0 + (x - x0) * (y1 - y0) / (x1 - x0)


def ae1_interpolation(x, y):
    x_keys = list(ae1_table.keys())

    x1 = max([k for k in x_keys if k <= x])
    x2 = min([k for k in x_keys if k >= x])

    y1 = max([k for k in y_keys if k <= y])
    y2 = min([k for k in y_keys if k >= y])

    f_x1_y1 = ae1_table[x1][y_keys.index(y1)]
    f_x2_y1 = ae1_table[x2][y_keys.index(y1)]
    f_x1_y2 = ae1_table[x1][y_keys.index(y2)]
    f_x2_y2 = ae1_table[x2][y_keys.index(y2)]


    f_x_y1 = interpolate(x, x1, x2, f_x1_y1, f_x2_y1)
    f_x_y2 = interpolate(x, x1, x2, f_x1_y2, f_x2_y2)

    return interpolate(y, y1, y2, f_x_y1, f_x_y2)




import openpyxl
from config import file_name_excel
from utils import get_values_cell

wb = openpyxl.load_workbook(file_name_excel)
values = get_values_cell(file_name_excel, {'x':'F22', 'y':'F23'}, calculate=True)
print(values)
sheet = wb['Лист1']


# x = sheet['F22'].calculated_value
# y = sheet['F23'].calculated_value

# d_13 = sheet['D13'].value
# d_12 = sheet['D12'].value
# d_14 = sheet['D14'].value
# d_10 = sheet['D10'].value
# f_18 = d_10-2*d_12
#
# f16_value = d_12*d_13
# print(f_18,d_14)
# f17_value = f_18*d_14
# # print(f16_value, f17_value)
# # print(d_12, d_13, f_18, d_14)
# x = f16_value / f17_value
# f_19 = f17_value+2*f16_value
#
#
# y = (f16_value + f17_value)/f_19
#
# print(x,y)
# print(ae1_interpolation(x, y))
# Запись значения в ячейку A1

sheet['F21'] = ae1_interpolation(values['x'], values['y'])


# Сохранение изменений в файл
wb.save(file_name_excel)