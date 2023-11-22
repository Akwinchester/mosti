import numpy as np

from utils import get_values_cell

y_keys = [200, 240, 280, 320, 360, 400, 440, 480, 520, 560, 600, 640]

table = {
    10: [988, 987, 985, 984, 983, 982, 981, 980, 979, 978, 977, 977],
    20: [967, 962, 959, 955, 952, 949, 946, 943, 941, 938, 936, 934],
    30: [939, 931, 924, 917, 911, 905, 900, 895, 891, 887, 883, 879],
    40: [906, 894, 883, 873, 863, 854, 846, 849, 832, 825, 820, 814],
    50: [869, 852, 836, 822, 809, 796, 785, 775, 764, 746, 729, 712],
    60: [827, 805, 785, 766, 749, 721, 696, 672, 650, 628, 608, 588],
    70: [782, 754, 724, 687, 654, 623, 595, 568, 542, 518, 494, 470],
    80: [734, 686, 641, 602, 566, 532, 501, 471, 442, 414, 386, 359],
    90: [665, 612, 565, 522, 483, 447, 413, 380, 349, 326, 305, 287],
    100: [599, 542, 493, 448, 408, 369, 335, 309, 286, 267, 250, 235],
    110: [537, 478, 427, 381, 338, 306, 280, 258, 239, 223, 209, 197],
    120: [479, 419, 366, 321, 287, 260, 237, 219, 203, 190, 178, 167],
    130: [425, 364, 313, 276, 247, 223, 204, 189, 175, 163, 153, 145],
    140: [376, 315, 272, 240, 215, 195, 178, 164, 153, 143, 134, 126],
    150: [328, 276, 239, 211, 189, 171, 157, 145, 134, 126, 118, 111]
}


def interpolate(x, x0, x1, y0, y1):
    return y0 + (x - x0) * (y1 - y0) / (x1 - x0)


def intrp_f(x, y):
    x_keys = list(table.keys())
    print(x_keys)

    x1 = max([k for k in x_keys if k <= x])
    x2 = min([k for k in x_keys if k >= x])

    y1 = max([k for k in y_keys if k <= y])
    y2 = min([k for k in y_keys if k >= y])
    print(x1, x2)
    print(y1, y2)

    f_x1_y1 = table[x1][y_keys.index(y1)]
    f_x2_y1 = table[x2][y_keys.index(y1)]
    f_x1_y2 = table[x1][y_keys.index(y2)]
    f_x2_y2 = table[x2][y_keys.index(y2)]


    f_x_y1 = interpolate(x, x1, x2, f_x1_y1, f_x2_y1)
    f_x_y2 = interpolate(x, x1, x2, f_x1_y2, f_x2_y2)

    return interpolate(y, y1, y2, f_x_y1, f_x_y2)

# print(ae1_interpolation(25.641, 215))
#94.648

import openpyxl
from config import file_name_excel

wb = openpyxl.load_workbook(file_name_excel)

sheet = wb['Лист1']


# R_y = sheet['D18'].value
# print(f'R_y = {R_y}')
#
# B10 = sheet['B10'].value
# print(f'B10 = {B10}')
#
# F13_f = sheet['F13'].value
# print(f'F13_f = {F13_f}')
#
# F31_f = sheet['F31'].value
# print(f'F31_f = {F31_f}')
#
# F14_f = sheet['F14'].value
# print(f'F14_f = {F14_f}')
#
# F15_f = sheet['F15'].value
# print(f'F15_f = {F15_f}')
#
# B8 = sheet['B8'].value
# print(f'B8 = {B8}')
#
# B11 = sheet['B11'].value
# print(f'B11 = {B11}')
#
# D10 = sheet['D10'].value
# print(f'D10 = {D10}')
#
# D12 = sheet['D12'].value
# print(f'D12 = {D12}')
#
# D13 = sheet['D13'].value
# print(f'D13 = {D13}')
#
# D14 = sheet['D14'].value
# print(f'D14 = {D14}')
#
# D15 = sheet['D15'].value
# print(f'D15 = {D15}')
#
# F18_f = sheet['F18'].value
# print(f'F18_f = {F18_f}')
#
# F15 = D10/2
# print(f'F15 = {F15}')
#
#
# F37 = sheet['F37'].value
# print(f'F37 = {F37}')
#
# F34_f = sheet['F34'].value
# print(f'F34_f = {F34_f}')
#
# F35_f = sheet['F35'].value
# print(f'F35_f = {F35_f}')
#
# F36_f = sheet['F36'].value
# print(f'F36_f = {F36_f}')
#
# B11 = sheet['B11'].value
# print(f'B11 = {B11}')
#
#
# F18 = D10-2*D12
#
# F14 = (2*(D13-2*D15)/12)*D12**3+2*(D13-2*D15)*D12*((F18+D12)/2)**2+B8*((D14*F18**3)/12)
# print(f'F14 = {F14}')
# F13 = F14/F15
# print(f'F13 = {F13}')
#
# D6 = sheet['D6'].value
# print(f'D6 = {D6}')
#
# F34 =D6/4
# F35 = ((D12*D13**3)/6) + (D14**3*F18)/12
# F36 = (F18*D14**3+2*(D13*D12**3))/3
#
# F31 = F37/F34 * (B10*F35*B11*F36)**0.5
# print(f'F31 = {F31}')
# Labda_1_4 = 3.1415926535 *(B10*F13/F31)**0.5
#
# print(R_y, Labda_1_4)
#
# F30 = sheet['F30'].value

values = get_values_cell(file_name_excel, {'Labda_1_4':'F30', 'R_y':'D19'}, calculate=True)

sheet['F39'] = intrp_f(values['Labda_1_4'], values['R_y'])

# Labda_1_4 = sheet['F30'].value
# R_y = sheet['D18'].value
# sheet['F39'] = intrp_f(Labda_1_4, R_y)


# Сохранение изменений в файл
wb.save(file_name_excel)