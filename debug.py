from utils import get_values_cell
from config import file_name_excel

print(get_values_cell(file_name_excel, {'x':'F22', 'y':'F23'}, calculate=True))