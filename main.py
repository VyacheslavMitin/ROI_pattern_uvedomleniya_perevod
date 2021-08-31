#  Скрипт генерации уведомлений
# pip install pyautogui - для автоматизации печати, не будем использовать
# pip install openpyxl - для работы с эксельками

import openpyxl

path_to_xls = '+EXPORTS/'

wb_pattern = openpyxl.load_workbook('pattern.xlsx')
wb_data = openpyxl.load_workbook('data.xlsx')

B_sotrudniki_FIO_datelniyu = []
D_sotrudniki_doljnost_datelniyu = []
E_sotrudniki_doljnost_imintelniyu = []
F_podrazdelenie_roditelnom = []
G_trud_dog_nomer = []
H_trud_dog_data = []
J_stavka_ciframi = []
K_stavka_propisyui = []
sheet_data = wb_data.active


# Циклы чтения столбцов
for row in sheet_data.rows:
    string = ''
    column_a = sheet_data['B']
    for cell in column_a:
        string = str(cell.value)
        B_sotrudniki_FIO_datelniyu.append(string)

for row in sheet_data.rows:
    string = ''
    column_a = sheet_data['B']
    for cell in column_a:
        string = str(cell.value)
        B_sotrudniki_FIO_datelniyu.append(string)