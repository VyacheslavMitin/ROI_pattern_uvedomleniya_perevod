#  Скрипт генерации уведомлений
# pip install pyautogui - для автоматизации печати, не будем использовать
# pip install openpyxl - для работы с эксельками

import openpyxl
# переменные
path_to_xls = '+EXPORTS'
wb_pattern = openpyxl.load_workbook('pattern.xlsx')
wb_data = openpyxl.load_workbook('data.xlsx')

# кортежи
B_sotrudniki_FIO_datelniyu = []
D_sotrudniki_doljnost_datelniyu = []
E_sotrudniki_doljnost_imintelniyu = []
F_podrazdelenie_roditelnom = []
G_trud_dog_nomer = []
H_trud_dog_data = []
J_stavka_ciframi = []
K_stavka_propisyui = []

# Циклы чтения столбцов
sheet_data = wb_data.active  # активация эксельки с листом

for row in sheet_data.rows:
    string = ''
    column_b = sheet_data['B']
    for cell in column_b:
        string = str(cell.value)
        B_sotrudniki_FIO_datelniyu.append(string)

for row in sheet_data.rows:
    string = ''
    column_d = sheet_data['D']
    for cell in column_d:
        string = str(cell.value)
        D_sotrudniki_doljnost_datelniyu.append(string)

for row in sheet_data.rows:
    string = ''
    column_e = sheet_data['E']
    for cell in column_e:
        string = str(cell.value)
        E_sotrudniki_doljnost_imintelniyu.append(string)

for row in sheet_data.rows:
    string = ''
    column_f = sheet_data['F']
    for cell in column_f:
        string = str(cell.value)
        F_podrazdelenie_roditelnom.append(string)

for row in sheet_data.rows:
    string = ''
    column_g = sheet_data['G']
    for cell in column_g:
        string = str(cell.value)
        G_trud_dog_nomer.append(string)

for row in sheet_data.rows:
    string = ''
    column_h = sheet_data['H']
    for cell in column_h:
        string = str(cell.value)
        H_trud_dog_data.append(string)

for row in sheet_data.rows:
    string = ''
    column_j = sheet_data['J']
    for cell in column_j:
        string = str(cell.value)
        J_stavka_ciframi.append(string)

for row in sheet_data.rows:
    string = ''
    column_k = sheet_data['K']
    for cell in column_k:
        string = str(cell.value)
        K_stavka_propisyui.append(string)


# функции
def generate_uvedomleniya():
    sheet_pattern = wb_pattern.active
    cell_doljnost_datelnyui = sheet_pattern['H5']  # должность в дательном
    cell_podrazdeleniyu = sheet_pattern['H6']  # подразделение в родительном
    cell_fio_datelnom = sheet_pattern['H7']  # фио в дательном
    cell_truddog_nomer = sheet_pattern['F16']  # номер труд договора
    cell_truddog_data = sheet_pattern['A17']  # дата труд договора
    cell_doljnost_iminitelnyui = sheet_pattern['A19']  # должность в именительном
    cell_stavka_ciframi = sheet_pattern['F22']  # ставка суммой
    cell_stavka_propisyui = sheet_pattern['A23']  # ставка суммой
    for FIO, dolj_datet, dolj_imin, podrazdet, trud_nomer, trud_data, stavka_cifri, stavka_propis in zip(B_sotrudniki_FIO_datelniyu, D_sotrudniki_doljnost_datelniyu,
                                E_sotrudniki_doljnost_imintelniyu, F_podrazdelenie_roditelnom, G_trud_dog_nomer,
                                H_trud_dog_data, J_stavka_ciframi, K_stavka_propisyui):
        cell_doljnost_datelnyui.value = dolj_datet
        cell_podrazdeleniyu.value = podrazdet
        cell_fio_datelnom.value = FIO
        cell_truddog_nomer.value = trud_nomer
        cell_truddog_data.value = trud_data
        cell_doljnost_iminitelnyui.value = dolj_imin
        cell_stavka_ciframi.value = stavka_cifri
        cell_stavka_propisyui.value = f'({stavka_propis})'
        wb_pattern.save(f'{path_to_xls}/{FIO}.xlsx')
        if FIO == 'Ятманову Владимиру Степановичу':  # последний в списке брейкает цикл
            break


if __name__ == "__main__":
    generate_uvedomleniya()  # запуск функции

