#  Скрипт генерации уведомлений
# pip install pyautogui - для автоматизации печати
# pip install openpyxl - для работы с эксельками

import openpyxl

path_to_xls = '+EXPORTS/'

wb_pattern = openpyxl.load_workbook('pattern.xlsx')
wb_data = openpyxl.load_workbook('data.xlsx')

