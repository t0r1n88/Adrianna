import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import os
import time

# Настройки для отображения колонок в пандас
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', None)


def find_coordinate(sheet_d,column: str, row: int, text) :
    """
    Функция для поиска координат ячейки с заданным текстом
    param sheet_d: объект листа где нужно искать
    param column : название колонки
    param row : номер строки
    param text : значение ячейки координаты которой нужно найти.
    """


def defining_boundaries_table(wb_d: openpyxl.Workbook):
    """
    Функция для определения границ извлекаемых данных
    Где левое верхнее поле это поле в первом столбце со значением 1
    А правое нижнее пересечение строки ФК.00 и колонки Вар.часть.
    """
    # Переключаемся на лист План
    sheet = wb_d['План']
    top_left = None
    # Ищем в колонке 2(B) значение 1
    for cell in sheet['B']:
        if cell.value == '1':
            top_left = cell.coordinate
        else:
            continue
    print(top_left)


# Подготавливаем данные
name_curriculum = 'data/УП ПМ 2021.xlsx'
path_to_end_folder = 'output'
t = time.localtime()
current_time = time.strftime('%H_%M_%S', t)

# Открываем документ
wb = openpyxl.load_workbook(name_curriculum, data_only=True)

defining_boundaries_table(wb)
