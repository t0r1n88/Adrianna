import pandas as pd
import openpyxl
import os
import time
# Настройки для отображения колонок в пандас
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', None)

def defining_boundaries_table(file):
    """
    Функция для определения границ извлекаемых данных
    Где левое верхнее поле это поле в первом столбце со значением 1
    А правое нижнее пересечение строки ФК.00 и колонки Вар.часть.
    """
    return 1



# Подготавливаем данные
name_curriculum = 'data/УП ПМ 2021.xlsx'
path_to_end_folder = 'output'
t = time.localtime()
current_time = time.strftime('%H_%M_%S', t)

# Открываем документ

defining_boundaries_table(name_curriculum)
