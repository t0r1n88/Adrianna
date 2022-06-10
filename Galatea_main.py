"""
Скрипт для отработки функционала который будет использоваться в графическом интерфейсе
"""
import pandas as pd
import openpyxl
import os



def processing_data(path_to_file,name_file):
    """
    Функция для обработки данных

    :return:
    """
    df = pd.read_excel(f'{path_to_file}{name_file}')

    df.to_excel('c:/Users/1/PycharmProjects/Galatea/output/Result.xlsx',index=False)


# path_folder_data = 'data/'
# file_data_xlsx = 'data.xlsx'
path_folder_data = 'c:/Users/1/PycharmProjects/Galatea/data/'
file_data_xlsx = 'data.xlsx'


processing_data(path_folder_data,file_data_xlsx)
