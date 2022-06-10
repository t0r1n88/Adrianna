import pytest
import pandas as pd
import openpyxl
"""
Файл для фикстур
"""
PATH_FOLDER_DATA = 'c:/Users/1/PycharmProjects/Adrianna/data'
NAME_DATA_XLSX = 'data.xlsx'
NAME_CURRICULUM = 'УП ПМ 2021.xlsx'



@pytest.fixture
def test_df():
    """
    Фикстура для создания датафрейма из файла
    """
    df = pd.read_excel(f'{PATH_FOLDER_DATA}{NAME_DATA_XLSX}')
    return df

@pytest.fixture
def test_curriculum():
    """
    Фикстура для считывания пробного учебного плана
    """
    test_curr = openpyxl.load_workbook(f'{PATH_FOLDER_DATA}/{NAME_CURRICULUM}',read_only=True,data_only=True)
    return test_curr


