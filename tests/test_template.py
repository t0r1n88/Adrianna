import pytest
import pandas as pd
# для корректной проверки на тип колонок в датафрейме
from pandas.api.types import is_object_dtype, is_numeric_dtype, is_bool_dtype
import os
from Galatea_main import processing_data

PATH_FOLDER_DATA = 'c:/Users/1/PycharmProjects/Galatea/output/'

FILE_DATA_XLSX = 'Result.xlsx'
FILE_DATA_INCORRECT = 'incorrect_xlsx.xlsx'

@pytest.mark.skip(reason='Примеры')
class TestCorrectFIle:
    """
    Класс для проверки корректности итогового файла(правильные колонки, размер и пр.)
    """

    def test_shape_df(self, test_df):
        """
        Дано: датафрейм, результат работы функции processing_data
        Когда: датафрейм перед сохранением в файл
        Тогда: количество колонок в файле должно быть 4
        """
        assert test_df.shape[1] == 4, "Количество колонок в датафрейме должно быть равным 4"

    def test_name_columns(self, test_df):
        """
        Дано: датафрейм, результат работы функции processing_data
        Когда: датафрейм перед сохранением в файл
        Тогда: Названия колонок должны совпадать с образцом
        """
        # не забывай что test_df.columns это индекс и для сравнения со списком его нужно превратить в список
        assert list(test_df.columns) == ['ФИО', 'Серия паспорта', 'Номер паспорта', 'Код подразделения']

    def test_type_columns(self, test_df):
        """
        Дано: датафрейм, результат работы функции processing_data
        Когда: датафрейм перед сохранением в файл
        Тогда :Типы колонок должны совпадать с заданными
        """

        assert is_object_dtype(test_df['ФИО'])
        assert is_numeric_dtype(test_df['Серия паспорта'])
        assert not is_bool_dtype(test_df['Номер паспорта'])


@pytest.mark.skip(reason='Примеры')
class TestCorrectCreateFile:
    """
    Класс для тестирования корректности созданного в результате работы функции proccessing_data файла Excel
    """

    def test_create_output_file(self):
        """
        Дано:Путь к созданному файлу и название созданного файла
        Когда: Функция processing_data завершает свою работу
        Тогда: По указанному путь должен быть создан файл с указанным названием

        """
        processing_data(PATH_FOLDER_DATA, FILE_DATA_XLSX)
        assert os.path.isfile(f'{PATH_FOLDER_DATA}{FILE_DATA_XLSX}')

    def test_opened_output_xlsx_file(self):
        """
        Дано:Путь к созданному файлу и название созданного файла
        Когда: Функция processing_data завершает свою работу
        Тогда: Файл который создан после работы функции processing_data должен корректно открываться для обработки pandas.
        Для проверки используется метод isinstance
        """
        # test_df = pd.read_excel(f'{PATH_FOLDER_DATA}{FILE_DATA_INCORRECT}')
        test_df = pd.read_excel(f'{PATH_FOLDER_DATA}{FILE_DATA_XLSX}')
        assert isinstance(test_df, pd.core.frame.DataFrame), "Проверяем можно ли сгенерированный файл в пандас"
