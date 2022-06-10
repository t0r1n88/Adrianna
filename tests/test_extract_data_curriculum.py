import pytest
import pandas as pd
# для корректной проверки на тип колонок в датафрейме
from pandas.api.types import is_object_dtype, is_numeric_dtype, is_bool_dtype
import os
from extract_data_curriculum import *


class TestExtractDataCurriculum:
    """
    Клосс для проверки  корректности извлечения данных из файла учебного плана
    """
    def test_defining_boundaries_table(self,test_curriculum):
        assert defining_boundaries_table(test_curriculum) == 1