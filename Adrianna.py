import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from itertools import islice
import time
import os





def change_name_discipline(cell):
    # Функция для установки заглавной буквы для содержимого ячейки
    # очищаем от пробельных символов спереди и сзади
    value = cell.strip()
    # Делаем заглавным первый символ
    out_value = f'{value[0].upper()}{value[1:]}'

    return out_value


path_to_end_folder = 'data/'
path_data_folder = 'data/tarif_data/'

# Создаем общий датафрейм
col_out_df = ['Ф.И.О. преподавателя (полностью)', 'Занимаемая в ПОО должность', 'Квалификационная категория',
              'Учебная дисциплина',
              'Курс', 'Группа', 'Теория', 'ЛПЗ', 'Учебная практика', 'Производственная практика',
              'Преддипломная практика', 'Руководство над курсовым проектом'
    , 'Консультации', 'Контроль (экзамены, зачеты и т.д.)', 'Руководство ВКР', 'ИТОГО', 'Итого по тарификации']
out_df = pd.DataFrame(columns=col_out_df)

# Обрабатываем файл, пропускаем все не xlsx файлы и временные файлы
for file in os.listdir(path_data_folder):
    if not file.startswith('~') and file.endswith('.xlsx'):
        # Создаем датафрейм для соединения результата обработки таблицы и строки с суммой
        finish_df = pd.DataFrame()
        df = pd.read_excel(f'{path_data_folder}{file}', skiprows=6)
        df.columns = ['№/пп', 'Наименование группы', 'Наименование дисциплины', '1 семестр кол 1', '1 семестр кол 2',
                      '2 семестр кол 1', '2 семестр кол 2', 'Всего часов',
                      'Теория', 'ЛПЗ', 'Учебная практика', 'Производственная практика', 'Преддипломная практика',
                      'Руководство над курсовым проектом', 'Консультации', 'Контроль (экзамены, зачеты и т.д.)',
                      'Руководство ВКР', 'ИТОГО', 'Преподаватель', 'Промежуточные суммы']

        df = df[df['Наименование группы'].notna()]

        # Очищаем от пробелов перед и после слов
        df['Наименование группы'] = df['Наименование группы'].apply(lambda x: x.strip())
        df['Наименование дисциплины'] = df['Наименование дисциплины'].apply(
            lambda x: change_name_discipline(x) if type(x) == str else x)

        df = df[df['Наименование группы'] != 'ознакомлен']

        # Удаляем лишний столбец
        df.drop(columns=['Промежуточные суммы'], inplace=True)
        # Создаем книгу для того чтобы отобрать все строки до внебюджета
        wb = openpyxl.Workbook()
        ws = wb.active

        for r in dataframe_to_rows(df, index=True, header=True):
            if 'внебюджет' in r:
                break
            ws.append(r)

        # Загружаем обратно очищенный от внебюджетных дисциплин датафрейм
        data = ws.values
        cols = next(data)[1:]
        data = list(data)
        idx = [r[0] for r in data]
        data = (islice(r, 1, None) for r in data)
        clear_df = pd.DataFrame(data, index=idx, columns=cols)

        clear_df.sort_values(by='Наименование дисциплины', inplace=True)

        # Копируем данные  в датафрейм
        finish_df['Ф.И.О. преподавателя (полностью)'] = clear_df['Преподаватель']
        finish_df['Занимаемая в ПОО должность'] = ''
        finish_df['Квалификационная категория'] = ''
        finish_df['Учебная дисциплина'] = clear_df['Наименование дисциплины']
        finish_df['Курс'] = ''
        finish_df['Группа'] = clear_df['Наименование группы']
        finish_df['Теория'] = clear_df['Теория']
        finish_df['ЛПЗ'] = clear_df['ЛПЗ']
        finish_df['Учебная практика'] = clear_df['Учебная практика']
        finish_df['Производственная практика'] = clear_df['Производственная практика']
        finish_df['Преддипломная практика'] = clear_df['Преддипломная практика']
        finish_df['Руководство над курсовым проектом'] = clear_df['Руководство над курсовым проектом']
        finish_df['Консультации'] = clear_df['Консультации']
        finish_df['Контроль (экзамены, зачеты и т.д.)'] = clear_df['Контроль (экзамены, зачеты и т.д.)']
        finish_df['Руководство ВКР'] = clear_df['Руководство ВКР']
        finish_df['ИТОГО'] = clear_df['ИТОГО']
        finish_df['Итого по тарификации'] = ''

        finish_df.dropna(subset=['Ф.И.О. преподавателя (полностью)'], inplace=True)

        # Получаем сумму колонок
        sum_col = finish_df.sum(axis=0, numeric_only=True).to_frame().T
        sum_col['Ф.И.О. преподавателя (полностью)'] = 'Итого'

        finish_df = pd.concat([finish_df, sum_col], ignore_index=True)


        out_df = pd.concat([out_df, finish_df], ignore_index=True)

# Создаем книгу для итогового файла
out_wb = openpyxl.Workbook()
out_ws = out_wb.active

# Записываем финальный датафрейм в созданную книгу
for r in dataframe_to_rows(out_df, index=False, header=True):
    if len(r) != 1:
        out_ws.append(r)

t = time.localtime()
current_time = time.strftime('%H_%M_%S', t)
# Сохраняем итоговый файл
out_wb.save(f'{path_to_end_folder}Приложение №6 от {current_time}.xlsx')