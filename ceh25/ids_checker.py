import pandas as pd
from pathlib import Path
from openpyxl import Workbook

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

path_parent = Path(__file__).parent
path_footage = path_parent / 'footage'
files = path_footage.glob('*')
path_to_write = path_parent / 'Новые заявки.xlsx'


list_of_files = []
i = 0
for file in files:
    i += 1
    print(f'{i}. {file = }')
    list_of_files.append(file)

screening_file = int(input('Файл с отобранными заявками (число): '))
file_with_new_applications = int(input('Файл с новыми заявками: '))

with pd.ExcelFile(list_of_files[screening_file - 1]) as xls:
    sheets_actual = xls.sheet_names
    print()

    master_file = pd.read_excel(list_of_files[screening_file - 1], sheet_name=sheets_actual[0])
    columns_name = []

    col_i = 0
    for col in master_file.columns:
        col_i += 1
        columns_name.append(col)
        print(f'{col_i}. {col}')

    user_choice_column = int(input('Выбери столбец (номер): '))

    current_id = [uid for uid in master_file[columns_name[user_choice_column - 1]] if str(uid) != 'nan']
    print(f'id участников, которые отобранны: {current_id = }')

with pd.ExcelFile(list_of_files[file_with_new_applications - 1]) as xlsn:
    sheets_actual = xlsn.sheet_names
    print()

    master_file = pd.read_excel(list_of_files[file_with_new_applications - 1], sheet_name=sheets_actual[0], index_col=0)
    columns_name = []

    col_i = 0
    for col in master_file.columns:
        col_i += 1
        columns_name.append(col)
        print(f'{col_i}. {col}')

    col_index_start = len(master_file.index)

    for value in master_file.index:
        if value in current_id:
            master_file = master_file.drop(value)
    else:
        col_index_end = len(master_file.index)
        print(f'Старт: {col_index_start}. Конец: {col_index_end}')

    with pd.ExcelWriter(path_to_write, engine='openpyxl', mode='w') as writer:
        master_file.to_excel(writer, header=True)
