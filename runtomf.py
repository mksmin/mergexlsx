"""
Этот модуль открывает последнюю версию мастер-файла Excel и новую версию из выгрузки
И добавляет в первый все новые значения
"""

# Импорт
import pandas as pd
import os
from dotenv import load_dotenv


# Снимаем ограничения на показ строк и столбцов
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# Подгружаем .env файл
dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
load_dotenv(dotenv_path)

# Назначаем путь к списку файлов, которые нужно обработать
path_from = os.path.normpath(
    os.path.join(os.path.dirname(__file__),
                 f'.\\from\\'))

path_to_write = os.getenv('MFPATHTOWROTE')  # Получаем из .env файла путь к мастерфайлу, который будем перезаписывать
files = os.listdir(path_from)  # Получаем список всех файлов
with pd.ExcelFile(path_to_write) as xls:  # Получаем список актуальных листов в Мастерфайле
    sheets_actual = xls.sheet_names

cortej = set()
drop_counts = 0
nondroped = 0
for i in files:
    path_new = os.path.join(path_from, i)
    master_file = pd.read_excel(path_to_write)
    col_start = master_file.shape[0] + 1
    print(f'СТАРТ: {i}')
    current_id = list(master_file['ID'])  # собираем актуальные ID

    with pd.ExcelFile(path_new) as xls:
        value_from_unload = pd.read_excel(xls, index_col=0)
        col_index_start = len(value_from_unload.index)

        for value in value_from_unload['ID']:
            index_cell = value_from_unload[value_from_unload['ID'] == value].index
            cortej.add(value)

            if value in current_id:
                value_from_unload = value_from_unload.drop(index_cell)
                drop_counts += 1
            else:
                nondroped += 1

        with pd.ExcelWriter(path_to_write, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writers:
            value_from_unload.to_excel(writers, sheet_name=sheets_actual[0], startrow=col_start, header=False)

print(f'ВСЕГО ЗАЯВОК: {len(cortej)}')

print(f'УДАЛИЛ ЗАЯВОК: {drop_counts}')
print(f'ОСТАВИЛ ЗАЯВОК: {nondroped}')
