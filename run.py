"""
Это модуль открывает каждый Excel файл из выгрузки, а так же мастер-файл
Добавляет все новые значения (сравнивая их по id)
Добавляет их на лист в мастер-файле, который называется так же, как называется файл в из выгрузки
Если названия не совпадают - предлагает создать новый лист и подгрузить в него данные
"""


# Импорт
import pandas as pd
import os, re
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

needed_columns = re.split(r", ", os.environ['NEED_COLUMNS'])  # Получаем из .env файла список столбцов
path_to_write = os.getenv('PATHWRITE')  # Получаем из .env файла путь к мастерфайлу, который будем перезаписывать


files = os.listdir(path_from)  # Получаем список всех файлов
with pd.ExcelFile(path_to_write) as xls:  # Получаем список актуальных листов в Мастерфайле
    sheets_actual = xls.sheet_names

df = pd.read_excel(path_to_write) # Читаем мастерфайл

for i in files:
    path_new = os.path.join(path_from, i)  # Через цикл работаем с каждым файлом из папки FROM

    print(f'СТАРТ: {i}')

    sheet_name = i[:-5]  # Лист будет называться так как файл без .xlsx
    if len(sheet_name) > 31:  # У листа ограничение в названии в 31 символ
        new_digit = len(sheet_name) - 31
        new_sheet_name = sheet_name[:-new_digit]
    else:
        new_sheet_name = sheet_name

    if new_sheet_name in sheets_actual:
        master_file = pd.read_excel(path_to_write,
                                    sheet_name=new_sheet_name, skiprows=2)  # Открываем нужный лист в Мастерфайле
        current_id = [mf for mf in master_file['ID']] # Собираем актуальные ID
        print(f'АКТУАЛЬНЫЕ ID: {current_id}')

        with pd.ExcelFile(path_new) as xls:
            value_1 = pd.read_excel(xls, usecols=needed_columns, index_col=0) # Выгружаем все значения из нужных столбцов
            col_index_start = len(value_1.index)  # Считаем сколько значений было до удаления

            for value in value_1.index:
                if value in current_id:
                    value_1 = value_1.drop(value)  # Если ID есть в Мастерфайле, то удаляем
            else:
                col_index_end = len(value_1.index)
                print(f'ПРОВЕРИЛ ВСЕ СОВПАДЕНИЯ ИНДЕКСА.'
                      f'\n Было {col_index_start} // стало {col_index_end}'
                      f'\n Удалено {col_index_start - col_index_end}')

        list_col = pd.read_excel(path_to_write,
                                 sheet_name=new_sheet_name)
        col_start = list_col.shape[0] + 2

        with pd.ExcelWriter(path_to_write, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writers:
            value_1.to_excel(writers, sheet_name=new_sheet_name, startrow=col_start, startcol=3, header=False)
    else:
        with pd.ExcelFile(path_new) as xls:
            value_1 = pd.read_excel(xls, usecols=needed_columns, index_col=0)
        print(f'В мастерфайле нет листа с названием: {new_sheet_name}')
        create_sheet = input('Создать такой лист и загрузить в него данные? (да/нет): ').lower()

        continue_operation = False
        while not continue_operation:
            if create_sheet == 'да':
                with pd.ExcelWriter(path_to_write, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writers:
                    value_1.to_excel(writers, sheet_name=new_sheet_name,  header=False)
                continue_operation = True
            elif create_sheet == 'нет':
                print(f'ОПЕРАЦИЯ ОТМЕНЕНА')
                continue_operation = True
            else:
                create_sheet = input('Создать такой лист и загрузить в него данные? (да/нет): ').lower()


    print(f'ЗАКОНЧИЛ: {i}')
    print(' ')
