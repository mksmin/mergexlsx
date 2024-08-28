"""
Это модуль открывает каждый Excel файл из выгрузки, а так же мастер-файл
Считает количество заявок в каждой компетенции
и выводит общую табличку со статистикой
"""


# Импорт
import pandas as pd
import os, re
from dotenv import load_dotenv
import openpyxl
from openpyxl.styles import Alignment

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

files = os.listdir(path_from)  # Получаем список всех файлов

needed_columns = re.split(r", ", os.environ['NEED_COLUMNS'])  # Получаем из .env файла список столбцов
path_to_write = os.getenv('PATHWRITE')  # Получаем из .env файла путь к мастерфайлу, который будем перезаписывать
path_to_stats = os.path.join(os.path.dirname(__file__), 'FilesXlsx\\stats.xlsx')

data_dict = {}
names = []
values = []

data_dict_master = {}
names_mf = []
values_mf = []

actual_id_masterfile = []
with pd.ExcelFile(path_to_write) as xls:  # Получаем список актуальных листов в Мастерфайле
    sheets_actual = xls.sheet_names
    print(f'НАЧИНАЮ РАБОТАТЬ С МАСТЕРФАЙЛОМ')
    actual_id = 0
    for sheet in sheets_actual:
        master_file = pd.read_excel(path_to_write, sheet_name=sheet, skiprows=2)  # Открываем нужный лист в Мастерфайле
        current_id = [mf for mf in master_file['ID'] if str(mf) != 'nan']  # Собираем актуальные ID
        actual_id = actual_id + len(current_id)
        actual_id_masterfile = [*actual_id_masterfile, *current_id]
        names_mf.append(sheet)
        values_mf.append(len(current_id))

    else:
        names_mf.append(' ')
        values_mf.append(' ')
        names_mf.append('ВСЕГО ЗАЯВОК')
        values_mf.append(actual_id)

        data_dict_master['Компетенция'] = names_mf
        data_dict_master['Заявок'] = values_mf


print(f'\n\n\nЗАКОНЧИЛ С МАСТЕРФАЙЛОМ\n\n\n')
print(f'НАЧИНАЮ РАБОТУ С ВЫГРУЗКОЙ')

# Обрабатываю все файлы из выгрузки
all_id_list = []
all_id_list_filter = []
index = 0

for file in files:
    index = index + 1
    check_file = pd.read_excel(os.path.join(path_from, file))
    current_ids = [mf for mf in check_file['ID']]
    all_id_list = [*all_id_list, *current_ids]
    names.append(file)
    values.append(len(current_ids))
    current_ids_filter = [idfilt for idfilt in check_file['ID'] if idfilt not in all_id_list_filter]
    all_id_list_filter = [*all_id_list_filter, *current_ids_filter]

else:
    names.append(' ')
    values.append(' ')
    names.append('ВСЕГО ЗАЯВОК')
    values.append(len(all_id_list))

    data_dict['Файл'] = names
    data_dict['Заявок'] = values

    value_1 = pd.DataFrame(data_dict)
    value_2 = pd.DataFrame(data_dict_master)

    value_1.index += 1
    value_2.index += 1
    with pd.ExcelWriter(path_to_stats, engine='openpyxl', mode='a', if_sheet_exists='replace') as writers:
        value_1.to_excel(writers, sheet_name='Лист1', index_label='№')

    with pd.ExcelWriter(path_to_stats, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writers:
        value_2.to_excel(writers, sheet_name='Лист1', startcol=5, index_label='№')

    wb = openpyxl.load_workbook(path_to_stats)
    sheet = wb.active
    cell_align_1 = sheet['C']
    range = sheet['A1':'H50']
    for cell in range:
        for c in cell:
            c.alignment = Alignment(horizontal='center', vertical='center')

    range = sheet['B2':'B50']
    for cell in range:
        for c in cell:
            c.alignment = Alignment(horizontal='left', vertical='center')

    range = sheet['G2':'G50']
    for cell in range:
        for c in cell:
            c.alignment = Alignment(horizontal='left', vertical='center')

    sheet.column_dimensions['A'].width = 5
    sheet.column_dimensions['B'].width = 50
    sheet.column_dimensions['C'].width = 10
    sheet.column_dimensions['F'].width = 5
    sheet.column_dimensions['G'].width = 50
    sheet.column_dimensions['H'].width = 10
    wb.save(path_to_stats)

print(f'\nВСЕГО ФАЙЛОВ В ВЫГРУЗКЕ: {len(files)}')
print(f'ЗАКОНЧИЛ С ВЫГРУЗКОЙ\n\n\n')

print(f'ВСЕГО ЗАЯВОК В МАСТЕРФАЙЛЕ: {actual_id}')
print(f'ВСЕГО ЗАЯВОК В ВЫГРУЗКЕ ДО ФИЛЬТРА: {len(all_id_list)}')
print(f'ЗАЯВОК В ВЫГРУЗКЕ ПОСЛЕ ФИЛЬТРА: {len(all_id_list_filter)}')

print(f'Разница в заявках до фильтра: {len(actual_id_masterfile) - len(all_id_list)}')
print(f'Разница в заявках после фильтра: {len(actual_id_masterfile) - len(all_id_list_filter)}')
