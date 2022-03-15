#  Copyright (c) 2022. Tikhon Ostapenko
import sys
from pathlib import Path
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import datetime
import pandas as pd
import openpyxl as openpyxl


def fill_cell_names():
    """
    Заполнение словаря для обращения к ячейкам Excel 1:A, 2:B,... 27:AA, 28:AB и так далее до ZZZ

    :return Dictionary:
    """
    _count = 1
    _cell_names = {}

    for _i in range(65, 91):
        _cell_names[_count] = chr(_i)
        _count += 1
    for _i in range(65, 91):
        for _j in range(65, 91):
            _cell_names[_count] = chr(_i) + chr(_j)
            _count += 1
    for _i in range(65, 91):
        for _j in range(65, 91):
            for _k in range(65, 91):
                _cell_names[_count] = chr(_i) + chr(_j) + chr(_k)
                _count += 1
    return _cell_names


# Colors for print
class Color:
    """
    Содержит кодировки цветов для консольного вывода
    """
    PURPLE = '\033[95m'
    CYAN = '\033[96m'
    DARKCYAN = '\033[36m'
    BLUE = '\033[94m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    END = '\033[0m'


def print_debug(level, message):
    """
    Функция печати DEBUG сообщений.
    Принимает параметр номер и само сообщение.
    Сообщение должно быть в формате строки вывода

    :param level:
    :param message:
    """
    print(f'{Color.RED}DEBUG ({level}): \n{Color.END}{Color.YELLOW}{message}{Color.END}')


def read_from_dashboard(_url):
    """
    Читает данные JSON из url и сохраняет их в DataFrame

    :param _url:
    :return DataFrame:
    """
    print(f'Read data from: "{_url}"')
    try:
        _dashboard_data = pd.read_json(_url, convert_dates=('дата', 'Дата'))
    except Exception:
        print(f"ERROR: can't read data from url {_url}")
        sys.exit(1)
    return _dashboard_data


def sort_branch(_data_frame, _id, _branch):
    """
    Сортирует DataFrame и возвращает DataFrame с данными только по заданному филиала

    :param _data_frame:
    :param _id:
    :param _branch:
    :return DataFrame:
    """
    _data_frame = _data_frame[_data_frame[_id] == _branch]
    return _data_frame


def write_dataframe_to_file(_data_frame, _file_name, _sheet):
    """
    Записывает в Excel файл таблицы с данными

    :param _data_frame:
    :param _file_name:
    :param _sheet:
    """
    if Path(_file_name).is_file():
        with pd.ExcelWriter(_file_name, mode='a', if_sheet_exists="replace", datetime_format="DD.MM.YYYY",
                            engine='openpyxl') as writer:
            print(
                f'Append "{_sheet}" sheet to exist file: "{_file_name}"')
            _data_frame.to_excel(writer, sheet_name=_sheet, index=False)
    else:
        with pd.ExcelWriter(_file_name, mode='w', datetime_format="DD.MM.YYYY", engine='openpyxl') as writer:
            print(
                f'Write "{_sheet}" sheet to new file: "{_file_name}"')
            _data_frame.to_excel(writer, sheet_name=_sheet, index=False)


def format_table(_data_frame, _sheet, _file_name, _tables_names, _excel_cell_names, _table_style):
    """
    Форматирует таблицы для Excel файла и перезаписывает в файл в виде именованных Таблиц

    :param _table_style:
    :param _excel_cell_names:
    :param _data_frame:
    :param _sheet:
    :param _file_name:
    :param _tables_names:
    """
    if not _data_frame.empty:
        print(
            f'Read "{_sheet}" sheet from file: "{_file_name}"')
        _wb = openpyxl.load_workbook(filename=_file_name)
        tab = Table(displayName=_tables_names[_sheet],
                    ref=f'A1:{_excel_cell_names[len(_data_frame.columns)]}{len(_data_frame) + 1}')
        style = TableStyleInfo(name=_table_style, showRowStripes=True, showColumnStripes=True)
        tab.tableStyleInfo = style
        _wb[_sheet].add_table(tab)
        try:
            _ws = _wb[_sheet]
        except Exception:
            _ws = _wb.create_sheet(title=_sheet)
        _ws = adjust_columns_width(_ws)
        print(
            f'Write formatted "{_sheet}" sheet to file: "{_file_name}"')
        _wb.save(_file_name)
    else:
        pass


def convert_date(_data_frame, _columns):
    """
    Конвертирует поля с датами в формат datetime64.
    Возвращает конвертированный DataFrame

    :param _data_frame:
    :param _columns:
    :return DataFrame:
    """
    _columns_names = _data_frame.columns
    for _column_name in _columns_names:
        for _column in _columns:
            if _column.lower() in _column_name.lower():
                _data_frame[_column_name] = pd.to_datetime(_data_frame[_column_name], dayfirst=True, format="%d.%m.%Y")
            else:
                pass
    return _data_frame


def convert_int(_data_frame, _columns):
    """
    Конвертирует поля с целыми в формат int32.
    Возвращает конвертированный DataFrame

    :param _data_frame:
    :param _columns:
    :return DataFrame:
    """
    _columns_names = list(_data_frame)
    for _column_name in _columns_names:
        for _column in _columns:
            if _column.lower() in _column_name.lower():
                _data_frame = _data_frame.astype({_column_name: 'int32'})
            else:
                pass
    return _data_frame


def sort_by_column(_data_frame, _column):
    """
    Сортирует DataFrame и возвращает DataFrame с данными отсортированными по возрастанию

    :param _data_frame:
    :param _column:
    :return DataFrame:
    """
    _data_frame = _data_frame.sort_values(by=_column)
    return _data_frame


def last_day_of_month(_date):
    if _date.month == 12:
        return _date.replace(day=31)
    return _date.replace(month=_date.month + 1, day=1) - datetime.timedelta(days=1)


def sum_sort_events(_data_frame, _column, _condition):
    _sum_sort = 0
    for _sum_data in _data_frame[_column]:
        if _sum_data in _condition:
            _sum_sort += 1
    return _sum_sort


def sum_done_events(_data_frame, _ks_date, _commissioning_date, _ks_status, _commissioning_status, _condition, _month, _last_days_of_month):
    _sum_sort = 0
    _sort_frame = _data_frame[[_ks_date, _commissioning_date, _ks_status, _commissioning_status]]
    for _row in _sort_frame.values:
        if pd.Timestamp(_row[0]) <= _last_days_of_month[_month] and pd.Timestamp(_row[1]) <= _last_days_of_month[_month] and _row[2] in _condition and _row[3] in _condition:
            _sum_sort += 1
    return _sum_sort


def sum_sort_month_events(_data_frame, _column, _month, _last_days_of_month):
    _sum_sort = 0
    for _sum_data in _data_frame[_column]:
        if pd.Timestamp(_sum_data) <= _last_days_of_month[_month]:
            _sum_sort += 1
    return _sum_sort


def write_report_table_to_file(_data_frame, _file_name, _sheet, _excel_tables_names, _excel_cell_names, _table_style):
    write_dataframe_to_file(_data_frame, _file_name, _sheet)
    format_table(_data_frame, _sheet, _file_name, _excel_tables_names, _excel_cell_names, _table_style)


def adjust_columns_width(_dataframe):
    # Форматирование ширины полей отчётной таблицы
    for _col in _dataframe.columns:
        _max_length = 0
        _column = get_column_letter(_col[0].column)  # Get the column name
        for _cell in _col:
            if _cell.coordinate in _dataframe.merged_cells:  # not check merge_cells
                continue
            try:  # Necessary to avoid error on empty cells
                if len(str(_cell.value)) > _max_length:
                    _max_length = len(str(_cell.value))
            except Exception:
                pass
        _adjusted_width = _max_length + 3
        _dataframe.column_dimensions[_column].width = _adjusted_width
    return _dataframe
