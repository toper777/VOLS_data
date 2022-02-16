from datetime import date
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo


def fill_cell_names():
    """
    Заполнение словаря для обращения к ячейкам Excel 1:A, 2:B,... 27:AA, 28:AB и так далее до ZZZ

    :return Dictionary:
    """
    count = 1
    cell_names = {}

    for i in range(65, 91):
        cell_names[count] = chr(i)
        count += 1
    for i in range(65, 91):
        for j in range(65, 91):
            cell_names[count] = chr(i) + chr(j)
            count += 1
    for i in range(65, 91):
        for j in range(65, 91):
            for k in range(65, 91):
                cell_names[count] = chr(i) + chr(j) + chr(k)
                count += 1
    return cell_names


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


def read_from_dashboard(dashboard_url):
    """
    Читает данные JSON из url и сохраняет их в DataFrame

    :param dashboard_url:
    :return DataFrame:
    """
    print(f'Read data from: {Color.BLUE}"{dashboard_url}"{Color.END}')
    dashboard_data = pd.read_json(url)
# Leave index
#    dashboard_data = dashboard_data.set_index(work_index)
    return dashboard_data


def sort_branch(data_frame_sort, id_sort, branch_sort):
    """
    Сортирует DataFrame и возвращает DataFrame с данными только по заданному филиала

    :param data_frame_sort:
    :param id_sort:
    :param branch_sort:
    :return DataFrame:
    """
    data_frame_sort = data_frame_sort[data_frame_sort[id_sort] == branch_sort]
    return data_frame_sort


def write_dataframe_to_file(write_frame, write_file_name, write_sheet):
    """
    Записывает в Excel файл таблицы с данными

    :param write_frame:
    :param write_file_name:
    :param write_sheet:
    """
    if Path(write_file_name).is_file():
        with pd.ExcelWriter(write_file_name, mode='a', if_sheet_exists="replace", datetime_format="DD.MM.YYYY", engine='openpyxl') as writer:
            print(
                f'Writing {Color.GREEN}"{write_sheet}"{Color.END} sheet to file: {Color.CYAN}"{write_file_name}"{Color.END}')
            write_frame.to_excel(writer, sheet_name=write_sheet, index=False)
    else:
        with pd.ExcelWriter(write_file_name, mode='w', datetime_format="DD.MM.YYYY", engine='openpyxl') as writer:
            print(
                f'Writing {Color.GREEN}"{write_sheet}"{Color.END} sheet to file: {Color.CYAN}"{write_file_name}"{Color.END}')
            write_frame.to_excel(writer, sheet_name=write_sheet, index=False)



def format_table(format_frame, format_sheet, format_file_name, format_tables_names):
    """
    Форматирует таблицы для Excel файла и перезаписывает в файл в виде именованных Таблиц

    :param format_frame:
    :param format_sheet:
    :param format_file_name:
    :param format_tables_names:
    """
    print(
        f'Read {Color.GREEN}"{format_sheet}"{Color.END} sheet from file: {Color.CYAN}"{format_file_name}"{Color.END}')
    wb = openpyxl.load_workbook(filename=format_file_name)
    tab = Table(displayName=format_tables_names[format_sheet],
                ref=f'A1:{excel_cell_names[len(format_frame.columns)]}{len(format_frame) + 1}')
    style = TableStyleInfo(name=table_style, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    wb[format_sheet].add_table(tab)
    print(
        f'Writing FORMATTED {Color.GREEN}"{format_sheet}"{Color.END} sheet to file: {Color.CYAN}"{format_file_name}"{Color.END}')
    wb.save(format_file_name)


def convert_date(convert_frame, convert_columns):
    """
    Конвертирует поля с датами в формат datetime64.
    Возвращает конвертированный DataFrame

    :param convert_frame:
    :param convert_columns:
    :return DataFrame:
    """
    columns_names = list(convert_frame)
    for column_name in columns_names:
        for column in convert_columns:
            if column.lower() in column_name.lower():
                convert_frame = convert_frame.astype({column_name: 'datetime64[ns]'})
            else:
                pass
    return convert_frame


def convert_int(convert_frame, convert_columns):
    """
    Конвертирует поля с целыми в формат int32.
    Возвращает конвертированный DataFrame

    :param convert_frame:
    :param convert_columns:
    :return DataFrame:
    """
    columns_names = list(convert_frame)
    for column_name in columns_names:
        for column in convert_columns:
            if column.lower() in column_name.lower():
                convert_frame = convert_frame.astype({column_name: 'int32'})
            else:
                pass
    return convert_frame


def sort_by_id(data_frame_sort, id_sort):
    """
    Сортирует DataFrame и возвращает DataFrame с данными отсортированными по возрастанию

    :param data_frame_sort:
    :param id_sort:
    :return DataFrame:
    """
    data_frame_sort = data_frame_sort.sort_values(by=id_sort)
    return data_frame_sort


# program and version
program_name = "GDC_VOLS"
program_version = "0.1.1"

# main variables
urls = {'Строительство гор.ВОЛС 2022': "https://gdc-rts/api/test-table/vw_2022_FOCL_Common_Build_City",
        'Реконструкция гор.ВОЛС 2022': "https://gdc-rts/api/test-table/vw_2022_FOCL_Common_Rebuild_City",
        'Строительство зон.ВОЛС 2022': "https://gdc-rts/api/test-table/vw_2022_FOCL_Common_Build_Zone",
        'Реконструкция зон.ВОЛС 2022': "https://gdc-rts/api/test-table/vw_2022_FOCL_Common_Rebuild_Zone"
        }
excel_tables_names = {'Строительство гор.ВОЛС 2022': "Urban_VOLS_Build",
                      'Реконструкция гор.ВОЛС 2022': "Urban_VOLS_Reconstruction",
                      'Строительство зон.ВОЛС 2022': "Zone_VOLS_Build",
                      'Реконструкция зон.ВОЛС 2022': "Zone_VOLS_Reconstruction"}
excel_cell_names = fill_cell_names()

# Стиль таблицы Excel
table_style = "TableStyleMedium2"
# Наименования колонок для преобразования даты
columns_dates = ['Планируемая дата окончания', 'Дата ввода', '_дата']
# Наименования колонок для преобразования числа
columns_digit = ['ID']
# Наименование колонки для сортировки по возрастанию
columns_for_sort = ['ID']
today_date = date.today().strftime("%Y%m%d")  # YYYYMMDD format today date
vols_dir = 'y:\\Блок №4\\ВОЛС\\2022\\'
vols_file = f'{today_date} Отчет по строительству и реконструкции ВОЛС КФ 2022.xlsx'
file_name = f'{vols_dir}{vols_file}'
id_branch = "Филиал"
work_branch = "Кавказский филиал"


if __name__ == '__main__':
    for sheet, url in urls.items():
        data_frame = read_from_dashboard(url)
        data_frame = sort_branch(data_frame, id_branch, work_branch)
        data_frame = convert_date(data_frame, columns_dates)
        data_frame = convert_int(data_frame, columns_digit)
        data_frame = sort_by_id(data_frame, columns_for_sort)
        write_dataframe_to_file(data_frame, file_name, sheet)
        format_table(data_frame, sheet, file_name, excel_tables_names)

