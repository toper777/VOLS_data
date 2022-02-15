from datetime import date
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo


def fill_cell_names():
    k = 1
    cell_names = {}

    for i in range(65, 91):
        cell_names[k] = chr(i)
        k += 1
    for i in range(65, 91):
        for j in range(65, 91):
            cell_names[k] = chr(i) + chr(j)
            k += 1
    return cell_names


# Colors for print
class Color:
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
    print(f'{Color.RED}DEBUG ({level}): \n{Color.END}{Color.YELLOW}{message}{Color.END}')


def read_from_dashboard(dashboard_url):
    print(f'Read data from: {Color.BLUE}"{dashboard_url}"{Color.END}')
    dashboard_data = pd.read_json(url)
# Leave index
#    dashboard_data = dashboard_data.set_index(work_index)
    return dashboard_data


def sort_branch(data_frame_sort, id_sort, branch_sort):
    data_frame_sort = data_frame_sort[data_frame_sort[id_sort] == branch_sort]
    return data_frame_sort


def write_to_file(write_frame, write_dir, write_file, write_sheet):
    file_name = f'{write_dir}{write_file}'
    if Path(file_name).is_file():
        with pd.ExcelWriter(file_name, mode='a', if_sheet_exists="replace", datetime_format="DD.MM.YYYY", engine='openpyxl') as writer:
            print(
                f'Writing {Color.GREEN}"{write_sheet}"{Color.END} sheet to file: {Color.CYAN}"{write_dir}{write_file}"{Color.END}')
            write_frame.to_excel(writer, sheet_name=write_sheet, index=False)
        format_table(write_frame, write_sheet, file_name, excel_tables_names)
    else:
        with pd.ExcelWriter(file_name, mode='w', datetime_format="DD.MM.YYYY", engine='openpyxl') as writer:
            print(
                f'Writing {Color.GREEN}"{write_sheet}"{Color.END} sheet to file: {Color.CYAN}"{write_dir}{write_file}"{Color.END}')
            write_frame.to_excel(writer, sheet_name=write_sheet, index=False)
        format_table(write_frame, write_sheet, file_name, excel_tables_names)


def format_table(format_frame, format_sheet, format_file_name, format_tables_names):
    print(
        f'Read {Color.GREEN}"{format_sheet}"{Color.END} sheet to file: {Color.CYAN}"{format_file_name}"{Color.END}')
    wb = openpyxl.load_workbook(filename=format_file_name)
#    print_debug(11, format_frame.columns)
#    ref_fields = f'A1:{chr(len(format_frame.columns) + 64).upper()}{len(format_frame) + 1}'
    print_debug(10, f'A1:{excel_cell_names[len(format_frame.columns) + 1]}{len(format_frame) + 1}')
    tab = Table(displayName=format_tables_names[format_sheet], ref=f'A1:{excel_cell_names[len(format_frame.columns)]}{len(format_frame) + 1}')
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    wb[format_sheet].add_table(tab)
    print(
        f'Writing FORMATTED {Color.GREEN}"{format_sheet}"{Color.END} sheet to file: {Color.CYAN}"{format_file_name}"{Color.END}')
    wb.save(format_file_name)


def convert_date(convert_frame, convert_columns):
    columns_names = list(convert_frame)
    for column_name in columns_names:
        for column in convert_columns:
            if column.lower() in column_name.lower():
                convert_frame = convert_frame.astype({column_name: 'datetime64[ns]'})
#                print_debug(5, f'{column_name} # {convert_frame[column_name].dtype} # {convert_frame[column_name].values}')
            else:
                pass
    return convert_frame


def convert_digit(convert_frame, convert_columns):
    columns_names = list(convert_frame)
    for column_name in columns_names:
        for column in convert_columns:
            if column.lower() in column_name.lower():
                convert_frame = convert_frame.astype({column_name: 'int32'})
            else:
                pass
    return convert_frame


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

# Наименования колонок для преобразования даты
columns_dates = ['Планируемая дата окончания', 'Дата ввода', '_дата']
# Наименования колонок для преобразования числа
columns_digit = ['ID']

today_date = date.today().strftime("%Y%m%d")  # YYYYMMDD format today date
vols_dir = 'y:\\Блок №4\\ВОЛС\\2022\\'
vols_file = f'{today_date} Отчет по строительству и реконструкции ВОЛС КФ 2022.xlsx'

work_index = "ID"
id_branch = "Филиал"
work_branch = "Кавказский филиал"


if __name__ == '__main__':
    for sheet, url in urls.items():
        data_frame = read_from_dashboard(url)
#        print_debug(1, data_frame.dtypes)
        data_frame = sort_branch(data_frame, id_branch, work_branch)
#        print_debug(2, data_frame.dtypes)
        data_frame = convert_date(data_frame, columns_dates)
#        print_debug(3, data_frame.dtypes)
        data_frame = convert_digit(data_frame, columns_digit)
#        print_debug(4, data_frame.dtypes)
        write_to_file(data_frame, vols_dir, vols_file, sheet)
