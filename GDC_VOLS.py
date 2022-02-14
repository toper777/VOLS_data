from datetime import date
from pathlib import Path
import pandas as pd

# program and version
program_name = "GDC_VOLS"
program_version = "0.1.1"

# main variables
urls = {'Строительство гор.ВОЛС 2022': "https://gdc-rts/api/test-table/vw_2022_FOCL_Common_Build_City",
        'Реконструкция гор.ВОЛС 2022': "https://gdc-rts/api/test-table/vw_2022_FOCL_Common_Rebuild_City",
        'Строительство зон.ВОЛС 2022': "https://gdc-rts/api/test-table/vw_2022_FOCL_Common_Build_Zone",
        'Реконструкция зон.ВОЛС 2022': "https://gdc-rts/api/test-table/vw_2022_FOCL_Common_Rebuild_Zone"}

today_date = date.today().strftime("%Y%m%d")  # YYYYMMDD format today date
vols_dir = "y:/Блок №4/ВОЛС/2022/"
vols_file = f'{today_date} Отчет по строительству и реконструкции ВОЛС КФ 2022.xlsx'

work_index = "ID"
id_branch = "Филиал"
work_branch = "Кавказский филиал"


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


def read_from_dashboard(dashboard_url):
    print(f'Read data from: {Color.BLUE}"{dashboard_url}"{Color.END}')
    dashboard_data = pd.read_json(url)
    dashboard_data = dashboard_data.set_index(work_index)
    return dashboard_data


def sort_branch(data_frame_sort, id_sort, branch_sort):
    data_frame_sort = data_frame_sort[data_frame_sort[id_sort] == branch_sort]
    return data_frame_sort


def write_to_file(write_frame, write_dir, write_file, write_sheet):
    if Path(write_dir+write_file).is_file():
        with pd.ExcelWriter(write_dir + write_file, mode='a', if_sheet_exists="replace") as writer:
            print(
                f'Writing {Color.GREEN}"{write_sheet}"{Color.END} sheet to file: {Color.CYAN}"{write_dir}{write_file}"{Color.END}')
            write_frame.to_excel(writer, sheet_name=write_sheet)
    else:
        with pd.ExcelWriter(write_dir + write_file, mode='w') as writer:
            print(
                f'Writing {Color.GREEN}"{write_sheet}"{Color.END} sheet to file: {Color.CYAN}"{write_dir}{write_file}"{Color.END}')
            write_frame.to_excel(writer, sheet_name=write_sheet)


def convert_date(convert_frame):
    columns_names = list(convert_frame)
    for column_name in columns_names:
        if 'Планируемая дата окончания' in column_name:
            convert_frame = convert_frame.astype({column_name: 'datetime64[ns]'})
        elif 'Дата ввода' in column_name:
            convert_frame = convert_frame.astype({column_name: 'datetime64[ns]'})
        elif "_дата" in column_name.lower():
            convert_frame = convert_frame.astype({column_name: 'datetime64[ns]'})
        else:
            pass
    return convert_frame


if __name__ == '__main__':
    for sheet, url in urls.items():
        data_frame = read_from_dashboard(url)
        data_frame = sort_branch(data_frame, id_branch, work_branch)
        data_frame = convert_date(data_frame)
        write_to_file(data_frame, vols_dir, vols_file, sheet)
