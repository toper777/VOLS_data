from pathlib import Path
import datetime
import pandas as pd
import openpyxl as opxl
from openpyxl.worksheet.table import Table, TableStyleInfo


def fill_cell_names():
    """
    Заполнение словаря для обращения к ячейкам Excel 1:A, 2:B,... 27:AA, 28:AB и так далее до ZZZ

    :return Dictionary:
    """
    inner_count = 1
    cell_names = {}

    for inner_i in range(65, 91):
        cell_names[inner_count] = chr(inner_i)
        inner_count += 1
    for inner_i in range(65, 91):
        for j in range(65, 91):
            cell_names[inner_count] = chr(inner_i) + chr(j)
            inner_count += 1
    for inner_i in range(65, 91):
        for j in range(65, 91):
            for k in range(65, 91):
                cell_names[inner_count] = chr(inner_i) + chr(j) + chr(k)
                inner_count += 1
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
    inner_dashboard_data = pd.read_json(url, convert_dates=('дата', 'Дата'))
    return inner_dashboard_data


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
        with pd.ExcelWriter(write_file_name, mode='a', if_sheet_exists="replace", datetime_format="DD.MM.YYYY",
                            engine='openpyxl') as writer:
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
    inner_wb = opxl.load_workbook(filename=format_file_name)
    tab = Table(displayName=format_tables_names[format_sheet],
                ref=f'A1:{excel_cell_names[len(format_frame.columns)]}{len(format_frame) + 1}')
    style = TableStyleInfo(name=table_style, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    inner_wb[format_sheet].add_table(tab)
    print(
        f'Writing FORMATTED {Color.GREEN}"{format_sheet}"{Color.END} sheet to file: {Color.CYAN}"{format_file_name}"{Color.END}')
    inner_wb.save(format_file_name)


def convert_date(convert_frame, convert_columns):
    """
    Конвертирует поля с датами в формат datetime64.
    Возвращает конвертированный DataFrame

    :param convert_frame:
    :param convert_columns:
    :return DataFrame:
    """
    columns_names = convert_frame.columns
    for column_name in columns_names:
        for column in convert_columns:
            if column.lower() in column_name.lower():
                convert_frame[column_name] = pd.to_datetime(convert_frame[column_name], dayfirst=True, format="%d.%m.%Y")
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


def last_day_of_month(date):
    if date.month == 12:
        return date.replace(day=31)
    return date.replace(month=date.month + 1, day=1) - datetime.timedelta(days=1)


def sum_all_events(sum_dataframe, sum_column):
    sum_all = 0
    for sum_data in sum_dataframe[sum_column]:
        sum_all += 1
    return sum_all


def sum_sort_events(sum_dataframe, sum_column, sum_condition):
    sum_sort = 0
    for sum_data in sum_dataframe[sum_column]:
        if sum_data in sum_condition:
            sum_sort += 1
    return sum_sort


def sum_done_events(sum_dataframe, sum_ks_date, sum_commissioning_date, sum_ks_status, sum_commissioning_status, sum_condition, sum_month):
    sum_sort = 0
    sort_frame = sum_dataframe[[sum_ks_date, sum_commissioning_date, sum_ks_status, sum_commissioning_status]]
    for row in sort_frame.values:
        # print_debug(3, row)
        if pd.Timestamp(row[0]) <= last_days_of_month[sum_month] and pd.Timestamp(row[1]) <= last_days_of_month[sum_month] and row[2] in sum_condition and row[3] in sum_condition:
            # print_debug(4, row)
            sum_sort += 1
    return sum_sort


def sum_sort_month_events(sum_dataframe, sum_column, sum_month):
    sum_sort = 0
    for sum_data in sum_dataframe[sum_column]:
        if pd.Timestamp(sum_data) <= last_days_of_month[sum_month]:
            sum_sort += 1
    return sum_sort


if __name__ == '__main__':
    # program and version
    program_name = "GDC_VOLS"
    program_version = "0.2.14"

    # Год анализа
    year = 2022

    # Месяц для анализа
    process_month = 2

    # main variables
    urls = {'Строительство гор.ВОЛС 2022': "https://gdc-rts/api/test-table/vw_2022_FOCL_Common_Build_City",
            'Реконструкция гор.ВОЛС 2022': "https://gdc-rts/api/test-table/vw_2022_FOCL_Common_Rebuild_City",
            'Строительство зон.ВОЛС 2022': "https://gdc-rts/api/test-table/vw_2022_FOCL_Common_Build_Zone",
            'Реконструкция зон.ВОЛС 2022': "https://gdc-rts/api/test-table/vw_2022_FOCL_Common_Rebuild_Zone"
            }

    report_sheet = "Отчетная таблица"

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
    today_date = datetime.date.today().strftime("%Y%m%d")  # YYYYMMDD format today date
    vols_dir = 'y:\\Блок №4\\ВОЛС\\2022\\'
    vols_file = f'{today_date} Отчет по строительству и реконструкции ВОЛС КФ 2022.xlsx'
    file_name = f'{vols_dir}{vols_file}'
    id_branch = "Филиал"
    work_branch = "Кавказский филиал"
    last_days_of_month = {}

    count = 0
    all_events = 0
    sort_events = 0

    process_columns_date = {'plan_date': 'Планируемая дата окончания',
                            'tz_date': 'Разработка ТЗ_дата',
                            'tz_date2': 'Разработка ТЗ ВОЛС_Дата',
                            'send_tz_date': 'Передача ТЗ подрядчику_дата',
                            'send_tz_date2': 'Передача ТЗ на ВОЛС подрядчику_Дата',
                            'received_tz_date': 'ТЗ принято подрядчиком_дата',
                            'received_tz_date2': 'ТЗ принято подрядчиком_дата',
                            'pir_smr_date': 'Заказ ПИР,СМР_дата',
                            'pir_smr_date2': 'Подписание договора (дс/заказа) на ПИР/ПИР+СМР_Дата',
                            'line_scheme_date': 'Линейная схема_дата',
                            'line_scheme_date2': 'Линейная схема_Дата',
                            'tu_date': 'Получение ТУ_дата',
                            'tu_date2': 'Получение ТУ_Дата',
                            'build_date': 'Строительство трассы_дата',
                            'build_date2': 'Строительство трассы_Дата',
                            'ks2_date': 'КС-2 (ПИР, СМР)_дата',
                            'ks2_date2': 'КС-2,3_Дата',
                            'commissioning_date': 'Приемка в эксплуатацию_дата',
                            'commissioning_date2': 'Приемка ВОЛС в эксплуатацию_Дата',
                            'complete_date': 'Дата ввода в эксплуатацию',
                            'complete_date2': 'Дата ввода ВОЛС в эксплуатацию',
                            'traffic_date': 'Запуск трафика_дата',
                            'traffic_date2': 'Запуск трафика_Дата'
                            }

    process_column_status = {'tz_status': 'Разработка ТЗ_статус',
                             'tz_status2': 'Разработка ТЗ ВОЛС_Статус',
                             'send_tz_status': 'Передача ТЗ подрядчику_статус',
                             'send_tz_status2': 'Передача ТЗ на ВОЛС подрядчику_Статус',
                             'received_tz_status': 'ТЗ принято подрядчиком_статус',
                             'received_tz_status2': 'ТЗ принято подрядчиком_статус',
                             'pir_smr_status': 'Заказ ПИР,СМР_статус',
                             'pir_smr_status2': 'Подписание договора (дс/заказа) на ПИР/ПИР+СМР_Статус',
                             'line_scheme_status': 'Линейная схема_статус',
                             'line_scheme_status2': 'Линейная схема_Статус',
                             'tu_status': 'Получение ТУ_статус',
                             'tu_status2': 'Получение ТУ_Статус',
                             'build_status': 'Строительство трассы_статус',
                             'build_status2': 'Строительство трассы_Статус',
                             'ks2_status': 'КС-2 (ПИР, СМР)_статус',
                             'ks2_status2': 'КС-2,3_Статус',
                             'commissioning_status': 'Приемка в эксплуатацию_статус',
                             'commissioning_status2': 'Приемка ВОЛС в эксплуатацию_Статус',
                             'traffic_status': 'Запуск трафика_статус',
                             'traffic_status2': 'Запуск трафика_Статус'
                             }

    # Получение исходных данных и запись форматированных данных
    for sheet, url in urls.items():
        data_frame = read_from_dashboard(url)
        data_frame = sort_branch(data_frame, id_branch, work_branch)
        data_frame = data_frame.reset_index(drop=True)
        data_frame = convert_date(data_frame, columns_dates)
        data_frame = convert_int(data_frame, columns_digit)
        data_frame = sort_by_id(data_frame, columns_for_sort)
        write_dataframe_to_file(data_frame, file_name, sheet)
        format_table(data_frame, sheet, file_name, excel_tables_names)

    # Создание отчёта
    print(
        f'Generate report sheet: {Color.GREEN}"{report_sheet}"{Color.END}')
    for i in range(1, 13):
        last_days_of_month[i] = pd.Timestamp(last_day_of_month(datetime.date(year, i, 1)))
    wb = opxl.load_workbook(filename=file_name)
    try:
        ws = wb[report_sheet]
    except:
        ws = wb.create_sheet(title=report_sheet)
    ws['A1'] = "Строительство городских ВОЛС"
    ws['A2'] = 'Всего мероприятий'
    ws['A4'] = 'Исполнение KPI ВОЛС КФ (накопительный итог)'
    ws['A6'] = 'Учтенных ВОЛС в KPI'
    ws['A8'] = 'Исполнение мероприятий в ЕСУП'
    ws['A9'] = 'Наименование мероприятия'
    ws['A10'] = 'Выпущены ТЗ'
    ws['A11'] = 'Переданы ТЗ в ПО'
    ws['A12'] = 'Приняты ТЗ ПО'
    ws['A13'] = 'Подписание договора на ПИР/ПИР+СМР'
    ws['A14'] = 'Линейная схема'
    ws['A15'] = 'Получено ТУ'
    ws['A16'] = 'Строительство трассы'
    ws['A17'] = 'Подготовка актов КС-2,3'
    ws['A18'] = 'Приёмка ВОЛС в эксплуатацию'
    ws['B9'] = 'Выполнено'
    ws['C9'] = 'Осталось'
    ws['A22'] = "Реконструкция городских ВОЛС"
    ws['A23'] = 'Всего мероприятий'
    ws['A25'] = 'Исполнение KPI ВОЛС КФ (накопительный итог)'
    ws['A27'] = 'Учтенных ВОЛС в KPI'
    ws['A29'] = 'Исполнение мероприятий в ЕСУП'
    ws['A30'] = 'Наименование мероприятия'
    ws['A31'] = 'Выпущены ТЗ'
    ws['A32'] = 'Переданы ТЗ в ПО'
    ws['A33'] = 'Приняты ТЗ ПО'
    ws['A34'] = 'Подписание договора на ПИР/ПИР+СМР'
    ws['A35'] = 'Линейная схема'
    ws['A36'] = 'Получено ТУ'
    ws['A37'] = 'Строительство трассы'
    ws['A38'] = 'Подготовка актов КС-2,3'
    ws['A39'] = 'Приёмка ВОЛС в эксплуатацию'
    ws['B30'] = 'Выполнено'
    ws['C30'] = 'Осталось'
    ws['B5'] = f'План, {datetime.datetime(year, process_month, 1).strftime("%b %Y")}'
    ws['C5'] = f'Факт, {datetime.datetime(year, process_month, 1).strftime("%b %Y")}'
    ws['D5'] = f'{chr(0x0394)}, {datetime.datetime(year, process_month, 1).strftime("%b %Y")}'
    ws['B26'] = f'План, {datetime.datetime(year, process_month, 1).strftime("%b %Y")}'
    ws['C26'] = f'Факт, {datetime.datetime(year, process_month, 1).strftime("%b %Y")}'
    ws['D26'] = f'{chr(0x0394)}, {datetime.datetime(year, process_month, 1).strftime("%b %Y")}'

    # Анализ строительства ВОЛС
    dashboard_data = pd.read_excel(file_name, sheet_name=list(urls.keys())[0])
    ws['B2'] = sum_all_events(dashboard_data, process_columns_date['plan_date'])
    ws['B6'] = sum_sort_month_events(dashboard_data, process_columns_date['plan_date'], process_month)
    ws['C6'] = sum_done_events(dashboard_data, process_columns_date['ks2_date'], process_columns_date['commissioning_date'], process_column_status['ks2_status'], process_column_status['commissioning_status'], ['Исполнена'], process_month)
    ws['D6'] = ws['C6'].value - ws['B6'].value
    for i, process in zip(range(10, 19), ['tz_status', 'send_tz_status', 'received_tz_status', 'pir_smr_status', 'line_scheme_status', 'tu_status', 'build_status', 'ks2_status', 'commissioning_status']):
        ws[f'B{i}'] = sum_sort_events(dashboard_data, process_column_status[process], ['Исполнена', 'Не требуется'])
    for i in range (10, 19):
        ws[f'C{i}'] = ws['B2'].value - ws[f'B{i}'].value

    # Анализ реконструкции ВОЛС
    dashboard_data = pd.read_excel(file_name, sheet_name=list(urls.keys())[1])
    ws['B23'] = sum_all_events(dashboard_data, process_columns_date['plan_date'])
    ws['B27'] = sum_sort_month_events(dashboard_data, process_columns_date['plan_date'], process_month)
    ws['C27'] = sum_done_events(dashboard_data, process_columns_date['ks2_date2'], process_columns_date['commissioning_date2'], process_column_status['ks2_status2'], process_column_status['commissioning_status2'], ['Исполнена'], process_month)
    ws['D27'] = ws['C27'].value - ws['B27'].value
    for i, process in zip(range(31, 40), ['tz_status2', 'send_tz_status2', 'received_tz_status2', 'pir_smr_status2', 'line_scheme_status2', 'tu_status2', 'build_status2', 'ks2_status2', 'commissioning_status2']):
        ws[f'B{i}'] = sum_sort_events(dashboard_data, process_column_status[process], ['Исполнена', 'Не требуется'])
    for i in range (31, 40):
        ws[f'C{i}'] = ws['B23'].value - ws[f'B{i}'].value
    print(
        f'Writing {Color.GREEN}"{report_sheet}"{Color.END} sheet to file: {Color.CYAN}"{file_name}"{Color.END}')
    wb.save(file_name)
