from pathlib import Path
import datetime
import pandas as pd
import openpyxl as openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, Side, PatternFill, Alignment, Border
import openpyxl.styles.borders as borders_style
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


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
    _dashboard_data = pd.read_json(_url, convert_dates=('дата', 'Дата'))
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


def format_table(_data_frame, _sheet, _file_name, _tables_names):
    """
    Форматирует таблицы для Excel файла и перезаписывает в файл в виде именованных Таблиц

    :param _data_frame:
    :param _sheet:
    :param _file_name:
    :param _tables_names:
    """
    print(
        f'Read "{_sheet}" sheet from file: "{_file_name}"')
    _wb = openpyxl.load_workbook(filename=_file_name)
    tab = Table(displayName=_tables_names[_sheet],
                ref=f'A1:{excel_cell_names[len(_data_frame.columns)]}{len(_data_frame) + 1}')
    style = TableStyleInfo(name=table_style, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    _wb[_sheet].add_table(tab)
    try:
        _ws = _wb[_sheet]
    except:
        _ws = _wb.create_sheet(title=_sheet)
    _ws = adjust_columns_width(_ws)
    print(
        f'Write formatted "{_sheet}" sheet to file: "{_file_name}"')
    _wb.save(_file_name)


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


def sort_by_id(_data_frame, _id):
    """
    Сортирует DataFrame и возвращает DataFrame с данными отсортированными по возрастанию

    :param _data_frame:
    :param _id:
    :return DataFrame:
    """
    _data_frame = _data_frame.sort_values(by=_id)
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


def sum_done_events(_data_frame, _ks_date, _commissioning_date, _ks_status, _commissioning_status, _condition, _month):
    _sum_sort = 0
    _sort_frame = _data_frame[[_ks_date, _commissioning_date, _ks_status, _commissioning_status]]
    for _row in _sort_frame.values:
        if pd.Timestamp(_row[0]) <= last_days_of_month[_month] and pd.Timestamp(_row[1]) <= last_days_of_month[_month] and _row[2] in _condition and _row[3] in _condition:
            _sum_sort += 1
    return _sum_sort


def sum_sort_month_events(_data_frame, _column, _month):
    _sum_sort = 0
    for _sum_data in _data_frame[_column]:
        if pd.Timestamp(_sum_data) <= last_days_of_month[_month]:
            _sum_sort += 1
    return _sum_sort


def write_report_table_to_file(_data_frame, _file_name, _sheet, _excel_tables_names):
    write_dataframe_to_file(_data_frame, _file_name, _sheet)
    format_table(_data_frame, _sheet, _file_name, _excel_tables_names)


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
            except:
                pass
        _adjusted_width = (_max_length + 2)
        _dataframe.column_dimensions[_column].width = _adjusted_width
    return _dataframe


if __name__ == '__main__':
    # program and version
    program_name = "gdc_vols"
    program_version = "0.3.3"

    # Год анализа. Если оставить 0, то берется текущий год
    process_year = 2022
    if process_year == 0:
        process_year = datetime.date.today().year

    # Месяц для анализа. Если оставить 0, то берется текущий месяц
    process_month = 2
    if process_month == 0:
        process_month = datetime.date.today().month

    # main variables
    urls = {f'Строительство гор.ВОЛС {process_year}': f'https://gdc-rts/api/test-table/vw_{process_year}_FOCL_Common_Build_City',
            f'Реконструкция гор.ВОЛС {process_year}': f'https://gdc-rts/api/test-table/vw_{process_year}_FOCL_Common_Rebuild_City',
            f'Строительство зон.ВОЛС {process_year}': f'https://gdc-rts/api/test-table/vw_{process_year}_FOCL_Common_Build_Zone',
            f'Реконструкция зон.ВОЛС {process_year}': f'https://gdc-rts/api/test-table/vw_{process_year}_FOCL_Common_Rebuild_Zone'
            }

    report_sheets = {'report': "Отчетная таблица",
                     'tz_build': 'Нет ТЗ Стр.',
                     'tz_reconstruction': 'Нет ТЗ Рек.',
                     'sending_po_build': "Нет передачи ТЗ Стр.",
                     'sending_po_reconstruction': "Нет передачи ТЗ Рек.",
                     'received_po_build': 'Не приняты ТЗ Стр.',
                     'received_po_reconstruction': 'Не приняты ТЗ Рек.',
                     'tz': 'Нет ТЗ',
                     'sending_po': "Нет передачи ТЗ",
                     'received_po': 'Не приняты ТЗ'
                     }

    excel_tables_names = {f'Строительство гор.ВОЛС {process_year}': "Urban_VOLS_Build",
                          f'Реконструкция гор.ВОЛС {process_year}': "Urban_VOLS_Reconstruction",
                          f'Строительство зон.ВОЛС {process_year}': "Zone_VOLS_Build",
                          f'Реконструкция зон.ВОЛС {process_year}': "Zone_VOLS_Reconstruction",
                          report_sheets['tz']: "tz_not_done",
                          report_sheets['sending_po']: "sending_po_not_done",
                          report_sheets['received_po']: "received_po_not_done",
                          }

    excel_cell_names = fill_cell_names()

    # Стиль таблицы Excel
    table_style = "TableStyleMedium2"
    # Наименования колонок для преобразования даты
    columns_dates = ['Планируемая дата окончания', 'Дата ввода', '_дата']
    # Наименования колонок для преобразования числа
    columns_digit = ['ID']
    # Наименование колонки для сортировки по возрастанию
    columns_for_sort = ['ID']
    work_branch = "Кавказский филиал"
    today_date = datetime.date.today().strftime("%Y%m%d")  # YYYYMMDD format today date
    vols_dir = f'y:\\Блок №4\\ВОЛС\\{process_year}\\'
#    vols_dir = f'.\\'
    vols_file = f'{today_date} Отчет по строительству и реконструкции ВОЛС {"".join(symbol[0].upper() for symbol in work_branch.split())} {process_year}.xlsx'
    file_name = f'{vols_dir}{vols_file}'
    id_branch = "Филиал"

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
                             'traffic_status2': 'Запуск трафика_Статус',
                             'region': 'Регион/Зона мероприятия'
                             }
    # Excel styles
    fn_bold = Font(bold=True)
    fn_red_bold = Font(color="FF0000", bold=True)
    fn_red = Font(color="B22222")
    fn_green = Font(color="006400")
    fn_mag = Font(color="6633FF")
    bd = Side(style='thick', color="000000")
    fill_red = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
    fill_yellow = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
    fill_green = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')
    align_center = Alignment(horizontal="center")
    border_medium = Border(left=Side(style=borders_style.BORDER_MEDIUM), right=Side(style=borders_style.BORDER_MEDIUM),
                           top=Side(style=borders_style.BORDER_MEDIUM), bottom=Side(style=borders_style.BORDER_MEDIUM))

    print(f'{program_name}: {program_version}')

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
        f'Generate report sheet: "{report_sheets["report"]}"')
    for i in range(1, 13):
        last_days_of_month[i] = pd.Timestamp(last_day_of_month(datetime.date(process_year, i, 1)))
    wb = openpyxl.load_workbook(filename=file_name)
    try:
        ws = wb[report_sheets['report']]
    except:
        ws = wb.create_sheet(title=report_sheets['report'])

    ws['A1'] = "Строительство городских ВОЛС"
    ws['A1'].font = fn_red_bold
    ws['A2'] = 'Всего мероприятий'
    ws['A2'].border = border_medium
    ws['A4'] = 'Исполнение KPI ВОЛС КФ (накопительный итог)'
    ws['A4'].font = fn_red_bold
    ws['A6'] = 'Учтенных ВОЛС в KPI'
    ws['A6'].border = border_medium
    ws['A8'] = 'Исполнение мероприятий в ЕСУП'
    ws['A8'].font = fn_red_bold
    ws['A9'] = 'Наименование мероприятия'
    ws['A9'].font = fn_bold
    ws['A9'].border = border_medium
    ws['A10'] = 'Выпущены ТЗ'
    ws['A10'].border = border_medium
    ws['A11'] = 'Переданы ТЗ в ПО'
    ws['A11'].border = border_medium
    ws['A12'] = 'Приняты ТЗ ПО'
    ws['A12'].border = border_medium
    ws['A13'] = 'Подписание договора на ПИР/ПИР+СМР'
    ws['A13'].border = border_medium
    ws['A14'] = 'Линейная схема'
    ws['A14'].border = border_medium
    ws['A15'] = 'Получено ТУ'
    ws['A15'].border = border_medium
    ws['A16'] = 'Строительство трассы'
    ws['A16'].border = border_medium
    ws['A17'] = 'Подготовка актов КС-2,3'
    ws['A17'].border = border_medium
    ws['A18'] = 'Приёмка ВОЛС в эксплуатацию'
    ws['A18'].border = border_medium
    ws['B9'] = 'Выполнено'
    ws['B9'].font = fn_bold
    ws['B9'].alignment = align_center
    ws['B9'].border = border_medium
    ws['C9'] = 'Осталось'
    ws['C9'].font = fn_bold
    ws['C9'].alignment = align_center
    ws['C9'].border = border_medium
    ws['F1'] = "Реконструкция городских ВОЛС"
    ws['F1'].font = fn_red_bold
    ws['F2'] = 'Всего мероприятий'
    ws['F2'].border = border_medium
    ws['F2'].border = border_medium
    ws['F4'] = 'Исполнение KPI ВОЛС КФ (накопительный итог)'
    ws['F4'].font = fn_red_bold
    ws['F6'] = 'Учтенных ВОЛС в KPI'
    ws['F6'].border = border_medium
    ws['F8'] = 'Исполнение мероприятий в ЕСУП'
    ws['F8'].font = fn_red_bold
    ws['F9'] = 'Наименование мероприятия'
    ws['F9'].font = fn_bold
    ws['F9'].border = border_medium
    ws['F10'] = 'Выпущены ТЗ'
    ws['F10'].border = border_medium
    ws['F11'] = 'Переданы ТЗ в ПО'
    ws['F11'].border = border_medium
    ws['F12'] = 'Приняты ТЗ ПО'
    ws['F12'].border = border_medium
    ws['F13'] = 'Подписание договора на ПИР/ПИР+СМР'
    ws['F13'].border = border_medium
    ws['F14'] = 'Линейная схема'
    ws['F14'].border = border_medium
    ws['F15'] = 'Получено ТУ'
    ws['F15'].border = border_medium
    ws['F16'] = 'Строительство трассы'
    ws['F16'].border = border_medium
    ws['F17'] = 'Подготовка актов КС-2,3'
    ws['F17'].border = border_medium
    ws['F18'] = 'Приёмка ВОЛС в эксплуатацию'
    ws['F18'].border = border_medium
    ws['G9'] = 'Выполнено'
    ws['G9'].font = fn_bold
    ws['G9'].alignment = align_center
    ws['G9'].border = border_medium
    ws['H9'] = 'Осталось'
    ws['H9'].font = fn_bold
    ws['H9'].alignment = align_center
    ws['H9'].border = border_medium
    ws['B5'] = f'План, {datetime.datetime(process_year, process_month, 1).strftime("%b %Y")}'
    ws['B5'].font = fn_bold
    ws['B5'].alignment = align_center
    ws['B5'].border = border_medium
    ws['C5'] = f'Факт, {datetime.datetime(process_year, process_month, 1).strftime("%b %Y")}'
    ws['C5'].font = fn_bold
    ws['C5'].alignment = align_center
    ws['C5'].border = border_medium
    ws['D5'] = f'{chr(0x0394)}, {datetime.datetime(process_year, process_month, 1).strftime("%b %Y")}'
    ws['D5'].font = fn_bold
    ws['D5'].alignment = align_center
    ws['D5'].border = border_medium
    ws['G5'] = f'План, {datetime.datetime(process_year, process_month, 1).strftime("%b %Y")}'
    ws['G5'].font = fn_bold
    ws['G5'].alignment = align_center
    ws['G5'].border = border_medium
    ws['H5'] = f'Факт, {datetime.datetime(process_year, process_month, 1).strftime("%b %Y")}'
    ws['H5'].font = fn_bold
    ws['H5'].alignment = align_center
    ws['H5'].border = border_medium
    ws['I5'] = f'{chr(0x0394)}, {datetime.datetime(process_year, process_month, 1).strftime("%b %Y")}'
    ws['I5'].font = fn_bold
    ws['I5'].alignment = align_center
    ws['I5'].border = border_medium

    # Анализ строительства ВОЛС
    dashboard_data = pd.read_excel(file_name, sheet_name=list(excel_tables_names.keys())[0])

    tz_build_dataframe = dashboard_data[dashboard_data[process_column_status['tz_status']] != 'Исполнена']
    sending_po_build_dataframe = dashboard_data[dashboard_data[process_column_status['send_tz_status']] != 'Исполнена']
    received_po_build_dataframe = dashboard_data[dashboard_data[process_column_status['received_tz_status']] != 'Исполнена']

    ws['B2'] = len(dashboard_data[process_columns_date['plan_date']])
    ws['B2'].font = fn_bold
    ws['B2'].alignment = align_center
    ws['B2'].border = border_medium
    ws['B6'] = sum_sort_month_events(dashboard_data, process_columns_date['plan_date'], process_month)
    ws['B6'].alignment = align_center
    ws['B6'].border = border_medium
    ws['C6'] = sum_done_events(dashboard_data, process_columns_date['ks2_date'], process_columns_date['commissioning_date'], process_column_status['ks2_status'], process_column_status['commissioning_status'], ['Исполнена'], process_month)
    ws['C6'].alignment = align_center
    ws['C6'].border = border_medium
    ws['D6'] = ws['C6'].value - ws['B6'].value
    ws['D6'].alignment = align_center
    ws['D6'].border = border_medium
    ws.conditional_formatting.add('D6', CellIsRule(operator='lessThan', formula=['0'], stopIfTrue=True, font=fn_red, fill=fill_red))
    ws.conditional_formatting.add('D6', CellIsRule(operator='greaterThan', formula=['0'], stopIfTrue=True, font=fn_green, fill=fill_green))
    ws.conditional_formatting.add('D6', CellIsRule(operator='equal', formula=['0'], stopIfTrue=True, font=fn_mag, fill=fill_yellow))

    for i, process in zip(range(10, 19), ['tz_status', 'send_tz_status', 'received_tz_status', 'pir_smr_status', 'line_scheme_status', 'tu_status', 'build_status', 'ks2_status', 'commissioning_status']):
        ws[f'B{i}'] = sum_sort_events(dashboard_data, process_column_status[process], ['Исполнена', 'Не требуется'])
        ws[f'B{i}'].alignment = align_center
        ws[f'B{i}'].border = border_medium
    for i in range(10, 19):
        # ws[f'C{i}'] = ws['B2'].value - ws[f'B{i}'].value
        ws[f'C{i}'] = f'=B2-B{i}'
        ws[f'C{i}'].alignment = align_center
        ws[f'C{i}'].border = border_medium
        ws.conditional_formatting.add(f'C{i}', CellIsRule(operator='greaterThan', formula=['0'], stopIfTrue=True, font=fn_red, fill=fill_red))
        ws.conditional_formatting.add(f'C{i}', CellIsRule(operator='lessThanOrEqual', formula=['0'], stopIfTrue=True, font=fn_green, fill=fill_green))

    # Анализ реконструкции ВОЛС
    dashboard_data = pd.read_excel(file_name, sheet_name=list(excel_tables_names.keys())[1])
    tz_reconstruction_dataframe = dashboard_data[dashboard_data[process_column_status['tz_status2']] != 'Исполнена']
    sending_po_reconstruction_dataframe = dashboard_data[dashboard_data[process_column_status['send_tz_status2']] != 'Исполнена']
    received_po_reconstruction_dataframe = dashboard_data[dashboard_data[process_column_status['received_tz_status2']] != 'Исполнена']

    ws['G2'] = len(dashboard_data[process_columns_date['plan_date']])
    ws['G2'].font = fn_bold
    ws['G2'].alignment = align_center
    ws['G2'].border = border_medium
    ws['G6'] = sum_sort_month_events(dashboard_data, process_columns_date['plan_date'], process_month)
    ws['G6'].alignment = align_center
    ws['G6'].border = border_medium
    ws['H6'] = sum_done_events(dashboard_data, process_columns_date['ks2_date2'], process_columns_date['commissioning_date2'], process_column_status['ks2_status2'], process_column_status['commissioning_status2'], ['Исполнена'], process_month)
    ws['H6'].alignment = align_center
    ws['H6'].border = border_medium
    ws['I6'] = ws['H6'].value - ws['G6'].value
    ws['I6'].alignment = align_center
    ws['I6'].border = border_medium
    ws.conditional_formatting.add('I6', CellIsRule(operator='lessThanOrEqual', formula=['0'], stopIfTrue=True, font=fn_red, fill=fill_red))
    ws.conditional_formatting.add('I6', CellIsRule(operator='greaterThan', formula=['0'], stopIfTrue=True, font=fn_green, fill=fill_green))
    ws.conditional_formatting.add('I6', CellIsRule(operator='equal', formula=['0'], stopIfTrue=True, font=fn_mag, fill=fill_yellow))

    for i, process in zip(range(10, 19), ['tz_status2', 'send_tz_status2', 'received_tz_status2', 'pir_smr_status2', 'line_scheme_status2', 'tu_status2', 'build_status2', 'ks2_status2', 'commissioning_status2']):
        ws[f'G{i}'] = sum_sort_events(dashboard_data, process_column_status[process], ['Исполнена', 'Не требуется'])
        ws[f'G{i}'].alignment = align_center
        ws[f'G{i}'].border = border_medium
        ws[f'G{i}'].border = border_medium
    for i in range(10, 19):
        ws[f'H{i}'] = f'=G2-G{i}'
        ws[f'H{i}'].alignment = align_center
        ws[f'H{i}'].border = border_medium
        ws.conditional_formatting.add(f'H{i}', CellIsRule(operator='greaterThan', formula=['0'], stopIfTrue=True, font=fn_red, fill=fill_red))
        ws.conditional_formatting.add(f'H{i}', CellIsRule(operator='lessThanOrEqual', formula=['0'], stopIfTrue=True, font=fn_green, fill=fill_green))
    ws = adjust_columns_width(ws)

    print(
        f'Write "{report_sheets["report"]}" sheets to file: "{file_name}"')
    wb.save(file_name)

    # Создание листов для рассылки
    #
    # Объединяем ТЗ стройки и реконструкции первые 5 полей
    tz_dataframe = pd.concat([tz_build_dataframe.iloc[:, :4], tz_reconstruction_dataframe.iloc[:, :4]], ignore_index=True).reset_index(drop=True)
    write_report_table_to_file(tz_dataframe, file_name, report_sheets['tz'], excel_tables_names)

    # Объединяем передачу ТЗ стройки и реконструкции первые 5 полей
    sending_po_dataframe = pd.concat([sending_po_build_dataframe.iloc[:, :4], sending_po_reconstruction_dataframe.iloc[:, :4]], ignore_index=True).reset_index(drop=True)
    # Убираем мероприятия с не выданными ТЗ
    sending_po_dataframe = pd.concat([sending_po_dataframe, tz_dataframe], ignore_index=True).drop_duplicates(keep=False).reset_index(drop=True)
    write_report_table_to_file(sending_po_dataframe, file_name, report_sheets['sending_po'], excel_tables_names)

    # Объединяем прием ТЗ стройки и реконструкции первые 5 полей
    received_po_dataframe = pd.concat([received_po_build_dataframe.iloc[:, :4], received_po_reconstruction_dataframe.iloc[:, :4]], ignore_index=True).reset_index(drop=True)
    # Убираем мероприятия с не выданными ТЗ и не переданные в ПО
    received_po_dataframe = pd.concat([received_po_dataframe, sending_po_dataframe, tz_dataframe], ignore_index=True).drop_duplicates(keep=False).reset_index(drop=True)
    write_report_table_to_file(received_po_dataframe, file_name, report_sheets['received_po'], excel_tables_names)
