#  Copyright (c) 2022. Tikhon Ostapenko
import argparse
import locale
import os

from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, Side, PatternFill, Alignment, Border
import openpyxl.styles.borders as borders_style
from vols_functions import *


if __name__ == '__main__':
    # program and version
    program_name = "gdc_vols"
    program_version = "0.4.8"

    # Стиль таблицы Excel
    table_style = "TableStyleMedium2"
    # Наименования колонок для преобразования даты
    columns_date = ['Планируемая дата окончания', 'Дата ввода', '_дата']
    # Наименования колонок для преобразования числа
    columns_digit = ['ID']
    # Наименование колонки для сортировки по возрастанию
    columns_for_sort = ['Регион/Зона мероприятия']
    work_branch = "Кавказский филиал"
    today_date = datetime.date.today().strftime("%Y%m%d")  # YYYYMMDD format today date
    last_days_of_month = {}

    # Set Russian localization
    locale.setlocale(locale.LC_TIME, "ru_RU")

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
    border_thin = Border(left=Side(style=borders_style.BORDER_THIN), right=Side(style=borders_style.BORDER_THIN),
                         top=Side(style=borders_style.BORDER_THIN), bottom=Side(style=borders_style.BORDER_THIN))

    # Parse command line arguments
    parser = argparse.ArgumentParser(description=f'{program_name} v.{program_version}')
    parser.add_argument("-y", "--year", type=int, help="year for processing")
    parser.add_argument("-m", "--month", type=int, help="month for processing")
    parser.add_argument("-r", "--report-file", help="report file name, must have .xlsx extension")
    parser.add_argument("-b", "--report-branch", help="Branch name", default=work_branch)
    parser.add_argument("-o", "--report-only", help="Don't get new online data. Generate report only", action='store_true')
    args = parser.parse_args()

    # Год анализа.
    if args.year is None:
        process_year = datetime.date.today().year
    else:
        process_year = args.year

        # Месяц для анализа.
    if args.month is None:
        process_month = datetime.date.today().month
    else:
        process_month = args.month

    if args.report_branch is not None:
        work_branch = args.report_branch

    if args.report_file is None:
        vols_dir = f'\\\\megafon.ru\\KVK\\KRN\\Files\\TelegrafFiles\\ОПРС\\!Проекты РЦРП\\Блок №4\\ВОЛС\\{process_year}'
        vols_file = f'{today_date} Отчет по строительству и реконструкции ВОЛС {"".join(symbol[0].upper() for symbol in work_branch.split())} {datetime.date(process_year, process_month, 1).strftime("%m.%Y")}.xlsx'
        file_name = f'{vols_dir}\\{vols_file}'
    else:
        file_name = args.report_file

    urls = {
        f'Расш. стр. гор.ВОЛС {process_year}': f'https://gdc-rts/api/test-table/vw_{process_year}_FOCL_Common_Build_City_211',
        f'Реконструкция гор.ВОЛС {process_year}': f'https://gdc-rts/api/test-table/vw_{process_year}_FOCL_Common_Rebuild_City',
        f'Строительство зон.ВОЛС {process_year}': f'https://gdc-rts/api/test-table/vw_{process_year}_FOCL_Common_Build_Zone',
        f'Реконструкция зон.ВОЛС {process_year}': f'https://gdc-rts/api/test-table/vw_{process_year}_FOCL_Common_Rebuild_Zone'}

    data_sheets = {'city_main_build': f'Осн. стр. гор.ВОЛС {process_year}',
                   'city_ext_build': f'Доп. стр. гор.ВОЛС {process_year}',
                   'city_reconstruction': f'Реконструкция гор.ВОЛС {process_year}',
                   'zone_build':  f'Строительство зон.ВОЛС {process_year}',
                   'zone_reconstruction': f'Реконструкция зон.ВОЛС {process_year}'}

    report_sheets = {'report': "Отчетная таблица",
                     'current_month': f'Активные мероприятия {datetime.date(process_year, process_month, 1).strftime("%m.%Y")}',
                     'tz': 'Нет ТЗ',
                     'sending_po': "Нет передачи ТЗ",
                     'received_po': 'Не приняты ТЗ'}

    excel_tables_names = {data_sheets['city_main_build']: "Urban_VOLS_Main_Build",
                          data_sheets['city_ext_build']: "Urban_VOLS_Ext_Build",
                          data_sheets['city_reconstruction']: "Urban_VOLS_Reconstruction",
                          data_sheets['zone_build']: "Zone_VOLS_Build",
                          data_sheets['zone_reconstruction']: "Zone_VOLS_Reconstruction",
                          report_sheets['current_month']: "current_month",
                          report_sheets['tz']: "tz_not_done",
                          report_sheets['sending_po']: "sending_po_not_done",
                          report_sheets['received_po']: "received_po_not_done"}

    excel_cell_names = fill_cell_names()

    process_columns = {'plan_date': 'Планируемая дата окончания',
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
                       'traffic_date2': 'Запуск трафика_Дата',
                       'tz_status': 'Разработка ТЗ_статус',
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
                       'id': 'ID',
                       'branch': 'Филиал',
                       'region': 'Регион/Зона мероприятия',
                       'name': 'Название',
                       'program': 'Программы'}

    print(f'{program_name}: {program_version}')

    # Получение исходных данных и запись форматированных данных

    # get_report - определяет получать ли внешние данные

    if not args.report_only:
        if Path(file_name).is_file():
            print(f'Remove old file {file_name}')
            os.remove(file_name)

        for sheet, url in urls.items():
            data_frame = read_from_dashboard(url)  # Читаем данные из сети
            data_frame = data_frame[data_frame[process_columns['branch']] == work_branch]  # Оставляем только отчётный филиал
            data_frame = data_frame.reset_index(drop=True)
            data_frame = convert_date(data_frame, columns_date)  # Переводим дату в формат datetime
            data_frame = convert_int(data_frame, columns_digit)  # Переводим ESUP_ID в числовой формат
            data_frame = data_frame.sort_values(by=columns_for_sort)  # Сортируем по заданному столбцу
            if sheet == f'Расш. стр. гор.ВОЛС {process_year}':
                extended_build_df = data_frame.copy(deep=True)  # keep extended data for analyses
                # Формируем таблицу основного строительства
                main_build_df = data_frame[data_frame['KPI ПТР текущего года, км'].notnull()]
                write_dataframe_to_file(main_build_df, file_name, data_sheets['city_main_build'])
                format_table(main_build_df, data_sheets['city_main_build'], file_name, excel_tables_names, excel_cell_names, table_style)
                # Формируем таблицу дополнительного строительства
                ext_build_df = data_frame[~data_frame['KPI ПТР текущего года, км'].notnull()]
                write_dataframe_to_file(ext_build_df, file_name, data_sheets['city_ext_build'])
                format_table(ext_build_df, data_sheets['city_ext_build'], file_name, excel_tables_names, excel_cell_names, table_style)
            else:
                write_dataframe_to_file(data_frame, file_name, sheet)
                format_table(data_frame, sheet, file_name, excel_tables_names, excel_cell_names, table_style)

    # Создание отчёта
    print(f'Generate report sheet: "{report_sheets["report"]}"')
    for i in range(1, 13):
        last_days_of_month[i] = pd.Timestamp(last_day_of_month(datetime.date(process_year, i, 1)))

    wb = openpyxl.load_workbook(filename=file_name)
    try:
        ws = wb[report_sheets['report']]
    except Exception:
        ws = wb.create_sheet(title=report_sheets['report'])

    # Формирование статических полей отчёта
    ws['A1'] = "Основное строительство ВОЛС"
    ws['A1'].font = fn_red_bold
    ws['A1'].border = border_thin
    ws['A2'] = 'Всего мероприятий'
    ws['A2'].border = border_medium
    ws['A4'] = 'Исполнение KPI ВОЛС КФ (накопительный итог)'
    ws['A4'].font = fn_red_bold
    ws['A4'].border = border_thin
    ws['A6'] = 'Учтенных ВОЛС в KPI'
    ws['A6'].border = border_medium
    ws['A8'] = 'Исполнение мероприятий в ЕСУП'
    ws['A8'].font = fn_red_bold
    ws['A8'].border = border_thin
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
    ws['C9'] = f'{chr(0x0394)}'
    ws['C9'].font = fn_bold
    ws['C9'].alignment = align_center
    ws['C9'].border = border_medium

    ws['A21'] = "Дополнительное строительство ВОЛС"
    ws['A21'].font = fn_red_bold
    ws['A21'].border = border_thin
    ws['A22'] = 'Всего мероприятий'
    ws['A22'].border = border_medium
    ws['A24'] = 'Исполнение KPI ВОЛС КФ (накопительный итог)'
    ws['A24'].font = fn_red_bold
    ws['A24'].border = border_thin
    ws['A26'] = 'Учтенных ВОЛС в KPI'
    ws['A26'].border = border_medium
    ws['A28'] = 'Исполнение мероприятий в ЕСУП'
    ws['A28'].font = fn_red_bold
    ws['A28'].border = border_thin
    ws['A29'] = 'Наименование мероприятия'
    ws['A29'].font = fn_bold
    ws['A29'].border = border_medium
    ws['A30'] = 'Выпущены ТЗ'
    ws['A30'].border = border_medium
    ws['A31'] = 'Переданы ТЗ в ПО'
    ws['A31'].border = border_medium
    ws['A32'] = 'Приняты ТЗ ПО'
    ws['A32'].border = border_medium
    ws['A33'] = 'Подписание договора на ПИР/ПИР+СМР'
    ws['A33'].border = border_medium
    ws['A34'] = 'Линейная схема'
    ws['A34'].border = border_medium
    ws['A35'] = 'Получено ТУ'
    ws['A35'].border = border_medium
    ws['A36'] = 'Строительство трассы'
    ws['A36'].border = border_medium
    ws['A37'] = 'Подготовка актов КС-2,3'
    ws['A37'].border = border_medium
    ws['A38'] = 'Приёмка ВОЛС в эксплуатацию'
    ws['A38'].border = border_medium
    ws['B29'] = 'Выполнено'
    ws['B29'].font = fn_bold
    ws['B29'].alignment = align_center
    ws['B29'].border = border_medium
    ws['C29'] = f'{chr(0x0394)}'
    ws['C29'].font = fn_bold
    ws['C29'].alignment = align_center
    ws['C29'].border = border_medium

    ws['F1'] = "Реконструкция ВОЛС"
    ws['F1'].font = fn_red_bold
    ws['F1'].border = border_thin
    ws['F2'] = 'Всего мероприятий'
    ws['F2'].border = border_medium
    ws['F2'].border = border_medium
    ws['F4'] = 'Исполнение KPI ВОЛС КФ (накопительный итог)'
    ws['F4'].font = fn_red_bold
    ws['F4'].border = border_thin
    ws['F6'] = 'Учтенных ВОЛС в KPI'
    ws['F6'].border = border_medium
    ws['F8'] = 'Исполнение мероприятий в ЕСУП'
    ws['F8'].font = fn_red_bold
    ws['F8'].border = border_thin
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
    ws['H9'] = f'{chr(0x0394)}'
    ws['H9'].font = fn_bold
    ws['H9'].alignment = align_center
    ws['H9'].border = border_medium

    # Формирование динамических полей отчёта
    ws['B5'] = f'План, {datetime.date(process_year, process_month, 1).strftime("%b %Y")}'
    ws['B5'].font = fn_bold
    ws['B5'].alignment = align_center
    ws['B5'].border = border_medium
    ws['C5'] = f'Факт, {datetime.date(process_year, process_month, 1).strftime("%b %Y")}'
    ws['C5'].font = fn_bold
    ws['C5'].alignment = align_center
    ws['C5'].border = border_medium
    ws['D5'] = f'{chr(0x0394)}, {datetime.date(process_year, process_month, 1).strftime("%b %Y")}'
    ws['D5'].font = fn_bold
    ws['D5'].alignment = align_center
    ws['D5'].border = border_medium

    ws['B25'] = f'План, {datetime.date(process_year, process_month, 1).strftime("%b %Y")}'
    ws['B25'].font = fn_bold
    ws['B25'].alignment = align_center
    ws['B25'].border = border_medium
    ws['C25'] = f'Факт, {datetime.date(process_year, process_month, 1).strftime("%b %Y")}'
    ws['C25'].font = fn_bold
    ws['C25'].alignment = align_center
    ws['C25'].border = border_medium
    ws['D25'] = f'{chr(0x0394)}, {datetime.date(process_year, process_month, 1).strftime("%b %Y")}'
    ws['D25'].font = fn_bold
    ws['D25'].alignment = align_center
    ws['D25'].border = border_medium

    ws['G5'] = f'План, {datetime.date(process_year, process_month, 1).strftime("%b %Y")}'
    ws['G5'].font = fn_bold
    ws['G5'].alignment = align_center
    ws['G5'].border = border_medium
    ws['H5'] = f'Факт, {datetime.date(process_year, process_month, 1).strftime("%b %Y")}'
    ws['H5'].font = fn_bold
    ws['H5'].alignment = align_center
    ws['H5'].border = border_medium
    ws['I5'] = f'{chr(0x0394)}, {datetime.date(process_year, process_month, 1).strftime("%b %Y")}'
    ws['I5'].font = fn_bold
    ws['I5'].alignment = align_center
    ws['I5'].border = border_medium

    # Анализ строительства ВОЛС
    if not args.report_only:
        dashboard_data = extended_build_df
    else:
        print(f'Read "{data_sheets["city_main_build"]}" sheet from file: "{file_name}"')
        df_main_build = pd.read_excel(file_name, sheet_name=data_sheets['city_main_build'])
        print(f'Read "{data_sheets["city_ext_build"]}" sheet from file: "{file_name}"')
        df_ext_build = pd.read_excel(file_name, sheet_name=data_sheets['city_ext_build'])
        dashboard_data = pd.concat([df_main_build, df_ext_build])

    build_dashboard_data = dashboard_data
    tz_build_dataframe = dashboard_data[dashboard_data[process_columns['tz_status']] != 'Исполнена']
    sending_po_build_dataframe = dashboard_data[dashboard_data[process_columns['send_tz_status']] != 'Исполнена']
    received_po_build_dataframe = dashboard_data[dashboard_data[process_columns['received_tz_status']] != 'Исполнена']

    main_build_df = dashboard_data[dashboard_data['KPI ПТР текущего года, км'].notnull()]
    ext_build_df = dashboard_data[~dashboard_data['KPI ПТР текущего года, км'].notnull()]

    ws['B2'] = main_build_df[process_columns['plan_date']].count()
    ws['B2'].font = fn_bold
    ws['B2'].alignment = align_center
    ws['B2'].border = border_medium
    ws['B6'] = sum_sort_month_events(main_build_df, process_columns['plan_date'], process_month, last_days_of_month)
    ws['B6'].alignment = align_center
    ws['B6'].border = border_medium
    ws['C6'] = sum_done_events(main_build_df, process_columns['ks2_date'],
                               process_columns['commissioning_date'], process_columns['ks2_status'],
                               process_columns['commissioning_status'], ['Исполнена'], process_month, last_days_of_month)
    ws['C6'].alignment = align_center
    ws['C6'].border = border_medium
    ws['D6'] = ws['C6'].value - ws['B6'].value
    ws['D6'].alignment = align_center
    ws['D6'].border = border_medium
    ws.conditional_formatting.add('D6', CellIsRule(operator='lessThan', formula=['0'], stopIfTrue=True, font=fn_red,
                                                   fill=fill_red))
    ws.conditional_formatting.add('D6',
                                  CellIsRule(operator='greaterThan', formula=['0'], stopIfTrue=True, font=fn_green,
                                             fill=fill_green))
    ws.conditional_formatting.add('D6', CellIsRule(operator='equal', formula=['0'], stopIfTrue=True, font=fn_mag,
                                                   fill=fill_yellow))

    ws['B22'] = ext_build_df[process_columns['plan_date']].count()
    ws['B22'].font = fn_bold
    ws['B22'].alignment = align_center
    ws['B22'].border = border_medium
    ws['B26'] = sum_sort_month_events(ext_build_df, process_columns['plan_date'], process_month, last_days_of_month)
    ws['B26'].alignment = align_center
    ws['B26'].border = border_medium
    ws['C26'] = sum_done_events(ext_build_df, process_columns['ks2_date'],
                                process_columns['commissioning_date'], process_columns['ks2_status'],
                                process_columns['commissioning_status'], ['Исполнена'], process_month, last_days_of_month)
    ws['C26'].alignment = align_center
    ws['C26'].border = border_medium
    ws['D26'] = ws['C26'].value - ws['B26'].value
    ws['D26'].alignment = align_center
    ws['D26'].border = border_medium
    ws.conditional_formatting.add('D26', CellIsRule(operator='lessThan', formula=['0'], stopIfTrue=True, font=fn_red,
                                                    fill=fill_red))
    ws.conditional_formatting.add('D26',
                                  CellIsRule(operator='greaterThan', formula=['0'], stopIfTrue=True, font=fn_green,
                                             fill=fill_green))
    ws.conditional_formatting.add('D26', CellIsRule(operator='equal', formula=['0'], stopIfTrue=True, font=fn_mag,
                                                    fill=fill_yellow))

    for i, process in zip(range(10, 19),
                          ['tz_status', 'send_tz_status', 'received_tz_status', 'pir_smr_status', 'line_scheme_status',
                           'tu_status', 'build_status', 'ks2_status', 'commissioning_status']):
        ws[f'B{i}'] = sum_sort_events(main_build_df, process_columns[process], ['Исполнена', 'Не требуется'])
        ws[f'B{i}'].alignment = align_center
        ws[f'B{i}'].border = border_medium
    for i in range(10, 19):
        ws[f'C{i}'] = ws[f'B{i}'].value - ws['B2'].value
        ws[f'C{i}'].alignment = align_center
        ws[f'C{i}'].border = border_medium
        ws.conditional_formatting.add(f'C{i}',
                                      CellIsRule(operator='greaterThanOrEqual', formula=['0'], stopIfTrue=True, font=fn_green,
                                                 fill=fill_green))
        ws.conditional_formatting.add(f'C{i}', CellIsRule(operator='lessThan', formula=['0'], stopIfTrue=True,
                                                          font=fn_red, fill=fill_red))

    for i, process in zip(range(30, 39),
                          ['tz_status', 'send_tz_status', 'received_tz_status', 'pir_smr_status', 'line_scheme_status',
                           'tu_status', 'build_status', 'ks2_status', 'commissioning_status']):
        ws[f'B{i}'] = sum_sort_events(ext_build_df, process_columns[process], ['Исполнена', 'Не требуется'])
        ws[f'B{i}'].alignment = align_center
        ws[f'B{i}'].border = border_medium
    for i in range(30, 39):
        ws[f'C{i}'] = ws[f'B{i}'].value - ws['B22'].value
        ws[f'C{i}'].alignment = align_center
        ws[f'C{i}'].border = border_medium
        ws.conditional_formatting.add(f'C{i}',
                                      CellIsRule(operator='greaterThanOrEqual', formula=['0'], stopIfTrue=True, font=fn_green,
                                                 fill=fill_green))
        ws.conditional_formatting.add(f'C{i}', CellIsRule(operator='lessThan', formula=['0'], stopIfTrue=True,
                                                          font=fn_red, fill=fill_red))

    # Анализ реконструкции ВОЛС
    print(f'Read "{data_sheets["city_reconstruction"]}" sheet from file: "{file_name}"')
    dashboard_data = pd.read_excel(file_name, sheet_name=data_sheets['city_reconstruction'])
    reconstruction_dashboard_data = dashboard_data
    tz_reconstruction_dataframe = dashboard_data[dashboard_data[process_columns['tz_status2']] != 'Исполнена']
    sending_po_reconstruction_dataframe = dashboard_data[
        dashboard_data[process_columns['send_tz_status2']] != 'Исполнена']
    received_po_reconstruction_dataframe = dashboard_data[
        dashboard_data[process_columns['received_tz_status2']] != 'Исполнена']

    ws['G2'] = dashboard_data[process_columns['plan_date']].count()
    ws['G2'].font = fn_bold
    ws['G2'].alignment = align_center
    ws['G2'].border = border_medium
    ws['G6'] = sum_sort_month_events(dashboard_data, process_columns['plan_date'], process_month, last_days_of_month)
    ws['G6'].alignment = align_center
    ws['G6'].border = border_medium
    ws['H6'] = sum_done_events(dashboard_data, process_columns['ks2_date2'],
                               process_columns['commissioning_date2'], process_columns['ks2_status2'],
                               process_columns['commissioning_status2'], ['Исполнена'], process_month, last_days_of_month)
    ws['H6'].alignment = align_center
    ws['H6'].border = border_medium
    ws['I6'] = ws['H6'].value - ws['G6'].value
    ws['I6'].alignment = align_center
    ws['I6'].border = border_medium
    ws.conditional_formatting.add('I6',
                                  CellIsRule(operator='lessThanOrEqual', formula=['0'], stopIfTrue=True, font=fn_red,
                                             fill=fill_red))
    ws.conditional_formatting.add('I6',
                                  CellIsRule(operator='greaterThan', formula=['0'], stopIfTrue=True, font=fn_green,
                                             fill=fill_green))
    ws.conditional_formatting.add('I6', CellIsRule(operator='equal', formula=['0'], stopIfTrue=True, font=fn_mag,
                                                   fill=fill_yellow))

    for i, process in zip(range(10, 19), ['tz_status2', 'send_tz_status2', 'received_tz_status2', 'pir_smr_status2',
                                          'line_scheme_status2', 'tu_status2', 'build_status2', 'ks2_status2',
                                          'commissioning_status2']):
        ws[f'G{i}'] = sum_sort_events(dashboard_data, process_columns[process], ['Исполнена', 'Не требуется'])
        ws[f'G{i}'].alignment = align_center
        ws[f'G{i}'].border = border_medium
        ws[f'G{i}'].border = border_medium
    for i in range(10, 19):
        ws[f'H{i}'] = ws[f'G{i}'].value - ws['G2'].value
        # ws[f'H{i}'] = f'=G{i}-G2'
        ws[f'H{i}'].alignment = align_center
        ws[f'H{i}'].border = border_medium
        ws.conditional_formatting.add(f'H{i}',
                                      CellIsRule(operator='greaterThanOrEqual', formula=['0'], stopIfTrue=True, font=fn_green,
                                                 fill=fill_green))
        ws.conditional_formatting.add(f'H{i}', CellIsRule(operator='lessThan', formula=['0'], stopIfTrue=True,
                                                          font=fn_red, fill=fill_red))
    ws = adjust_columns_width(ws)

    print(f'Write "{report_sheets["report"]}" sheets to file: "{file_name}"')
    wb.save(file_name)

    # Создание листов для рассылки

    # Создание листа Активные мероприятия строительства месяца отчёта
    # маска для текущего месяца
    curr_month_bool_mask = (build_dashboard_data[process_columns['plan_date']] <= last_days_of_month[process_month].strftime('%Y-%m-%d')) & (build_dashboard_data[process_columns['plan_date']] >= datetime.datetime(process_year, process_month, 1).strftime('%Y-%m-%d'))
    # маска для не "Исполнено" или не "Не требуется"
    curr_status_bool_mask = (~build_dashboard_data[process_columns['commissioning_status']].str.contains('Исполнено|Не требуется', regex=True)) & (~build_dashboard_data[process_columns['ks2_status']].str.contains('Исполнено|Не требуется', regex=True))
    # Выборка объектов строительства по маскам
    current_month_build_dataframe = build_dashboard_data[curr_month_bool_mask & curr_status_bool_mask]
    current_month_build_dataframe = current_month_build_dataframe[[process_columns['id'],
                                                                   process_columns['branch'],
                                                                   process_columns['region'],
                                                                   process_columns['name'],
                                                                   process_columns['program'],
                                                                   process_columns['plan_date']]]
    current_month_build_dataframe['БП'] = 'Строительство ВОЛС'

    # маска для текущего месяца
    curr_month_bool_mask = (reconstruction_dashboard_data[process_columns['plan_date']] <= last_days_of_month[process_month].strftime('%Y-%m-%d')) & (reconstruction_dashboard_data[process_columns['plan_date']] >= datetime.datetime(process_year, process_month, 1).strftime('%Y-%m-%d'))
    # маска для не "Исполнено" или не "Не требуется"
    curr_status_bool_mask = (~reconstruction_dashboard_data[process_columns['commissioning_status2']].str.contains('Исполнено|Не требуется', regex=True)) & (~reconstruction_dashboard_data[process_columns['ks2_status2']].str.contains('Исполнено|Не требуется', regex=True))
    # Выборка объектов реконструкции по маскам
    current_month_reconstruction_dataframe = reconstruction_dashboard_data[curr_month_bool_mask & curr_status_bool_mask]
    current_month_reconstruction_dataframe = current_month_reconstruction_dataframe[[process_columns['id'],
                                                                                     process_columns['branch'],
                                                                                     process_columns['region'],
                                                                                     process_columns['name'],
                                                                                     process_columns['program'],
                                                                                     process_columns['plan_date']]]
    current_month_reconstruction_dataframe['БП'] = 'Реконструкция ВОЛС'  # Добавляем столбец с бизнес-процессом

    # Объединяем стройку и реконструкцию
    current_month_dataframe = pd.concat([current_month_build_dataframe, current_month_reconstruction_dataframe], ignore_index=True).reset_index(drop=True).sort_values(by=columns_for_sort)
    write_report_table_to_file(current_month_dataframe, file_name, report_sheets['current_month'], excel_tables_names, excel_cell_names, table_style)

    # Создание листа Нет ТЗ
    # Формируем таблицы ТЗ для стройки и реконструкции
    tz_build_dataframe = tz_build_dataframe[[process_columns['id'],
                                             process_columns['branch'],
                                             process_columns['region'],
                                             process_columns['name'],
                                             process_columns['program'],
                                             process_columns['plan_date']]]
    tz_build_dataframe['БП'] = 'Строительство ВОЛС'
    tz_reconstruction_dataframe = tz_reconstruction_dataframe[[process_columns['id'],
                                                               process_columns['branch'],
                                                               process_columns['region'],
                                                               process_columns['name'],
                                                               process_columns['program'],
                                                               process_columns['plan_date']]]
    tz_reconstruction_dataframe['БП'] = 'Реконструкция ВОЛС'
    # Объединяем ТЗ стройки и реконструкции
    tz_dataframe = pd.concat([tz_build_dataframe, tz_reconstruction_dataframe], ignore_index=True).reset_index(drop=True).sort_values(by=columns_for_sort)
    write_report_table_to_file(tz_dataframe, file_name, report_sheets['tz'], excel_tables_names, excel_cell_names, table_style)

    # Создание листа Не переданы ТЗ в ПО
    # Формируем таблицы передачи в ПО для стройки и реконструкции
    sending_po_build_dataframe = sending_po_build_dataframe[[process_columns['id'],
                                                             process_columns['branch'],
                                                             process_columns['region'],
                                                             process_columns['name'],
                                                             process_columns['program'],
                                                             process_columns['plan_date']]]
    sending_po_build_dataframe['БП'] = 'Строительство ВОЛС'
    sending_po_reconstruction_dataframe = sending_po_reconstruction_dataframe[[process_columns['id'],
                                                                               process_columns['branch'],
                                                                               process_columns['region'],
                                                                               process_columns['name'],
                                                                               process_columns['program'],
                                                                               process_columns['plan_date']]]
    sending_po_reconstruction_dataframe['БП'] = 'Реконструкция ВОЛС'
    # Объединяем передачу ТЗ в ПО стройки и реконструкции
    sending_po_dataframe = pd.concat([sending_po_build_dataframe, sending_po_reconstruction_dataframe], ignore_index=True).reset_index(drop=True)
    # Убираем мероприятия с не выданными ТЗ
    sending_po_dataframe = pd.concat([sending_po_dataframe, tz_dataframe], ignore_index=True).drop_duplicates(keep=False).reset_index(drop=True).sort_values(by=columns_for_sort)
    write_report_table_to_file(sending_po_dataframe, file_name, report_sheets['sending_po'], excel_tables_names, excel_cell_names, table_style)

    # Создание листа ТЗ не принято ПО
    # Формируем таблицы не принято ПО для стройки и реконструкции
    received_po_build_dataframe = received_po_build_dataframe[[process_columns['id'], process_columns['branch'], process_columns['region'], process_columns['name'], process_columns['program'], process_columns['plan_date']]]
    received_po_build_dataframe['БП'] = 'Строительство ВОЛС'
    received_po_reconstruction_dataframe = received_po_reconstruction_dataframe[[process_columns['id'],
                                                                                 process_columns['branch'],
                                                                                 process_columns['region'],
                                                                                 process_columns['name'],
                                                                                 process_columns['program'],
                                                                                 process_columns['plan_date']]]
    received_po_reconstruction_dataframe['БП'] = 'Реконструкция ВОЛС'
    # Объединяем не принято в ПО стройки и реконструкции
    received_po_dataframe = pd.concat([received_po_build_dataframe, received_po_reconstruction_dataframe], ignore_index=True).reset_index(drop=True)
    # Убираем мероприятия с не выданными ТЗ и не переданные в ПО
    received_po_dataframe = pd.concat([received_po_dataframe, sending_po_dataframe, tz_dataframe], ignore_index=True).drop_duplicates(keep=False).reset_index(drop=True).sort_values(by=columns_for_sort)
    write_report_table_to_file(received_po_dataframe, file_name, report_sheets['received_po'], excel_tables_names, excel_cell_names, table_style)
