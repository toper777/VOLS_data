# Import section
import sys
from datetime import date

# import argparse and other for command-line arguments parser
import argparse

# import pandas for data operation
import pandas as pd

# program and version
program_name = "VOLS_data"
program_version = "0.0.4"

# main variables
today_date = date.today().strftime("%Y%m%d")
data_dir = "y:/Блок №3/2022 год/"
data_file = "!!!SQL Блок№3!!!  2022.xlsm"
data_sheet = "Массив"

# vols_dir = "c:/tmp/"
vols_dir = "y:/Блок №3/2022 год/"
vols_file_base = "VOLS KVK 2022.xlsx"
vols_file = f'{today_date} {vols_file_base}'
vols_sheet = "ВОЛС Кавказ 2022"
vols_q_sheet_begin = "ВОЛС "
vols_q_sheet_end = " кв."
vols_program = "Строительство ВОЛС (городская)"

work_program = "BP_ESUP"
work_ro = "RO"
work_prognoz_date = "PROGNOZ_DATE"
work_index = "ID_ESUP"

caucasian_region = ""


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


# Caucasian regions
regions = {
    "БО": "Белгородская область",
    "ВО": "Воронежская область",
    "КБР": "Кабардино-Балкарская республика",
    "КЧР": "Карачаево-Черкесская республика",
    "КК": "Краснодарский край",
    "ЛО": "Липецкая область",
    "РА": "Республика Адыгея",
    "РД": "Республика Дагестан",
    "РИ": "Республика Ингушетия",
    "РСО-А": "Республика Северная Осетия-Алания",
    "РО": "Ростовская область",
    "Сочи": "Сочи",
    "СК": "Ставропольский край",
    "ТО": "Тамбовская область",
    "ЧР": "Чеченская республика"
}

quarters = {
    "1_begin": "2022-01-01",
    "1_end": "2022-03-31",
    "2_begin": "2022-04-01",
    "2_end": "2022-06-30",
    "3_begin": "2022-07-01",
    "3_end": "2022-09-30",
    "4_begin": "2022-10-01",
    "4_end": "2022-12-31",
}


# command-line argument(s) parser
def create_parser():
    parser = argparse.ArgumentParser(
        prog=program_name,
        description='''Обработка данных по строительству ВОЛС''',
        epilog=f'{program_name} {program_version} (c) Tikhon Ostapenko 2021, 20222'
    )
    subparsers = parser.add_subparsers(dest='command')

    hello_parser = subparsers.add_parser('sql')
    hello_parser.add_argument('--sql-directory', '-s', default=[data_dir])
    hello_parser.add_argument('--sql-file', '-n', default=[data_file])

#    goodbye_parser = subparsers.add_parser('gdc')
#    goodbye_parser.add_argument('--gdc-directory', '-g', default=[gdc_dir])
#    goodbye_parser.add_argument('--gdc-file', '-g', default=[gdc_file])

    return parser


if __name__ == '__main__':

    # command-line parameters parser
    parser = create_parser()
    namespace = parser.parse_args(sys.argv[1:])
    print(namespace)

    # read data from file
    print(f'Read data file: {Color.CYAN}"{data_dir}{data_file}"{Color.END}')
    kvk_data = pd.read_excel(data_dir + data_file, sheet_name=data_sheet, index_col=work_index)

    # sorting for VOLS
    print(f'Sorting VOLS entity for {Color.CYAN}"{vols_program}"{Color.END}')
    vols = kvk_data[kvk_data[work_program] == vols_program]

    # write to file
    with pd.ExcelWriter(vols_dir + vols_file) as writer:
        print(f'Writing {Color.GREEN}"{vols_sheet}"{Color.END} sheet to file: {Color.CYAN}"{vols_dir}{vols_file}"{Color.END}')
        vols.to_excel(writer, sheet_name=vols_sheet)  # write YEAR sheet to file

        for i in range(1, 5):
            # sorting VOLS to quarters
            vols_q = vols[(vols[work_prognoz_date] > quarters[f'{i}_begin']) & (vols[work_prognoz_date] <= quarters[f'{i}_end'])]
            print(f'Writing {Color.GREEN}"{vols_q_sheet_begin}{i}{vols_q_sheet_end}"{Color.END} sheet to file: {Color.CYAN}"{vols_dir}{vols_file}"{Color.END}')
            vols_q.to_excel(writer, sheet_name=f'{vols_q_sheet_begin}{i}{vols_q_sheet_end}')  # write quarters sheets to file

        for caucasian_region in regions:
            # sorting VOLS to regions
            vols_region = vols[vols[work_ro] == regions[caucasian_region]]
            print(f'Writing {Color.GREEN}"{caucasian_region}"{Color.END} sheet to file: {Color.CYAN}"{vols_dir}{vols_file}"{Color.END}')
            vols_region.to_excel(writer, sheet_name=caucasian_region)  # write region sheets to file
