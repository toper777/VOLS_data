# Import section

# import os for file operation
# import os

# import pandas for data operation
import pandas as pd

# main variables
data_dir = "y:/Блок №3/2022 год/"
data_file = "!!!SQL Блок№3!!!  2022.xlsm"
data_sheet = "Массив"

# vols_dir = "c:/tmp/"
vols_dir = "y:/Блок №3/2022 год/"
vols_file = "VOLS KVK 2022.xlsx"
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
