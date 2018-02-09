import xlrd
import re

DATA_PATH = '../../data/pge_gas_small.xlsx'
SHEET = 'G-NR1 (2016-Present)'

CHARGE_INDEX = [i for i in range(2, 6)]
SUMMER = 12
WINTER = 14

CHARGE_INDEX.extend([SUMMER, WINTER])

'''Returns the xls sheet object'''
def get_xls_object(file_path, sheet):
    xl_workbook = xlrd.open_workbook(file_path, encoding_override="cp1252")
    return xl_workbook.sheet_by_name(sheet)

def main():
    xl_sheet = get_xls_object(DATA_PATH, SHEET)
    total_rows = xl_sheet.nrows

    re_date = '^\d+\.'
    for i in range(0, total_rows):
        cell_value = xl_sheet.cell_value(i, 0)
        re_match = re.match(re_date, str(cell_value))

        if re_match:
            p_date = xlrd.xldate_as_datetime(cell_value, 0)
            p_rate = []
            for index in CHARGE_INDEX:
                p_rate.append(xl_sheet.cell_value(i, index))
            print '{}{}|{}'.format(p_date.strftime('%y'), p_date.strftime('%m'), p_rate)


if __name__ == '__main__':
    main()
