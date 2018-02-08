import xlrd
import re

DATA_PATH = '../../data/pge_electric.xlsx'
SHEET = 'Comm\'l_170301-Present'


'''Returns the xls sheet object'''
def get_xls_object(file_path, sheet):
    xl_workbook = xlrd.open_workbook(file_path, encoding_override="cp1252")
    return xl_workbook.sheet_by_name(sheet)


'''Returns the index of the first match from the given list'''
def index_of_match(row_list, regex):
    try:
        match_text = next(i for i in row_list if regex.search(str(i)))
        return row_list.index(match_text)
    except:
        return -1

def main():

    xl_sheet = get_xls_object(DATA_PATH, SHEET)
    total_rows = xl_sheet.nrows

    '''Header Block'''
    header_row = xl_sheet.row_values(0)
    header_row_clean = [re.sub('\s+', ' ', item).strip().lower() for item in header_row]

    season = xl_sheet.col_slice(index_of_match(header_row_clean, re.compile('season')))

    demand_charge = xl_sheet.col_slice(index_of_match(header_row_clean, re.compile('^demand')))
    time_of_use_dc = xl_sheet.col_slice(index_of_match(header_row_clean, re.compile('^demand')) - 1)

    energy_charge = xl_sheet.col_slice(index_of_match(header_row_clean, re.compile('^total energy')))
    time_of_use_ec = xl_sheet.col_slice(index_of_match(header_row_clean, re.compile('^total energy')) - 1)

    '''Electric Rate Structure Block'''
    re_charge = '(^(A|E)-\d+(\s+)(TOU)?)'

    current_rate = None
    current_season = None
    current_demand_charge = None

    for i in range(0, total_rows):
        cell_value = xl_sheet.cell(i, 0).value
        re_match = re.match(re_charge, cell_value)
        if re_match:
            rate_clean = re.sub('\s+', ' ', re_match.group(1)).strip()
            current_rate = rate_clean

        if current_rate and energy_charge[i].value:
            if season[i].value:
                current_season = season[i].value

            if demand_charge[i].value:
                current_demand_charge = demand_charge[i].value

            print '{} | {}  | {} | {} | {} | {}'.\
                format(current_rate, current_season, time_of_use_dc[i].value, current_demand_charge,
                       time_of_use_ec[i].value, energy_charge[i].value)


if __name__ == '__main__':
    main()
