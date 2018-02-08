import xlrd
import re
import sys
import codecs
import csv

DATA_PATH = '../../data/pge_electric.xlsx'
SHEET = 'Comm\'l_170301-Present'

xl_sheet = None


'''Returns the xls sheet object'''
def get_xls_object(file_path, sheet):
    xl_workbook = xlrd.open_workbook(file_path, encoding_override="cp1252")
    return xl_workbook.sheet_by_name(sheet)

def main():

    xl_sheet = get_xls_object(DATA_PATH, SHEET)
    total_rows = xl_sheet.nrows

    rate_structure_block = {}
    demand_charge_block = {}
    total_energy_charge_block = {}

    '''Electric Rate Structure Block'''
    # 1. Find the Row Blocks for Each of the Rates ie [A-1 | A-1 TOU | A-6 TOU ... E-19]
    # 2. Identify the Start Index | End Index
    re_charge = '((A|E)-\d+(\s+)(TOU)?)'

    counter = 0
    for i in range(0, total_rows):
        cell_value = xl_sheet.cell(i, 0).value
        re_match = re.match(re_charge, cell_value)
        if re_match:
            rate_clean = re_match.group(1).strip()
            rate_structure_block[rate_clean] = {
                'start_index': counter,
                'end_index': i - 1
            }
            counter = i

    '''Demand Charge Block'''

    print rate_structure_block

if __name__ == '__main__':
    main()
