import xlrd
import re
import csv
from datetime import datetime

DATA_PATH = '../../data/sample_interval_data.xlsx'
SHEET = 'Sample Interval Data.csv'
OUTPUT_FILE = '../resource/interval_data.csv'
SEASON = range(1, 13)

usage = dict()
usage['summer-peak'] = 0.0
usage['summer-partial-peak'] = 0.0
usage['summer-off-peak'] = 0.0
usage['winter-partial-peak'] = 0.0
usage['winter-off-peak'] = 0.0


'''Validates the Time Range Between Two Specified Time'''
def in_between(now, start, end):
    if start < end:
        return start <= now < end
    elif end < start:
        return start <= now or now < end
    else:
        return True


'''Returns the xls sheet object'''
def get_xls_object(file_path, sheet):
    xl_workbook = xlrd.open_workbook(file_path, encoding_override="cp1252")
    return xl_workbook.sheet_by_name(sheet)


'''Checker Functions - Season | WeekDay'''
def is_summer(month):
    return month in SEASON[4:10]

def is_winter(month):
    return month in SEASON[0:4] + SEASON[10:12]

def is_weekend(day):
    return day in ['Saturday', 'Sunday']


'''Peak Hour Validators'''
def get_time(time):
    return datetime.strptime(time, '%H:%M').time()

def is_summer_peak(time):
    return in_between(time, get_time('12:00'), get_time('18:00'))

def is_summer_partial_peak(time):
    return in_between(time, get_time('8:30'), get_time('12:00')) or \
           in_between(time, get_time('18:00'), get_time('21:30'))

def is_summer_off_peak(time):
    return in_between(time, get_time('21:30'), get_time('8:30'))

def is_winter_partial_peak(time):
    return in_between(time, get_time('8:30'), get_time('9:30'))

def is_winter_off_peak(time):
    return in_between(time, get_time('21:00'), get_time('8:30'))


'''MAIN'''
def main():
    xl_sheet = get_xls_object(DATA_PATH, SHEET)
    total_rows = xl_sheet.nrows
    outgoing = []

    re_date = '^\d+\.'
    for i in range(0, total_rows):
        cell_value = xl_sheet.cell_value(i, 2)
        re_match = re.match(re_date, str(cell_value))

        if re_match:
            p_date = xlrd.xldate_as_datetime(cell_value, 0)
            p_usage_amount = xl_sheet.cell_value(i, 8)

            month = int(p_date.strftime('%m'))
            day_of_week = p_date.strftime('%A')
            time = get_time(p_date.strftime('%H:%M'))

            if is_summer(month):
                if is_weekend(day_of_week):
                    usage['summer-off-peak'] += p_usage_amount
                elif is_summer_peak(time):
                    usage['summer-peak'] += p_usage_amount
                elif is_summer_partial_peak(time):
                    usage['summer-partial-peak'] += p_usage_amount
                elif is_summer_off_peak(time):
                    usage['summer-off-peak'] += p_usage_amount
            elif is_winter(month):
                if is_weekend(day_of_week):
                    usage['winter-off-peak'] += p_usage_amount
                elif is_winter_partial_peak(time):
                    usage['winter-partial-peak'] += p_usage_amount
                elif is_winter_off_peak(time):
                    usage['winter-off-peak'] += p_usage_amount

    outgoing.append(usage.keys())
    outgoing.append([round(x, 2) for x in usage.values()])

    with open(OUTPUT_FILE, 'w') as file_obj:
        writer = csv.writer(file_obj, delimiter=',', quoting=csv.QUOTE_MINIMAL)
        for line in outgoing:
            writer.writerow(line)

if __name__ == '__main__':
    main()
