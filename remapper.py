import json
import xlrd

from os import listdir
from os.path import isfile, join

DATA_PATH = '../data/form/raw_input/'

def parse(data_row):
    """Loads the CSV data - Creates a Dictionary Map"""
    form_mapper = dict()
    form_mapper['geminiForm'] = []

    counter_section = 0
    counter_element = 0

    for row in data_row:

        if not row[1]:
            continue

        if row[0]:
            section = row[0]

            counter_section += 1
            counter_element = 0

            section_block = dict()
            section_block['section'] = section
            section_block['id'] = str(counter_section)
            section_block['index'] = counter_section
            section_block['elements'] = []

            form_mapper.get('geminiForm').append(section_block)

        element = row[1]
        element_type = row[2]
        element_options = ''
        element_validation = ''

        def map_option_id(option_line):
            options = option_line.split(',')
            for option_id in range(1, len(options) + 1):
                yield str(option_id) + ":" + str(options[option_id - 1]).strip()

        if element_type == 'PickerInputRow':
            element_options = ",".join(map_option_id(row[3]))

        counter_element += 1

        element_block = dict()
        element_block['id'] = str(counter_section) + str(counter_element)
        element_block['param'] = element
        element_block['index'] = int(str(counter_section) + str(counter_element))
        element_block['dataType'] = str.lower(element_type)
        element_block['defaultValues'] = element_options
        element_block['validation'] = element_validation

        section_block.get('elements').append(element_block)

    return form_mapper


def list_file_path():
    data_files = [files for files in listdir(DATA_PATH) if isfile(join(DATA_PATH, files))]

    '''Note: Removing this item as it does not comply with Company - Model Index'''
    try:
        data_files.remove('.DS_Store')
    except():
        print('unable to clean data files')
        pass

    return data_files


def build_config(file_path):
    wb = get_work_book(file_path)
    total_index = len(wb.sheets())

    for index in range(2, total_index):
        sheet = wb.sheet_by_index(index)
        name = sheet.name
        data = parse(build_json(sheet))

        write_json(data, name)


def build_json(sheet):
    total_rows = sheet.nrows
    total_columns = sheet.ncols
    for row in range(0, total_rows):
        if row == 0:
            continue
        data_row = sheet.row_values(row, start_colx=0, end_colx=total_columns)
        csv_data_row = [str(item) for item in data_row]
        yield csv_data_row


def write_json(data, file_name):
    output_file = '../data/form/output/{}.json'.format(file_name)

    print output_file
    print(json.dumps(data))

    '''Writing the JSON Output to a file'''
    with open(output_file, 'w') as fp:
        json.dump(data, fp)


def get_work_book(file_path):
    try:
        return xlrd.open_workbook(file_path, encoding_override='cp1252')
    except Exception, err:
        print('Exception in file {}'.format(file_path))
        print err


def main():
    data_files = list_file_path()
    for data_file in data_files:
        build_config(join(DATA_PATH, data_file))


if __name__ == "__main__":
    main()
