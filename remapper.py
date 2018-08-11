import csv
import json
import uuid

'''Loads the CSV data - Creates a Dictionary Map'''
def parse_csv(file_path):
    fh = open(file_path, 'rt')
    form_mapper = dict()
    form_mapper['gemini-form'] = []

    counter_section = 0
    counter_element = 0

    try:
        reader = csv.reader(fh)
        for index, row in enumerate(reader):

            if not row[1]:
                continue

            if row[0]:
                section = row[0]
                print section

                counter_section += 1
                counter_element = 0

                section_block = dict()
                section_block['section'] = section
                section_block['id'] = str(counter_section)
                section_block['index'] = counter_section    
                section_block['elements'] = []
                
                form_mapper.get('gemini-form').append(section_block)

            element = row[1]
            element_type = row[2]
            element_options = ''
            element_validation = ''

            if element_type == 'PickerInputRow':
                element_options = row[3]

            counter_element += 1
            
            element_block = dict()
            element_block['id'] = str(counter_section) + str(counter_element)
            element_block['param'] = element
            element_block['index'] = int(str(counter_section) + str(counter_element))
            element_block['data-type'] = str.lower(element_type)
            element_block['default-values'] = element_options
            element_block['validation'] = element_validation

            section_block.get('elements').append(element_block)

    finally:
        fh.close()

    return form_mapper


input_file = './form/combination_oven.csv'
output_file = './output/combination_oven.json'
data = parse_csv(input_file)

print(json.dumps(data))

'''Writing the JSON Output to a file'''
with open(output_file, 'w') as fp:
    json.dump(data, fp)

exit(0)
