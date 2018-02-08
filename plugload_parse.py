import os

instance = 'LOCAL'

if instance == 'LOCAL':
    os.environ["PARSE_API_ROOT"] = "http://localhost:1337/parse"
    APPLICATION_ID = 'NLI214vDqkoFTJSTtIE2xLqMme6Evd0kA1BbJ20S'
    MASTER_KEY = 'lgEhciURXhAjzITTgLUlXAEdiMJyIF4ZBXdwfpUr'

from parse_rest.connection import register
from parse_rest.datatypes import Object
from parse_rest.connection import ParseBatcher

register(APPLICATION_ID, '', master_key=MASTER_KEY)


class PlugLoad(Object):
    pass


## 1. Read the CSV file
## 2. Save the Header to a list
## 3. Iterate through each of the data lines
## 4. Create Plugload Object
## 5. Save it to the Parse Server

import xlrd
import re

import mysql.connector as mysql
from mysql.connector import Error
cnx = mysql.connect(user='root', password='password', host='127.0.0.1', database='gemini')

from os import listdir
from os.path import isfile, join

DATA_PATH = '../../data/PlugLoad-Equipment/'

op_hdr_lines = []
hdr_hash_map = dict()
hdr_metadata = list()

re_hdr_identifier = re.compile('company', re.IGNORECASE)
hdr_column_index = 1
tbl_column_index = 1


def main():
    data_files = list_file_path()

    '''Loop through each of the PlugLoad Equipment Files'''
    for data_file in data_files:
        # print data_file
        file_path = join(DATA_PATH, data_file)

        try:
            xlsx = xlrd.open_workbook(file_path, encoding_override='cp1252')
        except Exception, err:
            print('Exception in file {}'.format(data_file))
            print(err)

        data_file = data_file.replace('.xlsx', '').lower().replace('-', '_')
        sheet = xlsx.sheet_by_index(0)
        total_rows = sheet.nrows

        hdr_row_index = None

        '''Identify the Header Start Row Index'''
        for i in range(0, total_rows):
            if re.search(re_hdr_identifier, sheet.cell(i, hdr_column_index).value):
                hdr_row_index = i
                break

        '''Sanitize the Header Row - Create the Hash Map for each of the Header Items'''
        hdr_row = sheet.row(hdr_row_index)
        hdr_row.pop(0)  # This is because the first column is empty ****
        hdr_row_sanitized = [sanitize(item) for item in hdr_row]
        hdr_hash_map.update({data_file: [item for item in hdr_row_sanitized]})

        '''SQL DROP | CREATE TABLE - STATEMENT'''
        sql_create = "DROP TABLE IF EXISTS `{}`; CREATE TABLE {} ({} VARCHAR(255));".format(data_file, data_file, ' VARCHAR(255), '.join(hdr_row_sanitized))
        print sql_create

        for statement in sql_create.split(";"):
            try:
                cursor = cnx.cursor()
                cursor.execute(statement)
            except Error as e:
                print e
            finally:
                cnx.close

        '''Create the Data Row'''
        pl_data_collection = []
        for i in range(hdr_row_index + 1, total_rows):
            data_row = sheet.row_values(i, start_colx=tbl_column_index, end_colx=len(hdr_row_sanitized) + 1)

            ''' ***** SQL DATA INSERTION ***** '''

            sql_data_row = [str(item) for item in data_row]
            sql_insert = "INSERT INTO {} ({}) VALUES (\"{}\")".format(data_file, ','.join(hdr_row_sanitized), '","'.join(sql_data_row))
            print sql_insert

            try:
                cursor = cnx.cursor()
                cursor.execute(sql_insert)
            except Error as e:
                print e

            ''' ***** PARSE UPLOAD ***** '''

            ## 1. Initializing the Plugload Object
            plugload = PlugLoad(
                type=data_file
            )
            plugload.data = dict()

            for index, data in enumerate(data_row):
                field_id = hdr_hash_map.get(data_file)[index]
                plugload.data.update({
                    field_id: data
                })

            ## 3. Aggregating the Data to do a Batch Save
            pl_data_collection.append(plugload)

        cnx.commit()
        cursor.close()


        ## 4. Batch Upload - Parse Server
        ParseBatcher().batch_save(pl_data_collection)

def list_file_path():
    data_files = [files for files in listdir(DATA_PATH) if isfile(join(DATA_PATH, files))]
    # data_files = ['Solid_Door_Freezers_retrofits.xlsx']

    '''Note: Removing this item as it does not comply with Company - Model Index'''
    try:
        data_files.remove('televisions_retrofit_2017.xlsx')
        data_files.remove('.DS_Store')
    except:
        print('unable to clean data files')
        pass

    return data_files

def sanitize(item):
    tmp = item.value.encode('utf-8')
    tmp = tmp.replace('\n', ' ').replace('/', ' ').replace('&', 'n').strip()
    tmp = re.sub('\(.*?\)', '', tmp).strip()
    tmp = re.sub('\s', '_', re.sub('\s+', ' ', tmp)).strip()

    return tmp.lower()


if __name__ == "__main__":
    main()
