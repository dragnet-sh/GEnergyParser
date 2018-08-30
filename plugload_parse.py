import os

instance = 'PROD'

if instance == 'LOCAL':
    os.environ["PARSE_API_ROOT"] = "http://localhost:1337/parse"
    APPLICATION_ID = 'NLI214vDqkoFTJSTtIE2xLqMme6Evd0kA1BbJ20S'
    MASTER_KEY = 'lgEhciURXhAjzITTgLUlXAEdiMJyIF4ZBXdwfpUr'

if instance == 'PROD':
    os.environ["PARSE_API_ROOT"] = "http://ec2-18-220-200-115.us-east-2.compute.amazonaws.com:80/parse/"
    APPLICATION_ID = '47f916f7005d19ddd78a6be6b4bdba3ca49615a0'
    MASTER_KEY = 'NLI214vDqkoFTJSTtIE2xLqMme6Evd0kA1BbJ20S'

from parse_rest.connection import register
from parse_rest.datatypes import Object
from parse_rest.connection import ParseBatcher

register(APPLICATION_ID, '', master_key=MASTER_KEY)


class PlugLoad(Object):
    pass


class HVAC(Object):
    pass


class Motors(Object):
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
cnx = mysql.connect(user='root', password='', host='127.0.0.1', database='gemini_v1')

from os import listdir
from os.path import isfile, join

DATA_PATH = '../data/motors/'

op_hdr_lines = []
hdr_hash_map = dict()
hdr_metadata = list()

re_hdr_identifier = re.compile('company', re.IGNORECASE)
hdr_column_index = 0
tbl_column_index = 0  # *** Set this to whatever value the Table Column starts at ***


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
        # for i in range(0, total_rows):
        #     if re.search(re_hdr_identifier, sheet.cell(i, hdr_column_index).value):
        #         hdr_row_index = i
        #         break

        hdr_row_index = 0  # *** The regular expression search does this or set it manually ***

        '''Sanitize the Header Row - Create the Hash Map for each of the Header Items'''
        hdr_row = sheet.row(hdr_row_index)
        # hdr_row.pop(0)  # This is because the first column is empty **** IMP. CONFIGURE THIS !! ****
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
        hdr_offset = 1  # Set to 1 if the first line is a header
        data_collection = []
        for i in range(hdr_row_index + hdr_offset, total_rows):
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
            # plugload = PlugLoad(
            #     type=data_file
            # )
            # plugload.data = dict()

            to_upload = Motors(
                type=data_file
            )
            to_upload.data = dict()

            for index, data in enumerate(data_row):
                field_id = hdr_hash_map.get(data_file)[index]
                to_upload.data.update({
                    field_id: data
                })

            ## 3. Aggregating the Data to do a Batch Save
            data_collection.append(to_upload)

        cnx.commit()
        cursor.close()

        ## 4. Batch Upload - Parse Server
        ParseBatcher().batch_save(data_collection)

def list_file_path():
    # data_files = [files for files in listdir(DATA_PATH) if isfile(join(DATA_PATH, files))]
    data_files = ['motor_efficiency.xlsx']

    '''Note: Removing this item as it does not comply with Company - Model Index'''
    try:
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
