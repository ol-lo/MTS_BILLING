# -*- coding: utf-8 -*-
u"""
    Константы пакета
"""
import yaml

CP_1251 = 'cp1251'
UTF8 = 'utf-8'
TXT = u'txt'
PDF = u'pdf'
FIRST = ZERO = 0
SECOND = ONE = 1
READ_BINARY = 'rb'
ENCODE_TYPE = 'utf-16-le'
REPLACE_CHAR = b'\x00'
EMPTY_STRING = ''
CONFIG_FILE_NAME = 'config.yaml'


################################################################################
#   Константы для вынесения в конфиг
################################################################################


# SOURCE_FILE = (
#     u'W:\\IT\\maxim.nalimov\\PROJECTS\\MTS BILLING REPORT\\Эталон.xlsx')

TITLE_ROW = 'TITLE_ROW'
COL_NAMES = 'COL_NAMES'
DATA_ROW_HEADER = 'DATA_ROW_HEADER'
DATA_START = 'DATA_START'
DATA_END = 'DATA_END'
RANGE_FULL = 'RANGE_FULL'
PHONE = 'PHONE'
OWNER = 'OWNER'
REGION = 'REGION'
CONTRACT = 'CONTRACT'
WHO = 'WHO'
LIMIT = 'LIMIT'
START_CELL_DATE = 'START_CELL_DATE'
END_CELL_DATE = 'END_CELL_DATE'
START = 'START'
END = 'END'
RANGE_WITH_SUM = 'RANGE_WITH_SUM'
TEMPLATE_RANGE = '{cell_char_one}{digit_one}:{cell_char_two}{digit_two}'
AMOUNT_PER_NUMBER = "AMOUNT_PER_NUMBER"
MAP_EXCEL = {
    # Управление параметрами через словарь
    TITLE_ROW: 1,
    COL_NAMES: {
        REGION: 'A',
        OWNER: 'B',
        PHONE: 'C',
        CONTRACT: 'D',
        WHO: 'E',
        LIMIT: 'F'
    },
    START_CELL_DATE: 'G',
    END_CELL_DATE: 'AA',
    RANGE_WITH_SUM: {
        START: 2,
        END: 2000
    }
}

MAP_EXCEL.update(
    {
        DATA_ROW_HEADER:
            '{START_CELL_DATE}{TITLE_ROW}:{END_CELL_DATE}{TITLE_ROW}'.format(
                TITLE_ROW=MAP_EXCEL.get(TITLE_ROW),
                END_CELL_DATE=MAP_EXCEL.get(END_CELL_DATE),
                START_CELL_DATE=MAP_EXCEL.get(START_CELL_DATE)
            ),

        DATA_START:
            '{START_CELL_DATE}{START}:{END_CELL_DATE}{START}'.format(
                START=MAP_EXCEL.get(RANGE_WITH_SUM).get(START),
                START_CELL_DATE=MAP_EXCEL.get(START_CELL_DATE),
                END_CELL_DATE=MAP_EXCEL.get(END_CELL_DATE)
            ),
        DATA_END:
            '{START_CELL_DATE}{END}:{END_CELL_DATE}{END}'.format(
                START_CELL_DATE=MAP_EXCEL.get(START_CELL_DATE),
                END=MAP_EXCEL.get(RANGE_WITH_SUM).get(END),
                END_CELL_DATE=MAP_EXCEL.get(END_CELL_DATE)
            ),
        RANGE_FULL:
            '{START_CELL_DATE}{START}:{END_CELL_DATE}{END}'.format(
                START_CELL_DATE=MAP_EXCEL.get(START_CELL_DATE),
                END_CELL_DATE=MAP_EXCEL.get(END_CELL_DATE),
                START=MAP_EXCEL.get(RANGE_WITH_SUM).get(START),
                END=MAP_EXCEL.get(RANGE_WITH_SUM).get(END)
            )
    }
)

MAP_MONTH = {
    '12': 'December',
    '11': 'November',
    '10': 'October',
    '09': 'September',
    '08': 'August',
    '07': 'July',
    '06': 'June',
    '05': 'May',
    '04': 'April',
    '03': 'March',
    '02': 'February',
    '01': 'January'
}

with open(CONFIG_FILE_NAME) as config_file:
    config = yaml.load(config_file.read())
NAME_PAGE = config['FILE_PARAM']['PAGE_NAME']
NAME_FILE = config['FILE_PARAM']['PATH_TO_FILE_OUT']
PATH_TO_IN = config['FILE_PARAM']['PATH_TO_IN_FOLDER']
