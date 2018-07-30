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
ONE = 1
READ_BINARY = 'rb'
ENCODE_TYPE = 'utf-16-le'
EMPTY_CHAR_FOR_REPLACE = b'\x00'
YYYY_MM_DD_FORMAT = '%Y_%m_%d'
CONFIG_FILE_NAME = 'config.yaml'


################################################################################
#   Константы для вынесения в конфиг
################################################################################

[
    TITLE_ROW, COL_NAMES, DATA_ROW_HEADER, DATA_START, DATA_END, RANGE_FULL,
    PHONE, OWNER, REGION, CONTRACT, WHO, LIMIT, START_CELL_DATE, END_CELL_DATE,
    START, END, RANGE_WITH_SUM, AMOUNT_PER_NUMBER] = [
    'TITLE_ROW', 'COL_NAMES', 'DATA_ROW_HEADER', 'DATA_START', 'DATA_END',
    'RANGE_FULL', 'PHONE', 'OWNER', 'REGION', 'CONTRACT', 'WHO', 'LIMIT',
    'START_CELL_DATE', 'END_CELL_DATE', 'START', 'END', 'RANGE_WITH_SUM',
    'AMOUNT_PER_NUMBER'
]
TEMPLATE_RANGE = '{cell_char_one}{digit_one}:{cell_char_two}{digit_two}'
MAP_EXCEL = {
    # файл excel, является представлением и базой данных одновременно
    # Текущий словарь, является "картой", содержащую координаты ячеек и
    # и их физический смысл
    TITLE_ROW: 1,  # вторая координата заголовков
    COL_NAMES: {
        REGION: 'A',  # первая координата заголовка REGION
        OWNER: 'B',  # первая координата заголовка OWNER
        PHONE: 'C',  # первая координата заголовка PHONE
        CONTRACT: 'D',  # первая координата заголовка CONTRACT
        WHO: 'E',  # первая координата заголовка WHO
        LIMIT: 'F'  # первая координата заголовка LIMIT
    },
    START_CELL_DATE: 'G',  # первая координата, с которой начинается поиск дат
    END_CELL_DATE: 'AA',  # первая координата, конца поиска дат
    RANGE_WITH_SUM: {  # вторые координаты, начала и конца ячеек со значениями
        START: 2,
        END: 2000
    }
    # Принял решение разделить excel файл на G2:AA2000 - область с данными
    # A1:F1 - область заголовков
    # A2:F2000 - область базы данных
    # остальная часть, на странице с данными является неиспользуемой
}

MAP_EXCEL.update(
    {
        # Здесь происходит обновление, пар координат.
        # Каждая пара координат, это ряд, опряделяющий задачу ячейки
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
    #  Словарь месяцев
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
