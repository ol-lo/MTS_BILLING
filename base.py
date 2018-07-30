# -*- coding: utf-8 -*-
import os
from copy import copy

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import Color
from time import sleep

from constants import (
    NAME_FILE, NAME_PAGE, MAP_EXCEL, DATA_ROW_HEADER, FIRST, CP_1251,
    MAP_MONTH,
    PHONE, COL_NAMES, RANGE_WITH_SUM, START, TEMPLATE_RANGE, END,
    AMOUNT_PER_NUMBER, ONE, CONTRACT, PATH_TO_IN, XLSM)
from report import PDFReport, BillingErrorFileWrite


class AllReports(PDFReport):
    END_LINE = u'Итого по лицевому счету без скидки'
    DICT_ARGUMENTS = {
        12: AMOUNT_PER_NUMBER,
        2: PHONE
    }

    def __init__(self, data_unit, pattern_date, pattern_contract, out_xls):
        super(
            AllReports,
            self).__init__(data_unit, pattern_date, pattern_contract)
        self.out_file = out_xls

    def write_(self, out_file, result):

        book_excel = openpyxl.load_workbook(out_file)

        if XLSM in out_file:
            book_excel = openpyxl.load_workbook(out_file, keep_vba=True)

        page_for_report = book_excel[NAME_PAGE]
        day, month, year = self.date_.split('.')
        month_name = MAP_MONTH.get(month)
        range_with_data_headers = page_for_report[MAP_EXCEL.get(
            DATA_ROW_HEADER)
        ]
        cell_phone = MAP_EXCEL[COL_NAMES][PHONE]
        cell_phone_start = MAP_EXCEL[RANGE_WITH_SUM][START]
        cell_phone_end = MAP_EXCEL[RANGE_WITH_SUM][END]

        phone_range = TEMPLATE_RANGE.format(
            cell_char_one=cell_phone,
            digit_one=cell_phone_start,
            cell_char_two=cell_phone,
            digit_two=cell_phone_end
        )
        range_phone_all = page_for_report[phone_range]

        range_date_headers = filter(
            lambda x:
            x if x.value == ' '.join([month_name, year]) else None,
            range_with_data_headers[FIRST]
        )
        sys.stdout.write(
            u'\n Выборка дат из excel завершина \n '
        )

        range_phone = filter(lambda x: x[0].value is not None, range_phone_all)

        sys.stdout.write(u'\nВыборка телефонов завершина. \n')

        if range_date_headers:
            cell_cur = range_date_headers[0]
            sys.stdout.write(
                u"дата присуствует в"
                u" списке\n ячейка {}{} \n значение {}\n\n".format(
                    cell_cur.column, cell_cur.row, cell_cur.value)
            )

            phones_amount_dict = map(
                lambda x: {
                    int(x[PHONE]):
                        round(
                            float(
                                x[AMOUNT_PER_NUMBER].replace(',', '.')
                            ),
                            2
                        )
                },
                result)

            dict_res = {}
            for item in phones_amount_dict:
                dict_res.update(item)
            for item_cell in range_phone:
                var_temp_value = int(item_cell[0].value)
                var_temp_row_index = '{}{}'.format(
                    cell_cur.column, item_cell[0].row)
                try:
                    cur_value = float(
                        page_for_report[var_temp_row_index].value or 0.0
                    )
                except ValueError:
                    cur_value = 0.0

                amount_per_number = dict_res.get(var_temp_value, None)
                if amount_per_number is not None:
                    dict_res.pop(var_temp_value)
                    sys.stdout.write(
                        u"Подстановка значений {}+{} в ячейку {}\n".format(
                            amount_per_number, cur_value, var_temp_row_index)
                    )
                    page_for_report[var_temp_row_index].value = (
                            amount_per_number + cur_value)

                    left_cell = chr(ord(cell_cur.column)-1)
                    fill_ = page_for_report[
                        '{}{}'.format(left_cell, item_cell[0].row)].fill

                    color = fill_.fgColor.rgb
                    if color == '00000000':
                        # Замена черного на белый
                        color = 'FFFFFFFF'
                    try:
                        left_fill = PatternFill(
                            fill_type='solid',
                            fgColor=Color(color),
                        )
                    except TypeError:
                        # Не всегда удается определить цвет

                        new_cell = page_for_report[var_temp_row_index]
                        old_cell = page_for_report[
                            '{}{}'.format(left_cell, item_cell[0].row)]
                        # копируем значения стилей
                        new_cell.font = copy(old_cell.font)
                        new_cell.border = copy(old_cell.border)
                        new_cell.fill = copy(old_cell.fill)
                        new_cell.number_format = copy(old_cell.number_format)
                        new_cell.protection = copy(old_cell.protection)
                        new_cell.alignment = copy(old_cell.alignment)

                    else:
                        page_for_report[var_temp_row_index].fill = left_fill

            last_num = range_phone[-1][0].row
            if dict_res and not amount_per_number:
                for item_not_found in dict_res:
                    sys.stdout.write(
                        u"{}   нет в списке \n".format(item_not_found)
                    )
                    last_num += ONE
                    sys.stdout.write(
                        u'добавляю значение {} в {}{}\n'.format(
                            dict_res[item_not_found], cell_cur.column, last_num)
                    )
                    page_for_report[
                        '{}{}'.format(cell_cur.column, last_num)
                    ].value = dict_res[item_not_found]

                    sys.stdout.write(
                        u'добавляю значение {} в {}{}\n'.format(
                            item_not_found, MAP_EXCEL[COL_NAMES][PHONE],
                            last_num)
                    )

                    page_for_report[
                        '{}{}'.format(MAP_EXCEL[COL_NAMES][PHONE], last_num)
                    ].value = item_not_found

                    sys.stdout.write(
                        u'добавляю значение {} в {}{}\n'.format(
                            self.contract, MAP_EXCEL[COL_NAMES][CONTRACT],
                            last_num)
                    )

                    page_for_report[
                        '{}{}'.format(
                            MAP_EXCEL[COL_NAMES][CONTRACT], last_num)
                    ].value = self.contract

        else:
            sys.stdout.write(
                u"дата {} {} {} не найдена\n\nСледующий ...\n".format(
                    day, month, year)
            )
        try:
            book_excel.save(out_file)
            book_excel.close()

        except IOError:
            raise BillingErrorFileWrite


if __name__ == "__main__" and __package__ is None:
    from os import sys, path

    try:
        sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))
    except NameError:  # py2exe script, not a module
        import sys
        sys.path.append(path.dirname(path.dirname(sys.argv[0])))

    files = os.listdir(PATH_TO_IN)

    for file_ in files:
        sys.stdout.write(file_.decode(CP_1251) + '\n')
        try:
            AllReports(
                (
                    u'(Сетевой ресурс(\d{11})'
                    u'\s(Периодические\sуслуги\s\d{1,6},'
                    u'\d{4}\s{2})?(Скидка\soff-line'
                    u'\s\(Периодические услуги\)-\d{1,6},'
                    u'\d{4}\s{2})?(Разовые\sуслуги'
                    u'\s\d{1,6},\d{4})?(Скидка\soff-line\s'
                    u'\(Разовые\sуслуги\)-\d{1,6},\d{4}\s{2})'
                    u'?(Телефонные\sуслуги\s\d{1,6},'
                    u'\d{4}\s{0,2})?(Скидка\soff-line\s\(Телефонные\sуслуги\)-'
                    u'\d{1,6},\d{4}\s{2})?(НДС\s\d{1,6},\d{4})'
                    u'?(Итого\s\без\sскидки:'
                    u'\s\d{1,6},\d{4}\s{2})'
                    u'?(Итого\sскидка\soff-line:-\d{1,6},\d{4}\s{2})?'
                    u'Итого:\s(\d{1,6},\d{4})?)'
                ),

                u'.+(?P<date_>\d{2}\.\d{2}\.\d{4})',

                u'.+\d{12}\s{2}(?P<contract>\d{3}\s\d{3}\s\d{3}\s\d{3})',
                NAME_FILE

            ).find_values(os.path.join(PATH_TO_IN, file_))
        except Exception as err:
            sys.stdout.write(err.message)

    sys.stdout.write(u" Конец ... ")
    sys.stdout.write(u" Для выхода зажмите ctrl + C ")
    sleep(60)
