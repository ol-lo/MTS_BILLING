# -*- coding: utf-8 -*-
""" Интерфейc поиска в отчетах регулярных вырожений """

import re
import sre_constants
import sys

import PyPDF2

from constants import (
    ENCODE_TYPE, EMPTY_CHAR_FOR_REPLACE, READ_BINARY, ONE, CP_1251, UTF8, ZERO)
from logger import ReportLoggerMixin, LoggerAbstractMixin


class BillingErrorDataUnit(Exception):
    """
    Класс исключений, ошибка компиляции  выжения
    итоговых сумм в отчетах
    """
    _TEXT_UNIT = u'(data_unit)'
    _ARGS = (
        u'Регулярное вырожение не {} ' 
        u'может быть скомпилировано \n'.format(_TEXT_UNIT)
    )

    def __init__(self):
        self.__module__ = self.__class__.__name__
        self.args = self._ARGS


class BillingErrorPatternDate(BillingErrorDataUnit):
    """
    Класс исключений, ошибка компиляции регулярного вырожения
    для поиска даты отчета
    """
    _TEXT_UNIT = u'(pattern_date)'


class BillingErrorPatternContact(BillingErrorDataUnit):
    """
        Класс исключений, ошибка компиляции регулярного вырожения
        для поиска контакта
    """
    _TEXT_UNIT = u'(pattern_contract)'


class BillingErrorFileRead(BillingErrorPatternContact):
    """
    Класс ошибок, связанный с работой  операционной системы
    или ошибки доступа к файлу.
    """
    _ARGS = (
        u"Не удается прочитать файл.\n"
        u"Возможно, нет прав на прочтение \n"
        u" или файла не сущестует"
    )


class BillingErrorFileWrite(BillingErrorFileRead):
    """
       Класс ошибок, связанный с работой  операционной системы
       или ошибки доступа к файлу.
    """
    _ARGS = (
        u"Файл не может быть изменен, \n вероятно файл открыт ?"
    )


class InterfaceReportMTS(object):
    """ Интерфейс отчетов """

    def __init__(self, data_unit):
        """
        Ининциализация класса
        :param data_unit: Текст, использумый в качестве регулярного вырожения.
        """
        self.result = []
        try:
            try:
                self.data_unit = re.compile(data_unit, re.UNICODE)
            except sre_constants.error:
                raise BillingErrorDataUnit

        except ValueError as error:
            sys.stdout.write(
                error.message
            )
            sys.exit()

    def find_values(self, file_name):
        """ Метод запускает поиск регулярных выражений на странице"""
        raise NotImplementedError(
            u'Определите find_values в %s.' % self.__class__.__name__)


class PDFReport(InterfaceReportMTS, ReportLoggerMixin):

    END_LINE = None
    DICT_ARGUMENTS = {}

    def __init__(self, data_unit, pattern_date, pattern_contract):
        """
        :param data_unit:
            регулярное вырожение для поиска итоговых сумм по номерам
        :param pattern_date:
            регулярное вырожение для поиска даты
        :param pattern_contract:
            регулярное вырожение для номера контракта
        """
        self.out_file = None
        self.contract = ''
        self.date_ = ''
        super(PDFReport, self).__init__(data_unit)
        try:
            try:
                self.pattern_date = re.compile(pattern_date, re.UNICODE)
                self.write_log(LoggerAbstractMixin.INFO,
                               u"Успех: компиляции регулярных вырожений")
            except sre_constants.error:
                self.write_log(LoggerAbstractMixin.ERROR,
                               u"Ошибка: компиляции регулярного вырожения")
                raise BillingErrorPatternDate
            try:
                self.pattern_contract = re.compile(pattern_contract, re.UNICODE)
            except sre_constants.error:
                self.write_log(LoggerAbstractMixin.ERROR,
                               u"Ошибка: компиляции регулярного вырожения")
                raise BillingErrorPatternContact

        except (BillingErrorPatternContact, BillingErrorPatternDate) as error:
            self.write_log(LoggerAbstractMixin.ERROR,
                           u"error.message: {}".format(error.message))
        except ValueError as error:
            self.write_log(LoggerAbstractMixin.ERROR,
                           u"error.message: {}".format(error.message))
            sys.stdout.write(
                error.message
            )

    @staticmethod
    def count_pages(file_name):
        u"""
        Возвращает коичество страниц pdf файла
        :param file_name: полный путь к файлу
        :return: int
        """
        try:
            pdf_file_obj = open(file_name, READ_BINARY)
        except IOError:
            raise BillingErrorFileRead
        else:
            pdf_reader = PyPDF2.PdfFileReader(pdf_file_obj)
            return pdf_reader.numPages

    @staticmethod
    def open_file_and_return_text(file_name, num_page):
        u"""
        Возвращает текст с указанной станицы
        :param file_name: имя файла
        :param num_page: номер станицы
        :return: str
        """
        try:
            pdf_file_obj = open(file_name, READ_BINARY)
        except IOError:
            raise BillingErrorFileRead
        else:
            pdf_reader = PyPDF2.PdfFileReader(pdf_file_obj)
            page = pdf_reader.getPage(num_page)
            reading_data = page.extractText()
            first_text_page = reading_data.encode(ENCODE_TYPE)
            result = first_text_page.replace(
                EMPTY_CHAR_FOR_REPLACE, '')
            pdf_file_obj.close()
            return result

    def read_while_not_find_end_line(self, file_name):
        u"""
        Функция прочиывает файл со второй страницы,
        пока не дойдет до опредленного текста. В тексте ищет соотвествия
        регуляному вырожению. Если файле конфигураций задан номер поля
        и описание, то результат поиска сохраниться в self.result
        :param file_name: Имя файла
        :return: None
        """
        count_pages = self.count_pages(file_name)
        self.write_log(LoggerAbstractMixin.INFO,
                       u"Успех: Получено количество страниц")
        if count_pages > ONE:

            for page_num in xrange(ONE, count_pages):
                current_page_text = self.open_file_and_return_text(
                    file_name, page_num)
                unicode_text = current_page_text.decode(CP_1251).encode(UTF8)
                stop_flag = re.findall(
                    self.END_LINE, unicode_text.decode(UTF8))

                contract = self.pattern_contract.search(unicode_text)

                if contract:
                    self.contract = contract.group('contract').replace(' ', '')

                if not self.contract:
                    self.write_log(LoggerAbstractMixin.INFO,
                                   u"Информация: Не обнаружен контракт")

                date_ = self.pattern_date.search(unicode_text)

                if date_ and not self.date_:
                    self.date_ = date_.group('date_')

                if not self.date_:
                    self.write_log(LoggerAbstractMixin.INFO,
                                   u"Информация: Не обнаружена дата")

                for element_on_page in self.data_unit.findall(
                        unicode_text.decode(UTF8)):

                    groups = self.data_unit.match(element_on_page[ZERO])
                    details = {}
                    for key, value in self.DICT_ARGUMENTS.items():
                        details.update({value: groups.group(key)})
                    self.result.append(details)
                if stop_flag:
                    self.write_log(LoggerAbstractMixin.INFO,
                                   u"Успех: Сборка коллекции завершина")
                    return

    def write_(self, out_file, result):
        """
        Метод определяющий поведение записи данных в файл

        :param out_file: Файл для результата
        :param result: Результат
        :return:
        """
        return NotImplemented

    def find_values(self, file_name):
        u"""
        Функция определяет порядок запуска методов класса.
        По сути после использования этой функции свойствами класса
        становиться текст из МТС отчетов
        :param file_name: Иия файла
        :return: None
        """
        if self.END_LINE is None:
            raise ValueError(u"Не задано значение END_LINE")
        self.write_log(LoggerAbstractMixin.INFO,
                       u"Информация: Файл сканируется {}".format(
                           file_name.decode(CP_1251))
                       )
        self.read_while_not_find_end_line(file_name)

        if self.out_file:
            try:
                self.write_(self.out_file, self.result)
            except ValueError:
                self.write_log(LoggerAbstractMixin.ERROR,
                               u"Ошибка: Не достаточно данных для отчетов")
                self.write_log(LoggerAbstractMixin.INFO,
                               u"Информация: Файл пропущен {}".format(
                                   file_name)
                               )
                return
        else:
            sys.stdout.write(
                u"Не определен файл выходной файл {}".format(
                    self.__class__.__name__)
            )
