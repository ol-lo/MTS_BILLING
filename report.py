# -*- coding: utf-8 -*-
""" Интерфейc поиска в отчетах регулярных вырожений """

import codecs
import datetime
import os
import re
import sre_constants
import sys
from abc import abstractmethod

import PyPDF2

from constants import (
    ENCODE_TYPE, EMPTY_CHAR_FOR_REPLACE, READ_BINARY, ONE, CP_1251, TXT, UTF8,
    YYYY_MM_DD_FORMAT)


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
        u"Файл не может быть изменен, \n вероятно файд открыт"
    )


class LoggerAbstractMixin(object):
    u"""
    Класс Миксин логгера
    """
    INFO, WARNING, ERROR, DEBUG = range(0, 4)
    _LEVEL = {
        INFO: u'Информация',
        WARNING: u'Внимание',
        ERROR: u'Ошибка',
        DEBUG: u'Отладка'
    }

    def __init__(self):
        self._create_folder()

    @abstractmethod
    def _create_folder(self):
        u"""
        Метод создает цепочку из каталогов
        :return:
        """
        pass

    @abstractmethod
    def _write_log(self, level, message):
        u"""
        Метод записывает в файл
        и очищает содержимое
        :return:
        """
        pass


class ReportLoggerMixin(LoggerAbstractMixin):
    u"""
    Миксин для ведения журнала
    """
    SEPARATE_FOLDER = ''
    ROOT_FOLDER = ''
    FOLDER_IS_EXITS = False

    def _write_log(self, level, message):
        time = datetime.datetime.now().strftime(u'%H:%M:%S')
        string = u'{time:15}  {level:15}     {message:20} \r\n'.format(
            time=time,
            level=LoggerAbstractMixin._LEVEL.get(
                level, LoggerAbstractMixin.ERROR),
            message=message
        )
        with codecs.open(
                os.path.join(
                    self.path_to_dir_log,
                'a+', UTF8)
        ) as log_file:
            log_file.write(string)

    def _create_folder(self):
        self.path_to_dir_log = re.sub(
            '[\"\<>*\'?:]|',
            '',
            os.path.join(
                self.ROOT_FOLDER,
                getattr(self, self.SEPARATE_FOLDER),
                datetime.datetime.now().strftime(YYYY_MM_DD_FORMAT),
                TXT
            )
        )
        if os.path.exists(self.path_to_dir_log):
            return
        os.makedirs(self.path_to_dir_log)

    def write_log(self, level, message):
        if not self.FOLDER_IS_EXITS:
            self._create_folder()
        self._write_log(level, message)


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
            except sre_constants.error:
                raise BillingErrorPatternDate
            try:
                self.pattern_contract = re.compile(pattern_contract, re.UNICODE)
            except sre_constants.error:
                raise BillingErrorPatternContact

        except ValueError as error:
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
                    sys.stdout.write(u"Не обнаружен контракт \n")

                date_ = self.pattern_date.search(unicode_text)

                if date_ and not self.date_:
                    self.date_ = date_.group('date_')

                if not self.date_:
                    sys.stdout.write(u"Не обнаружена дата \n")

                for element_on_page in self.data_unit.findall(
                        unicode_text.decode(UTF8)):

                    groups = self.data_unit.match(element_on_page[0])
                    details = {}
                    for key, value in self.DICT_ARGUMENTS.items():
                        details.update({value: groups.group(key)})
                    self.result.append(details)
                if stop_flag:
                    sys.stdout.write(
                        u'\nСборка коллекции завершина\n '
                    )
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
        self.read_while_not_find_end_line(file_name)

        if self.out_file:
            try:
                self.write_(self.out_file, self.result)
            except ValueError:
                sys.stdout.write(u'Отсуствуют данные  \n')
                sys.stdout.write(u'файл пропущен {} \n'.format(self.out_file))
                return
        else:
            sys.stdout.write(
                u"Не определен файл выходной файл {}".format(
                    self.__class__.__name__)
            )
