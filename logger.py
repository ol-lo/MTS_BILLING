# -*- coding: utf-8 -*-
""" Интерфейс журналирования приложения  """

import codecs
import datetime
import os
import sys
from abc import abstractmethod
from os import path

from constants import UTF8, YYYY_MM_DD_FORMAT


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
    # TODO: EXCEPT ДЛЯ Logger
    LOG_FOLDER = 'LOGS'
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
                self.path_to_file_log + '\\LOGGER.TXT',
                'a+', 
                UTF8) as log_file:
            log_file.write(string)

    def _create_folder(self):
        root = path.dirname(path.dirname(sys.argv[0]))
        root = root or path.dirname(path.dirname(path.abspath(__file__)))
        self.ROOT_FOLDER = root

        self.path_to_file_log = os.path.join(
            self.ROOT_FOLDER,
            self.LOG_FOLDER,
            datetime.datetime.now().strftime(YYYY_MM_DD_FORMAT)
        )
        if os.path.exists(self.path_to_file_log):
            self.FOLDER_IS_EXITS = True
            return
        os.makedirs(self.path_to_file_log)

    def write_log(self, level, message):
        if not self.FOLDER_IS_EXITS:
            self._create_folder()
        self._write_log(level, message)
