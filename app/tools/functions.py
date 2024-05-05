import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QCompleter, QFileDialog, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QCheckBox
from PyQt5.QtCore import QDate
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMessageBox, QDialog

import json
import os

from loguru import logger

from ..db import Config

from .parse.__init__ import FNS_API, List_org

""" OPEN CONFIG """
file_path = 'data/config.json'
if os.path.exists(file_path):
    config_client = Config(file_path)
    config = config_client.get()
else:
    raise Exception(f'File {file_path} not found')

def search_company(inn):
    logger.debug('Поиск компании...')

    FNS_TOKEN = config['parse']['FNS_TOKEN']

    # Создаем обьект компании
    company = FNS_API(FNS_TOKEN)

    # Получаем INN
    INN = inn
    
    # Получаем данные о компании по INN
    data = company.get_company_data(INN)

    if data is not True and len(data['items']) > 0:

        # Получаем данные о названии комании
        company_name = data['items'][0]['ЮЛ']['НаимСокрЮЛ']

        # Получаем данные о юр.адресе компании
        legal_address = data['items'][0]['ЮЛ']['АдресПолн']
        
        logger.debug('end')
        
        return company_name, legal_address


    else:

        LIST_ORG = List_org()

        try:

            data = LIST_ORG.get_company_name_by_inn(INN)

            company_name = data['name']
            legal_address = data['address']
            
            return company_name, legal_address

        except Exception as e:
            logger.error(f'Ошибка при работе с List-org.ru: {e}')


            message_box = QMessageBox()
            message_box.setIcon(QMessageBox.Critical)
            message_box.setText("Не удалось получить данные о компании!\nПроверьте введенный ИНН")
            message_box.setWindowTitle("Ошибка")
            message_box.setStandardButtons(QMessageBox.Ok)
            message_box.exec_()