import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QCompleter, QFileDialog, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QCheckBox
from PyQt5.QtCore import QDate
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMessageBox, QDialog, QFrame, QAction, QMenu

import json
import os

from docx import Document

from loguru import logger


from .tools.db import get_data
from .tools import main_window_functions, main_window_functions_excel

from .tools import dialogs
from .tools.dialogs import Excel
from .tools import layouts

from . import strings

# Variables
class VarBoxes:
    combo_boxes = {}
    text_boxes = {}

class Buttons:
    Buttons = {}
    CheckableButtons = {}

class App(QWidget):
    def __init__(self):
        super().__init__()

        # Logger
        logger.add("log.txt")
        logger.info('Start initUI')
        logger.debug(strings.start_text)

        # Получаем данные из config.py
        file_name = 'app\\tools\\data\\config.json'
        data = get_data(file_name=file_name)

        window_title = data['window_title']

        window_size = data['window_size']

        length_window = window_size['length']
        width_window = window_size['width']

        self.setWindowTitle(window_title)
        self.move(0,0)

        
        # Хранилища для widgets
        self.var_boxes = VarBoxes()
        self.buttons = Buttons()

        self.var_boxes_excel = VarBoxes()
        self.buttons_excel = Buttons()

        self.text_voltage = None
        self.text_frequency = None

        # Layouts
        self.main_layout = QHBoxLayout()

        self.WordLayout = layouts.Word.get_layout(self=self)
        self.ExcelLayout = layouts.Excel.get_layout(self=self)
        
        self.initUI()

    def initUI(self):

        # Устнавливаем размеры
        width = 40
        max_length = 700
        min_lenght = 300

        for widget in self.var_boxes.text_boxes.values():
            widget.setMaximumSize(max_length,width)
            widget.setMinimumSize(min_lenght,width)

        for widget in self.var_boxes.combo_boxes.values():
            widget.setMaximumSize(max_length,width)
            widget.setMinimumSize(min_lenght,width)

        for widget in self.buttons.Buttons.values():
            widget.setMaximumSize(max_length,width)
            widget.setMinimumSize(min_lenght,width)

        for widget in self.buttons.CheckableButtons.values():
            widget.setMaximumSize(max_length,width)
            widget.setMinimumSize(min_lenght,width)

        # Создаем Tab Widget
        MainTab = QTabWidget()

        # Создаем tabs
        word_tab = QWidget()
        excel_tab = QWidget()

        # Добавляем layouts в tab
        word_tab.setLayout(self.WordLayout)
        excel_tab.setLayout(self.ExcelLayout)

        # Добавляем tabs в Tab Widget
        MainTab.addTab(word_tab,'Создание протокола Word')
        MainTab.addTab(excel_tab,'Создание протокола Excel')

        # Добавляем TabWidget в main layout
        self.main_layout.addWidget(MainTab)

        # Добавляем main_layout в окно
        self.setLayout(self.main_layout)
        self.show()

    # Functions
    
    def search_company(self):
        main_window_functions.search_company(self=self)

    def scale_changed(self):
        main_window_functions.scale_changed(self=self)

    def get_selected_table(self):
        main_window_functions.get_selected_table(self=self)

    def verificationer_changed(self):
        main_window_functions.verificationer_changed(self=self)

    def create_protocol(self):
        main_window_functions.create_protocol(self=self)

    def use_data(self):
        main_window_functions.use_data(self=self)

    def clean(self):
        main_window_functions.clean(self=self)

    def create_template(self):
        main_window_functions.create_template(self=self)


    # Dialogs:
        
    def create_protocol_from_excel(self):
        logger.debug('Создание протокола из excel (funct)')

        dialogs.CreateProtocolFromExcelDialog(self).exec_()

        logger.debug('stop')
        
        
    def show_inspection_address_setting(self):
        logger.debug('start')

        dialogs.ChooseInspectionAddressDialog(self).exec_()

        logger.debug('end')
        
    def choose_scale(self):
        logger.info('start')

        dialogs.ChooseScaleDialog(self).exec_()

        logger.debug('end')
        
    def add_path(self):
        logger.info('start')

        dialog = QDialog()
        self.text_path.setPlainText(QFileDialog.getExistingDirectory(dialog, "Выберите папку"))

        logger.debug('end')
        
    def settings(self):
        logger.debug('start')

        dialogs.SettingDialog().exec_()

        logger.debug('end')

    def add_path_to_excel(self):
        logger.info('start')

        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getOpenFileName(self, 'Выберите файл', '', 'All Files (*);;Text Files (*.txt)', options=options)

        self.text_path_to_excel.setPlainText(filePath)

        logger.debug('end')


    # Excel Protocol
        
    def search_company_excel(self):
        main_window_functions_excel.search_company(self=self)

    def get_selected_table_excel(self):
        main_window_functions_excel.get_selected_table(self=self)

    def verificationer_changed_excel(self):
        main_window_functions_excel.verificationer_changed(self=self)

    def create_protocol_excel(self):
        main_window_functions_excel.create_protocol(self=self)

    def use_data_excel(self):
        main_window_functions_excel.use_data(self=self)

    def clean_excel(self):
        main_window_functions_excel.clean(self=self)

    def create_template_excel(self):
        main_window_functions_excel.create_template(self=self)


    # Dialogs:
        
    def create_protocol_from_excel_excel(self):
        logger.debug('Создание протокола из excel (funct)')

        Excel.CreateProtocolFromExcelDialog(self).exec_()

        logger.debug('stop')
        
        
    def show_inspection_address_setting_excel(self):
        logger.debug('start')

        Excel.ChooseInspectionAddressDialog(self).exec_()

        logger.debug('end')
        
    def choose_scale_excel(self):
        logger.info('start')

        Excel.ChooseScaleDialog(self).exec_()

        logger.debug('end')
        
    def add_path_excel(self):
        logger.info('start')

        dialog = QDialog()
        self.text_path.setPlainText(QFileDialog.getExistingDirectory(dialog, "Выберите папку"))

        logger.debug('end')
        
    def settings_excel(self):
        logger.debug('start')

        Excel.SettingDialog().exec_()

        logger.debug('end')

    def add_path_to_excel_excel(self):
        logger.info('start')

        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getOpenFileName(self, 'Выберите файл', '', 'All Files (*);;Text Files (*.txt)', options=options)

        self.text_path_to_excel.setPlainText(filePath)

        logger.debug('end')


    # Events

    def resizeEvent(self, event):
        print("Window has been resized")
        super().resizeEvent(event)