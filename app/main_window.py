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
from .tools import main_window_functions

from .tools import dialogs
from .tools import layouts

from . import strings

# Variables
class VarBoxes:
    combo_boxes = {}
    text_boxes = {}
    labels = {}

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

        # Отображаем окно в полноэкранном режиме
        # Баг
        #self.setMaximumSize(length_window, width_window)

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
        company_name, legal_address = main_window_functions.search_company(self=self)
        
        # Устанавливаем в QPlainTextEdit
        self.text_company.setPlainText(company_name)
        
        # Устанавливаем в QPlainTextEdit
        self.text_legal_address.setPlainText(legal_address)
        
    def search_company_excel(self):
        company_name, legal_address = main_window_functions.search_company(self=self)
        
        # Устанавливаем в QPlainTextEdit
        self.text_company_excel.setPlainText(company_name)

    def scale_changed(self):
        main_window_functions.scale_changed(self=self)

    def verificationer_changed(self):
        result = main_window_functions.verificationer_changed(self=self, verificationer_combo=self.verificationer_combo)
        self.text_num_protocol.setPlainText = result
        
    def verificationer_changed_excel(self):
        result = main_window_functions.verificationer_changed(self=self, verificationer_combo=self.verificationer_combo_excel)
        self.text_num_protocol_excel.setPlainText = result

    def create_protocol(self):
        word = True
        
        main_window_functions.create_protocol(self=self, 
                                                word=word, 
                                                var_boxes=self.var_boxes, 
                                                buttons=self.buttons, 
                                                inspection_date=self.inspection_date)
        
    def create_protocol_excel(self):
        word = False
        
        main_window_functions.create_protocol(self=self, 
                                                word=word, 
                                                var_boxes=self.var_boxes_excel, 
                                                buttons=self.buttons_excel, 
                                                inspection_date=self.inspection_date_excel)

    def use_data(self):
        main_window_functions.use_data(self=self,
                                        tab_widget=self.tab_standarts, 
                                        var_boxes=self.var_boxes, 
                                        buttons=self.buttons, 
                                        inspection_date=self.inspection_date)
        
    def use_data_excel(self):
        main_window_functions.use_data(self=self,
                                        tab_widget=self.tab_standarts_excel, 
                                        var_boxes=self.var_boxes_excel, 
                                        buttons=self.buttons_excel, 
                                        inspection_date=self.inspection_date_excel)

    def clean(self):
        main_window_functions.clean(self=self,
                                    var_boxes=self.var_boxes, 
                                    buttons=self.buttons)
        
    def clean_excel(self):
        main_window_functions.clean(self=self,
                                    var_boxes=self.var_boxes_excel, 
                                    buttons=self.buttons_excel)

    def create_template(self):
        main_window_functions.create_template(self=self,
                                                word=True)
        
    def create_template_excel(self):
        main_window_functions.create_template(self=self,
                                                word=False)


    # Dialogs:
        
    def create_protocol_from_excel(self):
        logger.debug('Создание протокола из excel (funct)')

        dialogs.CreateProtocolFromExcelDialog(self, word=True).exec_()

        logger.debug('stop')
        
    def create_protocol_from_excel_excel(self):
        logger.debug('Создание протокола из excel (funct)')

        dialogs.CreateProtocolFromExcelDialog(self, word=False).exec_()

        logger.debug('stop')
        
    def show_inspection_address_setting(self):
        logger.debug('start')

        dialogs.ChooseInspectionAddressDialog(self,inspection=self.text_inspection_address, legal=self.text_legal_address, word=True).exec_()

        logger.debug('end')
        
    def show_inspection_address_setting_excel(self):
        logger.debug('start')

        dialogs.ChooseInspectionAddressDialog(self,inspection=self.text_inspection_address_excel, legal=self.text_legal_address_excel, word=False).exec_()

        logger.debug('end')
        
    def choose_scale(self):
        logger.info('start')

        dialogs.ChooseScaleDialog(text_scale_widget=self.text_scale,word=True).exec_()

        logger.debug('end')
        
    def choose_scale_excel(self):
        logger.info('start')

        dialogs.ChooseScaleDialog(text_scale_widget=self.text_scale_excel,word=False).exec_()

        logger.debug('end')
        
    def add_path(self):
        logger.info('start')

        dialog = QDialog()
        self.text_path.setPlainText(QFileDialog.getExistingDirectory(dialog, "Выберите папку"))

        logger.debug('end')
        
    def add_path_excel(self):
        logger.info('start')

        dialog = QDialog()
        self.text_path_excel.setPlainText(QFileDialog.getExistingDirectory(dialog, "Выберите папку"))

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
        
    def add_path_to_excel_excel(self):
        logger.info('start')

        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getOpenFileName(self, 'Выберите файл', '', 'All Files (*);;Text Files (*.txt)', options=options)

        self.text_path_to_excel_journal_excel.setPlainText(filePath)

        logger.debug('end')

    # Events

    def resizeEvent(self, event):
        print("Window has been resized")
        super().resizeEvent(event)