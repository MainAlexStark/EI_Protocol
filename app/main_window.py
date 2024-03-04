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


        # Создаем TabWidget
        self.MainTabWidget = QTabWidget(self)

        # Создаем вкладки
        self.WordTab = QWidget()
        self.ExcelTab = QWidget()

        # Добавляем layouts с виджетами в вкладки
        self.WordTab.setLayout(self.WordLayout)
        self.ExcelTab.setLayout(self.ExcelLayout)

        # Добавляем вкладки в TabWidget
        self.MainTabWidget.addTab(self.WordTab, 'Создание протокола word')
        self.MainTabWidget.addTab(self.ExcelTab, 'Создание протокола excel')

        # Добавляем TabWidget в main layout
        self.main_layout.addWidget(self.MainTabWidget)

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


    # Events

    def resizeEvent(self, event):
        print("Window has been resized")
        super().resizeEvent(event)