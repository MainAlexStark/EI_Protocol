import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QCompleter, QFileDialog, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QCheckBox
from PyQt5.QtCore import QDate
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMessageBox, QDialog, QFrame

import json
import os

from docx import Document

from loguru import logger


from .tools.db import get_data
from .tools import main_window_functions

from .tools import dialogs

from . import strings



class App(QWidget):
    def __init__(self):
        super().__init__()
        self.text_scale = None
        self.label_text_scale = None
        self.button_choose_scale = None
        self.label_text_frequency = None
        self.label_text_voltage = None
        self.text_voltage = None
        self.text_frequency = None
        self.tab_widget = None
        self.button_settings = None
        self.button_clean = None
        self.button_use_data = None
        self.button_use_excel = None
        self.button_create_protocol = None
        self.button_create_template = None
        self.text_pressure = None
        self.label_text_pressure = None
        self.text_humidity = None
        self.label_text_humidity = None
        self.text_temperature = None
        self.label_text_temperature = None
        self.inspection_date = None
        self.button_unfit = None
        self.text_inspection_address = None
        self.label_text_inspection_address = None
        self.text_legal_address = None
        self.label_text_legal_address = None
        self.text_INN = None
        self.label_text_INN = None
        self.text_company = None
        self.label_text_company = None
        self.verificationer_combo = None
        self.label_verificationer_combo = None
        self.text_num_scale = None
        self.label_text_num_scale = None
        self.text_num_protocol = None
        self.label_text_num_protocol = None
        self.button_path = None
        self.text_path = None
        self.label_text_path = None
        self.initUI()

    def initUI(self):

        # Logger
        logger.add("log.txt")
        logger.info('Start initUI')
        logger.debug(strings.start_text)


        # Получаем данные из config.py
        file_name = 'app/tools/data/config.json'
        data = get_data(file_name=file_name)

        window_title = data['window_title']

        window_size = data['window_size']

        length_window = window_size['length']
        width_window = window_size['width']

        self.setWindowTitle(window_title)
        #self.setGeometry(0, 0, length_window, width_window)
        self.move(0,0)

        # Layouts
        # Создание layouts ( контейнеры в которые погружаются остальные элементы )
        main_layout = QHBoxLayout()
        self.var_layout = QVBoxLayout()
        self.var_r_layout = QVBoxLayout()
        self.var_r2_layout = QVBoxLayout()

        self.var_layout.setContentsMargins(0, 10, 0, 10)
        self.var_r_layout.setContentsMargins(0, 10, 0, 10)

        # Variables
        class VarBoxes:
            combo_boxes = {}
            text_boxes = {}

        class Buttons:
            Buttons = {}
            CheckableButtons = {}

        self.var_boxes = VarBoxes()
        self.buttons = Buttons()

        # Инициализируем элементы ввода

        # Scale

        self.text_scale = QPlainTextEdit(self)
        self.text_scale.setPlaceholderText('Весы')

        self.text_scale.textChanged.connect(self.scale_changed)

        self.var_layout.addWidget(self.text_scale)

        self.var_boxes.text_boxes['scale'] = self.text_scale

        # Надпись
        self.button_choose_scale = QPushButton('Выбрать весы', self)
        self.button_choose_scale.clicked.connect(self.choose_scale)  # Привязываем функцию
        self.var_layout.addWidget(self.button_choose_scale)  # Добавляем в layout

        self.buttons.Buttons['choose_scale'] = self.button_choose_scale

        # path
        # Текстовое поле
        self.text_path = QPlainTextEdit(self)
        self.text_path.setPlaceholderText('Путь сохранения')
        # Добавляем элементы в layouts
        self.var_layout.addWidget(self.text_path)

        self.var_boxes.text_boxes['path'] = self.text_path
        
        # Создаем кнопку
        self.button_path = QPushButton('Выбрать путь сохранения', self)
        self.button_path.clicked.connect(self.add_path)  # Привязываем функцию
        self.var_layout.addWidget(self.button_path)  # Добавляем в layout

        self.buttons.Buttons['path'] = self.button_path
        
        
        # Work place num
        
        work_places = [
            '01 - Манометры',
            '02 - Гири эталонные и общего назначения',
            '03 - Весы',
            '04 - Счетчики воды',
            '05 - Автоматические ВСУ',
            '06 - Влагомеры',
            '07 - Гидрометры',
            '08 - Дозаторы обьемные',
            '09 - Дозаторы весовые'
        ]
        
        self.label_work_place_combo = QLabel("Выберите рабочее место:")
        self.work_place_combo = QComboBox()
        self.work_place_combo.addItems(work_places)
        self.work_place_combo.setCurrentIndex(-1)
        self.var_layout.addWidget(self.label_work_place_combo)
        self.var_layout.addWidget(self.work_place_combo)

        # Добавляем в словарь combo boxes
        self.var_boxes.combo_boxes['work_place'] = self.work_place_combo

        # num_protocol

        # Текстовое поле
        self.text_num_protocol = QPlainTextEdit(self)
        self.text_num_protocol.setPlaceholderText('Номер протокола')
        # Добавляем элементы в layouts
        self.var_layout.addWidget(self.text_num_protocol)

        self.var_boxes.text_boxes['num_protocol'] = self.text_num_protocol

        # num_scale
        self.text_num_scale = QPlainTextEdit(self)
        self.text_num_scale.setPlaceholderText('Номер весов')
        self.var_layout.addWidget(self.text_num_scale)

        self.var_boxes.text_boxes['num_scale'] = self.text_num_scale

        # Verificationer
        # Определите имя файла хранилища
        file_name = 'config.json'
        data = get_data('app/tools/data/config.json')

        # Массив с поверителями
        ver = []


        verificationers = data.get('verificationers', [])
        verificationers_text = ''
        for verificationer in verificationers.keys():
            verificationers_text += verificationer
        verificationers_text = '\n'.join(verificationers)

        for i in verificationers_text.split('\n'): ver.append(i)

        self.label_verificationer_combo = QLabel("Выберите поверителя:")
        self.verificationer_combo = QComboBox()
        self.verificationer_combo.addItems(verificationers)
        self.verificationer_combo.setCurrentIndex(-1)
        self.verificationer_combo.setFixedSize(500, 40)
        self.verificationer_combo.currentTextChanged.connect(self.verificationer_changed)
        self.var_layout.addWidget(self.label_verificationer_combo)
        self.var_layout.addWidget(self.verificationer_combo)

        # Добавляем в словарь combo boxes
        self.var_boxes.combo_boxes['verificationer'] = self.verificationer_combo


        # Line

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setMinimumSize(300,2)
        line.setMaximumSize(700,2)

        self.var_layout.addWidget(line)

        # INN
        self.text_INN = QPlainTextEdit(self)
        self.text_INN.setPlaceholderText("ИНН")
        self.var_layout.addWidget(self.text_INN)

        self.var_boxes.text_boxes['INN'] = self.text_INN

        self.button_search_company = QPushButton('Найти', self)
        self.button_search_company.clicked.connect(self.search_company)  # Привязываем функцию
        self.var_layout.addWidget(self.button_search_company)  # Добавляем в layout

        self.buttons.Buttons['search'] = self.button_search_company

        # Company
        self.text_company = QPlainTextEdit(self)
        self.text_company.setPlaceholderText("Компания")
        self.var_layout.addWidget(self.text_company)

        self.var_boxes.text_boxes['company'] = self.text_company

        # Legal address

        self.text_legal_address = QPlainTextEdit(self)
        self.text_legal_address.setPlaceholderText("Юридический адрес")
        self.var_layout.addWidget(self.text_legal_address)

        self.var_boxes.text_boxes['legal_address'] = self.text_legal_address


        # inspection address

        inspection_address_layout = QHBoxLayout()

        self.text_inspection_address = QPlainTextEdit(self)
        self.text_inspection_address.setPlaceholderText('Адрес поверки')
        inspection_address_layout.addWidget(self.text_inspection_address)

        self.var_boxes.text_boxes['inspection_address'] = self.text_inspection_address

        self.button_inspection_address_setting = QPushButton('...', self)
        self.button_inspection_address_setting.setFixedSize(50, 50)
        self.button_inspection_address_setting.clicked.connect(self.show_inspection_address_setting)  # Привязываем функцию
        inspection_address_layout.addWidget(self.button_inspection_address_setting)  # Добавляем в layout

        self.var_layout.addLayout(inspection_address_layout)

        # Unfit

        self.button_unfit = QPushButton('Несоответсвует', self)
        self.button_unfit.setCheckable(True)
        self.var_layout.addWidget(self.button_unfit)  # Добавляем в layout

        self.buttons.CheckableButtons['unfit'] = self.button_unfit

        # Line

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setMinimumSize(300,2)
        line.setMaximumSize(700,2)

        self.var_layout.addWidget(line)

        # PAth to excel 

        self.path_excel_layout = QHBoxLayout()

        self.text_path_to_excel = QPlainTextEdit(self)
        self.text_path_to_excel.setPlaceholderText("Путь к журналу excel")
        self.path_excel_layout.addWidget(self.text_path_to_excel)

        self.var_boxes.text_boxes['path_to_excel'] = self.text_path_to_excel

        self.button_path_to_excel_dialog = QPushButton('...', self)
        self.button_path_to_excel_dialog.setFixedSize(50, 50)
        self.button_path_to_excel_dialog.clicked.connect(self.add_path_to_excel)  # Привязываем функцию
        self.path_excel_layout.addWidget(self.button_path_to_excel_dialog)  # Добавляем в layout

        self.var_layout.addLayout(self.path_excel_layout)

        # Use excel

        use_excel_layout = QHBoxLayout()

        self.button_use_excel = QPushButton('Использовать Excel', self)
        self.button_use_excel.setCheckable(True)
        self.var_layout.addWidget(self.button_use_excel)  # Добавляем в layout

        self.buttons.CheckableButtons['use_excel'] = self.button_use_excel

        ####### var_r_layout

        # inspection_date

        self.inspection_date = QCalendarWidget(self)
        self.inspection_date.setFixedSize(500, 300)
        self.var_r_layout.addWidget(self.inspection_date)

        # interval

        nums = ['1','2']

        self.label_interval_combo = QLabel("Выберите интервал до следующей проверки:")
        self.interval_combo = QComboBox()
        self.interval_combo.addItems(nums)
        self.interval_combo.setCurrentIndex(0)
        self.var_r_layout.addWidget(self.label_interval_combo)
        self.var_r_layout.addWidget(self.interval_combo)

        # Добавляем в словарь combo boxes
        self.var_boxes.combo_boxes['interval'] = self.interval_combo


        # weather

        self.text_temperature = QPlainTextEdit(self)
        self.text_temperature.setPlaceholderText("Температура")

        self.var_r_layout.addWidget(self.text_temperature)

        self.var_boxes.text_boxes['temperature'] = self.text_temperature

        self.text_humidity = QPlainTextEdit(self)
        self.text_humidity.setPlaceholderText("Влажность")
        self.var_r_layout.addWidget(self.text_humidity)

        self.var_boxes.text_boxes['humidity'] = self.text_humidity

        self.text_pressure = QPlainTextEdit(self)
        self.text_pressure.setPlaceholderText("Давление")
        self.var_r_layout.addWidget(self.text_pressure)

        self.var_boxes.text_boxes['pressure'] = self.text_pressure

    

        # Создаем кнопки

        # Create template

        self.button_create_template = QPushButton('Создать шаблон', self)
        self.button_create_template.clicked.connect(self.create_template)  # Привязываем функцию
        self.var_r_layout.addWidget(self.button_create_template)  # Добавляем в layout

        self.buttons.Buttons['create_template'] = self.button_create_template

        # Create protocol

        self.button_create_protocol = QPushButton('Создать протокол', self)
        self.button_create_protocol.clicked.connect(self.create_protocol)  # Привязываем функцию
        self.var_r_layout.addWidget(self.button_create_protocol)  # Добавляем в layout

        self.buttons.Buttons['create_protocol'] = self.button_create_protocol

        # Line

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setMinimumSize(300,2)
        line.setMaximumSize(700,2)

        self.var_r_layout.addWidget(line)

        # Use data

        self.button_use_data = QPushButton('Использовать преведущие данные', self)
        self.button_use_data.clicked.connect(self.use_data)  # Привязываем функцию
        self.var_r_layout.addWidget(self.button_use_data)  # Добавляем в layout

        self.buttons.Buttons['use_data'] = self.button_use_data

        # Clean

        self.button_clean = QPushButton('Очистить', self)
        self.button_clean.clicked.connect(self.clean)  # Привязываем функцию
        self.var_r_layout.addWidget(self.button_clean)  # Добавляем в layout

        self.buttons.Buttons['clean'] = self.button_clean

        # Settings

        self.button_settings = QPushButton('Настройки', self)
        self.button_settings.clicked.connect(self.settings)  # Привязываем функцию
        self.var_r_layout.addWidget(self.button_settings)  # Добавляем в layout

        self.buttons.Buttons['settings'] = self.button_settings

        # 2 layout

        # Standarts

        self.tab_widget = QTabWidget(self)
        self.tab_widget.setMaximumSize(900, 950)
        self.tab_widget.setMinimumSize(850, 800)

        # Укажите путь к нужной папке
        folder_path = f'app/standarts'
        file_names = os.listdir(folder_path)

        for file_name in file_names:
            tab = QWidget()
            Qtable = QTableWidget(tab)

            # Настройте таблицу
            Qtable.setSelectionBehavior(QAbstractItemView.SelectRows)  # Выбор целой строки
            Qtable.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Запрещаем редактирование
            # Добавление столбцов и строк в таблицу
            # Открываем docx документ
            document = Document(f"app/standarts/{file_name}")

            # Получаем первую таблицу из документа
            table = document.tables[0]

            # Устанавливаем количество строк и столбцов в QTableWidget
            Qtable.setRowCount(len(table.rows))
            Qtable.setColumnCount(len(table.columns))

            # Копируем содержимое таблицы в QTableWidget
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    Qtable.setItem(i, j, QTableWidgetItem(cell.text))

            # Устанавливаем ширину для первого столбца
            Qtable.setColumnWidth(0, 400)

            layout = QVBoxLayout(tab)
            layout.addWidget(Qtable)
            tab.setLayout(layout)

            self.tab_widget.addTab(tab, file_name)

        self.var_r2_layout.addWidget(self.tab_widget)




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

        # Add layouts in main_layout
        main_layout.addLayout(self.var_layout)
        main_layout.addLayout(self.var_r_layout)
        main_layout.addLayout(self.var_r2_layout)

        # Добавляем main_layout в окно
        self.setLayout(main_layout)
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