import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QCompleter, QFileDialog, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QCheckBox
from PyQt5.QtCore import QDate
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMessageBox, QDialog, QFrame, QAction, QMenu

from ...tools.db import get_data
from ...tools import functions
from ...tools import dialogs

from docx import Document
from loguru import logger

import os

main_window = None

def get_layout(self):
    logger.debug('Get Word layout')

    global main_window
    main_layout = QHBoxLayout()

    var_layout = QVBoxLayout()
    var_r_layout = QVBoxLayout()
    var_r2_layout = QVBoxLayout()

    main_window = self

    # Инициализируем элементы ввода

    # Scale

    text_scale = QPlainTextEdit(self)
    text_scale.setPlaceholderText('Весы')

    text_scale.textChanged.connect(scale_changed)

    var_layout.addWidget(text_scale)

    self.text_boxes_word['scale'] = text_scale

    # Надпись
    choose_scale_button = QPushButton('Выбрать весы', self)
    choose_scale_button.clicked.connect(choose_scale)  # Привязываем функцию
    var_layout.addWidget(choose_scale_button)  # Добавляем в layout

    self.buttons_word['choose_scale_button'] = choose_scale_button

    # path
    # Текстовое поле
    text_path = QPlainTextEdit(self)
    text_path.setPlaceholderText('Путь сохранения')
    # Добавляем элементы в layouts
    var_layout.addWidget(text_path)

    self.text_boxes_word['save_path'] = text_path
    
    # Создаем кнопку
    path_button = QPushButton('Выбрать путь сохранения', self)
    path_button.clicked.connect(add_save_path)  # Привязываем функцию
    var_layout.addWidget(path_button)  # Добавляем в layout

    self.buttons_word['path_button'] = path_button
    
    
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
    
    label_work_place_combo = QLabel("Выберите рабочее место:")
    work_place_combo = QComboBox()
    work_place_combo.addItems(work_places)
    work_place_combo.setCurrentIndex(-1)
    var_layout.addWidget(label_work_place_combo)
    var_layout.addWidget(work_place_combo)

    # Добавляем в словарь combo boxes
    self.combo_boxes_word['work_place_combo'] = work_place_combo

    # num_protocol

    # Текстовое поле
    text_num_protocol = QPlainTextEdit(self)
    text_num_protocol.setPlaceholderText('Номер протокола')
    # Добавляем элементы в layouts
    var_layout.addWidget(text_num_protocol)

    self.text_boxes_word['num_protocol'] = text_num_protocol

    # num_scale
    text_num_scale = QPlainTextEdit(self)
    text_num_scale.setPlaceholderText('Номер весов')
    var_layout.addWidget(text_num_scale)

    self.text_boxes_word['num_scale'] = text_num_scale

    # Verificationer
    # Определите имя файла хранилища
    path_to_data = 'app/tools/data/config.json'
    data = get_data(path_to_data)

    # Массив с поверителями
    ver = []


    verificationers = data.get('verificationers', [])
    verificationers_text = ''
    for verificationer in verificationers.keys():
        verificationers_text += verificationer
    verificationers_text = '\n'.join(verificationers)

    for i in verificationers_text.split('\n'): ver.append(i)

    label_verificationer_combo = QLabel("Выберите поверителя:")
    verificationer_combo = QComboBox()
    verificationer_combo.addItems(verificationers)
    verificationer_combo.setCurrentIndex(-1)
    verificationer_combo.setFixedSize(500, 40)
    verificationer_combo.currentTextChanged.connect(verificationer_changed)
    var_layout.addWidget(label_verificationer_combo)
    var_layout.addWidget(verificationer_combo)

    # Добавляем в словарь combo boxes
    self.combo_boxes_word['verificationer_combo'] = verificationer_combo


    # Line

    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setMinimumSize(300,2)
    line.setMaximumSize(700,2)

    var_layout.addWidget(line)

    # INN
    text_INN = QPlainTextEdit(self)
    text_INN.setPlaceholderText("ИНН")
    var_layout.addWidget(text_INN)

    self.text_boxes_word['INN'] = text_INN

    search_company_button = QPushButton('Найти', self)
    search_company_button.clicked.connect(search_company)  # Привязываем функцию
    var_layout.addWidget(search_company_button)  # Добавляем в layout

    self.buttons_word['search_company_button'] = search_company_button

    # Company
    text_company = QPlainTextEdit(self)
    text_company.setPlaceholderText("Компания")
    var_layout.addWidget(text_company)

    self.text_boxes_word['company'] = text_company

    # Legal address

    text_legal_address = QPlainTextEdit(self)
    text_legal_address.setPlaceholderText("Юридический адрес")
    var_layout.addWidget(text_legal_address)

    self.text_boxes_word['legal_address'] = text_legal_address


    # inspection address

    inspection_address_layout = QHBoxLayout()

    text_inspection_address = QPlainTextEdit(self)
    text_inspection_address.setPlaceholderText('Адрес поверки')
    inspection_address_layout.addWidget(text_inspection_address)

    self.text_boxes_word['inspection_address'] = text_inspection_address

    button_inspection_address_setting = QPushButton('...', self)
    button_inspection_address_setting.setFixedSize(50, 50)
    button_inspection_address_setting.clicked.connect(show_inspection_address_setting)  # Привязываем функцию
    inspection_address_layout.addWidget(button_inspection_address_setting)  # Добавляем в layout

    var_layout.addLayout(inspection_address_layout)

    # Unfit

    unfit_button = QPushButton('Несоответсвует', self)
    unfit_button.setCheckable(True)
    var_layout.addWidget(unfit_button)  # Добавляем в layout

    self.checkable_buttons_word['unfit_button'] = unfit_button

    # Line

    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setMinimumSize(300,2)
    line.setMaximumSize(700,2)

    var_layout.addWidget(line)

    # PAth to excel

    path_excel_layout = QHBoxLayout()

    text_path_to_excel_journal = QPlainTextEdit(self)
    text_path_to_excel_journal.setPlaceholderText("Путь к журналу excel")
    path_excel_layout.addWidget(text_path_to_excel_journal)

    self.text_boxes_word['path_to_excel_jounal'] = text_path_to_excel_journal

    path_to_excel_dialog_button = QPushButton('...', self)
    path_to_excel_dialog_button.setFixedSize(50, 50)
    path_to_excel_dialog_button.clicked.connect(add_path_to_excel)  # Привязываем функцию
    path_excel_layout.addWidget(path_to_excel_dialog_button)  # Добавляем в layout

    var_layout.addLayout(path_excel_layout)

    # Use excel

    add_to_excel_button = QPushButton('Добавить протокол в Excel', self)
    add_to_excel_button.setCheckable(True)
    var_layout.addWidget(add_to_excel_button)  # Добавляем в layout

    self.checkable_buttons_word['add_to_excel_button'] = add_to_excel_button

    ####### var_r_layout

    # inspection_date

    inspection_date = QCalendarWidget(self)
    inspection_date.setFixedSize(500, 300)
    var_r_layout.addWidget(inspection_date)

    self.other_widgets_word['inspection_date'] = inspection_date

    # interval

    nums = ['1','2']

    label_interval_combo = QLabel("Выберите интервал до следующей проверки:")
    interval_combo = QComboBox()
    interval_combo.addItems(nums)
    interval_combo.setCurrentIndex(0)
    var_r_layout.addWidget(label_interval_combo)
    var_r_layout.addWidget(interval_combo)

    # Добавляем в словарь combo boxes
    self.combo_boxes_word['interval_combo'] = interval_combo


    # weather

    text_temperature = QPlainTextEdit(self)
    text_temperature.setPlaceholderText("Температура")

    var_r_layout.addWidget(text_temperature)

    self.text_boxes_word['temperature'] = text_temperature

    text_humidity = QPlainTextEdit(self)
    text_humidity.setPlaceholderText("Влажность")
    var_r_layout.addWidget(text_humidity)

    self.text_boxes_word['humidity'] = text_humidity

    text_pressure = QPlainTextEdit(self)
    text_pressure.setPlaceholderText("Давление")
    var_r_layout.addWidget(text_pressure)

    self.text_boxes_word['pressure'] = text_pressure



    # Создаем кнопки

    # Create template

    create_template_button = QPushButton('Создать шаблон', self)
    create_template_button.clicked.connect(create_template)  # Привязываем функцию
    var_r_layout.addWidget(create_template_button)  # Добавляем в layout

    self.buttons_word['create_template_button'] = create_template_button

    # Create protocol

    create_protocol_button = QPushButton('Создать протокол', self)
    create_protocol_button.clicked.connect(create_protocol)  # Привязываем функцию
    var_r_layout.addWidget(create_protocol_button)  # Добавляем в layout

    self.buttons_word['create_protocol_button'] = create_protocol_button

    # Create protocol from excel

    create_protocol_from_excel_button = QPushButton('Создать протокол из excel', self)
    create_protocol_from_excel_button.clicked.connect(create_protocol_from_excel)  # Привязываем функцию
    var_r_layout.addWidget(create_protocol_from_excel_button)  # Добавляем в layout

    self.buttons_word['create_protocol_from_excel_button'] = create_protocol_from_excel_button

    # Line

    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setMinimumSize(300,2)
    line.setMaximumSize(700,2)

    var_r_layout.addWidget(line)

    # Use data

    use_data_button = QPushButton('Использовать преведущие данные', self)
    use_data_button.clicked.connect(use_data)  # Привязываем функцию
    var_r_layout.addWidget(use_data_button)  # Добавляем в layout

    self.buttons_word['use_data_button'] = use_data_button

    # Clean

    clean_button = QPushButton('Очистить', self)
    clean_button.clicked.connect(clean)  # Привязываем функцию
    var_r_layout.addWidget(clean_button)  # Добавляем в layout

    self.buttons_word['clean_button'] = clean_button

    # Settings

    settings_button = QPushButton('Настройки', self)
    settings_button.clicked.connect(settings)  # Привязываем функцию
    var_r_layout.addWidget(settings_button)  # Добавляем в layout

    self.buttons_word['settings_button'] = settings_button

    # 2 layout

    # Standarts

    tab_standarts = QTabWidget(self)
    tab_standarts.setMaximumSize(900, 950)
    tab_standarts.setMinimumSize(850, 800)

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

        tab_standarts.addTab(tab, file_name)

    var_r2_layout.addWidget(tab_standarts)

    self.other_widgets_word['tab_standarts'] = tab_standarts

    # Устнавливаем размеры
    width = 40
    max_length = 700
    min_lenght = 300

    for widget in self.text_boxes_word.values():
        widget.setMaximumSize(max_length,width)
        widget.setMinimumSize(min_lenght,width)

    for widget in self.combo_boxes_word.values():
        widget.setMaximumSize(max_length,width)
        widget.setMinimumSize(min_lenght,width)

    for widget in self.buttons_word.values():
        widget.setMaximumSize(max_length,width)
        widget.setMinimumSize(min_lenght,width)

    for widget in self.checkable_buttons_word.values():
        widget.setMaximumSize(max_length,width)
        widget.setMinimumSize(min_lenght,width)

    # Add layouts in main_layout
    main_layout.addLayout(var_layout)
    main_layout.addLayout(var_r_layout)
    main_layout.addLayout(var_r2_layout)

    return main_layout





# FUNCTIONS

def search_company():
    self = main_window
    company_name, legal_address = functions.search_company(self=self, inn=self.text_boxes_word['INN'].toPlainText())
        
    # Устанавливаем в QPlainTextEdit
    self.text_boxes_word['company'].setPlainText(company_name)
    
    # Устанавливаем в QPlainTextEdit
    self.text_boxes_word['legal_address'].setPlainText(legal_address)

def scale_changed():
    self = main_window
    functions.scale_changed(self=self)

def verificationer_changed():
    self = main_window
    result = functions.verificationer_changed(self=self, verificationer_combo=self.combo_boxes_word['verificationer_combo'])
    self.text_boxes_word['num_protocol'].setPlainText = result

def create_protocol():
    self = main_window
    functions.create_protocol(self=self, 
                                word=True, 
                                other_widgets=self.other_widgets_word,
                                text_boxes=self.text_boxes_word,
                                combo_boxes=self.combo_boxes_word,
                                checkable_buttons=self.checkable_buttons_word)
        
def use_data():
    self = main_window
    functions.use_data(self=self,
                        other_widgets=self.other_widgets_word,
                        text_boxes=self.text_boxes_word,
                        combo_boxes=self.combo_boxes_word,
                        checkable_buttons=self.checkable_buttons_word)
        
def clean():
    self = main_window
    functions.clean(self=self,
                    text_boxes=self.text_boxes_word,
                    combo_boxes=self.combo_boxes_word,
                    checkable_buttons=self.checkable_buttons_word)
        
def create_template():
    self = main_window
    functions.create_template(self=self, word=True)


# Dialogs
        
def create_protocol_from_excel():
    logger.debug('Создание протокола из excel')
    self = main_window

    dialogs.CreateProtocolFromExcelDialog(self, word=True).exec_()

    logger.success('Успешное создание протокола из excel')

def show_inspection_address_setting():
        logger.debug('Показать настройки адреса поверки')
        self = main_window

        dialogs.ChooseInspectionAddressDialog(self,inspection=self.text_boxes_word['inspection_address'],\
                                            legal=self.text_boxes_word['legal_address'], word=True).exec_()

        logger.success('Успешно показаны настройки адреса поверки')

def choose_scale():
    self = main_window
    logger.debug('Выбор весов диалог')

    dialogs.ChooseScaleDialog(text_scale_widget=self.text_boxes_word['scale'],word=True).exec_()

    logger.success('Успешно выбор весов диалог')

def add_save_path():
    self = main_window
    logger.debug('Выбор пути сохранения')

    dialog = QDialog()
    self.text_boxes_word['save_path'].setPlainText(QFileDialog.getExistingDirectory(dialog, "Выберите папку"))

    logger.success('Успешно выбор пути сохранения')

def settings():
        logger.debug('Настройки диалог')

        dialogs.SettingDialog().exec_()

        logger.success('Успешно настройки диалог')

def add_path_to_excel():
    self = main_window
    logger.debug('Добавление пути до журнала Excel')

    options = QFileDialog.Options()
    filePath, _ = QFileDialog.getOpenFileName(self, 'Выберите файл', '', 'All Files (*);;Text Files (*.txt)', options=options)

    self.text_boxes_word['path_to_excel_jounal'].setPlainText(filePath)

    logger.success('Успешно Добавление пути до журнала Excel')