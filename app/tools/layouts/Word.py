import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QCompleter, QFileDialog, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QCheckBox
from PyQt5.QtCore import QDate
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMessageBox, QDialog, QFrame, QAction, QMenu

from ...tools.db import get_data

from docx import Document

import os

def get_layout(self):
    main_layout = QHBoxLayout()

    self.var_layout = QVBoxLayout()
    self.var_r_layout = QVBoxLayout()
    self.var_r2_layout = QVBoxLayout()

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

    self.buttons.Buttons['button_path'] = self.button_path
    
    
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

    # Create Excel 
    self.create_excel_check_box = QCheckBox('Создавать excel шаблон/протокол')
    self.var_r_layout.addWidget(self.create_excel_check_box)  # Добавляем в layout

    self.buttons.CheckableButtons['create_excel'] = self.create_excel_check_box


    # Create protocol from excel

    self.button_create_protocol_from_excel = QPushButton('Создать протокол из excel', self)
    self.button_create_protocol_from_excel.clicked.connect(self.create_protocol_from_excel)  # Привязываем функцию
    self.var_r_layout.addWidget(self.button_create_protocol_from_excel)  # Добавляем в layout

    self.buttons.Buttons['create_protocol_from_excel'] = self.button_create_protocol_from_excel

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

    # Add layouts in main_layout
    main_layout.addLayout(self.var_layout)
    main_layout.addLayout(self.var_r_layout)
    main_layout.addLayout(self.var_r2_layout)

    return main_layout