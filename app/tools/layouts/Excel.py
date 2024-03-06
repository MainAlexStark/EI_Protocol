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

    var_layout = QVBoxLayout()
    var_r_layout = QVBoxLayout()
    var_r2_layout = QVBoxLayout()

    # Scale

    self.text_scale_excel = QPlainTextEdit(self)
    self.text_scale_excel.setPlaceholderText('Весы')

    var_layout.addWidget(self.text_scale_excel)

    self.var_boxes_excel.text_boxes['scale_excel'] = self.text_scale_excel

    # Надпись
    self.button_choose_scale_excel = QPushButton('Выбрать весы', self)
    self.button_choose_scale_excel.clicked.connect(self.choose_scale_excel)  # Привязываем функцию
    var_layout.addWidget(self.button_choose_scale_excel)  # Добавляем в layout

    self.buttons_excel.Buttons['choose_scale_excel'] = self.button_choose_scale_excel

    # path
    # Текстовое поле
    self.text_path_excel = QPlainTextEdit(self)
    self.text_path_excel.setPlaceholderText('Путь сохранения')
    # Добавляем элементы в layouts
    var_layout.addWidget(self.text_path_excel)

    self.var_boxes_excel.text_boxes['path'] = self.text_path_excel
    
    # Создаем кнопку
    self.button_path_excel = QPushButton('Выбрать путь сохранения', self)
    self.button_path_excel.clicked.connect(self.add_path_excel)  # Привязываем функцию
    var_layout.addWidget(self.button_path_excel)  # Добавляем в layout

    self.buttons_excel.Buttons['button_path_excel'] = self.button_path_excel

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
    
    self.label_work_place_combo_excel = QLabel("Выберите рабочее место:")
    self.work_place_combo_excel = QComboBox()
    self.work_place_combo_excel.addItems(work_places)
    self.work_place_combo_excel.setCurrentIndex(-1)
    var_layout.addWidget(self.label_work_place_combo_excel)
    var_layout.addWidget(self.work_place_combo_excel)

    # Добавляем в словарь combo boxes
    self.var_boxes_excel.combo_boxes['work_place_excel'] = self.work_place_combo_excel

    # num_protocol

    # Текстовое поле
    self.text_num_protocol_excel = QPlainTextEdit(self)
    self.text_num_protocol_excel.setPlaceholderText('Номер протокола')
    # Добавляем элементы в layouts
    var_layout.addWidget(self.text_num_protocol_excel)

    self.var_boxes_excel.text_boxes['num_protocol_excel'] = self.text_num_protocol_excel

    # num_scale
    self.text_num_scale_excel = QPlainTextEdit(self)
    self.text_num_scale_excel.setPlaceholderText('Номер весов')
    var_layout.addWidget(self.text_num_scale_excel)

    self.var_boxes_excel.text_boxes['num_scale_excel'] = self.text_num_scale_excel

    # Verificationer
    data = get_data('app/tools/data/config.json')

    # Массив с поверителями
    ver = []


    verificationers = data.get('verificationers', [])
    verificationers_text = ''
    for verificationer in verificationers.keys():
        verificationers_text += verificationer
    verificationers_text = '\n'.join(verificationers)

    for i in verificationers_text.split('\n'): ver.append(i)

    self.label_verificationer_combo_excel = QLabel("Выберите поверителя:")
    self.verificationer_combo_excel = QComboBox()
    self.verificationer_combo_excel.addItems(verificationers)
    self.verificationer_combo_excel.setCurrentIndex(-1)
    self.verificationer_combo_excel.setFixedSize(500, 40)
    self.verificationer_combo_excel.currentTextChanged.connect(self.verificationer_changed_excel)
    var_layout.addWidget(self.label_verificationer_combo_excel)
    var_layout.addWidget(self.verificationer_combo_excel)

    # Добавляем в словарь combo boxes
    self.var_boxes_excel.combo_boxes['verificationer_excel'] = self.verificationer_combo_excel

    # Company
    self.text_company_excel = QPlainTextEdit(self)
    self.text_company_excel.setPlaceholderText("Компания")
    var_layout.addWidget(self.text_company_excel)

    self.var_boxes_excel.text_boxes['company_excel'] = self.text_company_excel

    # inspection address

    inspection_address_layout = QHBoxLayout()

    self.text_inspection_address_excel = QPlainTextEdit(self)
    self.text_inspection_address_excel.setPlaceholderText('Адрес поверки')
    inspection_address_layout.addWidget(self.text_inspection_address_excel)

    self.var_boxes_excel.text_boxes['inspection_address_excel'] = self.text_inspection_address_excel

    self.button_inspection_address_setting_excel = QPushButton('...', self)
    self.button_inspection_address_setting_excel.setFixedSize(50, 50)
    self.button_inspection_address_setting_excel.clicked.connect(self.show_inspection_address_setting_excel)  # Привязываем функцию
    inspection_address_layout.addWidget(self.button_inspection_address_setting_excel)  # Добавляем в layout

    var_layout.addLayout(inspection_address_layout)

    # Line

    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setMinimumSize(300,2)
    line.setMaximumSize(700,2)

    var_layout.addWidget(line)

    # PAth to excel 

    self.path_excel_layout = QHBoxLayout()

    self.text_path_to_excel_excel = QPlainTextEdit(self)
    self.text_path_to_excel_excel.setPlaceholderText("Путь к журналу excel")
    self.path_excel_layout.addWidget(self.text_path_to_excel_excel)

    self.var_boxes_excel.text_boxes['path_to_excel_excel'] = self.text_path_to_excel_excel

    self.button_path_to_excel_dialog_excel = QPushButton('...', self)
    self.button_path_to_excel_dialog_excel.setFixedSize(50, 50)
    self.button_path_to_excel_dialog_excel.clicked.connect(self.add_path_to_excel_excel)  # Привязываем функцию
    self.path_excel_layout.addWidget(self.button_path_to_excel_dialog_excel)  # Добавляем в layout

    var_layout.addLayout(self.path_excel_layout)

    # Use excel

    self.button_use_excel_excel = QPushButton('Использовать Excel', self)
    self.button_use_excel_excel.setCheckable(True)
    var_layout.addWidget(self.button_use_excel_excel)  # Добавляем в layout

    self.buttons_excel.CheckableButtons['use_excel_excel'] = self.button_use_excel_excel

    
    ####### var_r_layout

    # inspection_date

    self.inspection_date_excel = QCalendarWidget(self)
    self.inspection_date_excel.setFixedSize(500, 300)
    var_r_layout.addWidget(self.inspection_date_excel)

    # weather

    self.text_temperature_excel = QPlainTextEdit(self)
    self.text_temperature_excel.setPlaceholderText("Температура")

    var_r_layout.addWidget(self.text_temperature_excel)

    self.var_boxes_excel.text_boxes['temperature_excel'] = self.text_temperature_excel

    self.text_humidity_excel = QPlainTextEdit(self)
    self.text_humidity_excel.setPlaceholderText("Влажность")
    var_r_layout.addWidget(self.text_humidity_excel)

    self.var_boxes_excel.text_boxes['humidity_excel'] = self.text_humidity_excel

    self.text_pressure_excel = QPlainTextEdit(self)
    self.text_pressure_excel.setPlaceholderText("Давление")
    var_r_layout.addWidget(self.text_pressure_excel)

    self.var_boxes_excel.text_boxes['pressure_excel'] = self.text_pressure_excel

    ##############################################

    # Create template

    self.button_create_template_excel = QPushButton('Создать шаблон', self)
    self.button_create_template_excel.clicked.connect(self.create_template_excel)  # Привязываем функцию
    var_r_layout.addWidget(self.button_create_template_excel)  # Добавляем в layout

    self.buttons_excel.Buttons['create_template_excel'] = self.button_create_template_excel

    # Create protocol

    self.button_create_protocol_excel = QPushButton('Создать протокол', self)
    self.button_create_protocol_excel.clicked.connect(self.create_protocol_excel)  # Привязываем функцию
    var_r_layout.addWidget(self.button_create_protocol_excel)  # Добавляем в layout

    self.buttons_excel.Buttons['create_protocol_excel'] = self.button_create_protocol_excel

    # Create Excel 
    self.create_excel_check_box_excel = QCheckBox('Создавать excel шаблон/протокол')
    var_r_layout.addWidget(self.create_excel_check_box_excel)  # Добавляем в layout

    self.buttons_excel.CheckableButtons['create_excel_excel'] = self.create_excel_check_box_excel


    # Create protocol from excel

    self.button_create_protocol_from_excel_excel = QPushButton('Создать протокол из excel', self)
    self.button_create_protocol_from_excel_excel.clicked.connect(self.create_protocol_from_excel_excel)  # Привязываем функцию
    var_r_layout.addWidget(self.button_create_protocol_from_excel_excel)  # Добавляем в layout

    self.buttons_excel.Buttons['create_protocol_from_excel_excel'] = self.button_create_protocol_from_excel_excel

    # Line

    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setMinimumSize(300,2)
    line.setMaximumSize(700,2)

    var_r_layout.addWidget(line)

    # Use data

    self.button_use_data_excel = QPushButton('Использовать преведущие данные', self)
    self.button_use_data_excel.clicked.connect(self.use_data_excel)  # Привязываем функцию
    var_r_layout.addWidget(self.button_use_data_excel)  # Добавляем в layout

    self.buttons_excel.Buttons['use_data_excel'] = self.button_use_data_excel

    # Clean

    self.button_clean_excel = QPushButton('Очистить', self)
    self.button_clean_excel.clicked.connect(self.clean_excel)  # Привязываем функцию
    var_r_layout.addWidget(self.button_clean_excel)  # Добавляем в layout

    self.buttons_excel.Buttons['clean_excel'] = self.button_clean_excel

    # Settings

    self.button_settings_excel = QPushButton('Настройки', self)
    self.button_settings_excel.clicked.connect(self.settings_excel)  # Привязываем функцию
    var_r_layout.addWidget(self.button_settings_excel)  # Добавляем в layout

    self.buttons_excel.Buttons['settings_excel'] = self.button_settings_excel

    # 2 layout

    # Standarts

    self.tab_widget_excel = QTabWidget(self)
    self.tab_widget_excel.setMaximumSize(900, 950)
    self.tab_widget_excel.setMinimumSize(850, 800)

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

        self.tab_widget_excel.addTab(tab, file_name)

    var_r2_layout.addWidget(self.tab_widget_excel)

    # Add layouts in main_layout
    main_layout.addLayout(var_layout)
    main_layout.addLayout(var_r_layout)
    main_layout.addLayout(var_r2_layout)


    return main_layout