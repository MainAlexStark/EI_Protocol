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


class SettingDialog(QDialog):
    def __init__(self):
        super().__init__()

        logger.info('SettingDialog(QDialog): __init__')

        self.setGeometry(200, 200, 400, 400)
        self.setWindowTitle('Настройки')

        main_layout = QHBoxLayout()
        verificationers_layout = QVBoxLayout()
        use_data_layout = QVBoxLayout()

        # Поверители
        self.label_text_verificationers = QLabel("Поверители:")
        self.table_verificationers = QTableWidget()
        self.table_verificationers.setRowCount(10)
        self.table_verificationers.setColumnCount(2)
        verificationers_layout.addWidget(self.label_text_verificationers)
        verificationers_layout.addWidget(self.table_verificationers)

        # Настройки использования преведущих данных
        self.label_use_data = QLabel("Использовать преведущие данные:")
        use_data_layout.addWidget(self.label_use_data)

        self.use_data_check_boxes = {'scale': QCheckBox("Весы"),
                                     'num_scale': QCheckBox("Номер весов"),
                                     'num_protocol': QCheckBox("Номер протокола"),
                                     'path': QCheckBox("Путь сохранения"),
                                     'verificationer': QCheckBox('Поверитель'),
                                     'company': QCheckBox('Компания'),
                                     'INN': QCheckBox('ИНН'),
                                     'legal_address': QCheckBox('Юридический адрес'),
                                     'inspection_address': QCheckBox('Адрес поверки'),
                                     'inspection_date': QCheckBox('Дата'),
                                     'standarts': QCheckBox('Набор эталонов (Может работать неккоректно)'),
                                     'use_excel': QCheckBox('Использовать excel'),
                                     'unfit': QCheckBox('Соответсвует/Несоответсвует')
                                     }

        for widget in self.use_data_check_boxes.values():
            use_data_layout.addWidget(widget)

        # Определите имя файла хранилища
        file_name = 'app/tools/data/config.json'

        # Откройте файл хранилища
        with open(file_name, 'r+') as file:
            # Загрузите данные из файла
            data = json.load(file)

            verificationers = data['verificationers']

            i = 0
            for verificationer in verificationers.keys():
                self.table_verificationers.setItem(i, 0, QTableWidgetItem(verificationer))
                self.table_verificationers.setItem(i, 1, QTableWidgetItem(verificationers[verificationer]))
                i += 1

            use_data = data['use_data']

            for key in use_data.keys():
                self.use_data_check_boxes[key].setChecked(use_data[key])




        # Кнопки

        # Сохранить
        self.button_save = QPushButton('Сохранить', self)
        self.button_save.setFixedSize(500, 40)
        self.button_save.clicked.connect(self.save)  # Привязываем функцию
        verificationers_layout.addWidget(self.button_save)  # Добавляем в layout

        # Use_all
        self.button_use_all = QPushButton('Выбрать все', self)
        self.button_use_all.setFixedSize(500, 40)
        self.button_use_all.setCheckable(True)
        self.button_use_all.clicked.connect(self.use_all)  # Привязываем функцию
        use_data_layout.addWidget(self.button_use_all)  # Добавляем в layout

        main_layout.addLayout(verificationers_layout)
        main_layout.addLayout(use_data_layout)
        self.setLayout(main_layout)

    def use_all(self):
        if self.button_use_all.isChecked():
            for widjet in self.use_data_check_boxes.values():
                widjet.setChecked(True)
        else:
            for widjet in self.use_data_check_boxes.values():
                widjet.setChecked(False)

    def save(self):
        try:
            # Определите имя файла хранилища
            file_name = 'app/tools/data/config.json'

            # Откройте файл хранилища
            with open(file_name, 'r+') as file:
                # Загрузите данные из файла
                data = json.load(file)

                # Очистите предыдущие данные
                data['verificationers'] = {}

                dictionary = data['verificationers']

                for i in range(self.table_verificationers.rowCount()):
                    item_0 = self.table_verificationers.item(i, 0)
                    item_1 = self.table_verificationers.item(i, 1)
                    if item_0 is not None and item_1 is not None:
                        dictionary[item_0.text()] = item_1.text()

                # Очистите предыдущие данные
                data['use_data'] = {}

                for key in self.use_data_check_boxes.keys():
                    data['use_data'][key] = self.use_data_check_boxes[key].isChecked()

                # Запишите обновленные данные обратно в файл
                file.seek(0)  # Переместите курсор в начало файла
                json.dump(data, file)
                file.truncate()  # Обрежьте файл, если новые данные занимают меньше места, чем предыдущие

        except Exception as e:
            logger.error(f'Ошибка при сохранении в config.json: e')

        self.close()
