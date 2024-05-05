""" third party imports """
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QCompleter, QFileDialog, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QCheckBox
from PyQt5.QtCore import QDate
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMessageBox, QDialog

import json
import os

""" internal imports """
from ....db import Config


""" OPEN CONFIG """
file_path = 'data/config.json'
if os.path.exists(file_path):
    config_client = Config(file_path)
else:
    raise Exception(f'File {file_path} not found')

main_window = None

class SaveDataDialog(QDialog):
    def __init__(self, app: QWidget):
        """
        Этот диалог для настройки сохраняемых значений виджетов
        Диалог при при нажатии на кнопку сохранить изменяет данные в config['save_widgets']
        В диалог занесены в ручную вкладки в главном окне, при изменении кол-ва вкладок требуется изменить этот диалог
        """
        super().__init__()

        config = config_client.get()

        global main_window
        main_window = app

        self.setGeometry(200, 200, 400, 400)
        self.setWindowTitle('Настройки')

        main_layout = QVBoxLayout()
        pages_layput = QHBoxLayout()
        word_layout = QVBoxLayout()
        excel_layout = QVBoxLayout()

        self.word_widgets = {}
        self.excel_widgets = {}

        self.dialog_word_widgets = {}
        self.dialog_excel_widgets = {}

        for name, widget in app.word_widgets.text_boxes.items():
            self.word_widgets[name] = widget
        for name, widget in app.word_widgets.combo_boxes.items():
            self.word_widgets[name] = widget
        for name, widget in app.word_widgets.check_boxes.items():
            self.word_widgets[name] = widget

        for name, widget in app.excel_widgets.text_boxes.items():
            self.excel_widgets[name] = widget
        for name, widget in app.excel_widgets.combo_boxes.items():
            self.excel_widgets[name] = widget
        for name, widget in app.excel_widgets.check_boxes.items():
            self.excel_widgets[name] = widget
        
        for name, widget in self.word_widgets.items():
            # Создаем виджет
            check_box_widget = QCheckBox(name, app)
            self.dialog_word_widgets[name] = check_box_widget
            # добавляем виджет в layout
            word_layout.addWidget(check_box_widget)

        for name, widget in self.excel_widgets.items():
            # Создаем виджет
            check_box_widget = QCheckBox(name, app)
            self.dialog_excel_widgets[name] = check_box_widget
            # добавляем виджет в layout
            excel_layout.addWidget(check_box_widget)

        for name_save_widget, state in config['save_widgets']['word'].items():
            for name, widget in self.dialog_word_widgets.items():
                if name_save_widget == name and state == True: 
                    widget.setChecked(True)

        for name_save_widget, state in config['save_widgets']['excel'].items():
            for name, widget in self.dialog_excel_widgets.items():
                if name_save_widget == name and state == True: 
                    widget.setChecked(True)

        
        self.button_save = QPushButton('Сохранить', self)
        self.button_save.setFixedSize(500, 40)
        self.button_save.clicked.connect(self.save)  # Привязываем функцию

        pages_layput.addLayout(word_layout)
        pages_layput.addLayout(excel_layout)

        main_layout.addLayout(pages_layput)
        main_layout.addWidget(self.button_save)

        self.setLayout(main_layout)

    def save(self):
        config = config_client.get()

        data = {
            "word": {},
            "excel": {}
        }

        for name, widget in self.dialog_word_widgets.items():
            data['word'][name] = widget.isChecked()

        for name, widget in self.dialog_excel_widgets.items():
            data['excel'][name] = widget.isChecked()

        config['save_widgets'] = data
        config_client.post(data=config)

        self.close()