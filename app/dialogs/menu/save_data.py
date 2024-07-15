""" third party imports """
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QCompleter, QFileDialog, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QCheckBox
from PyQt5.QtCore import QDate
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMessageBox, QDialog
from loguru import logger

import json
import os

""" internal imports """
from ...ConfigClient import Config
from ...tools.ClassStorage import ClassStorage


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

        # Создаем layouts
        main_layout = QVBoxLayout()
        word_layout = QVBoxLayout()
        excel_layout = QVBoxLayout()
        
        # Получаем виджеты
        word_widgets: ClassStorage = app.word_widgets
        excel_widgets: ClassStorage = app.excel_widgets
        
        self.word_widgets_check_boxes = {}
        self.excel_widgets_check_boxes = {}
        
        logger.debug(f"permissions={config['widgets']['permissions']['word']}")
        
        for widget_class, widgets in word_widgets.storage.items():
            for name, widget in widgets.items():
                # Создаем виджет
                check_box_widget = QCheckBox(name)
                self.word_widgets_check_boxes[name] = check_box_widget
                if name in config['widgets']['permissions']['word'].keys():
                    if config['widgets']['permissions']['word'][name]: check_box_widget.setChecked(True)
                # добавляем виджет в layout
                word_layout.addWidget(check_box_widget)
        
        for widget_class, widgets in excel_widgets.storage.items():
            for name, widget in widgets.items():
                # Создаем виджет
                check_box_widget = QCheckBox(name)
                self.excel_widgets_check_boxes[name] = check_box_widget
                if name in config['widgets']['permissions']['word'].keys():
                    if config['widgets']['permissions']['excel'][name]: check_box_widget.setChecked(True)
                # добавляем виджет в layout
                excel_layout.addWidget(check_box_widget)
        
        # Создаем Tab Widget
        MainTab = QTabWidget()
        
        # Создаем tabs
        word_tab = QWidget()
        excel_tab = QWidget()

        # Добавляем layouts в tab
        word_tab.setLayout(word_layout)
        excel_tab.setLayout(excel_layout)

        # Добавляем tabs в Tab Widget
        MainTab.addTab(word_tab, 'Создание протокола Word')
        MainTab.addTab(excel_tab, 'Создание протокола Excel')
        
        main_layout.addWidget(MainTab)
        
        button_save = QPushButton('Сохранить', self)
        button_save.setFixedSize(500, 40)
        button_save.clicked.connect(self.save)  # Привязываем функцию
        
        main_layout.addWidget(button_save)

        self.setLayout(main_layout)

    def save(self):
        config = config_client.get()

        for name, widget in self.word_widgets_check_boxes.items():
            config['widgets']['permissions']['word'][name] = widget.isChecked()
            
        for name, widget in self.excel_widgets_check_boxes.items():
            config['widgets']['permissions']['excel'][name] = widget.isChecked()
       
        config_client.post(data=config)

        self.close()