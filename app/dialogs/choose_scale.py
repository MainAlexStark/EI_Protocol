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


class ChooseScaleDialog(QDialog):
    selected_scale = ''

    def __init__(self, text_scale_widget, word):
        """
        Этот диалог для поиска нужного элемента в data/templates
        В диалог занесены в ручную вкладки в главном окне, при изменении кол-ва вкладок требуется изменить этот диалог
        """
        try:
            super().__init__()
            
            self.word = word

            self.text_scale_widget = text_scale_widget

            length_window = 1000
            width_window = 600

            self.setGeometry(200, 200, length_window, width_window)
            self.setWindowTitle('Выбрать весы')

            self.main_layout = QVBoxLayout()

            # Достаем все имена файлов шаблонов и добавляем в выпадающий список
            self.files = []
            if word:
                for file in os.listdir('data\\templates\\Word'):
                    self.files.append(file.replace('.docx', ''))
            else:
                for file in os.listdir('data\\templates\\Excel'):
                    self.files.append(file.replace('.xlsx', ''))

            # Поиск

            self.label_text_scale = QLabel("Поиск:")
            self.text_scale = QPlainTextEdit(self)
            self.text_scale.setFixedSize(500, 40)

            self.text_scale.textChanged.connect(self.text_scale_changed)

            self.main_layout.addWidget(self.label_text_scale)
            self.main_layout.addWidget(self.text_scale)

            # Таблица с весами
            self.label_text_table_scales = QLabel("Весы:")
            self.table_scales = QTableWidget()
            self.table_scales.setRowCount(len(self.files))
            self.table_scales.setColumnCount(1)

            self.table_scales.currentItemChanged.connect(self.item_changed)

            self.table_scales.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Запрещаем редактирование

            self.table_scales.setColumnWidth(0, 940)

            i = 0
            for file_name in self.files:
                self.table_scales.setItem(i,0, QTableWidgetItem(file_name))
                i += 1

            self.main_layout.addWidget(self.label_text_table_scales)
            self.main_layout.addWidget(self.table_scales)

            self.setLayout(self.main_layout)

        except Exception as e:
            logger.error(f'Error init ChooseScaleDialog(QDialog) : {e}')

    def text_scale_changed(self):
        search_text = self.text_scale.toPlainText().lower()  # Получаем текст поиска в нижнем регистре
        self.table_scales.setRowCount(0)  # Удаляем все строки из таблицы

        for file_name in self.files:
            if search_text in file_name.lower():  # Проверяем, соответствует ли имя файла поисковому запросу
                row_count = self.table_scales.rowCount()
                self.table_scales.insertRow(row_count)  # Вставляем новую строку
                self.table_scales.setItem(row_count, 0, QTableWidgetItem(file_name))


    def item_changed(self):
        
        self.text_scale_widget.setPlainText(self.table_scales.currentItem().text())

