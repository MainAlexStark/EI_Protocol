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

from ..parse.__init__ import FNS_API


class ChooseInspectionAddressDialog(QDialog):
    def __init__(self, inspection, legal, word: bool):
        try:
            super().__init__()
            
            self.word = word

            self.inspection_address = inspection
            self.legal_address = legal

            logger.info('ChooseInspectionAddressDialog(QDialog): __init__')

            length_window = 1000
            width_window = 600

            self.setGeometry(200, 200, length_window, width_window)
            self.setWindowTitle('Выберите нужный вариант')

            self.main_layout = QVBoxLayout()

            self.var = [
                'Использовать введенный юр.адрес',
                'Использовать юр.адрес ООО \"ЕДИНИЦА ИЗМЕРЕНИЯ\"'
            ]

            # Таблица с вариантами
            self.label = QLabel("Выберите нужный вариант:")
            self.table = QTableWidget()
            self.table.setRowCount(len(self.var))
            self.table.setColumnCount(1)

            self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Запрещаем редактирование

            self.table.setColumnWidth(0, 940)

            i = 0
            for var in self.var:
                self.table.setItem(i,0, QTableWidgetItem(var))
                i += 1

            self.table.setCurrentCell(-1,-1)
            self.table.currentItemChanged.connect(self.item_changed)

            self.main_layout.addWidget(self.label)
            self.main_layout.addWidget(self.table)


            self.setLayout(self.main_layout)

        except Exception as e:
            logger.error(f'Error init ChooseInspectionAddressDialog(QDialog) : {e}')

    def item_changed(self):
        logger.debug(f'item_changed')
        3

        if self.table.currentRow() == 0:
            text = self.text_legal_address.toPlainText()
            if len(text) > 0:
                self.inspection_address.setPlainText(text)
            else:
                message_box = QMessageBox()
                message_box.setIcon(QMessageBox.Critical)
                message_box.setText("Введите юр.адрес!")
                message_box.setWindowTitle("Ошибка")
                message_box.setStandardButtons(QMessageBox.Ok)
                message_box.exec_()

        else:
            # Устанавливаем
            self.inspection_address.setPlainText('610027, Россия, Кировская область, город Киров, улица Красноармейская, дом 43А, кв. помещение 1,21')

            self.close()