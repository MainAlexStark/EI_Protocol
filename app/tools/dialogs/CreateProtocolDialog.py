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


class CreateProtocolDialog(QDialog):

    def __init__(self, main_window, result):
        try:
            super().__init__()

            self.main_window = main_window

            logger.debug('CreateProtocolDialog(QDialog): __init__')

            length_window = 500
            width_window = 500

            self.setGeometry(200, 200, length_window, width_window)
            self.setWindowTitle('Создание протокола')

            self.main_layout = QVBoxLayout()

            # Label
            self.label = QLabel()
            self.main_layout.addWidget(self.label)

            if result:
                self.label.setText(result)
            else:
                self.label.setText('Возникла непредвиденная ошибка ошибка!')

            self.setLayout(self.main_layout)

        except Exception as e:
            logger.error(f'Error init ChooseScaleDialog(QDialog) : {e}')