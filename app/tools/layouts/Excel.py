import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QCompleter, QFileDialog, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QCheckBox
from PyQt5.QtCore import QDate
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMessageBox, QDialog, QFrame, QAction, QMenu



def get_layout(self):
    main_layout = QHBoxLayout()

    first_layout = QVBoxLayout()

    # Scale
    self.text_scale_excel = QPlainTextEdit(self)
    self.text_scale_excel.setPlaceholderText('Весы')

    self.text_scale_excel.textChanged.connect(self.scale_changed)

    first_layout.addWidget(self.text_scale_excel)

    self.var_boxes.text_boxes['scale_excel'] = self.text_scale_excel

    # Надпись
    self.button_choose_scale_excel = QPushButton('Выбрать весы', self)
    self.button_choose_scale_excel.clicked.connect(self.choose_scale)  # Привязываем функцию
    first_layout.addWidget(self.button_choose_scale_excel)  # Добавляем в layout

    self.buttons.Buttons['choose_scale_excel'] = self.button_choose_scale_excel


    main_layout.addLayout(first_layout)
    return main_layout
    