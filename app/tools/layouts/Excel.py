import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QCompleter, QFileDialog, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QCheckBox
from PyQt5.QtCore import QDate
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMessageBox, QDialog, QFrame, QAction, QMenu



def get_layout(self):
    main_layout = QHBoxLayout()

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

    return main_layout
    