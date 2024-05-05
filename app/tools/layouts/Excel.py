""" third party imports """
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QFrame, QDialog, QFileDialog
import os

""" internal imports """
from ...db import Config

""" OPEN CONFIG """
file_path = 'data/config.json'
if os.path.exists(file_path):
    config_client = Config(file_path)
    config = config_client.get()
else:
    raise Exception(f'File {file_path} not found')


def get_layout(app: QWidget):
    """
    Функция создает и возвращает layout по работе с excel-протколами
    Привязанные фукнции хранятся в этом же файле
    """

    main_layout = QHBoxLayout()

    # Scale

    scale = QPlainTextEdit()
    scale.setPlaceholderText('Счетчик')

    app.excel_widgets.add_widget('счетчик', scale)

    main_layout.addWidget(scale)

    return main_layout