from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QVBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QDialog, QPushButton, QHBoxLayout, \
    QCheckBox, QPlainTextEdit, QAbstractItemView
import json
import os

from loguru import logger

main_path = os.path.abspath(os.path.join(os.path.dirname("../__main__.py")))

class TableWidget(QTableWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.itemSelectionChanged.connect(self.handle_item_selection_changed)

    def handle_item_selection_changed(self):
        selected_items = self.selectedItems()
        for item in selected_items:
            # Выполняете необходимые действия с выбранным элементом, например
            row = item.row()
            column = item.column()
            value = item.text()
            print(f"Изменен элемент в строке {row}, столбце {column}: {value}")

        self.itemChanged.emit()


class ChooseScaleDialog(QDialog):

    selected_scale = ''

    def __init__(self, main_window):
        try:
            super().__init__()

            self.main_window = main_window

            logger.info('ChooseScaleDialog(QDialog)')

            length_window = 1000
            width_window = 600

            self.setGeometry(200, 200, length_window, width_window)
            self.setWindowTitle('Выбрать весы')

            self.main_layout = QVBoxLayout()

            # Достаем все имена файлов шаблонов и добавляем в выпадающий список
            self.files = []
            template_dir = os.path.join(main_path, 'templates')
            for file in os.listdir(template_dir):
                self.files.append(file.replace('.docx', ''))

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

        self.main_window.text_scale.setPlainText(self.table_scales.currentItem().text())

