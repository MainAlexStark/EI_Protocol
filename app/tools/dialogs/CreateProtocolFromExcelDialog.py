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

from .CreateProtocolDialog import CreateProtocolDialog

from ..db import get_data

from ..office import Word
from ..office import Excel

WORD = Word()
EXCEL = Excel()

class CreateProtocolFromExcelDialog(QDialog):

    def __init__(self, main_window, word):
        try:
            super().__init__()
            
            self.word = word

            self.main_window = main_window

            logger.info('CreateProtocolFromExcelDialog(QDialog): __init__')

            length_window = 500
            width_window = 500

            self.setGeometry(200, 200, length_window, width_window)
            self.setWindowTitle('Создание протокола из excel')


            data = get_data('app/tools/data/storage.json')

            self.main_layout = QVBoxLayout()

            # PAth to excel 

            self.path_excel_layout = QHBoxLayout()

            self.text_path_to_excel = QPlainTextEdit(self)
            self.text_path_to_excel.setPlaceholderText("Путь к журналу excel")
            self.text_path_to_excel.setPlainText(data['path_to_excel'])
            self.text_path_to_excel.setFixedSize(500,40)

            self.path_excel_layout.addWidget(self.text_path_to_excel)

            self.button_path_to_excel_dialog = QPushButton('...', self)
            self.button_path_to_excel_dialog.setFixedSize(50, 50)
            self.button_path_to_excel_dialog.clicked.connect(self.add_path_to_excel)  # Привязываем функцию
            self.path_excel_layout.addWidget(self.button_path_to_excel_dialog)  # Добавляем в layout

            self.main_layout.addLayout(self.path_excel_layout)


            # Path
            self.path_layout = QHBoxLayout()

            self.text_path = QPlainTextEdit(self)
            self.text_path.setPlaceholderText("Путь к журналу excel")
            self.text_path.setPlainText(data['path'])
            self.text_path.setFixedSize(500,40)

            self.path_layout.addWidget(self.text_path)

            self.button_path_dialog = QPushButton('...', self)
            self.button_path_dialog.setFixedSize(50, 50)
            self.button_path_dialog.clicked.connect(self.add_path)  # Привязываем функцию
            self.path_layout.addWidget(self.button_path_dialog)  # Добавляем в layout

            self.main_layout.addLayout(self.path_layout)

            # Указать строки

            self.text_num_row = QPlainTextEdit(self)
            self.text_num_row.setPlaceholderText("Номер строки")
            self.text_num_row.setFixedSize(500,40)

            self.main_layout.addWidget(self.text_num_row)

            self.text_num_row2 = QPlainTextEdit(self)
            self.text_num_row2.setPlaceholderText("Номер строки 2")
            self.text_num_row2.setFixedSize(500,40)

            self.main_layout.addWidget(self.text_num_row2)


            # Checkbox выбора

            self.CheckBox_use_date = QCheckBox('Использовать дату')
            self.CheckBox_use_date.stateChanged.connect(self.CheckBoxChenged)

            self.main_layout.addWidget(self.CheckBox_use_date)


            # Calendar
            self.inspection_date = QCalendarWidget(self)
            self.inspection_date.setFixedSize(500, 300)
            self.inspection_date.setEnabled(False)
            self.main_layout.addWidget(self.inspection_date)

            # Botton

            self.button_create_protocol = QPushButton('Создать протокол', self)
            self.button_create_protocol.clicked.connect(self.create_protocol)  # Привязываем функцию
            self.button_create_protocol.setFixedSize(500,40)
            self.main_layout.addWidget(self.button_create_protocol)  # Добавляем в layout

            self.setLayout(self.main_layout)

        except Exception as e:
            logger.error(f'Error init ChooseScaleDialog(QDialog) : {e}')


    def create_protocol(self, word):
        logger.debug('start')


        if self.CheckBox_use_date.isChecked():
            date = str(self.inspection_date.selectedDate().toString("dd.MM.yyyy")).strip()
            path = self.text_path_to_excel.toPlainText()
            rows = EXCEL.get_row_num(path=path ,date_to_find=date)

            if len(rows) > 0:
                for row in rows:
                    
                    result_get = EXCEL.get_from_excel(self.text_path_to_excel.toPlainText(),row=str(row))

                    if not result_get:

                        message_box = QMessageBox()
                        message_box.setIcon(QMessageBox.Critical)
                        message_box.setText(f"Ошибка при работе с Excel (Журналом)")
                        message_box.setWindowTitle("Ошибка")
                        message_box.setStandardButtons(QMessageBox.Ok)
                        message_box.exec_()
                    else:
                        args = result_get

                        args['path'] = self.text_path.toPlainText()

                        template_path = 'app\\templates\\Word\\' + args['scale'] + ' ' + args['FIF'] + '.docx'

                        if os.path.exists(template_path):
                            # Создаем протокол
                            result = WORD.make_new_protocol(args)

                            CreateProtocolDialog(self, result=result).exec_()

                            logger.debug('end')
                        else:
                            message_box = QMessageBox()
                            message_box.setIcon(QMessageBox.Critical)
                            message_box.setText(f"Нет нужного шаблона проткола")
                            message_box.setWindowTitle("Ошибка")
                            message_box.setStandardButtons(QMessageBox.Ok)
                            message_box.exec_()

                        
            else:
                message_box = QMessageBox()
                message_box.setIcon(QMessageBox.Critical)
                message_box.setText(f"Нет выбранной даты!")
                message_box.setWindowTitle("Ошибка")
                message_box.setStandardButtons(QMessageBox.Ok)
                message_box.exec_()

                logger.error("Нет выбранной даты!")

        else:

            for row in range(int(self.text_num_row.toPlainText()),int(self.text_num_row2.toPlainText()) + 1):

                try:

                    result_get = EXCEL.get_args_from_excel(self.text_path_to_excel.toPlainText(),row=str(row))

                    if not result_get:

                        message_box = QMessageBox()
                        message_box.setIcon(QMessageBox.Critical)
                        message_box.setText(f"Ошибка при работе с Excel (Журналом)")
                        message_box.setWindowTitle("Ошибка")
                        message_box.setStandardButtons(QMessageBox.Ok)
                        message_box.exec_()
                    else:
                        args = result_get

                        args['path'] = self.text_path.toPlainText()

                        template_path = 'app\\templates\\Word\\' + args['scale'] + ' ' + args['FIF'] + '.docx'

                        if os.path.exists(template_path):
                            # Создаем протокол
                            
                            if self.word:
                                result = WORD.make_new_protocol(args=args)
                            else:
                                result = EXCEL.make_new_protocol(args=args)

                            CreateProtocolDialog(self, result=result).exec_()

                            logger.debug('end')
                        else:
                            message_box = QMessageBox()
                            message_box.setIcon(QMessageBox.Critical)
                            message_box.setText(f"Нет нужного шаблона проткола")
                            message_box.setWindowTitle("Ошибка")
                            message_box.setStandardButtons(QMessageBox.Ok)
                            message_box.exec_()

                except Exception as e:
                    logger.error(f'Ошибка при создании протколов {e}')
                    message_box = QMessageBox()
                    message_box.setIcon(QMessageBox.Critical)
                    message_box.setText(f"Ошибка при создании протколов")
                    message_box.setWindowTitle("Ошибка")
                    message_box.setStandardButtons(QMessageBox.Ok)
                    message_box.exec_()



    def add_path_to_excel(self):
        logger.debug('start')

        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getOpenFileName(self, 'Выберите файл', '', 'All Files (*);;Text Files (*.txt)', options=options)

        self.text_path_to_excel.setPlainText(filePath)


    def CheckBoxChenged(self,state):

        if state == 2: # Если состояние чекбокса - отмечен
            self.text_num_row.setEnabled(False)
            self.text_num_row2.setEnabled(False)

            self.inspection_date.setEnabled(True)
        else:
            self.text_num_row.setEnabled(True)
            self.text_num_row2.setEnabled(True)

            self.inspection_date.setEnabled(False)


    def add_path(self):
        logger.info('start')

        dialog = QDialog()
        self.text_path.setPlainText(QFileDialog.getExistingDirectory(dialog, "Выберите папку"))

        logger.debug('end')