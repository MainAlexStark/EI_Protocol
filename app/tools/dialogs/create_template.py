import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QAction, QFileDialog, QDialog, QVBoxLayout, QPushButton
from PyQt5.QtGui import QTextCursor
from docx import Document

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QPushButton, QShortcut
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QFrame, QDialog, QFileDialog, QCheckBox, QMessageBox, QMenuBar
from PyQt5.QtGui import QTextCharFormat, QColor
from PyQt5 import QtGui, QtCore
import asyncio

from transliterate import translit


from ..office.Word import Word

class CreateTemplateDialog(QDialog):
    def __init__(self, path: str, widgets_names: list):
        super().__init__()

        self._word = Word(self)

        self.path = path
        self._widgets_names = widgets_names

        layout = QVBoxLayout()
        texts_layout = QHBoxLayout()

        # Create menu-bar
        menubar = QMenuBar()

        saveFileAction = QAction("Сохранить файл", self)
        saveFileAction.triggered.connect(self.saveDocx)
        menubar.addAction(saveFileAction)

        layout.addWidget(menubar)

        # Create widgets
        self.textEdit = QTextEdit()
        self.textEdit.setMinimumSize(500, 400)
        self.textEdit.setFontPointSize(12)
        self.textEdit.setReadOnly(True)
        texts_layout.addWidget(self.textEdit)

        self.result = QTextEdit()
        self.result.setMinimumSize(500, 400)
        self.result.setFontPointSize(12)
        self.result.setReadOnly(True)
        texts_layout.addWidget(self.result)

        layout.addLayout(texts_layout)

        self.fif = QTextEdit()
        self.fif.setFixedSize(500,50)
        self.fif.setPlaceholderText('Номер ФИФ')
        layout.addWidget(self.fif)


        self.setGeometry(100, 100, 800, 600)
        self.setWindowTitle('Docx Editor')


        # Создаем сочетание клавиш: Ctrl+B
        shortcut = QShortcut(Qt.CTRL + Qt.Key_F, self)

        # Соединяем сочетание клавиш с вызовом метода нажатия кнопки
        shortcut.activated.connect(self.key)

        # Create buttons
        # key - buttons text ; value - text 
        self.buttons = {}
        for name in widgets_names:
            self.buttons[name] = name.replace(' ','_').upper()

        self.values = {}

        self.values_text = {}

        self.setLayout(layout)

        self.openDocx()

        self.show()

    def openDocx(self):
        fileName = self.path
        if fileName:
            doc = Document(fileName)
            text = ''
            for para in doc.paragraphs:
                text += para.text + '\n'
            self.textEdit.setText(text)

    def saveDocx(self):

        err = ''

        self.values = {'НОМЕР_ПРОТОКОЛА': ' 173 ', 'ВЕСЫ': ' Весы электронные медицинские ВЭМ-150-«Масса-К» ', 'НОМЕР_ВЕСОВ': ' АЗ-14268 ', 'КОМПАНИЯ': ' КОГБУЗ «Детский клинический консультативно-диагностический центр» ', 'ЮРИДИЧЕСКИЙ_АДРЕС': ' Карла Маркса 42 ', 'АДРЕС_ПОВЕРКИ': ' периодическая ', 'ИНН': ' Методика ', 'ТЕМПЕРАТУРА': ': 20 ⁰', 'ВЛАЖНОСТЬ': ': 54 %', 'ДАВЛЕНИЕ': ': 99,6 к', 'ЭТАЛОНЫ': 'Гиря 1 кг М1 №741738;Гиря 2 кг М1\t№742074;Гиря 2 кг М1\t№742076;Гиря 5 кг М1\t№13251;Гиря 10 кг М1 №0577;Набор гирь (10 mg - 500 g) М1 №37925166;Гиря 20 кг М1 №А565;Гиря 20 кг М1 №А405;Гиря 20 кг М1 №А524;Гиря 20 кг М1 №А211;Гиря 20 кг М1 №А344;Гиря 20 кг М1 №А526;Гиря 20 кг М1 №А437;Гиря 20 кг М1 №А232;Гиря 20 кг М1 №А467', 'ПРИГОДНОСТЬ': ' СИ  соответствует установленным в описании типа метрологическим требованиям и пригодно к применению. ', 'ПОВЕРИТЕЛЬ': ' С.В. Стариков ', 'ДАТА_ПОВЕРКИ': ' 05 апреля 2024 '}

        standarts = self.values['ЭТАЛОНЫ'].split(';')

        for name in self.buttons.keys():
            if not(self.buttons[name] in self.values.keys()):
                err += name + '\n'

                
        if len(err) > 0:
            QMessageBox.warning(self, 'Ошибка, укажите ', err)
        elif len(self.fif.toPlainText()) == 0:
            QMessageBox.warning(self, 'Ошибка, укажите ', 'ФИФ')
        else:
            file_name = f"{self.values['ВЕСЫ']} {self.fif.toPlainText()}"

            self._word.create_template(self.path, self.values, file_name, standarts)

            self.close()

    def key(self):

        cursor = self.textEdit.textCursor()

        start_pos = self.textEdit.textCursor().selectionStart()
        end_pos = self.textEdit.textCursor().selectionEnd()

        cursor.setPosition(start_pos - 1)
        cursor.setPosition(end_pos + 1, QTextCursor.KeepAnchor)

        selected_text = cursor.selectedText()

        if start_pos != end_pos:
            self.showDialog(selected_text, start_pos, end_pos)

    def showDialog(self, selectedText, start_pos, end_pos):
        dialog = QDialog()
        layout = QVBoxLayout()
        for text, value in self.buttons.items():
            button = QPushButton(text)
            button.clicked.connect(lambda checked, text=text: self.saveTextAndPosition(text, selectedText, dialog))
            layout.addWidget(button)
        dialog.setLayout(layout)
        dialog.exec_()

    def saveTextAndPosition(self, buttonName, selectedText, dialog):
        # Save button text and selected text position to dictionary
        self.values[self.buttons[buttonName]] = selectedText
        self.values_text[buttonName] = selectedText

        print(self.values)

        text = ''
        for key, value in self.values_text.items():
            text += f"{key} -> {value}\n"
        self.result.setPlainText(text)

        dialog.close()