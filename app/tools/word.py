""" third party imports """
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QFrame, QDialog, QFileDialog, QCheckBox, QMessageBox
from PyQt5.QtGui import QFont
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
import docx
import os

from loguru import logger


class Word():

    def __init__(self, app: QWidget) -> None:
        self._app = app
    
    def open_document(self,path: str):
        
        try:
            doc = Document(path)
            return doc
            
        except Exception as e:
            logger.error(e)
            return False
        
    def make_new_protocol(self, args: dict):
        
        doc = self.open_document()

    def create_template(self, path: str, values: dict, file_name: str, standarts: list):

        try:
            doc = self.open_document(path)

            print(standarts)

            # Заменить все вхождения "старого_значения" на "новое_значение"
            i = 0
            for paragraph in doc.paragraphs:
                paragraph.text = '  ' + paragraph.text + '  '
                print(paragraph.text)
                for key, value in values.items():
                    if value in paragraph.text:
                        paragraph.text = paragraph.text.replace(value[2:-2], key).strip()
                for stand in standarts[1:]:
                    if stand in paragraph.text:
                        p = paragraph._element
                        p.getparent().remove(p)
                        p._p = p._element = None

                paragraph.text = paragraph.text.replace(standarts[0], "ЭТАЛОНЫ").replace(';','').strip()

                i += 1

            # Сохранить измененный файл
            doc.save(f"data/templates/Word/{file_name}.docx")

            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("шаблон создан успешно!")
            msg.setWindowTitle("Успех!")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()

        except Exception as e:
            QMessageBox.warning(self, 'Ошибка создания шаблона', e)