from PyQt5.QtWidgets import QWidget, QHBoxLayout , QFileDialog, QTabWidget, QDialog

from docx import Document

from loguru import logger


from .tools.db import get_data
from .tools import functions

from .tools import dialogs
from .tools import layouts

from . import strings

# Variables
class VarBoxes:
    combo_boxes = {}
    text_boxes = {}
    labels = {}

class Buttons:
    Buttons = {}
    CheckableButtons = {}

class Other_Widgets:
    widgets = {}

class App(QWidget):
    def __init__(self):
        super().__init__()

        # Logger
        logger.add("log.txt")
        logger.info('Start initUI')
        logger.success(strings.start_text)

        # Получаем данные из config.py
        path_to_data = 'app\\tools\\data\\config.json'
        data = get_data(file_name=path_to_data)

        window_title = data['window_title']

        self.setWindowTitle(window_title)
        self.move(0,0)

        
        # Хранилища для widgets
        self.other_widgets_word = {}
        self.text_boxes_word = {}
        self.combo_boxes_word = {}
        self.buttons_word = {}
        self.checkable_buttons_word = {}

        self.other_widgets_excel = {}
        self.text_boxes_excel = {}
        self.combo_boxes_excel = {}
        self.buttons_excel = {}
        self.checkable_buttons_excel = {}

        # Layouts
        self.main_layout = QHBoxLayout()

        self.WordLayout = layouts.Word.get_layout(self=self)
        self.ExcelLayout = layouts.Excel.get_layout(self=self)
        
        self.initUI()

    def initUI(self):

        # Создаем Tab Widget
        MainTab = QTabWidget()

        # Создаем tabs
        word_tab = QWidget()
        excel_tab = QWidget()

        # Добавляем layouts в tab
        word_tab.setLayout(self.WordLayout)
        excel_tab.setLayout(self.ExcelLayout)

        # Добавляем tabs в Tab Widget
        MainTab.addTab(word_tab,'Создание протокола Word')
        MainTab.addTab(excel_tab,'Создание протокола Excel')

        # Добавляем TabWidget в main layout
        self.main_layout.addWidget(MainTab)

        # Добавляем main_layout в окно
        self.setLayout(self.main_layout)
        self.show()

    # Events

    def resizeEvent(self, event):
        print("Window has been resized")
        super().resizeEvent(event)