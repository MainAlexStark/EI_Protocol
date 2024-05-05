""" third party imports """
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QFrame, QDialog, QFileDialog, QCheckBox, QMenu, QAction, QMenuBar
from docx import Document
from loguru import logger
import os

""" internal imports """
from .db import Config
# from .tools import functions
# from .tools import dialogs
from .tools import layouts
from . import strings
from .tools.EI_API import API_Interface


""" OPEN CONFIG """
file_path = 'data/config.json'
if os.path.exists(file_path):
    config_client = Config(file_path)
    config = config_client.get()
else:
    raise Exception(f'File {file_path} not found')

class Widgets:
    def __init__(self) -> None:
        self.text_boxes = {}
        self.combo_boxes = {}
        self.check_boxes = {}
        self.buttons = {}
        self.calendars = {}
        self.tab_widgets = {}

    def get_text_boxes_values(self) -> dict:
        """
        Функция возвращает текст введенный в текст-боксы
        Возвращает словарь (dict)
        """
        ...

    def get_combo_boxes_values(self) -> dict:
        """
        Функция возвращает текст выбранный в комбо-боксах
        Возвращает словарь (dict)
        """
        ...

    def get_check_boxes_values(self) -> dict:
        """
        Функция возвращает положения 2-х фазных кнопок
        Возвращает словарь (dict)
        """
        ...

    def get_checkable_buttons_values(self) -> dict:
        """
        Функция возвращает положения check boxes
        Возвращает словарь (dict)
        """
        ...

    def get_calendars_values(self) -> dict:
        """
        Функция возвращает даты из календарей
        Возвращает словарь (dict)
        """
        ...

    def get_tab_widgets_values(self) -> QTabWidget:
        """
        Функция возвращает tabwidget
        Возвращает (QTabWidget)
        """
        ...

    def add_widget(self, name: str, widget: QWidget) -> None:
        """
        Метод добавляет виджет к соответствующей категории (text_boxes, combo_boxes, check_boxes)

        Параметры
        ----------
        name: str
            имя виджета
        widget: QWidget
            виджет   
        """
        if isinstance(widget, QPlainTextEdit):
            self.text_boxes[name] = widget
        elif isinstance(widget, QComboBox):
            self.combo_boxes[name] = widget
        elif isinstance(widget, QCheckBox):
            self.check_boxes[name] = widget
        elif isinstance(widget, QPushButton):
            self.buttons[name] = widget
        elif isinstance(widget, QCalendarWidget):
            self.calendars[name] = widget
        elif isinstance(widget, QTabWidget):
            self.tab_widgets[name] = widget
        else:
            print("Widget type not supported")
        
class App(QWidget):
    def __init__(self, login: str):
        super().__init__()

        self.login = login
        self.ei_api = API_Interface(login)

        # Logger
        logger.add("log.txt")
        logger.info('Start initUI')
        logger.success(strings.start_text)

        window_title = config['window_settings']['title']

        self.setWindowTitle(window_title)
        self.move(0,0)

        # Vidgets
        self.word_widgets = Widgets()
        self.excel_widgets = Widgets()

        # Layouts
        self.main_layout = layouts.get_main_layout(app=self)
        
        self.initUI()

    def initUI(self):

        # Добавляем main_layout в окно
        self.setLayout(self.main_layout)
        self.show()

    # Events

    def resizeEvent(self, event):
        print("Window has been resized")
        super().resizeEvent(event)