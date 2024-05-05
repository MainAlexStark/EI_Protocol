""" third party imports """
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QFrame, QDialog, QFileDialog, QCheckBox, QMenu, QAction, QMenuBar

from . import Word, Excel
from ..dialogs.menu.save_data import SaveDataDialog

main_window = None

def get_main_layout(app: QWidget):
    """
    Функция возвращает layout с tab-widget:
    1. Word (из файла Word)
    2. Excel (из файла Excel)

    Функция инициализирует layouts и добаляет их в tab-widget, а после в main_layout, который в последствии и возвращает

    Layouts используются в dialogs/menu/save_data.py
    """

    global main_window
    main_window = app

    # Создаем MenuBar
    menubar = QMenuBar()

    # Аккаунт меню
    account_menu = menubar.addMenu('Аккаунт')

    staristicAction = QAction("Статистика", app)
    myProtocolsAction = QAction("Мои протоколы", app)
    myCodeAction = QAction("Мой код поверителя", app)

    account_menu.addAction(staristicAction)
    account_menu.addAction(myProtocolsAction)
    account_menu.addAction(myCodeAction)

    # Настройки меню 
    settings_menu = menubar.addMenu('Настройки')

    SaveDataAction = QAction("Сохраняемые данные", app)
    SaveDataAction.triggered.connect(save_data)

    settings_menu.addAction(SaveDataAction)

    main_layout = QHBoxLayout()

    # Помощь
    helpAction = QAction("Помощь", app)
    menubar.addAction(helpAction)

    # Создаем Tab Widget
    MainTab = QTabWidget()

    # Создаем tabs
    word_tab = QWidget()
    excel_tab = QWidget()

    # Добавляем layouts в tab
    word_tab.setLayout(Word.get_layout(app=app))
    excel_tab.setLayout(Excel.get_layout(app=app))

    # Добавляем tabs в Tab Widget
    MainTab.addTab(word_tab,'Создание протокола Word')
    MainTab.addTab(excel_tab,'Создание протокола Excel')

    # Создаем main_layout и добавляем в него MenuBar и TabWidge
    main_layout = QVBoxLayout()
    main_layout.addWidget(menubar)
    main_layout.addWidget(MainTab)


    return main_layout


# Обработчики нажатий на кнопки меню
def save_data():
    SaveDataDialog(main_window).exec_()