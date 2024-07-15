""" third party imports """
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QFrame, QDialog, QFileDialog, QCheckBox, QMenu, QAction, QMenuBar
from PyQt5.QtWidgets import QApplication, QMessageBox
from loguru import logger
import os, sys, subprocess

from ..dialogs.menu.save_data import SaveDataDialog

from ..ConfigClient import Config


""" OPEN CONFIG """
file_path = 'data/config.json'
if os.path.exists(file_path):
    config_client = Config(file_path)
else:
    raise Exception(f'File {file_path} not found')

main_window = None

def get_layout(app: QWidget):
    global main_window
    main_window = app
    
    config = config_client.get()

    # Создаем MenuBar
    menubar = QMenuBar()

    # Аккаунт меню
    account_menu = menubar.addMenu('Аккаунт')

    staristicAction = QAction("Статистика", app)
    staristicAction.triggered.connect(statistic)
    myProtocolsAction = QAction("Мои протоколы", app)
    myCodeAction = QAction("Мой код поверителя", app)
    exitAction = QAction("Выйти", app)

    account_menu.addAction(staristicAction)
    account_menu.addAction(myProtocolsAction)
    account_menu.addAction(myCodeAction)
    account_menu.addAction(exitAction)
    
    exitAction.triggered.connect(exit)

    # Настройки меню 
    settings_menu = menubar.addMenu('Настройки')

    SaveDataAction = QAction("Сохраняемые данные", app)
    SaveDataAction.triggered.connect(save_data)

    settings_menu.addAction(SaveDataAction)

    # Помощь
    helpAction = QAction("Помощь", app)
    menubar.addAction(helpAction)
    
    # Создаем layout, который будем возвращать
    main_layout = QHBoxLayout()
    main_layout.addWidget(menubar)
    
    return main_layout
    
# Обработчики нажатий на кнопки меню
def save_data():
    SaveDataDialog(main_window).exec_()
    
def statistic():
    # StatisticDialog(main_window).exec_()
    ...
    
def exit():
    config = config_client.get()
    reply = QMessageBox.question(None, "Вопрос", "Вы точно хотите выйти из аккаунта?",
                                 QMessageBox.Yes | QMessageBox.No,
                                 QMessageBox.Yes)

    if reply == QMessageBox.Yes:
        config['app']['email'] = ""
        config['app']['password'] = ""
        config_client.post(config)
        # Закрываем текущее приложение
        main_window.close()
        # Запускаем новый процесс с текущим скриптом
        subprocess.Popen([sys.executable, sys.argv[0]])
        # Выходим из текущего процесса
        sys.exit()