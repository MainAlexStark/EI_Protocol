""" third party imports """
import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QFrame, QDialog, QFileDialog, QCheckBox, QMenu, QAction, QMenuBar
import os
from loguru import logger
import time


""" internal imports """
from app.dialogs.login import LoginDialog
from app.api import Client
from app.pages import word, excel, menu
from app.ConfigClient import Config

# Настройка логов
logger.add("log.log", rotation="10 MB", compression="zip")
logger.add(sink=sys.stdout, level="INFO")

""" OPEN CONFIG """
file_path = 'data/config.json'
if os.path.exists(file_path):
    config_client = Config(file_path)
else:
    raise Exception(f'File {file_path} not found')

class App(QWidget):
    def __init__(self):
        super().__init__()
        logger.info("Создание UI главного окна")

        # Создаем Tab Widget
        MainTab = QTabWidget()

        # Создаем tabs
        word_tab = QWidget()
        excel_tab = QWidget()

        # Добавляем layouts в tab
        word_tab.setLayout(word.get_layout(app=self))
        excel_tab.setLayout(excel.get_layout(app=self))

        # Добавляем tabs в Tab Widget
        MainTab.addTab(word_tab,'Создание протокола Word')
        MainTab.addTab(excel_tab,'Создание протокола Excel')

        # Создаем main_layout и добавляем в него MenuBar и TabWidge
        main_layout = QVBoxLayout()
        main_layout.addLayout(menu.get_layout(app=self))
        main_layout.addWidget(MainTab)

        # Устанавливаем layout для виджета
        self.setLayout(main_layout)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    # Создаем экземпляр класса App
    logger.debug("Создаем экземпляр класса App...")
    main_window = App()
    
    login_dialog = LoginDialog(app=main_window)
    if login_dialog.exec_() == LoginDialog.Accepted:
        
        logger.debug("Отображение окна...")

        # Отображаем виджет
        main_window.showMaximized()
        
        sys.exit(app.exec_())