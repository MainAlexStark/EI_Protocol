""" third party imports """
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QFrame, QDialog, QFileDialog, QCheckBox, QMenu, QAction, QMenuBar
from loguru import logger
from PyQt5.QtWidgets import QWidget, QApplication, QLabel, QVBoxLayout, QScrollArea

from ..tools.ClassStorage import ClassStorage

main_window = None

def get_layout(app: QWidget):
    logger.debug("Создание вкладки 'excel'")
    
    # Создаем хранилище виджетов
    app.excel_widgets = ClassStorage()
    
    # Создаем layouts
    main_layout = QHBoxLayout()
    scroll_layout = QHBoxLayout()
    
    # Создаем QScrollArea и добавляем в него layout
    scroll_area = QScrollArea()
    scroll_area.setWidgetResizable(True) # разрешаем QScrollArea изменять размер своего виджета
    scroll_area.setWidget(QWidget()) # Создаем пустой виджет для QScrollArea
    scroll_area.widget().setLayout(scroll_layout)
    
    # Добавляем scroll_area в main layout
    main_layout.addWidget(scroll_area)
    
    return main_layout