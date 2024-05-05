""" third party imports """
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QFrame, QDialog, QFileDialog, QCheckBox, QMessageBox
from PyQt5.QtGui import QFont
import os
from loguru import logger
from docx import Document

""" internal imports """
from ...db import Config
from .. import functions
from ..dialogs.choose_scale import ChooseScaleDialog
from ..dialogs.create_template import CreateTemplateDialog

""" OPEN CONFIG """
file_path = 'data/config.json'
if os.path.exists(file_path):
    config_client = Config(file_path)
else:
    raise Exception(f'File {file_path} not found')


main_window = None

def get_layout(app: QWidget):
    """
    Функция создает и возвращает layout по работе с word-протколами
    Привязанные фукнции хранятся в этом же файле
    """

    global main_window

    main_window = app

    main_layout = QHBoxLayout()

    var_layout = QVBoxLayout()
    var_r_layout = QVBoxLayout()
    var_r2_layout = QVBoxLayout()

    # Scale

    scale = QPlainTextEdit()
    scale.setPlaceholderText('Весы')
    app.word_widgets.add_widget('весы', scale)
    var_layout.addWidget(scale)

    # Надпись
    choose_scale_button = QPushButton('Выбрать весы', app)
    choose_scale_button.clicked.connect(choose_scale)  # Привязываем функцию
    var_layout.addWidget(choose_scale_button)  # Добавляем в layout
    app.word_widgets.add_widget('весы', choose_scale_button)

    # Путь сохранения

    path_layout = QHBoxLayout()

    save_path = QPlainTextEdit(app)
    save_path.setPlaceholderText('Путь сохрания')
    path_layout.addWidget(save_path)

    app.word_widgets.add_widget('путь сохранения', save_path)

    save_path_button = QPushButton('...', app)
    save_path_button.setFixedSize(50, 50)
    save_path_button.clicked.connect(add_save_path)  # Привязываем функцию
    path_layout.addWidget(save_path_button)  # Добавляем в layout

    var_layout.addLayout(path_layout)


    # num_protocol
    text_num_protocol = QPlainTextEdit(app)
    text_num_protocol.setPlaceholderText('Номер протокола')
    var_layout.addWidget(text_num_protocol)
    app.word_widgets.add_widget('номер протокола', text_num_protocol)

    # num_scale
    text_num_scale = QPlainTextEdit(app)
    text_num_scale.setPlaceholderText('Номер весов')
    var_layout.addWidget(text_num_scale)
    app.word_widgets.add_widget('номер весов', text_num_scale)


    # Поверитель
    login = app.login
    verificationers = app.ei_api.get_verificationers()
    ver = []

    label_verificationer_combo = QLabel("Выберите поверителя:")
    verificationer_combo = QComboBox()
    verificationer_combo.addItems(verificationers.keys())
    verificationer_combo.setCurrentText(login)
    verificationer_combo.setFixedSize(500, 40)
    verificationer_combo.currentTextChanged.connect(verificationer_changed)
    var_layout.addWidget(label_verificationer_combo)
    var_layout.addWidget(verificationer_combo)

    app.word_widgets.add_widget('поверитель', verificationer_combo)


    # Line

    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setMinimumSize(300,2)
    line.setMaximumSize(700,2)

    var_layout.addWidget(line)

    # INN
    INN = QPlainTextEdit(app)
    INN.setPlaceholderText("ИНН")
    var_layout.addWidget(INN)

    app.word_widgets.add_widget('инн', INN)

    search_company_button = QPushButton('Найти', app)
    search_company_button.clicked.connect(search_company)  # Привязываем функцию
    var_layout.addWidget(search_company_button)  # Добавляем в layout

    app.word_widgets.add_widget('найти компанию', search_company_button)


    # Company
    company = QPlainTextEdit(app)
    company.setPlaceholderText("Компания")
    var_layout.addWidget(company)

    app.word_widgets.add_widget('компания', company)


    # Legal address
    legal_address = QPlainTextEdit(app)
    legal_address.setPlaceholderText("Юридический адрес")
    var_layout.addWidget(legal_address)

    app.word_widgets.add_widget('юридический адрес', legal_address)


    # inspection address

    inspection_address_layout = QHBoxLayout()

    inspection_address = QPlainTextEdit(app)
    inspection_address.setPlaceholderText('Адрес поверки')
    inspection_address_layout.addWidget(inspection_address)

    app.word_widgets.add_widget('адрес поверки', inspection_address)

    button_inspection_address_setting = QPushButton('...', app)
    button_inspection_address_setting.setFixedSize(50, 50)
    button_inspection_address_setting.clicked.connect(show_inspection_address_setting)  # Привязываем функцию
    inspection_address_layout.addWidget(button_inspection_address_setting)  # Добавляем в layout

    var_layout.addLayout(inspection_address_layout)

    # Unfit
    unfit = QCheckBox('Несоответсвует', app)
    var_layout.addWidget(unfit)  # Добавляем в layout

    app.word_widgets.add_widget('пригодность', unfit)

    # Line
    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setMinimumSize(300,2)
    line.setMaximumSize(700,2)
    var_layout.addWidget(line)

    # path to excel

    path_excel_layout = QHBoxLayout()

    path_to_excel_journal = QPlainTextEdit(app)
    path_to_excel_journal.setPlaceholderText("Путь к журналу excel")
    path_excel_layout.addWidget(path_to_excel_journal)

    app.word_widgets.add_widget('путь к журналу', path_to_excel_journal)

    path_to_excel_dialog_button = QPushButton('...', app)
    path_to_excel_dialog_button.setFixedSize(50, 50)
    path_to_excel_dialog_button.clicked.connect(add_path_to_excel)  # Привязываем функцию
    path_excel_layout.addWidget(path_to_excel_dialog_button)  # Добавляем в layout

    var_layout.addLayout(path_excel_layout)

    # Use Excel
    use_excel = QCheckBox('Добавить протокол в журнал', app)
    var_layout.addWidget(use_excel)  # Добавляем в layout

    app.word_widgets.add_widget('добавить протокол в журнал', use_excel)


    """     Var r Layout    """

    # inspection_date
    inspection_date = QCalendarWidget(app)
    inspection_date.setFixedSize(500, 300)
    var_r_layout.addWidget(inspection_date)

    app.word_widgets.add_widget('дата поверки', inspection_date)

    # interval
    nums = ['1','2']

    label_interval_combo = QLabel("Выберите интервал до следующей проверки:")
    interval_combo = QComboBox()
    interval_combo.addItems(nums)
    interval_combo.setCurrentIndex(0)
    var_r_layout.addWidget(label_interval_combo)
    var_r_layout.addWidget(interval_combo)

    app.word_widgets.add_widget('интервал между поверками', interval_combo)

    # weather
    text_temperature = QPlainTextEdit(app)
    text_temperature.setPlaceholderText("Температура")

    var_r_layout.addWidget(text_temperature)

    app.word_widgets.add_widget('температура', text_temperature)

    text_humidity = QPlainTextEdit(app)
    text_humidity.setPlaceholderText("Влажность")
    var_r_layout.addWidget(text_humidity)

    app.word_widgets.add_widget('влажность', text_humidity)

    text_pressure = QPlainTextEdit(app)
    text_pressure.setPlaceholderText("Давление")
    var_r_layout.addWidget(text_pressure)

    app.word_widgets.add_widget('давление', text_pressure)


    # Создаем кнопки

    # Create template
    create_template_button = QPushButton('Создать шаблон', app)
    create_template_button.clicked.connect(create_template)  # Привязываем функцию
    var_r_layout.addWidget(create_template_button)  # Добавляем в layout

    app.word_widgets.add_widget('создать шаблон', create_template_button)

    # Create protocol
    create_protocol_button = QPushButton('Создать протокол', app)
    create_protocol_button.clicked.connect(create_protocol)  # Привязываем функцию
    var_r_layout.addWidget(create_protocol_button)  # Добавляем в layout

    app.word_widgets.add_widget('создать протокол', create_protocol_button)


    # Create protocol from excel
    create_protocol_from_excel_button = QPushButton('Создать протокол из журнала excel', app)
    create_protocol_from_excel_button.clicked.connect(create_protocol_from_excel)  # Привязываем функцию
    var_r_layout.addWidget(create_protocol_from_excel_button)  # Добавляем в layout

    app.word_widgets.add_widget('создать протокол из журнала', create_protocol_from_excel_button)


    # Line
    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setMinimumSize(300,2)
    line.setMaximumSize(700,2)
    var_r_layout.addWidget(line)


    # Use data
    use_data_button = QPushButton('Использовать преведущие данные', app)
    use_data_button.clicked.connect(use_data)  # Привязываем функцию
    var_r_layout.addWidget(use_data_button)  # Добавляем в layout

    app.word_widgets.add_widget('использовать преведущие данные', use_data_button)


    # Clean
    clean_button = QPushButton('Очистить', app)
    clean_button.clicked.connect(clean)  # Привязываем функцию
    var_r_layout.addWidget(clean_button)  # Добавляем в layout

    app.word_widgets.add_widget('очистить', clean_button)


    # 2 layout

    # Standarts

    tab_standarts = QTabWidget(app)
    tab_standarts.setMaximumSize(900, 950)
    tab_standarts.setMinimumSize(850, 800)

    # Укажите путь к нужной папке
    folder_path = f'data/standarts'
    file_names = os.listdir(folder_path)

    for file_name in file_names:
        tab = QWidget()
        Qtable = QTableWidget(tab)

        # Настройте таблицу
        Qtable.setSelectionBehavior(QAbstractItemView.SelectRows)  # Выбор целой строки
        Qtable.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Запрещаем редактирование
        # Добавление столбцов и строк в таблицу
        # Открываем docx документ
        document = Document(f"data/standarts/{file_name}")

        # Получаем первую таблицу из документа
        table = document.tables[0]

        # Устанавливаем количество строк и столбцов в QTableWidget
        Qtable.setRowCount(len(table.rows))
        Qtable.setColumnCount(len(table.columns))

        # Копируем содержимое таблицы в QTableWidget
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                Qtable.setItem(i, j, QTableWidgetItem(cell.text))

        # Устанавливаем ширину для первого столбца
        Qtable.setColumnWidth(0, 400)

        layout = QVBoxLayout(tab)
        layout.addWidget(Qtable)
        tab.setLayout(layout)

        tab_standarts.addTab(tab, file_name)

    var_r2_layout.addWidget(tab_standarts)

    app.word_widgets.add_widget('эталоны', tab_standarts)

    for widget in app.word_widgets.text_boxes.values():
        widget.setMinimumSize(300,40)
        widget.setMaximumSize(700,40)
    for widget in app.word_widgets.combo_boxes.values():
        widget.setMinimumSize(300,40)
        widget.setMaximumSize(700,40)
    for widget in app.word_widgets.check_boxes.values():
        widget.setMinimumSize(300,40)
        widget.setMaximumSize(700,40)
    for widget in app.word_widgets.buttons.values():
        widget.setMinimumSize(300,40)
        widget.setMaximumSize(700,40)

    # Add layouts in main_layout
    main_layout.addLayout(var_layout)
    main_layout.addLayout(var_r_layout)
    main_layout.addLayout(var_r2_layout)

    return main_layout


# FUNCTIONS

def search_company():
    try:
        company_name, legal_address = functions.search_company(self=main_window, inn=main_window.word_widgets.text_boxes['инн'].toPlainText())

        main_window.word_widgets.text_boxes['компания'].setPlainText(company_name)
        main_window.word_widgets.text_boxes['юридический адрес'].setPlainText(legal_address)
    except:
        ...

def scale_changed():
    ...

def verificationer_changed():
    ...

def create_protocol():
    ...
        
def use_data():
    config = config_client.get()

    errors = []

    # Словари с данными
    values_word = config['data_widgets']['word']
    values_excel = config['data_widgets']['excel']

    # Словари с bool значением
    data_word_widgets = config['save_widgets']['word']
    data_excel_widgets = config['save_widgets']['excel']

    # Word
    # text_boxes
    for name, widget in main_window.word_widgets.text_boxes.items():
        try:
            if name in data_word_widgets.keys(): 
                if data_word_widgets[name]:
                    widget.setPlainText(values_word['name'])
        except:
            errors.append(name)

    # combo_boxes
    for name, widget in main_window.word_widgets.combo_boxes.items():
        try:
            if name in data_word_widgets.keys(): 
                if data_word_widgets[name]:
                    widget.setPlainText(values_word['name'])
        except:
            errors.append(name)

    # check_boxes
    for name, widget in main_window.word_widgets.check_boxes.items():
        try:
            if name in data_word_widgets.keys(): 
                if data_word_widgets[name]:
                    widget.setChecked(values_word['name'])
        except:
            errors.append(name)

    # Excel
    # text_boxes
    for name, widget in main_window.excel_widgets.text_boxes.items():
        try:
            if name in data_excel_widgets.keys(): 
                if data_excel_widgets[name]:
                    widget.setPlainText(values_word['name'])
        except:
            errors.append(name)

    # combo_boxes
    for name, widget in main_window.excel_widgets.combo_boxes.items():
        try:
            if name in data_excel_widgets.keys(): 
                if data_excel_widgets[name]:
                    widget.setPlainText(values_word['name'])
        except:
            errors.append(name)

    # check_boxes
    for name, widget in main_window.excel_widgets.check_boxes.items():
        try:
            if name in data_excel_widgets.keys(): 
                if data_excel_widgets[name]:
                    widget.setChecked(values_word['name'])
        except:
            errors.append(name)

    if len(errors) > 0:
        msg = ''
        for err in errors:
            msg += err + '\n'

        QMessageBox.warning(main_window, 'Ошибка. Не удалось получить:', msg)
    
        
def clean():
    for widget in main_window.word_widgets.text_boxes.values():
        widget.clear()
    for widget in main_window.word_widgets.combo_boxes.values():
        widget.setCurrentIndex(-1)
    for widget in main_window.word_widgets.check_boxes.values():
        widget.setChecked(False)

    for widget in main_window.excel_widgets.text_boxes.values():
        widget.clear()
    for widget in main_window.excel_widgets.combo_boxes.values():
        widget.setCurrentIndex(-1)
    for widget in main_window.excel_widgets.check_boxes.values():
        widget.setChecked(False)
        
def create_template():
    names = []
    for widget in main_window.word_widgets.text_boxes.keys():
        if widget == 'путь к журналу' or widget == 'путь сохранения': continue
        names.append(widget)
    for widget in main_window.word_widgets.combo_boxes.keys():
        if widget == 'интервал между поверками': continue
        names.append(widget)
    for widget in main_window.word_widgets.check_boxes.keys():
        if widget == 'добавить протокол в журнал': continue
        names.append(widget)
    for widget in main_window.word_widgets.calendars.keys():
        names.append(widget)
    for widget in main_window.word_widgets.tab_widgets.keys():
        names.append(widget)

    options = QFileDialog.Options()
    filePath, _ = QFileDialog.getOpenFileName(main_window, 'Выберите файл', '', 'Word Files (*.docx);', options=options)
    if filePath:
        CreateTemplateDialog(path=filePath, widgets_names=names).exec_()

# Dialogs
        
def create_protocol_from_excel():
    ...

def show_inspection_address_setting():
    class CreateProtocolFromExcelDialog(QDialog):
        def __init__(self):
            super().__init__()

            self.main_layout = QVBoxLayout()

            length_window = 1000
            width_window = 150

            self.setGeometry(200, 400, length_window, width_window)
            self.setWindowTitle('Выберите нужный вариант')

            self.var = [
                'Использовать введенный юр.адрес',
                'Использовать юр.адрес ООО \"ЕДИНИЦА ИЗМЕРЕНИЯ\"'
            ]

            # Таблица с вариантами
            self.combo = QLabel("Выберите нужный вариант:")
            self.combo = QTableWidget()
            self.combo.setRowCount(len(self.var))
            self.combo.setColumnCount(1)

            self.combo.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Запрещаем редактирование

            self.combo.setColumnWidth(0, 940)

            i = 0
            for var in self.var:
                self.combo.setItem(i,0, QTableWidgetItem(var))
                i += 1

            self.combo.setCurrentCell(-1,-1)
            self.combo.currentItemChanged.connect(self.change)

            self.main_layout.addWidget(self.combo)

            self.setLayout(self.main_layout)

        def change(self):
            if self.combo.currentRow() == 0:
                main_window.word_widgets.text_boxes['адрес поверки'].setPlainText(main_window.word_widgets.text_boxes['юридический адрес'].toPlainText())
            else:
                main_window.word_widgets.text_boxes['адрес поверки'].setPlainText("610027, Россия, Кировская область, город Киров, улица Красноармейская, дом 43А") 

            self.close()
            
    CreateProtocolFromExcelDialog().exec_()

def choose_scale():
    ChooseScaleDialog(text_scale_widget=main_window.word_widgets.text_boxes['весы'],word=True).exec_()

def add_save_path():
    dialog = QDialog()
    main_window.word_widgets.text_boxes['путь сохранения'].setPlainText(QFileDialog.getExistingDirectory(dialog, "Выберите папку"))


def settings():
    ...

def add_path_to_excel():
    options = QFileDialog.Options()
    filePath, _ = QFileDialog.getOpenFileName(main_window, 'Выберите файл', '', 'Excel Files (*.xlsx);', options=options)

    main_window.word_widgets.text_boxes['путь к журналу'].setPlainText(filePath)