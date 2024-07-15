""" third party imports """
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QFrame, QDialog, QFileDialog, QCheckBox, QMenu, QAction, QMenuBar, QMessageBox
from loguru import logger
from PyQt5.QtWidgets import QWidget, QApplication, QLabel, QVBoxLayout, QScrollArea, QSizePolicy
from PyQt5.QtCore import QDate
import os
from docx import Document

from ..tools.ClassStorage import ClassStorage
from ..dialogs.choose_scale import ChooseScaleDialog
from ..dialogs.create_template import CreateTemplateDialog

from ..ConfigClient import Config


""" OPEN CONFIG """
file_path = 'data/config.json'
if os.path.exists(file_path):
    config_client = Config(file_path)
else:
    raise Exception(f'File {file_path} not found')

main_window = None

def get_layout(app: QWidget):
    logger.debug("Создание вкладки 'word'")
    
    config = config_client.get()
    
    global main_window
    main_window = app
    
    # Создаем хранилище виджетов
    app.word_widgets = ClassStorage()
    
    # Создаем layouts
    main_layout = QHBoxLayout()
    scroll_layout = QHBoxLayout()
    var_layouts = [QVBoxLayout(),QVBoxLayout(),QVBoxLayout()]
    
    # Scale

    scale = QPlainTextEdit()
    scale.setPlaceholderText('Весы')
    app.word_widgets.add_value('весы_надпись', scale)
    var_layouts[0].addWidget(scale)

    # Надпись
    choose_scale_button = QPushButton('Выбрать весы', app)
    choose_scale_button.clicked.connect(choose_scale)  # Привязываем функцию
    var_layouts[0].addWidget(choose_scale_button)  # Добавляем в layout
    app.word_widgets.add_value('весы', scale)
    
    # Путь сохранения

    path_layout = QHBoxLayout()

    save_path = QPlainTextEdit(app)
    save_path.setPlaceholderText('Путь сохрания')
    path_layout.addWidget(save_path)

    app.word_widgets.add_value('путь сохранения', save_path)

    save_path_button = QPushButton('...', app)
    save_path_button.setFixedSize(50, 50)
    save_path_button.clicked.connect(add_save_path)  # Привязываем функцию
    path_layout.addWidget(save_path_button)  # Добавляем в layout

    var_layouts[0].addLayout(path_layout)
    
    
    # num_protocol
    text_num_protocol = QPlainTextEdit(app)
    text_num_protocol.setPlaceholderText('Номер протокола')
    var_layouts[0].addWidget(text_num_protocol)
    app.word_widgets.add_value('номер протокола', text_num_protocol)

    # num_scale
    text_num_scale = QPlainTextEdit(app)
    text_num_scale.setPlaceholderText('Номер весов')
    var_layouts[0].addWidget(text_num_scale)
    app.word_widgets.add_value('номер весов', text_num_scale)
    
    # Поверитель
    label_verificationer_combo = QLabel("Выберите поверителя:")
    verificationer_combo = QComboBox()
    verificationer_combo.setFixedSize(500, 40)
    verificationer_combo.currentTextChanged.connect(verificationer_changed)
    var_layouts[0].addWidget(label_verificationer_combo)
    var_layouts[0].addWidget(verificationer_combo)

    app.word_widgets.add_value('поверитель', verificationer_combo)
    app.word_widgets.add_value('поверитель_надпись', label_verificationer_combo)
    
    # Line

    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setMinimumSize(300,2)
    line.setMaximumSize(700,2)

    var_layouts[0].addWidget(line)
    
    # INN
    INN = QPlainTextEdit(app)
    INN.setPlaceholderText("ИНН")
    var_layouts[0].addWidget(INN)

    app.word_widgets.add_value('инн', INN)

    search_company_button = QPushButton('Найти', app)
    search_company_button.clicked.connect(search_company)  # Привязываем функцию
    var_layouts[0].addWidget(search_company_button)  # Добавляем в layout

    app.word_widgets.add_value('найти компанию', search_company_button)
    
    # Company
    company = QPlainTextEdit(app)
    company.setPlaceholderText("Компания")
    var_layouts[0].addWidget(company)

    app.word_widgets.add_value('компания', company)
    
    # Legal address
    legal_address = QPlainTextEdit(app)
    legal_address.setPlaceholderText("Юридический адрес")
    var_layouts[0].addWidget(legal_address)

    app.word_widgets.add_value('юридический адрес', legal_address)
    
    # inspection address

    inspection_address_layout = QHBoxLayout()

    inspection_address = QPlainTextEdit(app)
    inspection_address.setPlaceholderText('Адрес поверки')
    inspection_address_layout.addWidget(inspection_address)

    app.word_widgets.add_value('адрес поверки', inspection_address)

    button_inspection_address_setting = QPushButton('...', app)
    button_inspection_address_setting.setFixedSize(50, 50)
    button_inspection_address_setting.clicked.connect(show_inspection_address_setting)  # Привязываем функцию
    inspection_address_layout.addWidget(button_inspection_address_setting)  # Добавляем в layout

    var_layouts[0].addLayout(inspection_address_layout)
    
    # Unfit
    unfit = QCheckBox('Несоответсвует', app)
    var_layouts[0].addWidget(unfit)  # Добавляем в layout

    app.word_widgets.add_value('пригодность', unfit)
    
    # Line
    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setMinimumSize(300,2)
    line.setMaximumSize(700,2)
    var_layouts[0].addWidget(line)
    
    # path to excel

    path_excel_layout = QHBoxLayout()

    path_to_excel_journal = QPlainTextEdit(app)
    path_to_excel_journal.setPlaceholderText("Путь к журналу excel")
    path_excel_layout.addWidget(path_to_excel_journal)
    app.word_widgets.add_value('путь к журналу', path_to_excel_journal)

    path_to_excel_dialog_button = QPushButton('...', app)
    path_to_excel_dialog_button.setFixedSize(50, 50)
    path_to_excel_dialog_button.clicked.connect(add_path_to_excel)  # Привязываем функцию
    path_excel_layout.addWidget(path_to_excel_dialog_button)  # Добавляем в layout

    var_layouts[0].addLayout(path_excel_layout)
    
    # Use Excel
    use_excel = QCheckBox('Добавить протокол в журнал', app)
    var_layouts[0].addWidget(use_excel)  # Добавляем в layout

    app.word_widgets.add_value('добавить протокол в журнал', use_excel)
    
    """   Second layout   """
    
    # inspection_date
    inspection_date = QCalendarWidget(app)
    inspection_date.setFixedSize(500, 300)
    var_layouts[1].addWidget(inspection_date)

    app.word_widgets.add_value('дата поверки', inspection_date)
    
    # interval
    nums = ['1','2']

    label_interval_combo = QLabel("Выберите интервал до следующей проверки:")
    interval_combo = QComboBox()
    interval_combo.addItems(nums)
    interval_combo.setCurrentIndex(0)
    var_layouts[1].addWidget(label_interval_combo)
    var_layouts[1].addWidget(interval_combo)

    app.word_widgets.add_value('интервал между поверками', interval_combo)
    app.word_widgets.add_value('интервал_надпись', label_interval_combo)
    
    # weather
    text_temperature = QPlainTextEdit(app)
    text_temperature.setPlaceholderText("Температура")

    var_layouts[1].addWidget(text_temperature)

    app.word_widgets.add_value('температура', text_temperature)

    text_humidity = QPlainTextEdit(app)
    text_humidity.setPlaceholderText("Влажность")
    var_layouts[1].addWidget(text_humidity)

    app.word_widgets.add_value('влажность', text_humidity)

    text_pressure = QPlainTextEdit(app)
    text_pressure.setPlaceholderText("Давление")
    var_layouts[1].addWidget(text_pressure)

    app.word_widgets.add_value('давление', text_pressure)
    
    ### Создаем кнопки
    # Create template
    create_template_button = QPushButton('Создать шаблон', app)
    create_template_button.clicked.connect(create_template)  # Привязываем функцию
    var_layouts[1].addWidget(create_template_button)  # Добавляем в layout

    app.word_widgets.add_value('создать шаблон', create_template_button)

    # Create protocol
    create_protocol_button = QPushButton('Создать протокол', app)
    create_protocol_button.clicked.connect(create_protocol)  # Привязываем функцию
    var_layouts[1].addWidget(create_protocol_button)  # Добавляем в layout

    app.word_widgets.add_value('создать протокол', create_protocol_button)


    # Create protocol from excel
    create_protocol_from_excel_button = QPushButton('Создать протокол из журнала excel', app)
    create_protocol_from_excel_button.clicked.connect(create_protocol_from_excel)  # Привязываем функцию
    var_layouts[1].addWidget(create_protocol_from_excel_button)  # Добавляем в layout

    app.word_widgets.add_value('создать протокол из журнала', create_protocol_from_excel_button)
    
    
    # Line
    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setMinimumSize(300,2)
    line.setMaximumSize(700,2)
    var_layouts[1].addWidget(line)
    
    # Use data
    use_data_button = QPushButton('Использовать преведущие данные', app)
    use_data_button.clicked.connect(use_data)  # Привязываем функцию
    var_layouts[1].addWidget(use_data_button)  # Добавляем в layout

    app.word_widgets.add_value('использовать преведущие данные', use_data_button)


    # Clean
    clean_button = QPushButton('Очистить', app)
    clean_button.clicked.connect(clean)  # Привязываем функцию
    var_layouts[1].addWidget(clean_button)  # Добавляем в layout

    app.word_widgets.add_value('очистить', clean_button)
    
    
    """ third layout """
    # Standarts

    tab_standarts = QTabWidget(app)
    tab_standarts.setMaximumSize(900, 950)
    tab_standarts.setMinimumSize(850, 800)
    
    tab_standarts.currentWidget
    

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

    var_layouts[2].addWidget(tab_standarts)

    app.word_widgets.add_value('эталоны', tab_standarts)
    
    # Устанавливаем минимальный размер для виджетов
    for name, widget in app.word_widgets.get_values_by_class(QPlainTextEdit).items():
        # Устанавливаем размерную политику для виджета
        widget.setMaximumWidth(500)
        widget.setMaximumHeight(40)
        
    for name, widget in app.word_widgets.get_values_by_class(QLabel).items():
        # Устанавливаем размерную политику для виджета
        widget.setMaximumWidth(500)
        widget.setMaximumHeight(40)
        
    # Добавляем все layouts в scroll_area
    for layout in var_layouts:
        scroll_layout.addLayout(layout)
    # Создаем QScrollArea и добавляем в него layout
    scroll_area = QScrollArea()
    scroll_area.setWidgetResizable(True) # разрешаем QScrollArea изменять размер своего виджета
    #scroll_layout.adjustSize()
    scroll_area.setWidget(QWidget()) # Создаем пустой виджет для QScrollArea
    scroll_area.widget().setLayout(scroll_layout)
    
    # Добавляем scroll_area в main layout
    main_layout.addWidget(scroll_area)
    
    logger.debug(f"Word widgets=\n{app.word_widgets}")
    
    return main_layout


def choose_scale():
    ChooseScaleDialog(text_scale_widget=main_window.word_widgets.get_value("весы"),word=True).exec_()
    
    
def search_company():
    ...

def scale_changed():
    ...

def verificationer_changed():
    if main_window.word_widgets.get_value('поверитель').currentIndex() != -1:
        main_window.word_widgets.get_value('номер протокола').setPlainText(f"ЕИ-03-{main_window.verificationers[main_window.word_widgets.get_value('поверитель').currentText()]}-")
    
def get_selected_table(tab_widget):
    try:
    
        current_tab_index = tab_widget.currentIndex()
        current_tab = tab_widget.widget(current_tab_index)
        table = current_tab.findChild(QTableWidget)

        return table

    except Exception as e:
        
        logger.error(e)
    
def create_protocol():
    config = config_client.get()
    
    tab_standarts = main_window.word_widgets.get_value("эталоны")
    selected_table = get_selected_table(tab_standarts)
    current_row = selected_table.currentRow()
   
    
    if current_row == -1:
        QMessageBox.warning(main_window, 'Ошибка ввода', f'Выберите эталоны')
        return
    else:
        # Получаем standarts
        standarts = selected_table.item(current_row, 1).text()  
        standarts_briefly = selected_table.item(current_row, 4).text()
    
    for cls, values in main_window.word_widgets.storage.items():
            for name, widget in values.items():
                if isinstance(widget, QWidget):
                    # Обновление виджета в зависимости от его типа
                    if hasattr(widget, 'setCurrentText'):
                        if widget.CurrentIndex() == -1:
                            QMessageBox.warning(main_window, 'Ошибка ввода', f'Поле "{name}" не может быть пустым.')
                            return
                    elif hasattr(widget, 'setPlainText'):
                        if widget.toPlainText().strip() == '':
                            QMessageBox.warning(main_window, 'Ошибка ввода', f'Поле "{name}" не может быть пустым.')
                            return
    
    # Сохраняем данные в config
    config['widgets']['data']['word'] = main_window.word_widgets.to_json()
    config_client.post(config)
    
    logger.debug(f"Сохраняемые данные={main_window.word_widgets.to_json()}")
    
    # Полчаем введенные данные
    data = {}
    for cls, values in main_window.word_widgets.storage.items():
        for name, value in values.items():
            if isinstance(value, QWidget):
                if hasattr(value, 'text'):
                    data[name] = value.text()
                elif hasattr(value, 'currentText'):
                    data[name] = value.currentText()
                elif hasattr(value, 'isChecked'):
                    data[name] = value.isChecked()
                elif hasattr(value, 'selectedDate'):
                    data[name] = value.selectedDate().toString("yyyy-MM-dd")
                elif hasattr(value, 'toPlainText'):
                    data[name] = value.toPlainText()
                    
    logger.debug(f"Введенные данные=\n{data}")
    
    
        
def use_data():
    config = config_client.get()
                    
    # Меняем значения виджетов
    app_storage=main_window.word_widgets
    update_list=config['widgets']['permissions']['word']
    widgets_data = config['widgets']['data']['word']
    
    logger.debug(f"widgets_data={widgets_data}\nupdate_list={update_list}")
                        
    for widget_class, widgets in widgets_data.items():
        for name, value in widgets.items():
            if name in update_list.keys():
                if update_list[name]:
                    widget = app_storage.get_value(name)
                    class_name = type(widget).__name__
                    if isinstance(widget, QWidget):
                        # Обновление виджета в зависимости от его типа
                        if hasattr(widget, 'setText'):
                            widget.setText(value)
                        elif hasattr(widget, 'setCurrentText'):
                            widget.setCurrentText(value)
                        elif hasattr(widget, 'setChecked'):
                            widget.setChecked(value)
                        elif hasattr(widget, 'setSelectedDate'):
                            date_object = QDate.fromString(value, 'yyyy-MM-dd')
                            widget.setSelectedDate(date_object)
                        elif hasattr(widget, 'setPlainText'):
                            widget.setPlainText(value)
        
def clean():
    for cls, values in main_window.word_widgets.storage.items():
            for name, widget in values.items():
                if isinstance(widget, QWidget):
                    # Обновление виджета в зависимости от его типа
                    if hasattr(widget, 'setCurrentText'):
                        widget.setCurrentIndex(-1)
                    elif hasattr(widget, 'setChecked'):
                        widget.setChecked(False)
                    elif hasattr(widget, 'setPlainText'):
                        widget.setPlainText('')
            
    name = main_window.user._first_name + ' ' + main_window.user._last_name
    main_window.word_widgets.get_value("поверитель").setCurrentText(name)
    
        
def create_template():
    names:list = main_window.word_widgets.widgets_names
    logger.debug(f"Names={names}")
    i = 0
    for name in names:
        try:
            if "путь к журналу" in names: names.remove("путь к журналу")
            if "путь сохранения" in names: names.remove("путь сохранения")
            if "интервал между поверками" in names: names.remove("интервал между поверками")
            if "добавить протокол в журнал" in names: names.remove("добавить протокол в журнал")
            if "поверитель_надпись" in names: names.remove("поверитель_надпись")
            if "найти компанию" in names: names.remove("найти компанию")
            if "интервал_надпись" in names: names.remove("интервал_надпись")
            if "создать шаблон" in names: names.remove("создать шаблон")
            if "создать протокол" in names: names.remove("создать протокол")
            if "создать протокол из журнала" in names: names.remove("создать протокол из журнала")
            if "использовать преведущие данные" in names: names.remove("использовать преведущие данные")
            if "очистить" in names: names.remove("очистить")
            if "весы_надпись" in names: names.remove("весы_надпись")
        except Exception as e:
            logger.error(f"{e}: {name}")
        i += 1
    logger.debug(f"Names={names}")
    options = QFileDialog.Options()
    filePath, _ = QFileDialog.getOpenFileName(main_window, 'Выберите файл', '', 'Word Files (*.docx);', options=options)
    if filePath:
        CreateTemplateDialog(path=filePath, widgets_names=names).exec_()
        
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
                main_window.word_widgets.get_value('адрес поверки').setPlainText(main_window.word_widgets.get_value('юридический адрес').toPlainText())
            else:
                main_window.word_widgets.get_value('адрес поверки').setPlainText("610027, Россия, Кировская область, город Киров, улица Красноармейская, дом 43А") 

            self.close()
            
    CreateProtocolFromExcelDialog().exec_()

def add_save_path():
    dialog = QDialog()
    main_window.word_widgets.get_value('путь сохранения').setPlainText(QFileDialog.getExistingDirectory(dialog, "Выберите папку"))

def add_path_to_excel():
    options = QFileDialog.Options()
    filePath, _ = QFileDialog.getOpenFileName(main_window, 'Выберите файл', '', 'Excel Files (*.xlsx);', options=options)

    main_window.word_widgets.get_value('путь к журналу').setPlainText(filePath)