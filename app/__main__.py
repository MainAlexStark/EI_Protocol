import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QCompleter, QDialog, QFileDialog, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView
from PyQt5.QtCore import QDate
from PyQt5 import QtCore

import json
import os

from docx import Document

import functions
from dialogs import settings, choose_scale

from loguru import logger

main_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.text_scale = None
        self.label_text_scale = None
        self.button_choose_scale = None
        self.label_text_frequency = None
        self.label_text_voltage = None
        self.text_voltage = None
        self.text_frequency = None
        self.tab_widget = None
        self.button_settings = None
        self.button_clean = None
        self.button_use_data = None
        self.button_use_excel = None
        self.button_create_protocol = None
        self.button_create_template = None
        self.text_pressure = None
        self.label_text_pressure = None
        self.text_humidity = None
        self.label_text_humidity = None
        self.text_temperature = None
        self.label_text_temperature = None
        self.inspection_date = None
        self.button_unfit = None
        self.text_inspection_address = None
        self.label_text_inspection_address = None
        self.text_legal_address = None
        self.label_text_legal_address = None
        self.text_INN = None
        self.label_text_INN = None
        self.text_company = None
        self.label_text_company = None
        self.verificationer_combo = None
        self.label_verificationer_combo = None
        self.text_num_scale = None
        self.label_text_num_scale = None
        self.text_num_protocol = None
        self.label_text_num_protocol = None
        self.button_path = None
        self.text_path = None
        self.label_text_path = None
        self.initUI()

    def initUI(self):

        # Logger
        logger.add("log.txt")
        logger.info('Start initUI')


        file_path = os.path.abspath(__file__)
        main_path = os.path.dirname(os.path.dirname(file_path))


        # Определите имя файла хранилища
        file_name = 'config.json'

        # Откройте файл хранилища
        with open(file_name, 'r+') as file:
            # Загрузите данные из файла
            data = json.load(file)

            window_size = data['window_size']

        length_window = window_size['length']
        width_window = window_size['width']

        self.setWindowTitle('AUTO Listings')
        self.setGeometry(0, 0, length_window, width_window)

        # Layouts
        # Создание layouts ( контейнеры в которые погружаются остальные элементы )
        main_layout = QHBoxLayout()
        self.var_layout = QVBoxLayout()
        self.var_r_layout = QVBoxLayout()
        self.var_r2_layout = QVBoxLayout()

        # Variables

        class VarBoxes:
            combo_boxes = {}
            text_boxes = {}

        class Buttons:
            Buttons = {}
            CheckableButtons = {}

        self.var_boxes = VarBoxes()
        self.buttons = Buttons()

        # Инициализируем элементы ввода

        # Scale

        self.label_text_scale = QLabel("Весы:")
        self.text_scale = QPlainTextEdit(self)
        self.text_scale.setFixedSize(500, 40)

        self.text_scale.textChanged.connect(self.scale_changed)

        self.var_layout.addWidget(self.label_text_scale)
        self.var_layout.addWidget(self.text_scale)

        self.var_boxes.text_boxes['scale'] = self.text_scale

        # Надпись
        self.button_choose_scale = QPushButton('Выбрать весы', self)
        self.button_choose_scale.setFixedSize(500, 40)
        self.button_choose_scale.clicked.connect(self.choose_scale)  # Привязываем функцию
        self.var_layout.addWidget(self.button_choose_scale)  # Добавляем в layout

        self.buttons.Buttons['choose_scale'] = self.button_choose_scale

        # path

        # Надпись
        self.label_text_path = QLabel("Выберите путь сохранения:")
        # Текстовое поле
        self.text_path = QPlainTextEdit(self)
        self.text_path.setFixedSize(500, 40)  # Указывем размер
        # Добавляем элементы в layouts
        self.var_layout.addWidget(self.label_text_path)
        self.var_layout.addWidget(self.text_path)

        self.var_boxes.text_boxes['path'] = self.text_path

        # Создаем кнопку
        self.button_path = QPushButton('Выбрать путь сохранения', self)
        self.button_path.setFixedSize(500, 40)
        self.button_path.clicked.connect(self.add_path)  # Привязываем функцию
        self.var_layout.addWidget(self.button_path)  # Добавляем в layout

        self.buttons.Buttons['path'] = self.button_path

        # num_protocol
        # Надпись
        self.label_text_num_protocol = QLabel("Выберите номер протокола:")
        self.text_num_protocol = QPlainTextEdit(self)
        # Текстовое поле
        self.text_num_protocol.setFixedSize(500, 40)  # Указывем размер
        # Добавляем элементы в layouts
        self.var_layout.addWidget(self.label_text_num_protocol)
        self.var_layout.addWidget(self.text_num_protocol)

        self.var_boxes.text_boxes['num_protocol'] = self.text_num_protocol

        # num_scale
        self.label_text_num_scale = QLabel("Выберите номер весов:")
        self.text_num_scale = QPlainTextEdit(self)
        self.text_num_scale.setFixedSize(500, 40)
        self.var_layout.addWidget(self.label_text_num_scale)
        self.var_layout.addWidget(self.text_num_scale)

        self.var_boxes.text_boxes['num_scale'] = self.text_num_scale

        # Verificationer
        # Определите имя файла хранилища
        file_name = 'config.json'

        # Массив с поверителями
        ver = []

        # Откройте файл хранилища
        with open(file_name, 'r+') as file:
            # Загрузите данные из файла
            data = json.load(file)

            verificationers = data.get('verificationers', [])
            verificationers_text = ''
            for verificationer in verificationers.keys():
                verificationers_text += verificationer
            verificationers_text = '\n'.join(verificationers)

            for i in verificationers_text.split('\n'): ver.append(i)

        self.label_verificationer_combo = QLabel("Выберите поверителя:")
        self.verificationer_combo = QComboBox()
        self.verificationer_combo.addItems(verificationers)
        self.verificationer_combo.setCurrentIndex(-1)
        self.verificationer_combo.setFixedSize(500, 40)
        self.verificationer_combo.currentTextChanged.connect(self.verificationer_changed)
        self.var_layout.addWidget(self.label_verificationer_combo)
        self.var_layout.addWidget(self.verificationer_combo)

        # Добавляем в словарь combo boxes
        self.var_boxes.combo_boxes['verificationer'] = self.verificationer_combo

        # Company

        self.label_text_company = QLabel("Выберите компанию:")
        self.text_company = QPlainTextEdit(self)
        self.text_company.setFixedSize(500, 40)
        self.var_layout.addWidget(self.label_text_company)
        self.var_layout.addWidget(self.text_company)

        self.var_boxes.text_boxes['company'] = self.text_company

        # INN

        self.label_text_INN = QLabel("Выберите ИНН компании:")
        self.text_INN = QPlainTextEdit(self)
        self.text_INN.setFixedSize(500, 40)
        self.var_layout.addWidget(self.label_text_INN)
        self.var_layout.addWidget(self.text_INN)

        self.var_boxes.text_boxes['INN'] = self.text_INN

        # Legal address

        self.label_text_legal_address = QLabel("Выберите юридический адрес компании:")
        self.text_legal_address = QPlainTextEdit(self)
        self.text_legal_address.setFixedSize(500, 40)
        self.var_layout.addWidget(self.label_text_legal_address)
        self.var_layout.addWidget(self.text_legal_address)

        self.var_boxes.text_boxes['legal_address'] = self.text_legal_address

        # inspection address

        self.label_text_inspection_address = QLabel("Выберите адрес поверки:")
        self.text_inspection_address = QPlainTextEdit(self)
        self.text_inspection_address.setFixedSize(500, 40)
        self.var_layout.addWidget(self.label_text_inspection_address)
        self.var_layout.addWidget(self.text_inspection_address)

        self.var_boxes.text_boxes['inspection_address'] = self.text_inspection_address

        # Unfit

        self.button_unfit = QPushButton('Несоответсвует', self)
        self.button_unfit.setCheckable(True)
        self.button_unfit.setFixedSize(500, 40)
        self.var_layout.addWidget(self.button_unfit)  # Добавляем в layout

        self.buttons.CheckableButtons['unfit'] = self.button_unfit

        ####### var_r_layout

        # inspection_date

        self.inspection_date = QCalendarWidget(self)
        self.inspection_date.setFixedSize(500, 300)
        self.var_r_layout.addWidget(self.inspection_date)

        # weather

        self.label_text_temperature = QLabel("Выберите температуру:")
        self.text_temperature = QPlainTextEdit(self)
        self.text_temperature.setFixedSize(500, 40)
        self.var_r_layout.addWidget(self.label_text_temperature)
        self.var_r_layout.addWidget(self.text_temperature)

        self.var_boxes.text_boxes['temperature'] = self.text_temperature

        self.label_text_humidity = QLabel("Выберите влажность:")
        self.text_humidity = QPlainTextEdit(self)
        self.text_humidity.setFixedSize(500, 40)
        self.var_r_layout.addWidget(self.label_text_humidity)
        self.var_r_layout.addWidget(self.text_humidity)

        self.var_boxes.text_boxes['humidity'] = self.text_humidity

        self.label_text_pressure = QLabel("Выберите давление:")
        self.text_pressure = QPlainTextEdit(self)
        self.text_pressure.setFixedSize(500, 40)
        self.var_r_layout.addWidget(self.label_text_pressure)
        self.var_r_layout.addWidget(self.text_pressure)

        self.var_boxes.text_boxes['pressure'] = self.text_pressure

        # Создаем кнопки

        # Create template

        self.button_create_template = QPushButton('Создать шаблон', self)
        self.button_create_template.setFixedSize(500, 40)
        self.button_create_template.clicked.connect(self.create_template)  # Привязываем функцию
        self.var_r_layout.addWidget(self.button_create_template)  # Добавляем в layout

        self.buttons.Buttons['create_template'] = self.button_create_template

        # Create protocol

        self.button_create_protocol = QPushButton('Создать протокол', self)
        self.button_create_protocol.setFixedSize(500, 40)
        self.button_create_protocol.clicked.connect(self.create_protocol)  # Привязываем функцию
        self.var_r_layout.addWidget(self.button_create_protocol)  # Добавляем в layout

        self.buttons.Buttons['create_protocol'] = self.button_create_protocol

        # Use excel

        self.button_use_excel = QPushButton('Использовать Excel', self)
        self.button_use_excel.setCheckable(True)
        self.button_use_excel.setFixedSize(500, 40)
        self.var_r_layout.addWidget(self.button_use_excel)  # Добавляем в layout

        self.buttons.CheckableButtons['use_excel'] = self.button_use_excel

        # Use data

        self.button_use_data = QPushButton('Использовать преведущие данные', self)
        self.button_use_data.setFixedSize(500, 40)
        self.button_use_data.clicked.connect(self.use_data)  # Привязываем функцию
        self.var_r_layout.addWidget(self.button_use_data)  # Добавляем в layout

        self.buttons.Buttons['use_data'] = self.button_use_data

        # Clean

        self.button_clean = QPushButton('Очистить', self)
        self.button_clean.setFixedSize(500, 40)
        self.button_clean.clicked.connect(self.clean)  # Привязываем функцию
        self.var_r_layout.addWidget(self.button_clean)  # Добавляем в layout

        self.buttons.Buttons['clean'] = self.button_clean

        # Settings

        self.button_settings = QPushButton('Настройки', self)
        self.button_settings.setFixedSize(500, 40)
        self.button_settings.clicked.connect(self.settings)  # Привязываем функцию
        self.var_r_layout.addWidget(self.button_settings)  # Добавляем в layout

        self.buttons.Buttons['settings'] = self.button_settings

        # 2 layout

        # Standarts

        self.tab_widget = QTabWidget(self)
        self.tab_widget.setFixedSize(850, 800)

        # Укажите путь к нужной папке
        folder_path = f'{main_path}/standarts'
        file_names = os.listdir(folder_path)

        for file_name in file_names:
            tab = QWidget()
            Qtable = QTableWidget(tab)

            # Настройте таблицу
            Qtable.setSelectionBehavior(QAbstractItemView.SelectRows)  # Выбор целой строки
            Qtable.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Запрещаем редактирование
            # Добавление столбцов и строк в таблицу
            # Открываем docx документ
            document = Document(f"{main_path}/standarts/{file_name}")

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

            self.tab_widget.addTab(tab, file_name)

        self.var_r2_layout.addWidget(self.tab_widget)

        # Add layouts in main_layout
        main_layout.addLayout(self.var_layout)
        main_layout.addLayout(self.var_r_layout)
        main_layout.addLayout(self.var_r2_layout)

        # Добавляем main_layout в окно
        self.setLayout(main_layout)
        self.show()
        
        
    def choose_scale(self):
        logger.info('choose_scale')

        choose_scale.ChooseScaleDialog(self).exec_()



        

    def scale_changed(self):
        logger.info('scale_changed')

        if 'Влагомеры' in self.text_scale.toPlainText() and self.label_text_voltage is None:
            self.label_text_voltage = QLabel("Выберите напряжение:")
            self.text_voltage = QPlainTextEdit(self)
            self.text_voltage.setFixedSize(500, 40)
            self.var_r_layout.addWidget(self.label_text_voltage)
            self.var_r_layout.addWidget(self.text_voltage)

            self.var_boxes.text_boxes['voltage'] = self.text_voltage

            self.label_text_frequency = QLabel("Выберите частоту:")
            self.text_frequency = QPlainTextEdit(self)
            self.text_frequency.setFixedSize(500, 40)
            self.var_r_layout.addWidget(self.label_text_frequency)
            self.var_r_layout.addWidget(self.text_frequency)

            self.var_boxes.text_boxes['frequency'] = self.text_frequency
        else:
            if self.label_text_voltage is not None:
                self.label_text_voltage.close()
                self.text_voltage.close()
                self.label_text_frequency.close()
                self.text_frequency.close()

                del self.var_boxes.text_boxes['frequency']
                del self.var_boxes.text_boxes['voltage']

    def get_selected_table(self):
        current_tab_index = self.tab_widget.currentIndex()
        current_tab = self.tab_widget.widget(current_tab_index)
        table = current_tab.findChild(QTableWidget)
        return table

    def add_path(self):
        logger.info('add_path')
        dialog = QDialog()
        self.text_path.setPlainText(QFileDialog.getExistingDirectory(dialog, "Выберите папку"))

    def verificationer_changed(self):
        logger.info('verificationer_changed')
        # Определите имя файла хранилища
        file_name = 'config.json'

        # Откройте файл хранилища
        with open(file_name, 'r+') as file:
            # Загрузите данные из файла
            data = json.load(file)

            verificationers = data['verificationers']

            self.text_num_protocol.setPlainText(verificationers[self.verificationer_combo.currentText()])

    def create_protocol(self):
        logger.info('create_protocol')

        args = {}

        # Получаем standarts
        standarts = self.get_selected_table().item(self.get_selected_table().currentRow(), 1).text()

        # Определите имя файла хранилища
        file_name = 'storage.json'

        # Проверьте, существует ли файл, если нет, создайте пустой словарь
        if not os.path.exists(file_name):
            with open(file_name, 'w') as file:
                json.dump({}, file)

        # Откройте файл хранилища
        with open(file_name, 'r+') as file:

            # Загрузите данные из файла
            data = json.load(file)

            for text_widget_name in self.var_boxes.text_boxes.keys():
                args[text_widget_name] = data[text_widget_name] = str(self.var_boxes.text_boxes[text_widget_name].toPlainText()).strip()
            for combo_widget_name in self.var_boxes.combo_boxes.keys():
                args[combo_widget_name] =data[combo_widget_name] = str(self.var_boxes.combo_boxes[combo_widget_name].currentText()).strip()
            for button_widget_name in self.buttons.CheckableButtons.keys():
                args[button_widget_name] =data[button_widget_name] = self.buttons.CheckableButtons[button_widget_name].isChecked()

            # Внесите изменения в данные
            args['inspection_date'] = data['inspection_date'] = str(self.inspection_date.selectedDate().toString("dd.MM.yyyy")).strip()
            args['standarts'] = data['standarts'] = standarts

            # Запишите обновленные данные обратно в файл
            file.seek(0)  # Переместите курсор в начало файла
            json.dump(data, file)
            file.truncate()  # Обрежьте файл, если новые данные занимают меньше места, чем предыдущие

        # Создаем протокол
        functions.make_new_protocol(args)

        # Очищаем
        self.text_num_protocol.clear()
        self.text_num_scale.clear()

        self.verificationer_changed()

    def use_data(self):
        logger.info('use_data')

        # Определите имя файла хранилища
        file_name = 'config.json'

        # Откройте файл хранилища
        with open(file_name, 'r+') as file:
            # Загрузите данные из файла
            data = json.load(file)
            use_data = data['use_data']


        # Определите имя файла хранилища
        file_name = 'storage.json'

        # Откройте файл хранилища
        with open(file_name, 'r') as file:
            # Загрузите данные из файла
            data = json.load(file)


            # Перебираем все key в сохраненных значениях
            for name in data.keys():
                # Если в config напротив значения False, то пропускаем его
                if name in use_data.keys():
                    if not use_data[name]: continue
                try:
                    # Перебираем все key текстовых виджетов
                    for widget_name in self.var_boxes.text_boxes.keys():
                        # Если названия сходятся
                        if name == widget_name:
                            # Устанавливаем значение из storage.json в текст виджета
                            self.var_boxes.text_boxes[widget_name].setPlainText(data[name])
                except:
                    logger.error('Ошибка при установлении значений в self.var_boxes.text_boxes')

                try:
                    # Перебираем все key комбо виджетов
                    for widget_name in self.var_boxes.combo_boxes.keys():
                        # Если названия сходятся
                        if name == widget_name:
                            # Устанавливаем значение из storage.json в текст виджета
                            self.var_boxes.combo_boxes[widget_name].setCurrentText(data[name])
                except:
                    logger.error('Ошибка при установлении значений в self.var_boxes.combo_boxes')

                try:
                    # Перебираем все key комбо виджетов
                    for widget_name in self.buttons.CheckableButtons.keys():
                        # Если названия сходятся
                        if name == widget_name:
                            # Устанавливаем значение из storage.json в текст виджета
                            self.buttons.CheckableButtons[widget_name].setChecked(data[name])
                except:
                    logger.error('Ошибка при установлении значений в self.buttons.CheckableButtons')


            if use_data['inspection_date']:
                inspection_date_str = data.get('inspection_date', '')
                if inspection_date_str:
                    inspection_date_parts = inspection_date_str.split('.')
                    if len(inspection_date_parts) == 3:
                        day = int(inspection_date_parts[0])
                        month = int(inspection_date_parts[1])
                        year = int(inspection_date_parts[2])
                        self.inspection_date.setSelectedDate(QDate(year, month, day))

            if use_data['standarts']:
                # Перебор всех вкладок
                for index in range(self.tab_widget.count()):

                    current_tab = self.tab_widget.widget(index)  # Получить виджет текущей вкладки

                    table = current_tab.findChild(QTableWidget)

                    for row in range(table.rowCount()):
                        item = table.item(row, 1)
                        if item.text() == data['standarts']:
                            self.tab_widget.setCurrentIndex(index)  # Установить текущую вкладку
                            # Выделяем найденную строку
                            table.selectRow(row)


    def clean(self):
        logger.info('clean')

        for text_widget in self.var_boxes.text_boxes.values():
            text_widget.clear()
        for combo_widget in self.var_boxes.combo_boxes.values():
            combo_widget.setCurrentIndex(-1)
        for button_widget in self.buttons.CheckableButtons.values():
            button_widget.setChecked(False)

    def settings(self):
        logger.info('settings')

        settings.SettingDialog().exec_()

    def create_template(self):
        logger.info('create_template')

        # Открываем диалоговое окно проводника
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)

        # Словарь с реультатами файлов
        files = {}

        # Если пользователь выбрал файлы и нажал кнопку "Открыть"
        if file_dialog.exec_() == QFileDialog.Accepted:
            # Получаем выбранные файлы
            selected_files = file_dialog.selectedFiles()
            for file in selected_files:

                result = functions.create_template(file)
                if result == False:
                    files[file] = 'Успешно'
                else:
                    files[file] = result

                # Дополнительная проверка на формат
                if os.path.splitext(file)[1] == '.doc':
                    files[file] = 'Неверный формат файла (Нужен .docx)'

        class Create_template_Dialog(QDialog):
            def __init__(self):

                super().__init__()
                super().__init__()

                logger.info('Create_template_Dialog(QDialog)')

                self.setGeometry(200, 200, 400, 400)
                self.setWindowTitle('Создание шаблона')

                main_layout = QVBoxLayout()

                # path
                self.label_text_result = QLabel("Результат:")
                self.text_result = QPlainTextEdit(self)
                self.text_result.setFixedSize(500, 600)
                self.text_result.setReadOnly(True)

                for key in files.keys():
                    self.text_result.appendPlainText(f"{key}: {files[key]}\n")

                main_layout.addWidget(self.label_text_result)
                main_layout.addWidget(self.text_result)

                self.setLayout(main_layout)

        Create_template_Dialog().exec_()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
