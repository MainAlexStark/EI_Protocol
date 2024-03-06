import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit, QComboBox, \
    QLabel, QCalendarWidget, QCompleter, QFileDialog, QTableWidget, QTableWidgetItem, \
    QTabWidget, QAbstractItemView, QCheckBox
from PyQt5.QtCore import QDate
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMessageBox, QDialog

import json
import os

from loguru import logger

from .db import get_data
 
from . import dialogs

from .parse.__init__ import FNS_API, List_org

from .word.Word import Word

WORD = Word()



def search_company(self):
    logger.debug('start')

    if len(self.text_INN.toPlainText()) > 0:

        # Определите имя файла хранилища
        file_name = 'app/tools/data/config.json'
        data = get_data(file_name=file_name)

        FNS_TOKEN = data.get('FNS_TOKEN', [])

        # Создаем обьект компании
        company = FNS_API(FNS_TOKEN)

        # Получаем INN
        INN = self.text_INN.toPlainText()
        
        # Получаем данные о компании по INN
        data = company.get_company_data(INN)

        if data is not True and len(data['items']) > 0:

            # Получаем данные о названии комании
            company_name = data['items'][0]['ЮЛ']['НаимСокрЮЛ']

            # Устанавливаем в QPlainTextEdit
            self.text_company.setPlainText(company_name)

            # Получаем данные о юр.адресе компании
            legal_address = data['items'][0]['ЮЛ']['АдресПолн']

            # Устанавливаем в QPlainTextEdit
            self.text_legal_address.setPlainText(legal_address)

            logger.debug('end')

        else:

            LIST_ORG = List_org()

            try:

                data = LIST_ORG.get_company_name_by_inn(INN)

                name = data['name']
                address = data['address']

                # Устанавливаем в QPlainTextEdit
                self.text_company.setPlainText(name)

                # Устанавливаем в QPlainTextEdit
                self.text_legal_address.setPlainText(address)

            except Exception as e:
                logger.error('Ошибка при работе с List-org.ru')


                message_box = QMessageBox()
                message_box.setIcon(QMessageBox.Critical)
                message_box.setText("Не удалось получить данные о компании!\nПроверьте введенный ИНН")
                message_box.setWindowTitle("Ошибка")
                message_box.setStandardButtons(QMessageBox.Ok)
                message_box.exec_()

                logger.debug('end')

    else:
        message_box = QMessageBox()
        message_box.setIcon(QMessageBox.Critical)
        message_box.setText(f"Введите ИНН!")
        message_box.setWindowTitle("Ошибка")
        message_box.setStandardButtons(QMessageBox.Ok)
        message_box.exec_()




def scale_changed(self):
    logger.info('start')

    if 'Влагомеры' in self.text_scale.toPlainText() and self.text_voltage is None:
        self.text_voltage = QPlainTextEdit(self)
        self.text_voltage.setPlaceholderText("Выберите напряжение:")
        self.text_voltage.setFixedSize(500, 40)
        self.var_r_layout.addWidget(self.text_voltage)

        self.var_boxes.text_boxes['voltage'] = self.text_voltage

        self.text_frequency = QPlainTextEdit(self)
        self.text_frequency.setPlaceholderText("Выберите частоту:")
        self.text_frequency.setFixedSize(500, 40)
        self.var_r_layout.addWidget(self.text_frequency)

        self.var_boxes.text_boxes['frequency'] = self.text_frequency
    else:
        if self.text_voltage is not None:
            self.text_voltage.close()
            self.text_frequency.close()

            del self.var_boxes.text_boxes['frequency']
            del self.var_boxes.text_boxes['voltage']

    logger.debug('end')






def get_selected_table(self):
    logger.debug('start')
    
    try:
    
        current_tab_index = self.tab_widget.currentIndex()
        current_tab = self.tab_widget.widget(current_tab_index)
        table = current_tab.findChild(QTableWidget)
        
        logger.debug(table)

        logger.debug('end')

        return table

    except Exception as e:
        
        logger.error(e)

    






def verificationer_changed(self):
    logger.debug('start')

    # Определите имя файла хранилища
    file_name = 'app/tools/data/config.json'
    data = get_data(file_name=file_name)

    verificationers = data['verificationers']

    self.text_num_protocol.setPlainText(verificationers[self.verificationer_combo.currentText()])

    logger.debug('end')







def create_protocol(self):
    logger.info('start')

    args = {}

    # Получаем standarts
    standarts = get_selected_table(self).item(get_selected_table(self).currentRow(), 1).text()
    standarts_briefly = get_selected_table(self).item(get_selected_table(self).currentRow(), 4).text()
    
    # Флаг если какая либо переменная не заполнена
    var_flag = False

    # Определите имя файла хранилища
    file_name = 'app/tools/data/storage.json'

    # Проверьте, существует ли файл, если нет, создайте пустой словарь
    if not os.path.exists(file_name):
        with open(file_name, 'w') as file:
            json.dump({}, file)

    # Откройте файл хранилища
    with open(file_name, 'r+') as file:

        # Загрузите данные из файла
        data = json.load(file)

        for text_widget_name in self.var_boxes.text_boxes.keys():
            
            text = str(self.var_boxes.text_boxes[text_widget_name].toPlainText()).strip()
            
            PlaceHolderText = self.var_boxes.text_boxes[text_widget_name].placeholderText
            
            if len(text) > 0:
                args[text_widget_name] = data[text_widget_name] = text
            else:
                message_box = QMessageBox()
                message_box.setIcon(QMessageBox.Critical)
                message_box.setText(f"Введите {PlaceHolderText}!")
                message_box.setWindowTitle("Ошибка")
                message_box.setStandardButtons(QMessageBox.Ok)
                message_box.exec_()
                
                var_flag = True
            
        for combo_widget_name in self.var_boxes.combo_boxes.keys():
            
            text = str(self.var_boxes.combo_boxes[combo_widget_name].currentText()).strip()
            
            if not self.var_boxes.combo_boxes[combo_widget_name].currentIndex() == -1:
                
                args[combo_widget_name] =data[combo_widget_name] = text
                
            else:
                message_box = QMessageBox()
                message_box.setIcon(QMessageBox.Critical)
                message_box.setText(f"Не выбранно значение в одном из выпадающих списков!")
                message_box.setWindowTitle("Ошибка")
                message_box.setStandardButtons(QMessageBox.Ok)
                message_box.exec_()
                
                var_flag = True

        for button_widget_name in self.buttons.CheckableButtons.keys():
            args[button_widget_name] =data[button_widget_name] = self.buttons.CheckableButtons[button_widget_name].isChecked()

        # Внесите изменения в данные
        args['inspection_date'] = data['inspection_date'] = str(self.inspection_date.selectedDate().toString("dd.MM.yyyy")).strip()
        args['standarts'] = data['standarts'] = standarts
        args['standarts_briefly'] = standarts_briefly

        # Запишите обновленные данные обратно в файл
        file.seek(0)  # Переместите курсор в начало файла
        json.dump(data, file)
        file.truncate()  # Обрежьте файл, если новые данные занимают меньше места, чем предыдущие

    if not var_flag:
        if args['create_excel'].isChecked():
            ...
        else:
            # Создаем протокол
            result = WORD.make_new_protocol(args)

        dialogs.CreateProtocolDialog(self, result=result).exec_()

        # Очищаем
        self.text_num_protocol.clear()
        self.text_num_scale.clear()

        self.verificationer_changed()


        logger.debug('end')
        
        
    logger.debug('end')


def use_data(self):
    logger.info('start')

    # Определите имя файла хранилища
    file_name = 'app/tools/data/config.json'
    data = get_data(file_name=file_name)

    use_data = data['use_data']


    # Определите имя файла хранилища
    file_name = 'app/tools/data/storage.json'
    data = get_data(file_name=file_name)

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


    logger.debug('end')









def clean(self):
    logger.info('start')

    for text_widget in self.var_boxes.text_boxes.values():
        text_widget.clear()
    for combo_widget in self.var_boxes.combo_boxes.values():
        combo_widget.setCurrentIndex(-1)
    for button_widget in self.buttons.CheckableButtons.values():
        button_widget.setChecked(False)


    logger.debug('end')







def create_template(self):
    logger.debug('start')

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

            result = WORD.create_template(file)
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

            logger.debug('Create_template_Dialog(QDialog): __init__')

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

            logger.debug('__init__ end')

    Create_template_Dialog().exec_()

    logger.debug('end')
