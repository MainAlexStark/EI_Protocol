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

from .office import Word, Excel

WORD = Word()
EXCEL = Excel()



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

            # Получаем данные о юр.адресе компании
            legal_address = data['items'][0]['ЮЛ']['АдресПолн']
            
            logger.debug('end')
            
            return company_name, legal_address


        else:

            LIST_ORG = List_org()

            try:

                data = LIST_ORG.get_company_name_by_inn(INN)

                company_name = data['name']
                legal_address = data['address']
                
                return company_name, legal_address

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

    if 'Влагомеры' in self.text_boxes_word['scale'].toPlainText() and self.text_boxes_word['voltage'] is None:
        text_voltage = QPlainTextEdit(self)
        text_voltage.setPlaceholderText("Выберите напряжение:")
        text_voltage.setFixedSize(500, 40)
        self.var_r_layout.addWidget(text_voltage)

        self.text_boxes_word['voltage'] = text_voltage

        text_frequency = QPlainTextEdit(self)
        text_frequency.setPlaceholderText("Выберите частоту:")
        text_frequency.setFixedSize(500, 40)
        self.var_r_layout.addWidget(text_frequency)

        self.text_boxes_word['frequency'] = text_frequency
    else:
        if 'voltage' in self.text_boxes_word:
            self.text_boxes_word['voltage'].close()
            self.text_boxes_word['frequency'].close()

            del self.text_boxes_word['frequency']
            del self.text_boxes_word['voltage']

    logger.debug('end')






def get_selected_table(self, tab_widget):
    logger.debug('start')
    
    try:
    
        current_tab_index = tab_widget.currentIndex()
        current_tab = tab_widget.widget(current_tab_index)
        table = current_tab.findChild(QTableWidget)
        
        logger.debug(f'TAB_WIDGET={tab_widget}')
        logger.debug(f'TABLE={table}')

        logger.debug('end')

        return table

    except Exception as e:
        
        logger.error(e)

    






def verificationer_changed(self, verificationer_combo):
    logger.debug('Изменение поверителя')

    # Определите имя файла хранилища
    path_to_data = 'app/tools/data/config.json'
    data = get_data(file_name=path_to_data)

    verificationers = data['verificationers']

    logger.debug('Поверитель изменен')
    
    return verificationers[verificationer_combo.currentText()]







def create_protocol(self, word:bool,other_widgets, text_boxes, combo_boxes, checkable_buttons):
    logger.info('start')

    args = {}

    tab_standarts = other_widgets['tab_standarts']
    selected_table = get_selected_table(self, tab_widget=tab_standarts)
    current_row = selected_table.currentRow()
    
    logger.debug(f'OTHER_WIDGETS={other_widgets}')
    logger.debug(f'Selected table={selected_table}')
    logger.debug(f'current_row={current_row}')

    # Получаем standarts
    standarts = selected_table.item(current_row, 1).text()  
    standarts_briefly = selected_table.item(current_row, 4).text()
    
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

        for text_widget_name in text_boxes.keys():
            
            text = str(text_boxes[text_widget_name].toPlainText()).strip()
            
            PlaceHolderText = text_boxes[text_widget_name].placeholderText
            
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
            
        for combo_widget_name in combo_boxes.keys():
            
            text = str(combo_boxes[combo_widget_name].currentText()).strip()
            
            if not combo_boxes[combo_widget_name].currentIndex() == -1:
                
                args[combo_widget_name] =data[combo_widget_name] = text
                
            else:
                message_box = QMessageBox()
                message_box.setIcon(QMessageBox.Critical)
                message_box.setText(f"Не выбранно значение в одном из выпадающих списков!")
                message_box.setWindowTitle("Ошибка")
                message_box.setStandardButtons(QMessageBox.Ok)
                message_box.exec_()
                
                var_flag = True

        for button_widget_name in checkable_buttons.keys():
            args[button_widget_name] =data[button_widget_name] = checkable_buttons[button_widget_name].isChecked()

        # Внесите изменения в данные
        args['inspection_date'] = data['inspection_date'] = str(other_widgets['inspection_date'].selectedDate().toString("dd.MM.yyyy")).strip()
        args['standarts'] = data['standarts'] = standarts
        args['standarts_briefly'] = standarts_briefly

        # Запишите обновленные данные обратно в файл
        file.seek(0)  # Переместите курсор в начало файла
        json.dump(data, file)
        file.truncate()  # Обрежьте файл, если новые данные занимают меньше места, чем предыдущие

    if not var_flag:
        if word:
            # Создаем протокол
            result = WORD.make_new_protocol(args=args)
        else:
            result = EXCEL.make_new_protocol(args=args)

        dialogs.CreateProtocolDialog(self, result=result).exec_()

        # Очищаем
        self.text_boxes_excel['num_protocol'].clear()
        self.text_boxes_excel['num_scale'].clear()


        logger.debug('end')
        
        
    logger.debug('end')


def use_data(self,other_widgets, text_boxes, combo_boxes, checkable_buttons):
    logger.info('start')

    # Определите имя файла хранилища
    file_name = 'app/tools/data/config.json'
    data = get_data(file_name=file_name)

    use_data = data['use_data']

    tab_standarts = other_widgets['tab_standarts']
    inspection_date_widget = other_widgets['inspection_date']


    # Определите имя файла хранилища
    path_to_data = 'app/tools/data/storage.json'
    data = get_data(file_name=path_to_data)

    # Перебираем все key в сохраненных значениях
    for name in data.keys():
        # Если в config напротив значения False, то пропускаем его
        if name in use_data.keys():
            if not use_data[name]: continue
        try:
            # Перебираем все key текстовых виджетов
            for widget_name in text_boxes.keys():
                # Если названия сходятся
                if name == widget_name:
                    # Устанавливаем значение из storage.json в текст виджета
                    text_boxes[widget_name].setPlainText(data[name])
        except Exception as e:
            logger.error(f'Ошибка при установлении значений в var_boxes.text_boxes: {e}')

        try:
            # Перебираем все key комбо виджетов
            for widget_name in combo_boxes.keys():
                # Если названия сходятся
                if name == widget_name:
                    # Устанавливаем значение из storage.json в текст виджета
                    combo_boxes[widget_name].setCurrentText(data[name])
        except:
            logger.error('Ошибка при установлении значений в var_boxes.combo_boxes')

        try:
            # Перебираем все key комбо виджетов
            for widget_name in checkable_buttons.keys():
                # Если названия сходятся
                if name == widget_name:
                    # Устанавливаем значение из storage.json в текст виджета
                    checkable_buttons[widget_name].setChecked(data[name])
        except:
            logger.error('Ошибка при установлении значений в buttons.CheckableButtons')


    if use_data['inspection_date']:
        inspection_date_str = data.get('inspection_date', '')
        if inspection_date_str:
            inspection_date_parts = inspection_date_str.split('.')
            if len(inspection_date_parts) == 3:
                day = int(inspection_date_parts[0])
                month = int(inspection_date_parts[1])
                year = int(inspection_date_parts[2])
                inspection_date_widget.setSelectedDate(QDate(year, month, day))

    if use_data['standarts']:
        # Перебор всех вкладок
        for index in range(tab_standarts.count()):

            current_tab = tab_standarts.widget(index)  # Получить виджет текущей вкладки

            table = current_tab.findChild(QTableWidget)

            for row in range(table.rowCount()):
                item = table.item(row, 1)
                if item.text() == data['standarts']:
                    tab_standarts.setCurrentIndex(index)  # Установить текущую вкладку
                    # Выделяем найденную строку
                    table.selectRow(row)


    logger.debug('end')


def clean(self,text_boxes, combo_boxes, checkable_buttons):
    logger.info('Очистить все виджеты')

    try:

        for key, text_widget in text_boxes.items():
            try:
                text_widget.clear()
            except Exception as e:
                logger.error(f'Ошибка при очищении значения {key}: {e}')

        for key, combo_widget in combo_boxes.items():
            try:
                combo_widget.setCurrentIndex(-1)
            except Exception as e:
                logger.error(f'Ошибка при очищении значения {key}: {e}')

        for key, button_widget in checkable_buttons.values():
            try:
                button_widget.setChecked(False)
            except Exception as e:
                logger.error(f'Ошибка при очищении значения {key}: {e}')


        logger.success('Успешно очищены все виджеты')

    except Exception as e:
        logger.error('Ошибка при очищении значений всех виджетов')


    







def create_template(self, word):
    logger.debug('Создание шаблона начальная функция')

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

            result = None
            
            if word:
                
                # Дополнительная проверка на формат
                if file.rsplit('.',1)[1] == 'docx':
                    result = WORD.create_template(file)
                else:
                    logger.error('Неверный формат файла (Нужен .docx)')
                    files[file] = 'Неверный формат файла (Нужен .docx)'
            else:
                logger.debug(f'Проверка на формат .xlsx. result={file.rsplit('.',1)[1]}')
                # Дополнительная проверка на формат
                if file.rsplit('.',1)[1] == 'xlsx':
                    result = EXCEL.create_template(file)
                else:
                    logger.error('Неверный формат файла (Нужен .xlsx)')
                    files[file] = 'Неверный формат файла (Нужен .xlsx)'
                
            if result:
                files[file] = 'Успешно'
            elif result == None:
                files[file] = "Непредвиденная ошибка"
            else:
                files[file] = result
                
            

            

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
