from PyQt5.QtWidgets import QVBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QDialog, QPushButton
import json


class SettingDialog(QDialog):
    def __init__(self):
        super().__init__()

        self.setGeometry(200, 200, 400, 400)
        self.setWindowTitle('Настройки')

        main_layout = QVBoxLayout()

        # Поверители
        self.label_text_verificationers = QLabel("Поверители:")
        self.table_verificationers = QTableWidget()
        self.table_verificationers.setRowCount(10)
        self.table_verificationers.setColumnCount(2)
        main_layout.addWidget(self.label_text_verificationers)
        main_layout.addWidget(self.table_verificationers)

        # Определите имя файла хранилища
        file_name = 'config.json'

        # Откройте файл хранилища
        with open(file_name, 'r+') as file:
            # Загрузите данные из файла
            data = json.load(file)

            verificationers = data['verificationers']

            i = 0
            for verificationer in verificationers.keys():
                self.table_verificationers.setItem(i, 0, QTableWidgetItem(verificationer))
                self.table_verificationers.setItem(i, 1, QTableWidgetItem(verificationers[verificationer]))
                i += 1

        # Кнопки

        # Сохранить
        self.button_save = QPushButton('Сохранить', self)
        self.button_save.setFixedSize(500, 40)
        self.button_save.clicked.connect(self.save)  # Привязываем функцию
        main_layout.addWidget(self.button_save)  # Добавляем в layout

        self.setLayout(main_layout)

    def save(self):
        # Определите имя файла хранилища
        file_name = 'config.json'

        # Откройте файл хранилища
        with open(file_name, 'r+') as file:
            # Загрузите данные из файла
            data = json.load(file)

            # Очистите предыдущие данные
            data['verificationers'] = {}

            dictionary = data['verificationers']

            for i in range(self.table_verificationers.rowCount()):
                item_0 = self.table_verificationers.item(i, 0)
                item_1 = self.table_verificationers.item(i, 1)
                if item_0 is not None and item_1 is not None:
                    dictionary[item_0.text()] = item_1.text()

            # Запишите обновленные данные обратно в файл
            file.seek(0)  # Переместите курсор в начало файла
            json.dump(data, file)
            file.truncate()  # Обрежьте файл, если новые данные занимают меньше места, чем предыдущие

        self.close()