from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QMessageBox, QDialog
from ..ConfigClient import Config
import os
from loguru import logger
import time

from ..api import Client

""" OPEN CONFIG """
file_path = 'data/config.json'
if os.path.exists(file_path):
    config_client = Config(file_path)
else:
    raise Exception(f'File {file_path} not found')

main_window = None

class LoginDialog(QDialog):
    def __init__(self, app: QWidget):
        """
        Диалог входа и регистрации
        При нажатии на кнопку регистрации добавляются дополнительные окна для ввода, нужных для регистрации (func: get_registration)
        """
        super().__init__()
        logger.info("Создание UI Login Dialog")
        
        config = config_client.get()
        
        global main_window
        main_window = app
        
        self.setWindowTitle('Вход в аккаунт')
        
        # Создаем layout (Один в этом диалоге)
        self.layout = QVBoxLayout()
        
        # Создаем виджеты ввода обязательных данных для входа
        self.login_widgets = {}
        for name, type in config['api']['login_requirements'].items():
            # Создаем виджет
            if type == "QLineEdit": 
                widget = QLineEdit()
                widget.setPlaceholderText(name)
            # Добавляем виджет в словарь
            self.login_widgets[name] = widget
            # Добавляем виджет в layout
            self.layout.addWidget(widget)
            
        logger.debug(f"login_widgets=\n{self.login_widgets}")
        
        # Создаем кнопки
        self.login_button = QPushButton('Войти')
        self.login_button.clicked.connect(self.login)
        self.reg_button = QPushButton('Зарегистрироваться')
        self.reg_button.clicked.connect(self.get_registration)
        
        # Добавляем виджеты на layout
        self.layout.addWidget(self.login_button)
        self.layout.addWidget(self.reg_button)
        
        self.setLayout(self.layout)
        
        # Если есть сохраненные данные, то входим по ним
        email = config['app']['email']
        password = config['app']['password']
        
        if email != "" and password != "":
            self.login_widgets['email'].setText(email)
            self.login_widgets['password'].setText(password)
        
        
        
    def login(self):
        logger.info("Отправка запроса на login")
        config = config_client.get()
        
        # Проверяем, что все обязательные поля заполнены
        for name, widget in self.login_widgets.items():
            if not widget.text().strip():
                QMessageBox.warning(self, 'Ошибка ввода', f'Поле "{name}" не может быть пустым.')
                return
        
        email = self.login_widgets['email'].text()
        password = self.login_widgets['password'].text()
        
        main_window.user = Client(email=email, password=password)
        res = main_window.user._login()
        
        if res['retCode'] == 0:
            config['app']['email'] = self.login_widgets['email'].text()
            config['app']['password'] = self.login_widgets['password'].text()
            config_client.post(config)
            self.accept()
            logger.debug("Вход выполнен успешно")
            
            # Вставляем нужные значения в main_window
            main_window.word_widgets.get_value("номер протокола").setPlainText(f"ЕИ-03-{main_window.user._workplace_number}-")
            
            name = main_window.user._first_name + ' ' + main_window.user._last_name
            names = main_window.user.get_users_names()['result']
            main_window.verificationers = {}
            for user_names in names:
                main_window.verificationers[user_names['first_name'] + ' ' + user_names['last_name']] = user_names['workplace_number']
                
            main_window.word_widgets.get_value("поверитель").addItems(main_window.verificationers.keys())
            main_window.word_widgets.get_value("поверитель").setCurrentText(name)
        else:
            QMessageBox.warning(self, 'Ошибка входа', str(res['retMsg']))
            return True
        
    def registration(self):
        logger.info("Отправка запроса на регистрацию")
        
        # Проверяем, что все обязательные поля заполнены
        for name, widget in self.registration_widgets.items():
            if not widget.text().strip():
                QMessageBox.warning(self, 'Ошибка ввода', f'Поле "{name}" не может быть пустым.')
                return
        
        email = self.registration_widgets['email'].text()
        password = self.registration_widgets['password'].text()
        first_name = self.registration_widgets['first_name'].text()
        last_name = self.registration_widgets['last_name'].text()
        workplace_number = self.registration_widgets['workplace_number'].text()
        
        main_window.user = Client(email=email, password=password)
        res = main_window.user.register(first_name, last_name, workplace_number)
        
        if res['retCode'] == 0:
            QMessageBox.information(self, 'Успешная регистрация', 'Регистрация прошла успешно.')
            self.accept()
            logger.debug("Регистрация выполнена успешно")
        else:
            QMessageBox.warning(self, 'Ошибка регистрации', str(res['retMsg']))
        
    def get_registration(self):
        logger.debug("Добавление виджетов регистрации")
        config = config_client.get()
        
        # Удаляем виджет кнопки "Войти"
        self.login_button.setParent(None)
        # Удаляем виджет кнопки "Зарегистрироваться"
        self.reg_button.setParent(None)
        # Удаляем виджеты ввода
        for name, widget in self.login_widgets.items():
            self.layout.removeWidget(widget)
            widget.deleteLater()
            
        # Устанавливаем новые виджеты
        self.registration_widgets = {}
        for name, type in config['api']['registration_requirements'].items():
            # Создаем виджет
            if type == "QLineEdit": 
                widget = QLineEdit()
                widget.setPlaceholderText(name)
            # Добавляем виджет в словарь
            self.registration_widgets[name] = widget
            # Добавляем виджет в layout
            self.layout.addWidget(widget)
            
        logger.debug(f"registration_widgets=\n{self.registration_widgets}")
        logger.debug(f"login_widgets=\n{self.login_widgets}")
            
        # Создаем кнопку "Зарегистрироваться"
        self.reg_button = QPushButton('Зарегистрироваться')
        self.reg_button.clicked.connect(self.registration)
        self.layout.addWidget(self.reg_button)
