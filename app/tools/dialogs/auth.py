from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QMessageBox, QDialog

class LoginDialog(QDialog):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle('Вход в аккаунт')
        
        layout = QVBoxLayout()
        
        self.username_input = QLineEdit()
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        
        login_button = QPushButton('Login')
        login_button.clicked.connect(self.login)
        
        layout.addWidget(QLabel('Username:'))
        layout.addWidget(self.username_input)
        layout.addWidget(QLabel('Password:'))
        layout.addWidget(self.password_input)
        layout.addWidget(login_button)
        
        self.setLayout(layout)
        
    def login(self):
        username = self.username_input.text()
        password = self.password_input.text()
        
        # Проверка логина и пароля
        if username == '1' and password == '1':
            self.accept()
        else:
            QMessageBox.warning(self, 'Login Error', 'Invalid username or password.')