""" third party imports """
import sys
from PyQt5.QtWidgets import QApplication

""" internal imports """
from app.main_window import App
from app.tools.dialogs.auth import LoginDialog

if __name__ == '__main__':
    app = QApplication(sys.argv)

    login_dialog = LoginDialog()
    if login_dialog.exec_() == LoginDialog.Accepted:
        username = login_dialog.username_input.text()
        ex = App(login=username)
        sys.exit(app.exec_())
