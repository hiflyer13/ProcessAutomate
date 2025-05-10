import sys
from PyQt5.QtWidgets import QApplication
from main_menu import MainMenu
import xlrd

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_menu = MainMenu()
    main_menu.show()
    sys.exit(app.exec_())