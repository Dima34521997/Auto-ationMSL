import locale
import sys
import time

from PySide6 import QtWidgets
from Window import Window

locale.setlocale(locale.LC_ALL, ('ru_RU', 'UTF-8'))

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = Window()
    window.show()
    sys.exit(app.exec())
