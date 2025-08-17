# -*- coding: utf-8 -*-
import sys, os
from PyQt6.QtWidgets import QApplication
from finsuite.main_window import MainWindow

def main():
    app = QApplication(sys.argv)
    os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")
    win = MainWindow(); win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
