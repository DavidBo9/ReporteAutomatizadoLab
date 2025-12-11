# main.py
import sys
from PyQt5.QtWidgets import QApplication
from gui_app import ReportGeneratorApp

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ReportGeneratorApp()
    ex.show()
    sys.exit(app.exec_())