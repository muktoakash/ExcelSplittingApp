"""./main.py

This is an app that splits up an Excel file into many based on the current sheets in the spreadsheet."""

# Imports
import os, sys, re
import win32com.client as win32
from openpyxl import load_workbook
from pywintypes import com_error
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (QApplication, QWidget, QPushButton, QListWidget, QLineEdit, QStatusBar, QListWidgetItem, QFileDialog, QFormLayout, QHBoxLayout, QVBoxLayout)

def get_worksheet_names(file_path : str):
    """
    Loads all the sheet names from a given file path.
    Return type: list?
    """
    workbook = load_workbook(file_path)
    return workbook.sheetnames

class AppWindow(QWidget):
    """"""
    def __init__(self, parent: QWidget | None = ..., flags: Qt.WindowType = ...) -> None:
        super().__init__(parent, flags)
        self.setWindowTitle('Excel File Splitter')
        self.window_width, self.window_height = 700, 50
        self.setMinimumSize(self.window_width, self.window_height)
        self.setWindowIcon(QIcon('icon.png'))
        self.setStyleSheet('''
                           QWidget {
                           font-size: 14px;
                           }
                           '''
                           )
        self.setObjectName('AppWindow')
        self.setLayout(self.layout['main'])

        self.initUI()
        self.config_signals()

        def initUI(self):
            self.layout['form'] = QFormLayout()
            self.layout['main'].addLayout(self.layout['form'])

            self.layout['browse_file'] = QHBoxLayout()

            self.file_path = QLineEdit()
            self.layout['browse_file'].addWidget(self.file_path)


# Run
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    app.setStyleSheet(open('style.css').read())

    myApp = AppWindow()
    myApp.show()

    try:
        sys.exit(app.exec())
    except SystemExit:
        print("Closing Window ... ")
