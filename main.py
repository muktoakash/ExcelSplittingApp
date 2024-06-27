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
