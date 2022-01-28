# -*- coding: cp1251 -*-
import threading
import subprocess
import PyQt5
import sys
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import QApplication, QDialog, QMainWindow, QPushButton
from PyQt5.QtCore import QDir
from PyQt5.QtWidgets import QFileDialog
from BOM_interface import Ui_mainWindow
from excel import ExcelFile
import os
from PyQt5.QtWidgets import QTreeWidgetItem


class Main(QtWidgets.QMainWindow, Ui_mainWindow):
    def __init__(self):
        super(Main, self).__init__()
        self.setupUi(self)
        self.textBrowser.setHeaderLabels(['', ''])
        self.textBrowser.header().resizeSection(0, int(self.width()/10*8.5))
        self.textBrowser.header().resizeSection(1, int(self.width()/10))
        self.add_func()

    def add_func(self):
        self.openFile.clicked.connect(lambda: self.open_file_thread())
        self.launch_but.clicked.connect(lambda: self.check())

    def open_file(self):
        dlg = QFileDialog(filter="*.xlsx *.xls")
        dlg.setFileMode(QFileDialog.AnyFile)

        if dlg.exec_():
            self.textBrowser.clear()
            self.filenames = dlg.selectedFiles()
            os.startfile(self.filenames[0])

    def open_file_thread(self):
        t = threading.Thread(target=self.open_file())
        t.start()

    def check(self):
        try:
            self.textBrowser.clear()
            self.excel = ExcelFile(self.filenames[0], self.textBrowser)
        except AttributeError:
            QTreeWidgetItem(self.textBrowser, ["Файл не выбран!"])

    def closeEvent(self, a0: QtGui.QCloseEvent) -> None:
        super(Main, self).closeEvent(a0)
        # os.system('taskkill /f /im excel.exe')


app = QtWidgets.QApplication([])
win = Main()

win.show()
sys.exit(app.exec())
