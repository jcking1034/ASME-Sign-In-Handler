#!/usr/bin/env python3
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
import datetime as dt

import utils

"""
To Implement:
    "add": add_spreadsheet,
    "output": create_excel,
    "delete": delete_event,
"""

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        
        self.setMinimumSize(QtCore.QSize(500, 500))
        self.setWindowTitle("ASME Sign In Handler")

        self.name_label = QtWidgets.QLabel(self)
        self.name_label.setText("ASME Sign In Handler")

        self.help_button = QtWidgets.QPushButton('Help', self)
        self.help_button.clicked.connect(lambda: self.new_message(utils.display_help()))

        self.add_file_button = QtWidgets.QPushButton('Add Spreadsheet', self)
        self.add_file_button.clicked.connect(lambda: self.new_message(self.validate(utils.add_sheets, *QtWidgets.QFileDialog.getOpenFileNames(self,"Select File(s)"))))

        self.delete_event_button = QtWidgets.QPushButton('Delete Event', self)
        self.delete_event_button.clicked.connect(lambda: self.new_message(self.validate(utils.delete_event_with_fname, *QtWidgets.QInputDialog.getText(self, 'File to remove?', 'Enter file name:'))))

        self.output_file_button = QtWidgets.QPushButton('Output', self)
        self.output_file_button.clicked.connect(lambda: self.new_message(self.validate(utils.create_excel_with_fname, *QtWidgets.QInputDialog.getText(self, 'Name of Output File?', 'Enter file name:'))))

        self.list_events_button = QtWidgets.QPushButton('List Events', self)
        self.list_events_button.clicked.connect(lambda: self.new_message(utils.list_events()))

        self.reset_button = QtWidgets.QPushButton('Reset', self)
        self.reset_button.clicked.connect(lambda: self.reset_warning())

        self.output_box = QtWidgets.QTextEdit(self)
        self.output_box.setReadOnly(True)
        self.output_box.setObjectName("Output Box")

        widget = QtWidgets.QWidget(self)
        layout = QtWidgets.QVBoxLayout(widget)
        layout.addWidget(self.name_label)
        layout.addWidget(self.help_button)
        layout.addWidget(self.add_file_button)
        layout.addWidget(self.delete_event_button)
        layout.addWidget(self.output_file_button)
        layout.addWidget(self.list_events_button)
        layout.addWidget(self.reset_button)
        layout.addWidget(self.output_box)
        self.setCentralWidget(widget)


    def new_message(self, msg):
        msg = str(msg)
        print(msg)
        self.output_box.append(f"{str(dt.datetime.now())}\n{msg}\n")
        self.output_box.repaint()


    def validate(self, func, text, ok):
        return func(text) if text and ok else "Cancelled"


    def reset_warning(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)

        msg.setText("Do you really want to reset?")
        msg.setInformativeText("This deletes all records")
        msg.setWindowTitle("Reset Warning")

        msg.setStandardButtons(QtWidgets.QMessageBox.Ok | QtWidgets.QMessageBox.Cancel)
        val = msg.exec_()
        if val == QtWidgets.QMessageBox.Ok:
            self.new_message(utils.reset())
        else:
            self.new_message("No reset performed")


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.setGeometry(0, 0, 750, 750)
    mainWin.show()
    sys.exit(app.exec_())
