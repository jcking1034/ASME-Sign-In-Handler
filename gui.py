from PyQt5 import QtCore, QtGui, QtWidgets
import numpy as np
import sys
import datetime as dt

import main

"""
To Implement:
    "add": add_spreadsheet,
    "output": create_excel,
    "delete": delete_event,
"""

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        
        self.setMinimumSize(QtCore.QSize(1000, 1000))
        self.setWindowTitle("ASME Sign In Handler")

        self.name_label = QtWidgets.QLabel(self)
        self.name_label.setText("ASME Sign In Handler")
        self.name_label.setGeometry(QtCore.QRect(0, 0, 200, 25))

        self.click_button = QtWidgets.QPushButton('Click me', self)
        self.click_button.resize(200, 25)
        self.click_button.move(0, 25)       # x, y
        self.click_button.clicked.connect(self.click_method)

        self.help_button = QtWidgets.QPushButton('Help', self)
        self.help_button.resize(200, 25)
        self.help_button.move(0, 50)
        self.help_button.clicked.connect(lambda: self.new_message(main.display_help()))

        # Need to reimplement add
        self.add_file_button = QtWidgets.QPushButton('Add Spreadsheet', self)
        self.add_file_button.resize(200, 25)
        self.add_file_button.move(0, 75)
        self.add_file_button.clicked.connect(lambda: self.new_message(QtWidgets.QFileDialog.getOpenFileName(self,"TEST", "/")[0]))

        self.delete_event_button = QtWidgets.QPushButton('Delete Event', self)
        self.delete_event_button.resize(200, 25)
        self.delete_event_button.move(0, 100)
        self.delete_event_button.clicked.connect(lambda: self.new_message("Not Implemented Yet"))

        self.output_file_button = QtWidgets.QPushButton('Output', self)
        self.output_file_button.resize(200, 25)
        self.output_file_button.move(0, 125)
        self.output_file_button.clicked.connect(lambda: self.new_message("Not Implemented Yet"))

        self.list_events_button = QtWidgets.QPushButton('List Events', self)
        self.list_events_button.resize(200, 25)
        self.list_events_button.move(0, 150)
        self.list_events_button.clicked.connect(lambda: self.new_message(main.list_events()))

        self.reset_button = QtWidgets.QPushButton('Reset', self)
        self.reset_button.resize(200, 25)
        self.reset_button.move(0, 250)
        self.reset_button.clicked.connect(lambda: self.reset_warning())

        self.output_box = QtWidgets.QTextEdit(self)
        self.output_box.setReadOnly(True)
        self.output_box.setGeometry(QtCore.QRect(0, 300, 500, 250))
        self.output_box.setObjectName("Output Box")

        self.show()


    def reset_warning(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)

        msg.setText("Do you really want to reset?")
        msg.setInformativeText("This deletes all records")
        msg.setWindowTitle("Reset Warning")

        msg.setStandardButtons(QtWidgets.QMessageBox.Ok | QtWidgets.QMessageBox.Cancel)
        val = msg.exec_()
        if val == QtWidgets.QMessageBox.Ok:
            self.new_message(main.reset())
        else:
            self.new_message("No reset performed")


    def click_method(self):
        self.new_message("click_method called")


    def new_message(self, msg):
        msg = str(msg)
        print(msg)
        self.output_box.append(f"{str(dt.datetime.now())}\n{msg}\n")
        self.output_box.repaint()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mainWin = MainWindow()
    sys.exit(app.exec_())

