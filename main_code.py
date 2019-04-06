# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'read_excel_main_code.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import xlrd
import xlwt
import xlsxwriter
from xlutils.copy import copy


class Ui_MainWindow(object):
    # GUI design
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1104, 638)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(160, 200, 113, 22))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(320, 200, 113, 22))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_3.setGeometry(QtCore.QRect(160, 260, 113, 22))
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.lineEdit_4 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_4.setGeometry(QtCore.QRect(320, 260, 113, 22))
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(70, 200, 61, 20))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(70, 260, 51, 16))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(290, 200, 21, 16))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setScaledContents(False)
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(290, 260, 21, 16))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.lineEdit_5 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_5.setGeometry(QtCore.QRect(290, 90, 241, 21))
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.lineEdit_6 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_6.setGeometry(QtCore.QRect(770, 90, 241, 22))
        self.lineEdit_6.setObjectName("lineEdit_6")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(50, 90, 231, 20))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(580, 90, 191, 16))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(710, 420, 93, 28))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.listWidget = QtWidgets.QListWidget(self.centralwidget)
        self.listWidget.setGeometry(QtCore.QRect(50, 340, 591, 192))
        self.listWidget.setObjectName("listWidget")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(890, 420, 93, 28))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.lineEdit_7 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_7.setGeometry(QtCore.QRect(700, 200, 113, 22))
        self.lineEdit_7.setObjectName("lineEdit_7")
        self.lineEdit_8 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_8.setGeometry(QtCore.QRect(860, 200, 113, 22))
        self.lineEdit_8.setObjectName("lineEdit_8")
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setGeometry(QtCore.QRect(610, 260, 51, 16))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.lineEdit_9 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_9.setGeometry(QtCore.QRect(700, 260, 113, 22))
        self.lineEdit_9.setObjectName("lineEdit_9")
        self.lineEdit_10 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_10.setGeometry(QtCore.QRect(860, 260, 113, 22))
        self.lineEdit_10.setObjectName("lineEdit_10")
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setGeometry(QtCore.QRect(830, 260, 21, 16))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_8.setFont(font)
        self.label_8.setObjectName("label_8")
        self.label_9 = QtWidgets.QLabel(self.centralwidget)
        self.label_9.setGeometry(QtCore.QRect(830, 200, 21, 16))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_9.setFont(font)
        self.label_9.setScaledContents(False)
        self.label_9.setObjectName("label_9")
        self.label_10 = QtWidgets.QLabel(self.centralwidget)
        self.label_10.setGeometry(QtCore.QRect(610, 200, 61, 20))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_10.setFont(font)
        self.label_10.setObjectName("label_10")
        self.lineEdit_11 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_11.setGeometry(QtCore.QRect(270, 150, 161, 22))
        self.lineEdit_11.setObjectName("lineEdit_11")
        self.lineEdit_12 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_12.setGeometry(QtCore.QRect(810, 150, 161, 22))
        self.lineEdit_12.setObjectName("lineEdit_12")
        self.label_11 = QtWidgets.QLabel(self.centralwidget)
        self.label_11.setGeometry(QtCore.QRect(70, 150, 201, 20))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_11.setFont(font)
        self.label_11.setObjectName("label_11")
        self.label_12 = QtWidgets.QLabel(self.centralwidget)
        self.label_12.setGeometry(QtCore.QRect(610, 150, 171, 20))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_12.setFont(font)
        self.label_12.setObjectName("label_12")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1104, 25))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        # As button pushed, call these def
        self.pushButton.clicked.connect(self.read_and_create)
        self.pushButton_2.clicked.connect(self.read_and_write)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "Row"))
        self.label_2.setText(_translate("MainWindow", "Col"))
        self.label_3.setText(_translate("MainWindow", "~"))
        self.label_4.setText(_translate("MainWindow", "~"))
        self.label_5.setText(_translate("MainWindow", "Data Source Address"))
        self.label_6.setText(_translate("MainWindow", "Final File Address"))
        self.pushButton.setText(_translate("MainWindow", "Create"))
        self.pushButton_2.setText(_translate("MainWindow", "Write"))
        self.label_7.setText(_translate("MainWindow", "Col"))
        self.label_8.setText(_translate("MainWindow", "~"))
        self.label_9.setText(_translate("MainWindow", "~"))
        self.label_10.setText(_translate("MainWindow", "Row"))
        self.label_11.setText(_translate("MainWindow", "Data Source Sheet"))
        self.label_12.setText(_translate("MainWindow", "Final File Sheet"))

    # it will read out data and put on new excel file
    def read_and_create(self):
        # be sure that application will not freeze
        try:
            # open read file(excel) from form, lineEdit_5
            data_source_book = xlrd.open_workbook(self.lineEdit_5.text())
            # get all sheets of the read file in list(data_source_sheet)
            data_source_sheet = data_source_book.sheets()
            # configure which sheet in excel file the user selected
            sheet_index = 0
            for sheet_index_buffer in range(len(data_source_sheet)):
                if data_source_sheet[sheet_index_buffer] == self.lineEdit_11.text():
                    sheet_index = sheet_index_buffer

            # create a new excel file in specified address
            final_file_book = xlsxwriter.Workbook(self.lineEdit_6.text())
            # create a sheet in new excel file
            final_file_sheet = final_file_book.add_worksheet(self.lineEdit_12.text())

            # global list to get top row and left column value
            top_row_value = []
            left_col_value = []

            # make sure how many rows read and run
            # index_buffer is buffer index for syntax
            for index_buffer in range(int(self.lineEdit_2.text()) - int(self.lineEdit.text()) + 1):
                # index is real location(row) we read
                index = index_buffer + int(self.lineEdit.text())
                # read
                row_value = data_source_sheet[sheet_index].row_values(index, int(self.lineEdit_3.text()),
                                                                      int(self.lineEdit_4.text()) + 1)
                # read top row value
                if self.lineEdit.text() != 0:
                    top_row_value = data_source_sheet[sheet_index].row_values(0, int(self.lineEdit_3.text()),
                                                                              int(self.lineEdit_4.text()) + 1)
                # read lef column value
                if self.lineEdit_3.text() != 0:
                    left_col_value = data_source_sheet[sheet_index].row_values(index, 0)

                # write row values in order
                for sub_index_buffer in range(int(self.lineEdit_4.text()) - int(self.lineEdit_3.text()) + 1):
                    # configure which user choose column range including 0
                    if int(self.lineEdit_3.text() != 0):
                        # write first row or not
                        if (int(self.lineEdit.text()) != 0) and (index_buffer == 0):
                            final_file_sheet.write(index_buffer, sub_index_buffer + 1, top_row_value[sub_index_buffer])
                        else:
                            # write left cell of each row
                            if sub_index_buffer == 0:
                                final_file_sheet.write(index_buffer, 0, left_col_value[0])
                            final_file_sheet.write(index_buffer, sub_index_buffer + 1, row_value[sub_index_buffer])
                    else:
                        # write first row or not
                        if (int(self.lineEdit.text()) != 0) and (index_buffer == 0):
                            final_file_sheet.write(index_buffer, sub_index_buffer, top_row_value[sub_index_buffer])
                        else:
                            final_file_sheet.write(index_buffer, sub_index_buffer, row_value[sub_index_buffer])

            final_file_book.close()

            # display condition
            self.listWidget.clear()
            self.listWidget.addItem("Create Successful")
        except:
            # display condition
            self.listWidget.clear()
            self.listWidget.addItem("Have some errors(create), go to ask Acer.")

    # modify value in excel file or transfer data between two file
    def read_and_write(self):
        # be sure that application will not freeze
        try:
            # open read file(excel) from form, lineEdit_5
            data_source_book = xlrd.open_workbook(self.lineEdit_5.text())
            # get all sheets of the read file in list(data_source_sheet)
            data_source_sheet = data_source_book.sheets()
            # configure which sheet in excel file the user selected
            sheet_index = 0
            for sheet_index_buffer in range(len(data_source_sheet)):
                if data_source_sheet[sheet_index_buffer] == self.lineEdit_11.text():
                    sheet_index = sheet_index_buffer

            # open any existing excel file, including original file(ready to modified)
            read_final_file_book = xlrd.open_workbook(self.lineEdit_6.text())
            # backup first
            final_file_book = copy(read_final_file_book)
            # configure if the file has the sheet
            try:
                # add new sheet because there is no the sheet
                final_file_sheet = final_file_book.add_sheet(self.lineEdit_12.text())
            except:
                # there is the sheet, and get the object of sheet
                final_file_sheet = final_file_book.get_sheet(self.lineEdit_12.text())

            # make sure how many rows read and run
            # index_buffer is buffer index for syntax
            for index_buffer in range(int(self.lineEdit_2.text()) - int(self.lineEdit.text()) + 1):
                index = index_buffer + int(self.lineEdit.text())
                row_value = data_source_sheet[sheet_index].row_values(index, int(self.lineEdit_3.text()),
                                                                      int(self.lineEdit_4.text()) + 1)
                # write
                for sub_index_buffer in range(int(self.lineEdit_10.text()) - int(self.lineEdit_9.text()) + 1):
                    written_index = index_buffer + int(self.lineEdit_7.text())
                    sub_index = sub_index_buffer + int(self.lineEdit_9.text())
                    print(row_value)
                    final_file_sheet.write(written_index, sub_index, row_value[sub_index_buffer])

            final_file_book.save(self.lineEdit_6.text())

            # display condition
            self.listWidget.clear()
            self.listWidget.addItem("Write Successful")
        except:
            # display condition
            self.listWidget.clear()
            self.listWidget.addItem("Have some errors(write), go to ask Acer.")


# open application
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
