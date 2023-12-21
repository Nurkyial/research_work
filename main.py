# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QComboBox
import pandas as pd
import math
import numpy as np
import docx
import copy


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.upload_label = QtWidgets.QLabel(self.centralwidget)
        self.upload_label.setGeometry(QtCore.QRect(10, 20, 151, 51))
        font = QtGui.QFont()
        font.setPointSize(20)
        self.upload_label.setFont(font)
        self.upload_label.setObjectName("upload_label")
        self.text_label = QtWidgets.QLabel(self.centralwidget)
        self.text_label.setGeometry(QtCore.QRect(10, 80, 241, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.text_label.setFont(font)
        self.text_label.setObjectName("text_label")

        self.file1 = QtWidgets.QLabel(self.centralwidget)
        self.file1.setGeometry(QtCore.QRect(10, 140, 55, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.file1.setFont(font)
        self.file1.setObjectName("file1")

        self.file2 = QtWidgets.QLabel(self.centralwidget)
        self.file2.setGeometry(QtCore.QRect(10, 200, 55, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.file2.setFont(font)
        self.file2.setObjectName("file2")

        # uploading file1
        self.browse1 = QtWidgets.QPushButton(self.centralwidget)
        self.browse1.setGeometry(QtCore.QRect(360, 140, 93, 28))
        self.browse1.setObjectName("browse1")
        self.browse1.clicked.connect(self.browse1_open)
        self.browse1.setToolTip("Add student's marks. It must be a docx file")

        # uploading file2
        self.browse2 = QtWidgets.QPushButton(self.centralwidget)
        self.browse2.setGeometry(QtCore.QRect(360, 200, 93, 28))
        self.browse2.setObjectName("browse2")
        self.browse2.clicked.connect(self.browse2_open)
        self.browse2.setToolTip("Add a study plan. It must be a xlsx file.")

        # match button for two files
        self.match = QtWidgets.QPushButton(self.centralwidget)
        self.match.setGeometry(QtCore.QRect(160, 400, 93, 28))
        self.match.setObjectName("match")
        self.match.setStyleSheet('QPushButton {background-color: #00bfff}')
        self.match.clicked.connect(self.match_button)
        self.match.setToolTip("Eliminate the academic difference.")

        # Initialize the file_path attribute
        self.file_path = None

        # test labels for browses
        self.test_label1 = QtWidgets.QLabel(self.centralwidget)
        self.test_label1.setGeometry(QtCore.QRect(60, 140, 55, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.test_label1.setFont(font)
        self.test_label1.setObjectName("test_label")

        self.test_label2 = QtWidgets.QLabel(self.centralwidget)
        self.test_label2.setGeometry(QtCore.QRect(60, 200, 55, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.test_label2.setFont(font)
        self.test_label2.setObjectName("test_label")


        self.sem = QtWidgets.QLabel(self.centralwidget)
        self.sem.setGeometry(QtCore.QRect(10, 290, 90, 30))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.sem.setFont(font)
        self.sem.setObjectName("semester")
        #self.sem.adjustSize()

        #  creating combobox for list of semesters
        self.comboBox = QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(150, 290, 200, 50))
        self.comboBox.addItem("1")
        self.comboBox.addItem("2")
        self.comboBox.addItem("3")
        self.comboBox.addItem("4")
        self.comboBox.addItem("5")
        self.comboBox.addItem("6")
        self.comboBox.addItem("7")
        self.comboBox.addItem("8")

        MainWindow.setCentralWidget(self.centralwidget)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)


    def resize(self, text):
        max_word = 30
        return text[:max_word] + '...'


    def browse1_open(self):
        file, name = QtWidgets.QFileDialog.getOpenFileName(self.centralwidget, "Open File", "", "All Files (*);; pdf files (*.pdf)")
        path = str(file)

        if name:
            self.test_label1.setText(self.resize(file))
            self.test_label1.adjustSize()
            # Update the file path attribute
            self.file_path1 = path


    def browse2_open(self):
        file, name = QtWidgets.QFileDialog.getOpenFileName(self.centralwidget, "Open File", "", "All Files (*);; pdf files (*.pdf)")
        path = str(file)
        if name:
            self.test_label2.setText(self.resize(file))
            self.test_label2.adjustSize()
            self.file_path2 = path



    def match_button(self):
        # get the file path from the browse functions
        self.matching = Match()
        if (self.file_path1 and self.file_path2) is not None:
            if self.file_path1.endswith('.docx') and self.file_path2.endswith('.xlsx'):
                self.docx = DocxReader(self, self.file_path1)
                self.docx_data = self.docx.read_docx()
                self.xls_data = XLSReader(self, self.file_path2)

                #print(self.matching)

                complete_match, partially_match, mismatch_xls, mismatch_docx = self.matching.match(self.docx_data, self.xls_data.data_dict)
                print(80*'-')
                print("Complete Match:", complete_match)
                print(80 * '-')
                print("Partially Match:", partially_match)
                print(80 * '-')
                print("Mismatch for excel file:", mismatch_xls)
                print(80 * '-')
                print("Mismatch for docx file:", mismatch_docx)

            elif self.file_path2.endswith('.docx') and self.file_path1.endswith('.xlsx'):
                self.docx_data = DocxReader(self, self.file_path2)
                self.xls_data = XLSReader(self, self.file_path1)


                complete_match, partially_match, mismatch_xls, mismatch_docx = self.matching.match(self.docx_data, self.xls_data.data_dict)
                print(80 * '-')
                print("Complete Match:", complete_match)
                print(80 * '-')
                print("Partially Match:", partially_match)
                print(80 * '-')
                print("Mismatch for excel file:", mismatch_xls)
                print(80 * '-')
                print("Mismatch for docx file:", mismatch_docx)

            else:
                print("Invalid file types. Please choose a DOCX and an Excel file.")

        else:
            print("No file selected. Please choose a file")

    # finding the current item in the combobox
    def find(self):
        self.content = self.comboBox.currentText()
        return int(self.content)


    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.upload_label.setText(_translate("MainWindow", "Upload"))
        self.text_label.setText(_translate("MainWindow", "Upload two files (.pdf or .xls)"))
        self.test_label1.setText(_translate("MainWindow", ""))
        self.test_label2.setText(_translate("MainWindow", ""))
        self.file1.setText(_translate("MainWindow", "File1"))
        self.file2.setText(_translate("MainWindow", "File2"))
        self.browse1.setText(_translate("MainWindow", "Browse"))
        self.browse2.setText(_translate("MainWindow", "Browse"))
        self.match.setToolTip(_translate("MainWindow", "<html><head/><body><p>Match two files</p></body></html>"))
        self.match.setText(_translate("MainWindow", "Match"))
        self.sem.setText(_translate("MainWindow", "semester"))




class XLSReader:
    def __init__(self, ui_instance, file):
        self.ui_instance = ui_instance
        self.xls = pd.ExcelFile(file)
        self.course = 6 + self.ui_instance.find() / 2
        #self.column = [2, 11, 12, 21, 22]
        self.start = 5
        self.end = 30  # under question???
        self.save_to_dict(self.course, row_start=self.start, row_end=self.end)

    def save_to_dict(self, course, row_start=None, row_end=None):
        self.column = [2, 11, 12, 21, 22]
        self.data_dict = {}
        cnt = 0
        if course == 6.5:
            self.column_indices = self.column[0:3]
            self.tcourse = int(course) + 1
            print(self.tcourse, course)
        elif course - int(course) == 0.5:
            self.column_indices = self.column
            self.tcourse = int(course) + 1
            flag = False
        else:
            self.tcourse = int(course)
            self.column_indices = self.column


        # Iterate through each sheet
        for sheet_name in self.xls.sheet_names[6:self.tcourse]:
            cnt += 1
            if cnt == self.tcourse - 6 and course - int(course) == 0.5:
                self.column_indices = self.column[0:3]
            # Read the sheet into a DataFrame
            self.df = pd.read_excel(self.xls, sheet_name)

            # If column_indices is not provided, use all the columns
            if self.column_indices is None:
                column_indices = range(len(self.df.columns))

            # Validate column indices
            if any(idx >= len(self.df.columns) for idx in self.column_indices):
                raise ValueError("Invalid column index provided.")

            # Iterate over rows within the specified range and save it to the dictionary
            for index, row in self.df.iterrows():
                if row_start is not None and index < row_start:
                    continue  # skip rows before the specified start index

                if row_end is not None and index > row_end:
                    break

                for i in range(1, len(self.column_indices), 2):
                    key = row.iloc[self.column_indices[0]]
                    values = list(row.iloc[i] for i in self.column_indices[i:i + 2])  # Use the rest as values


                    if pd.isna(key):
                        continue
                    #print(f'Index: {index}, key: {key}, values: {values}')
                    if key not in self.data_dict:
                        if values[0] != 0:
                            self.data_dict[key] = values
                        else:
                            continue
                    else:
                        if values[0] != 0:
                            #print(f'Index: {index}, key: {key}, values: {values}')
                            self.data_dict[key][0] += values[0]
                            self.data_dict[key][1] = values[1]
                        else:
                            continue

        return self.data_dict


class DocxReader:
    def __init__(self, ui_instance, docx_path):
        self.ui_instance = ui_instance
        self.docx_path = docx_path
        self.all_subjects = []
        self.semester = self.ui_instance.find()
        #self.read_docx(semester)

    def read_docx(self):
        cnt_s = 0

        # Open the Word document
        doc = docx.Document(self.docx_path)

        for table in doc.tables:
            sem_subjects = {}
            # Extract data from cells in the first row
            first_row = table.rows[0].cells[0]
            first_row_data = first_row.text
            sem = str(cnt_s + 1) + " " + "семестр"

            #print(first_row_data)

            # Checking if the semester we need is in the table
            if sem in first_row_data:
                cnt_s += 1
                # Find the column indices
                header_row = table.rows[1]
                date_column_index = None
                marks_column_index = None

                for i, cell in enumerate(header_row.cells):
                    if cell.text.strip().lower() == 'дата сдачи':
                        date_column_index = i
                    elif cell.text.strip().lower() == 'оценка (ects)':
                        marks_column_index = i

                for row in table.rows[2:]:
                    discipline = row.cells[0].text.strip()
                    if "диф. зачёт" in discipline:
                        control = "диф. З"
                        discipline1 = discipline.replace("(диф. зачёт)", '')
                    elif "зачёт" in discipline:
                        control = "З"
                        discipline1 = discipline.replace("(зачёт)", '')
                    elif "экзамен" in discipline:
                        control = "Э"
                        discipline1 = discipline.replace("(экзамен)", '')
                    elif "аттестация всех разделов" in discipline:
                        control = "аттестовано"
                        discipline1 = discipline.replace("(аттестация всех разделов)", '')
                    else:
                        control = ''
                        discipline1 = discipline

                    # Getting the date and marks using identified column indices
                    date = row.cells[date_column_index].text.strip() if date_column_index is not None else ''
                    marks = row.cells[marks_column_index].text.strip() if marks_column_index is not None else ''

                    # Create a dictionary
                    sem_subjects[discipline1.strip()] = (control, date, marks)

            if sem_subjects:
                self.all_subjects.append(sem_subjects)
            if cnt_s == self.semester:
                break

        return self.all_subjects



class Match:
    def __init__(self):
        self.complete_match = []
        self.partially_match = []
        self.mismatch = []
        self.complete_match_dct = {}
        self.partially_match_dct = {}
        self.mismatch_docx = {}
        self.mismatch_xlsx = {}
        # xls_data is a dictionary, values are lists
        # docx_data is a list which consists several dictionaries


    def match(self, docx_data, xls_data):
        print("data from excel file: ", xls_data)
        print("data from docx file: ", docx_data)
        self.xls_data_mismatch = copy.deepcopy(xls_data)
        self.docx_data_mismatch = copy.deepcopy(docx_data)

        # first I get a first element of docx_data and compare it with xlsx_data
        for docx_semester, docx_subjects in enumerate(docx_data):
            #print(docx_semester, docx_subjects)
            # checking every subject in docx_data if it is in xls_data
            for docx_subject, docx_info in docx_subjects.items():
                #print(docx_subject)
                if docx_subject in xls_data:
                    xls_subject = docx_subject
                    self.xls_data_mismatch.pop(xls_subject, None)
                    self.docx_data_mismatch[docx_semester].pop(docx_subject, None)
                else:
                    xls_subject = None
                #print(docx_subject, xls_subject)




                if xls_subject:

                    # information for dictionary, if subjects are matching
                    hours = xls_data[xls_subject][0]
                    kr_type_xls = xls_data[xls_subject][1]
                    kr_type_docx = docx_data[docx_semester][docx_subject][0]
                    # you should add information about semester from excel file
                    mark = docx_data[docx_semester][docx_subject][1]
                    date = docx_data[docx_semester][docx_subject][2]


                    if docx_subject == xls_subject and docx_info[0] == xls_data[xls_subject][1]:

                        # checking if the subject is already in the list or not
                        if (docx_subject, xls_subject) in self.complete_match:
                            # find the index of the subject
                            index = self.complete_match.index((docx_subject, xls_subject))

                            # replace the subject
                            self.complete_match[index] = (docx_subject, xls_subject)

                            # appending elements to the dictionary (maybe in future i will change this part of code)
                            self.complete_match_dct[docx_subject] = (hours, kr_type_xls, mark, date)

                        else:
                            self.complete_match.append((docx_subject, xls_subject))

                            # appending elements to the dictionary
                            self.complete_match_dct[docx_subject] = (hours, kr_type_xls, mark, date)
                    else:
                        if (docx_subject, xls_subject) in self.partially_match:
                            # find the index of the subject
                            index = self.partially_match.index((docx_subject, xls_subject))

                            # replace the subject
                            self.partially_match[index] = (docx_subject, xls_subject)

                            # appending to the final dictionary all the information about subjects
                            self.partially_match_dct[docx_subject] = (hours, kr_type_xls, kr_type_docx, mark, date)
                        else:
                            self.partially_match.append((docx_subject, xls_subject))

                            # appending to the final dictionary all the information about subjects
                            self.partially_match_dct[docx_subject] = (hours, kr_type_xls, kr_type_docx, mark, date)
                else:
                    self.mismatch.append(docx_subject)
        #return self.complete_match, self.partially_match, self.mismatch
        return self.complete_match_dct, self.partially_match_dct, self.xls_data_mismatch, self.docx_data_mismatch



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle(QtWidgets.QStyleFactory.create('Fusion'))
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)

    MainWindow.show()
    sys.exit(app.exec_())
