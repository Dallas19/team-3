import sys
import os
import copy
import xlrd
from collections import defaultdict
from PyQt5.QtGui import (QIcon, QStandardItemModel, QPixmap, QFont)
from PyQt5.QtCore import (QDate, QDateTime, QRegExp, QSortFilterProxyModel,
                          Qt, QTime, QMetaObject, QSize, QRect, QCoreApplication, QObject, pyqtSignal)
from PyQt5.QtWidgets import (QApplication, QCheckBox, QComboBox, QGridLayout,
                             QGroupBox, QHBoxLayout, QLabel, QLineEdit, QTreeView, QVBoxLayout, QTabWidget,
                             QWidget, QPushButton, QFrame, QTextEdit, QToolButton, QMenuBar,
                             QMainWindow, QStatusBar, QMenu, QFileDialog, QAction, QTextBrowser, QAbstractItemView, QHeaderView)


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'File Explorer'
        self.left = 10
        self.top = 10
        self.width = 640
        self.height = 480
        self.Msg = ""
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.saveFileDialog()

    def saveFileDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(
            self, "QFileDialog.getOpenFileName)", "", "Excel Workbook (*.xlsx)", options=options)
        self.Msg = fileName


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):

        self.studentrank_file = ""
        self.companyrank_file = ""

        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1009, 863)
        icon = QIcon()
        icon.addPixmap(
            QPixmap("Working Directory/flame_eBQ_icon.ico"), QIcon.Normal, QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setStyleSheet("background: rgb(245, 235, 227)")
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.frame_2 = QFrame(self.centralwidget)
        self.frame_2.setMinimumSize(QSize(0, 140))
        self.frame_2.setMaximumSize(QSize(16777215, 140))
        self.frame_2.setFrameShape(QFrame.NoFrame)
        self.frame_2.setFrameShadow(QFrame.Plain)
        self.frame_2.setObjectName("frame_2")
        self.groupBox = QGroupBox(self.frame_2)
        self.groupBox.setGeometry(QRect(20, 0, 280, 130))
        self.groupBox.setMinimumSize(QSize(40, 130))
        self.groupBox.setMaximumSize(QSize(280, 16777215))
        self.groupBox.setObjectName("groupBox")

        ### PB1 ###
        self.pushButton = QPushButton(self.groupBox)
        self.pushButton.setGeometry(QRect(10, 40, 93, 28))
        self.pushButton.setObjectName("pushButton")
        ### PB2 ###
        self.pushButton_2 = QPushButton(self.groupBox)
        self.pushButton_2.setGeometry(QRect(10, 90, 93, 28))
        self.pushButton_2.setObjectName("pushButton_2")

        ### Line Edit 1 ###
        self.lineEdit = QLineEdit(self.groupBox)
        self.lineEdit.setGeometry(QRect(110, 40, 141, 31))
        self.lineEdit.setReadOnly(True)
        self.lineEdit.setClearButtonEnabled(False)
        self.lineEdit.setObjectName("lineEdit")

        ### Line Edit 2 ###
        self.lineEdit_2 = QLineEdit(self.groupBox)
        self.lineEdit_2.setGeometry(QRect(110, 90, 141, 31))
        self.lineEdit_2.setText("")
        self.lineEdit_2.setReadOnly(True)
        self.lineEdit_2.setObjectName("lineEdit_2")

        ### Button Events ##
        self.pushButton.clicked.connect(self.setLineText)
        self.pushButton_2.clicked.connect(self.setLineText_2)

        self.textBrowser = QTextBrowser(self.frame_2)
        self.textBrowser.setGeometry(QRect(680, 10, 301, 121))
        self.textBrowser.setFrameShape(QFrame.NoFrame)
        self.textBrowser.setFrameShadow(QFrame.Plain)
        self.textBrowser.setObjectName("textBrowser")
        self.label_9 = QLabel(self.frame_2)
        self.label_9.setGeometry(QRect(650, 20, 21, 91))
        self.label_9.setText("")
        self.label_9.setPixmap(QPixmap("Asset 3.png"))
        self.label_9.setObjectName("label_9")
        self.verticalLayout.addWidget(self.frame_2)
        self.treeView = QTreeView(self.centralwidget)
        self.treeView.setMinimumSize(QSize(0, 80))
        self.treeView.setFrameShape(QFrame.WinPanel)
        self.treeView.setFrameShadow(QFrame.Plain)
        self.treeView.setDragEnabled(True)
        self.treeView.setAlternatingRowColors(True)
        self.treeView.setSelectionMode(QAbstractItemView.MultiSelection)
        self.treeView.setSortingEnabled(True)
        self.treeView.setObjectName("treeView")
        self.treeView.header().setHighlightSections(True)
        self.treeView.header().setSortIndicatorShown(True)
        self.verticalLayout.addWidget(self.treeView)
        self.frame = QFrame(self.centralwidget)
        self.frame.setMinimumSize(QSize(0, 120))
        self.frame.setMaximumSize(QSize(16777215, 190))
        self.frame.setFrameShape(QFrame.NoFrame)
        self.frame.setFrameShadow(QFrame.Plain)
        self.frame.setObjectName("frame")
        self.groupBox_2 = QGroupBox(self.frame)
        self.groupBox_2.setGeometry(QRect(10, 0, 151, 121))
        self.groupBox_2.setObjectName("groupBox_2")
        self.label_3 = QLabel(self.groupBox_2)
        self.label_3.setGeometry(QRect(10, 30, 61, 16))
        self.label_3.setObjectName("label_3")
        self.label_4 = QLabel(self.groupBox_2)
        self.label_4.setGeometry(QRect(10, 50, 61, 16))
        self.label_4.setObjectName("label_4")
        self.label = QLabel(self.groupBox_2)
        self.label.setGeometry(QRect(10, 70, 61, 16))
        self.label.setObjectName("label")
        self.label_5 = QLabel(self.groupBox_2)
        self.label_5.setGeometry(QRect(10, 90, 81, 16))
        self.label_5.setObjectName("label_5")
        self.label_2 = QLabel(self.groupBox_2)
        self.label_2.setGeometry(QRect(90, 30, 31, 16))
        self.label_2.setObjectName("label_2")
        self.label_6 = QLabel(self.groupBox_2)
        self.label_6.setGeometry(QRect(90, 50, 31, 16))
        self.label_6.setObjectName("label_6")
        self.label_7 = QLabel(self.groupBox_2)
        self.label_7.setGeometry(QRect(90, 70, 55, 16))
        self.label_7.setObjectName("label_7")
        self.label_8 = QLabel(self.groupBox_2)
        self.label_8.setGeometry(QRect(90, 90, 55, 16))
        self.label_8.setObjectName("label_8")
        self.pushButton_3 = QPushButton(self.frame)
        self.pushButton_3.setGeometry(QRect(850, 0, 131, 28))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.clicked.connect(self.extract)

        self.verticalLayout.addWidget(self.frame)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Matching Tool"))
        self.groupBox.setTitle(_translate("MainWindow", "Import Data Set"))
        self.pushButton.setText(_translate("MainWindow", "File"))
        self.pushButton_2.setText(_translate("MainWindow", "File"))
        self.lineEdit.setPlaceholderText(_translate(
            "MainWindow", "Student Excel Data Set"))
        self.lineEdit_2.setPlaceholderText(
            _translate("MainWindow", "Employer Data Set"))
        self.textBrowser.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
                                            "<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
                                            "p, li { white-space: pre-wrap; }\n"
                                            "</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
                                            "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:16pt; font-weight:600; color:#f5b01c;\">E D U C A T I O N</span></p>\n"
                                            "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:16pt; font-weight:600; color:#f5b01c;\">     </span><span style=\" font-size:16pt; font-style:italic; color:#000000;\">i s</span><span style=\" font-size:16pt; font-weight:600; color:#000000;\">  </span><span style=\" font-size:16pt; font-weight:600; color:#d44f2a;\">F R E E D O M</span></p>\n"
                                            "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:16pt; font-weight:600; color:#d44f2a;\">   </span><span style=\" font-size:16pt; font-weight:600; color:#000000;\">Find Your Future</span></p></body></html>"))
        self.groupBox_2.setTitle(_translate("MainWindow", "Totals"))
        self.label_3.setText(_translate("MainWindow", "Positions:"))
        self.label_4.setText(_translate("MainWindow", "Students:"))
        self.label.setText(_translate("MainWindow", "Matches:"))
        self.label_5.setText(_translate("MainWindow", "Not Matched:"))
        self.label_2.setText(_translate("MainWindow", "Enum"))
        self.label_6.setText(_translate("MainWindow", "Enum"))
        self.label_7.setText(_translate("MainWindow", "Enum"))
        self.label_8.setText(_translate("MainWindow", "Enum"))
        self.pushButton_3.setText(_translate("MainWindow", "Extract Matches"))
        self.treeView.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        font = QFont()
        font.setPointSize(12)
        self.treeView.header().setFont(font)
        self.treeView.header().setDefaultAlignment(Qt.AlignCenter)

    def addData(self, model, data):
        for i, j in data.items():
            _translate = QCoreApplication.translate
            model.insertRow(0)
            model.setData(model.index(0, 0), i)
            model.setData(model.index(0, 1), str(j[0]))
            model.setData(model.index(0, 2), str(j[1]))
            model.setData(model.index(0, 3), str(j[2]))
            self.label_2.setText(_translate("MainWindow", str(len(data))))
            self.label_6.setText(_translate("MainWindow", str(len(data))))
            self.label_7.setText(_translate("MainWindow", str(len(data))))
            self.label_8.setText(_translate("MainWindow", str(0)))

    def createModel(self, parent):
        model = QStandardItemModel(0, 4)
        model.setHeaderData(0, Qt.Horizontal, "Job Title")
        model.setHeaderData(1, Qt.Horizontal, "Companys Rating")

        model.setHeaderData(2, Qt.Horizontal, "Student Match")
        model.setHeaderData(3, Qt.Horizontal, "Students Rating of Company")
        return model

    def setLineText(self):
        Path = App().Msg
        s, k = os.path.split(Path)
        self.lineEdit.setText(k)
        self.studentrank_file = Path

        self.wb_student = xlrd.open_workbook(self.studentrank_file)
        self.sheet_student = self.wb_student.sheet_by_index(0)
        del Path

    def setLineText_2(self):
        Path = App().Msg
        s, k = os.path.split(Path)
        self.lineEdit_2.setText(k)
        self.companyrank_file = Path
        # Finds Excel path for match maker to extract

        self.wb_company = xlrd.open_workbook(self.companyrank_file)
        self.sheet_company = self.wb_company.sheet_by_index(0)

    def extract(self):
        model = self.createModel(self)
        self.treeView.setModel(model)
        self.companyRank = {}
        self.studentRank = {}
        self.companySlots = {}

        for r in range(self.sheet_student.nrows):
            self.studentChoices = []
            if r == 0:
                continue
            self.student = str(self.sheet_student.cell_value(r, 0))
            self.student.strip()
            for c in range(1, 6):
                comp = self.sheet_student.cell_value(r, c)
                self.studentChoices.append(comp)

            self.studentRank[self.student] = self.studentChoices

        for r in range(self.sheet_company.nrows):
            self.companyChoices = []
            self.company = self.sheet_company.cell_value(r, 0)
            if r == 0:
                continue
            for c in range(self.sheet_company.ncols):
                self.student = self.sheet_company.cell_value(r, c)
                self.student.strip()
                if c == 0:
                    continue
                if self.student == "":
                    break

                self.companyChoices.append(self.student)

            self.companySlots[self.company] = 1
            self.companyRank[self.company] = self.companyChoices

        self.students = list(self.studentRank.keys())
        self.companies = list(self.companyRank.keys())

        (matched, studentslost) = self.matchmaker()
        self.check(matched, studentslost)

        for self.company in matched.keys():
            student = matched[self.company][0]
            cr = self.companyRank[self.company].index(student)+1
            if self.company in self.studentRank[self.student]:
                sr = self.studentRank[self.student].index(self.company)+1
            else:
                sr = 0
            matched[self.company].append(sr)
            matched[self.company].append(cr)

        self.addData(model, matched)

    def matchmaker(self):
        unmatchedStudents = self.students[:]
        # print(unmatchedStudents)
        studentslost = []
        matched = {}
        for companyName in self.companies:
            if companyName not in matched:
                matched[companyName] = list()
        studentRank2 = copy.deepcopy(self.studentRank)
        companyRank2 = copy.deepcopy(self.companyRank)
        while unmatchedStudents:
            self.student = unmatchedStudents.pop(0)
            print("%s is available" % (self.student))
            studentRankList = studentRank2[self.student]
            # print(studentRankList)
            while studentRankList:
                self.company = studentRankList.pop(0)
                if self.student in self.companyRank[self.company]:
                    print("  %s (company's #%s) is checking out %s (student's #%s)" % (
                        self.student, (self.companyRank[self.company].index(self.student)+1), self.company, (self.studentRank[self.student].index(self.company)+1)))
                    # get list of tentative matches for current company
                    tempmatch = matched.get(self.company)
                    # check if there's still a slot open and add student to matched list
                    if len(tempmatch) < self.companySlots.get(self.company):
                        if self.student not in matched[self.company]:
                            matched[self.company].append(self.student)
                            print("    There's a spot! Now matched: %s and %s" % (
                                self.student.upper(), self.company.upper()))
                            break
                    # when no slot open, check whether the current student is higher ranked than the students in tempmatch
                    else:
                        companieslist = companyRank2[self.company]
                        # [(0, 'Grace'), (1, 'Bob')] enumerate returns the index and element as a tuple from the iterable given
                        for (i, matchedAlready) in enumerate(tempmatch):
                            # Compare current student and old student rank
                            if companieslist.index(matchedAlready) > companieslist.index(self.student):
                                if self.student not in matched[self.company]:
                                    matched[self.company][i] = self.student
                                    print("  %s dumped %s (company's #%s) for %s (company's #%s)" % (self.company.upper(), matchedAlready, (
                                        self.companyRank[self.company].index(matchedAlready)+1), self.student.upper(), (self.companyRank[self.company].index(self.student)+1)))
                                    # check if old student has more companies to try
                                    if studentRank2[matchedAlready]:
                                        unmatchedStudents.append(
                                            matchedAlready)
                                    else:
                                        studentslost.append(matchedAlready)
                            else:
                                # when company still prefers old match and current student has no more companies to try
                                print("  %s would rather stay with %s (their #%s) over %s (their #%s)" % (self.company, matchedAlready, (
                                    self.companyRank[self.company].index(matchedAlready)+1), self.student, (self.companyRank[self.company].index(self.student)+1)))
                                if not studentRankList:
                                    studentslost.append(self.student)
        for lostsoul in studentslost:
            print('%s did not match' % lostsoul)
        return (matched, studentslost)

    def check(self, matched, lostSouls):
        modified = False
        while True:
            for company in matched.keys():
                modified = False
                for i in range(len(lostSouls)):
                    if lostSouls[i] in self.companyRank[company]:
                        if self.companyRank[company].index(lostSouls[i]) < self.companyRank[company].index(matched[company][0]):
                            temp = matched[company][0]
                            matched[company] = [lostSouls[i]]
                            lostSouls[i] = temp
                            modified = True
            if not modified:
                break


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
