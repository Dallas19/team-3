import sys
import os
import copy
from collections import defaultdict
from PyQt5.QtGui import (QIcon, QStandardItemModel, QPixmap, QFont)
from PyQt5.QtCore import (QDate, QDateTime, QRegExp, QSortFilterProxyModel,
                          Qt, QTime, QMetaObject, QSize, QRect, QCoreApplication, QObject, pyqtSignal)
from PyQt5.QtWidgets import (QApplication, QCheckBox, QComboBox, QGridLayout,
                             QGroupBox, QHBoxLayout, QLabel, QLineEdit, QTreeView, QVBoxLayout, QTabWidget,
                             QWidget, QPushButton, QFrame, QTextEdit, QToolButton, QMenuBar,
                             QMainWindow, QStatusBar, QMenu, QFileDialog, QAction)


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
        self.studentDataPath = ""
        self.employeerDataPath = ""
        self.studentprefers = {
            'Alex':  ['American', 'Mercy', 'County', 'Mission', 'General', 'Fairview', 'Saint Mark', 'City', 'Deaconess', 'Park'],
            'Brian':  ['County', 'Deaconess', 'American', 'Fairview', 'Mercy', 'Saint Mark', 'City', 'General', 'Mission', 'Park'],
            'Cassie':  ['Deaconess', 'Mercy', 'American', 'Fairview', 'City', 'Saint Mark', 'Mission', 'Park', 'County', 'General'],
            'Dana':  ['Mission', 'Saint Mark', 'Fairview', 'Park', 'Deaconess', 'Mercy', 'General', 'City', 'County', 'American'],
            'Edward':   ['General', 'Fairview', 'City', 'County', 'Saint Mark', 'Mercy', 'American', 'Mission', 'Deaconess', 'Park'],
            'Faith': ['City', 'American', 'Fairview', 'Park', 'Mercy', 'Mission', 'County', 'General', 'Deaconess', 'Saint Mark'],
            'George':  ['Park', 'Mercy', 'Mission', 'City', 'County', 'American', 'Fairview', 'Deaconess', 'General', 'Saint Mark'],
            'Hannah':  ['American', 'Mercy', 'Deaconess', 'Saint Mark', 'Mission', 'County', 'General', 'City', 'Park', 'Fairview'],
            'Ian':  ['Park', 'County', 'Fairview', 'Deaconess', 'City', 'American', 'Saint Mark', 'Mission', 'General', 'Mercy'],
            'Jessica':  ['American', 'Saint Mark', 'General', 'Park', 'Mercy', 'City', 'Fairview', 'County', 'Mission', 'Deaconess']}

        self.programprefers = {
            'American':  ['Brian', 'Faith', 'Jessica', 'George', 'Ian', 'Alex', 'Dana', 'Edward', 'Cassie', 'Hannah'],
            'City':  ['Brian', 'Alex', 'Cassie', 'Faith', 'George', 'Dana', 'Ian', 'Edward', 'Jessica', 'Hannah'],
            'County': ['Faith', 'Brian', 'Edward', 'George', 'Hannah', 'Cassie', 'Ian', 'Alex', 'Dana', 'Jessica'],
            'Fairview':  ['Faith', 'Jessica', 'Cassie', 'Alex', 'Ian', 'Hannah', 'George', 'Dana', 'Brian', 'Edward'],
            'Mercy':  ['Jessica', 'Hannah', 'Faith', 'Dana', 'Alex', 'George', 'Cassie', 'Edward', 'Ian', 'Brian'],
            'Saint Mark':  ['Brian', 'Alex', 'Edward', 'Ian', 'Jessica', 'Dana', 'Faith', 'George', 'Cassie', 'Hannah'],
            'Park':  ['Jessica', 'George', 'Hannah', 'Faith', 'Brian', 'Alex', 'Cassie', 'Edward', 'Dana', 'Ian'],
            'Deaconess': ['George', 'Jessica', 'Brian', 'Alex', 'Ian', 'Dana', 'Hannah', 'Edward', 'Cassie', 'Faith'],
            'Mission':  ['Ian', 'Cassie', 'Hannah', 'George', 'Faith', 'Brian', 'Alex', 'Edward', 'Jessica', 'Dana'],
            'General':  ['Edward', 'Hannah', 'George', 'Alex', 'Brian', 'Jessica', 'Cassie', 'Ian', 'Faith', 'Dana']}

        self.programSlots = {
            'American': 1,
            'City': 1,
            'County': 1,
            'Fairview': 1,
            'Mercy': 1,
            'Saint Mark': 2,
            'Park': 1,
            'Deaconess': 2,
            'Mission': 2,
            'General': 9}

        self.students = sorted(self.studentprefers.keys())
        self.programs = sorted(self.programprefers.keys())

        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1045, 682)

        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.groupBox = QGroupBox(self.centralwidget)
        self.groupBox.setGeometry(QRect(780, 10, 261, 131))
        self.groupBox.setObjectName("groupBox")

        ### Student Excel Button ###
        self.pushButton = QPushButton(self.groupBox)
        self.pushButton.setGeometry(QRect(10, 40, 93, 28))
        self.pushButton.setObjectName("pushButton")

        self.lineEdit = QLineEdit(self.groupBox)
        self.lineEdit.setGeometry(QRect(110, 40, 141, 31))
        self.lineEdit.setObjectName("lineEdit")

        ### Employer Excel Button ###
        self.pushButton_2 = QPushButton(self.groupBox)
        self.pushButton_2.setGeometry(QRect(10, 90, 93, 28))
        self.pushButton_2.setObjectName("pushButton_2")

        self.lineEdit_2 = QLineEdit(self.groupBox)
        self.lineEdit_2.setGeometry(QRect(110, 90, 141, 31))
        self.lineEdit_2.setObjectName("lineEdit_2")

        self.treeView = QTreeView(self.centralwidget)
        self.treeView.setGeometry(QRect(10, 170, 1021, 461))
        self.treeView.setObjectName("treeView")
        self.treeView.header().setVisible(True)


        ### Button Actions ###
        self.pushButton.clicked.connect(self.setLineText)
        self.pushButton_2.clicked.connect(self.setLineText_2)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setGeometry(QRect(0, 0, 1045, 26))
        self.menubar.setObjectName("menubar")
        self.menuEdit = QMenu(self.menubar)
        self.menuEdit.setObjectName("menuEdit")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.menubar.addAction(self.menuEdit.menuAction())

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
        self.menuEdit.setTitle(_translate("MainWindow", "Edit"))

    def addData(self, model, data):
        model.insertRow(0)

        model.setData(model.index(0, 0), data)
        model.setData(model.index(0, 1), data)

    def createModel(self, parent):
        model = QStandardItemModel(0, 2)
        model.setHeaderData(0, Qt.Horizontal, "Job Title ")
        model.setHeaderData(1, Qt.Horizontal, "Student Match")
        return model

    def setLineText(self):
        Path = App().Msg

        self.lineEdit.setText(Path)
        self.studentDataPath = Path

    def setLineText_2(self):
        Path = App().Msg
        self.lineEdit_2.setText(App().Msg)
        self.employeerDataPath = Path
        ## Finds Excel path for matcher maker to extract

        model = self.createModel(self)
        self.treeView.setModel(model)
        
        (matched, studentslost) = self.matchmaker()
        model.addData(matched)

    def matchmaker(self):
        studentsfree = self.students[:]
        studentslost = []
        matched = {}
        for programName in self.programs:
            if programName not in matched: 
                matched[programName] = list()
        studentprefers2 = copy.deepcopy(self.studentprefers)
        programprefers2 = copy.deepcopy(self.programprefers)
        while studentsfree:
            student = studentsfree.pop(0)
            print("%s is on the market" % (student))
            studentslist = studentprefers2[student]
            program = studentslist.pop(0)
            print("  %s (program's #%s) is checking out %s (student's #%s)" % (
                student, (self.programprefers[program].index(student)+1), program, (self.studentprefers[student].index(program)+1)))
            tempmatch = matched.get(program)
            if len(tempmatch) < self.programSlots.get(program):
                # Program's free
                if student not in matched[program]:
                    matched[program].append(student)
                    print("    There's a spot! Now matched: %s and %s" %
                          (student.upper(), program.upper()))
            else:
                # The student proposes to an full program!
                programslist = programprefers2[program]
                for (i, matchedAlready) in enumerate(tempmatch):
                    if programslist.index(matchedAlready) > programslist.index(student):
                        # Program prefers new student
                        if student not in matched[program]:
                            matched[program][i] = student
                            print("  %s dumped %s (program's #%s) for %s (program's #%s)" % (program.upper(), matchedAlready, (self.programprefers[program].index(
                                matchedAlready)+1), student.upper(), (self.programprefers[program].index(student)+1)))
                            if studentprefers2[matchedAlready]:
                                # Ex has more programs to try
                                studentsfree.append(matchedAlready)
                            else:
                                studentslost.append(matchedAlready)
                    else:
                        # Program still prefers old match
                        print("  %s would rather stay with %s (their #%s) over %s (their #%s)" % (program, matchedAlready, (
                            self.programprefers[program].index(matchedAlready)+1), student, (self.programprefers[program].index(student)+1)))
                        if studentslist:
                            # Look again
                            studentsfree.append(student)
                        else:
                            studentslost.append(student)
        print
        for lostsoul in studentslost:
            print('%s did not match' % lostsoul)
        print(matched)
        return (matched, studentslost)

    def check(self, matched):
        inversematched = defaultdict(list)
        for programName in matched.keys():
            for studentName in matched[programName]:
                inversematched[programName].append(studentName)

        for programName in matched.keys():
            for studentName in matched[programName]:

                programNamelikes = self.programprefers[programName]
                programNamelikesbetter = programNamelikes[:programNamelikes.index(
                    studentName)]
                helikes = self.studentprefers[studentName]
                helikesbetter = helikes[:helikes.index(programName)]
                for student in programNamelikesbetter:
                    for p in inversematched.keys():
                        if student in inversematched[p]:
                            studentsprogram = p
                    studentlikes = self.studentprefers[student]

                    try:
                        studentlikes.index(studentsprogram)
                    except ValueError:
                        continue

                    if studentlikes.index(studentsprogram) > studentlikes.index(programName):
                        print("%s and %s like each other better than "
                              "their present match: %s and %s, respectively"
                              % (programName, student, studentName, studentsprogram))
                        return False
                for program in helikesbetter:
                    programsstudents = matched[program]
                    programlikes = self.programprefers[program]
                    for programsstudent in programsstudents:
                        if programlikes.index(programsstudent) > programlikes.index(studentName):
                            print("%s and %s like each other better than "
                                  "their present match: %s and %s, respectively"
                                  % (studentName, program, programName, programsstudent))
                            return False
        return True


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
