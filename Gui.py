import sys
import os
from PyQt5.QtGui import (QIcon, QStandardItemModel, QPixmap)
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
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1045, 682)
        self.studentDataPath = ""
        self.employeerDataPath = ""

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

    def setLineText(self):
        Path = App().Msg

        self.lineEdit.setText(Path)
        self.studentDataPath = Path

    def setLineText_2(self):
        Path = App().Msg
        self.lineEdit_2.setText(App().Msg)
        self.employeerDataPath = Path

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
