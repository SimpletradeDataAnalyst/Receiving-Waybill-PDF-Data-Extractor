from PyQt5.QtWidgets import QMainWindow, QLabel, QApplication, QPushButton, QFileDialog
from PyQt5 import uic
import sys

class UI(QMainWindow):
    def _init_(self):
        super(UI, self).__init__()
        
        #Load the ui file
        uic.loadUi ("mainwindow.ui", self)

        #Define your widgets
        self.button = self.findChild (QPushButton, "openFile")
        self.label = self.findChild (QLabel, "filePath")
        
        #Click the dropdown
        self.button.clicked.connect(self.clicker)


        #show the application
        self.show()

    def clicker (self):
        self.label.setText ("Adrian")


app = QApplication (sys.argv)
UIWindow = UI ()
app.exec_()


