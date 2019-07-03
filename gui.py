import sys
#from PyQt5.QtWidgets import (QPushButton, QWidget, QLineEdit, QApplication, QPlainTextEdit)
#from PyQt5.QtCore import *
#from PyQt5.QtGui import *
#from PyQt5.QtWidgets import *
from PyQt5 import QtWidgets
import os

raw_data_file_path = ""

class Button(QtWidgets.QPushButton):
    def __init__(self, title, parent):
        super().__init__(title, parent)
        self.setAcceptDrops(True)
        
    def dragEnterEvent(self, e):
        m = e.mimeData()
        if m.hasUrls():
            e.accept()
        else:
            e.ignore()

    def dropEvent(self, e):
        global raw_data_file_path
        m = e.mimeData()
        if m.hasUrls():
            filepath = m.urls()[0].toLocalFile()
            print(F"filepath: {filepath}")
            if (os.path.splitext(filepath)[-1].lower() != ".xlsx"):
                self.setText("Seggl..\n\nBitte Rohdaten im Format *.xlsx wählen!")
            else:
                self.setText("Bitte Analyse starten\n\n.." + filepath[-40:])
                raw_data_file_path = filepath
        
def start_analyses():
    global raw_data_file_path
    
    if not os.path.isfile(raw_data_file_path):    
        print(F"Bitte wähle eine Datei aus.")
    else:
        print(F"button pressed: {raw_data_file_path}")
    
class MainGui(QtWidgets.QDialog):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        button = Button("Drag and Drop your file here", self)
        button.resize(330, 180)
        button.setStyleSheet("border: 2px dashed black;border-radius: 10px")
        button.move(10, 10)
        
        btn_run = QtWidgets.QPushButton("Analyse starten", self)
        btn_run.resize(330, 40)
        btn_run.move(10, 200)
        btn_run.clicked.connect(start_analyses)
        
        self.setWindowTitle('AMINOS v0.1')
        self.setGeometry(400, 400, 350, 250)
        self.setFixedSize(350, 260)
        
if __name__ == '__main__':    
    app = QtWidgets.QApplication(sys.argv)
    ex = MainGui()
    ex.show()
    app.exec_()