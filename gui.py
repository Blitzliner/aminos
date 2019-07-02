import sys
from PyQt5.QtWidgets import (QPushButton, QWidget, QLineEdit, QApplication, QPlainTextEdit)
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import logging

class Button(QPushButton):
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
        m = e.mimeData()
        if m.hasUrls():
            filepath = m.urls()[0].toLocalFile()
            self.parent().label.setText(filepath)

class QPlainTextEditLogger(logging.Handler):
    def __init__(self, parent):
        super().__init__()
        self.widget = QtGui.QPlainTextEdit(parent)
        self.widget.setReadOnly(True)    

    def emit(self, record):
        msg = self.format(record)
        self.widget.appendPlainText(msg) 
        

class Example(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        button = Button("Drag and Drop your file here", self)
        button.resize(220, 150)
        button.setStyleSheet("border: 2px dashed black;border-radius: 10px")
        button.move(5, 5)
        
        #ddlabel = DropLabel("Drag & Drop your file here", self)
        #ddlabel.move(50, 50)
        #ddlabel.setStyleSheet("background-color:#ff0000; margin-left: 10px; border-radius: 25px;");
        #ddlabel.setGeometry(QtCore.QRect(50, 50, 100, 100)) #(x, y, width, height)
        #ddlabel.setStyleSheet("QLabel { background-color : red; color : blue; margin-left: 10px; border-radius: 25px;}"
                              
        self.label = QLineEdit("", self)# QLabel("", self)
        self.label.setPlaceholderText("File to analyze..");
        self.label.setReadOnly(True);
        
        #QPalette *palette = new QPalette();
        #palette->setColor(QPalette::Base,Qt::gray);
        #palette->setColor(QPalette::Text,Qt::darkGray);
        #ui->lineEdit->setPalette(*palette);
        
        # logging output
        #logTextBox = QPlainTextEditLogger(self)
        #logTextBox.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        #logging.getLogger().addHandler(logTextBox)
        #logging.getLogger().setLevel(logging.INFO)

        #self.label.setEchoMode(QLineEdit::Normal); # NoEcho
        #self.label.setPixmap(QPixmap('example.jpg'))
        #self.label.move(5, 160)
        self.label.setGeometry(5, 160, 220, 20);
        
        self.setWindowTitle('AMINOS v0.1')
        self.setGeometry(400, 400, 400, 300) #x, y, w, h
        
#def display(cfg):
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    ex.show()
    app.exec_()