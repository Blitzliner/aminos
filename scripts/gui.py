import sys
from PyQt5 import QtWidgets, QtCore
import os
#import pickle needed for loading data from binary file
import pandas as pd
import aminos
import logging
import subprocess

_logger = logging.getLogger("gui")

class Button(QtWidgets.QPushButton):
    def __init__(self, title, parent):
        super().__init__(title, parent)
        self.setAcceptDrops(True)
        self.raw_data_file_path = "rohdaten_example.xlsx"
        
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
            print(F"filepath: {filepath}")
            if (os.path.splitext(filepath)[-1].lower() != ".xlsx"):
                self.setText("Netter Versuch..\n\nBitte Rohdaten im Format *.xlsx wählen!")
            else:
                self.setText("Bitte Analyse starten\n\n.." + filepath[-40:])
                self.raw_data_file_path = filepath
                
    def get_path(self):
        return self.raw_data_file_path
  
class MainGui(QtWidgets.QDialog):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.button = Button("Drag and Drop your file here", self)
        self.button.resize(330, 180)
        self.button.setStyleSheet("border: 2px dashed black;border-radius: 10px")
        self.button.move(10, 10)
        
        btn_run = QtWidgets.QPushButton("Analyse starten", self)
        btn_run.resize(330, 40)
        btn_run.move(10, 200)
        btn_run.clicked.connect(self.start_analyses)
        
        self.setWindowTitle('AMINOS v0.1')
        self.setGeometry(400, 400, 350, 250)
        self.setFixedSize(350, 260)

    def start_analyses(self):
        if not os.path.isfile(self.button.get_path()):    
            self.button.setText("Bitte wähle eine Datei aus.")
        else:
            
            #with open('data.pickle', 'rb') as handle:
                #results = pickle.load(handle)
            cfg = aminos.read_config()
            cfg["file_to_analyze"] = self.button.get_path()
            results = aminos.analyse(cfg)
            #open_path = results['export_excel_path']
            #subprocess.Popen(['explorer.exe', '/select,"{open_path}"'])
            conflicts, ret = DateDialog.ShowDialog(results)
            _logger.info(conflicts)
            
            if (ret == True):
                _logger.info("re-run with prefered control and AS")
                cfg['prefer_control'] = conflicts[0]
                cfg['prefer_aminos'] = conflicts[1]
                data = aminos.analyse(cfg)
                msgBox = QtWidgets.QMessageBox()
                msgBox.setText("Analyse erfolgreich durchgeführt.\nFenster wird geschlossen.");
                msgBox.exec();
                self.close()
            else:
                _logger.info("program finished")
                self.close()

class DateDialog(QtWidgets.QDialog):
    def __init__(self, results, parent = None):
        super(DateDialog, self).__init__(parent)
        self.setWindowTitle('Analyse Ergebnisse')
        
        dat = results['selected_control']['data']
        best_control = str(results['selected_control']['best_control_name'])
        best_control_score = str(results['selected_control']['best_control_score'])
        
        self.cb_control = QtWidgets.QComboBox()
        self.cb_control.addItems(dat.keys())
        self.cb_control.currentTextChanged.connect(self.on_control_changed)
        
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel, QtCore.Qt.Horizontal, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        
        main_layout =  QtWidgets.QGridLayout(self)
        l_control = QtWidgets.QLabel('Kontrolle:')
        export_dir = results['export_dir']#.replace('\\', '/')
        #export_excel_path = results['export_excel_path'].replace('\\', '/')
        description = F"Die Analyse mit der Kontrolle {best_control} ist fertig."
        l_description = QtWidgets.QLabel(description)
        l_export_dir = QtWidgets.QLabel(F"Die Ergebnisse liegen unter:\n{export_dir}\n\nDie besten Aminosäuren wurden bereits ausgetauscht. Es wurden jedoch gleichwertige Kontrollen gefunden. Falls die Analyse mit anderen Kontrollen durchlaufen werden soll, bitte auswählen:")
        l_export_dir.setWordWrap(True)
        l_new_analyse = QtWidgets.QLabel("Wähle OK für eine erneute Analyse mit den ausgewählten Parametern oder Cancel um das Program zu beenden. Bei OK werden die neuen Ergebnisse in einem neuen Ordner mit aktuellem Zeitstempel abgelegt.")
        l_new_analyse.setWordWrap(True)
        #l_export_dir.setOpenExternalLinks(True)
        #l_export_excel_path = QtWidgets.QLabel("<a href=\"file:///{export_excel_path}\">Öffne Excel Analyse</a>")
        #l_export_excel_path.setOpenExternalLinks(True)
        main_layout.addWidget(l_description      , 0, 0, 1, 2)
        main_layout.addWidget(l_export_dir       , 1, 0, 1, 2)
        main_layout.addWidget(l_control          , 2, 0)
        main_layout.addWidget(self.cb_control    , 2, 1)
        main_layout.addWidget(l_new_analyse      , 5, 0, 1, 2)
        main_layout.addWidget(buttons            , 6, 0, 1, 2)
        
        #out_str = ""
        self.gbs = {}
        self.aminos = {}
        gb_idx = 3
        first = 1
        for control in dat.keys():
            self.aminos[control] = []
            gb = QtWidgets.QGroupBox()
            if not first:
                gb.hide()
            first = 0
            
            self.gbs[control] = gb
            score = dat[control]['prios_score']
            gb.setTitle(F"Kontrolle {control} (Score: {score})")
            as_idx = 0
            layout =  QtWidgets.QGridLayout()
            for conflict in dat[control]['conflicts']:
                label = QtWidgets.QLabel(F"{conflict[0][0:3]}")
                combobox = QtWidgets.QComboBox()
                combobox.addItems(conflict)
                self.aminos[control].append(combobox)
                layout.addWidget(label,    as_idx, 0)
                layout.addWidget(combobox, as_idx, 1)
                as_idx += 1
            gb.setLayout(layout)    
            
            main_layout.addWidget(gb, gb_idx, 0, 1, 2)
            gb_idx += 1
    
        self.cb_control.setCurrentIndex(self.cb_control.findText(best_control))
            
    def on_control_changed(self):
        for key in self.gbs.keys():
            self.gbs[key].hide()
        self.gbs[self.cb_control.currentText()].show()

    def get_data(self):
        selected_control = self.cb_control.currentText()
        selected_as = []
        self.aminos[selected_control]
        for cb in self.aminos[selected_control]:
            selected_as.append(str(cb.currentText()))
        #self.gbs
        return selected_control, selected_as

    @staticmethod
    def ShowDialog(results, parent = None):
        dialog = DateDialog(results, parent)
        result = dialog.exec_()
        dat = dialog.get_data()
        return (dat, result == QtWidgets.QDialog.Accepted)


def show_main():          
    app = QtWidgets.QApplication(sys.argv)
    ex = MainGui()
    ex.show()
    app.exec_()
    
if __name__ == '__main__':  
    show_main()