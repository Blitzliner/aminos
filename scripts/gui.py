import sys
from PyQt5 import QtWidgets, QtCore
import os
#import pickle needed for loading data from binary file
import pandas as pd
import aminos
import logging
import traceback

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
            _logger.info(F"Filepath selected: {filepath}")
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
        self.button = Button("Drag und Drop deine Excel Rohdaten", self)
        self.button.resize(330, 180)
        self.button.setStyleSheet("border: 2px dashed black;border-radius: 10px")
        self.button.move(10, 10)
        
        btn_run = QtWidgets.QPushButton("Analyse starten", self)
        btn_run.resize(330, 40)
        btn_run.move(10, 200)
        btn_run.clicked.connect(self.start_analyses)
        
        self.setWindowTitle('AMINOS v0.1.2')
        self.setGeometry(400, 400, 350, 250)
        self.setFixedSize(350, 260)

    def start_analyses(self):
        if not os.path.isfile(self.button.get_path()):    
            self.button.setText("Bitte wähle eine Datei aus.")
        else:
            try:
                cfg = aminos.read_config()
                cfg["file_to_analyze"] = self.button.get_path()
                results = aminos.analyse(cfg)
                selected_control, ret = DateDialog.ShowDialog(results)
                _logger.info(f'Selected control: {selected_control}')
                
                if (ret == False):
                    _logger.info("re-run with prefered control and AS")
                    cfg['prefer_control'] = selected_control
                    data = aminos.analyse(cfg)
                    msgBox = QtWidgets.QMessageBox()
                    msgBox.setText("Analyse erfolgreich durchgeführt.\nFenster wird geschlossen.");
                    msgBox.exec();
                    self.close()
                else:
                    _logger.info("program finished")
                    self.close()
            except Exception as e:
                err_message = F"Unerwarteter Fehler: {e}\nBitte speicher die Rohdaten Exceltabelle als auch die datei 'logger.log' und kontaktiere den Softwareentwickler.\n{traceback.format_exc()}"
                _logger.error(err_message)
                msgBox = QtWidgets.QMessageBox()
                msgBox.setText(err_message);
                msgBox.exec();

class DateDialog(QtWidgets.QDialog):
    def __init__(self, results, parent = None):
        super(DateDialog, self).__init__(parent)
        self.setWindowTitle('Analyse Ergebnisse')
        best_control_name = results['checked_controls'][0]['name']
        all_controls_str = 'Alle Kontrollen in der Übersicht:\n'
        for cont in results['checked_controls']:
            all_controls_str += f"{cont['name']}: Score: {cont['coarse_score']}/20 ({cont['fine_score']})\n"
        all_controls = [cont['name'] for cont in results['checked_controls']]
        
        btn_ok = QtWidgets.QDialogButtonBox.Ok
        btn_cancel = QtWidgets.QDialogButtonBox.Cancel
        buttons = QtWidgets.QDialogButtonBox(btn_ok | btn_cancel, QtCore.Qt.Horizontal, self)
        buttons.buttons()[0].setText('Beenden')
        buttons.buttons()[1].setText('Analyse wiederholen')
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        
        l_description = QtWidgets.QLabel(F"Die Analyse mit der Kontrolle '{best_control_name}' ist fertig.")
        l_export_dir = QtWidgets.QLabel(F"Die Ergebnisse liegen unter:\n{results['export_dir']}\n\n\n{all_controls_str}")
        l_export_dir.setWordWrap(True)
        l_new_analyse = QtWidgets.QLabel("Wähle 'Analyse wiederholen' für eine erneute Analyse mit der ausgewählten Kontrolle oder 'Beenden' um das Program zu beenden. Bei 'Analyse wiederholen' werden die neuen Ergebnisse in einem neuen Ordner mit aktuellem Zeitstempel abgelegt.")
        l_new_analyse.setWordWrap(True)
        l_control = QtWidgets.QLabel('Gewählte Kontrolle:')
        self.cb_control = QtWidgets.QComboBox()
        self.cb_control.addItems(all_controls)
        
        main_layout =  QtWidgets.QGridLayout(self)
        main_layout.addWidget(l_description      , 0, 0, 1, 2)
        main_layout.addWidget(l_export_dir       , 1, 0, 1, 2)
        main_layout.addWidget(l_new_analyse      , 2, 0, 1, 2)
        main_layout.addWidget(l_control          , 5, 0)
        main_layout.addWidget(self.cb_control    , 5, 1)
        main_layout.addWidget(buttons            , 6, 0, 1, 2)
        
        self.cb_control.setCurrentIndex(self.cb_control.findText(best_control_name))
        
    def get_data(self):
        return self.cb_control.currentText()

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