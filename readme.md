# Aminosäure Analyse Tool
1. Anlegen von Projektordner, Rohdaten werden mit Zeitstempel versehen und kopiert.

2. Die Rohdaten sortieren:
    - Sigma und Phe werden nicht beachtet und können über das json file konfiguriert werden
    - Ko I werden alle nach oben genommen Ko I (61)
    - Ko II kommt darunter Ko II (RV-62)
    - Patienten aufsteigend sortieren

3. Kontrollen kontrollieren
    - alle AS einer Kontrolle mit der Grenzwert-Datei vergleichen Ala: 54, untere grenze: 53, obere Grenze 59
    - wenn zu hoch = rot markieren, zu niedrig = blau, okay = grün
    - anzeigen, wie viele von Ko I drin sind und wie viele von Ko 2 drin/zu hoch/niedrig sind
    - wenn eine Kontrolle nicht im Grenzbereich ist, schauen ob eine andere mit dem gleichen Namen drin ist (Thy oder Thy_ph oder Thy_hp4)
    - wenn es eine bessere gibt → austauschen
    - am Ende eine Kontrolle auswählen, die am besten ist Ko I (61) oder Ko II (62)

4. Daten übertragen
    - Patientennummern, die Werte und der Name der ausgetauschen AS in neues Sheet kopieren und transponieren
    - Immer 4 Patienten, eine Leerspalte, wieder 4 Patienten auf eine DINA4 Seite ablegen
    - Immer 3 Zeilen zusammenlegen, dann eine Leerzeil (letzte hat nur 2 AS)
    - Die grau markierten AS bleiben grau
    - Die anderen werden mit den Normwerten abgeglichen: zu hoch: rot, zu niedrig: blau
    - Bei Gleichstand der Kontrollen oder AS kann angegeben werden welche bevorzugt werden soll    

5. Excel wird exportiert
    - Die Analyse beinhaltet die Orginaldaten, die analyse der Kontrollen und die Ergebnisse der Patienten.

# Programm
run.bat führt eine Analyse mit den Parametern aus config.json aus.
runGui.bat führt die gleiche Analyse durch aber graphisch unterstützt. Bei Gleichstand der Kontrollen oder AS können im Folgenden Dialog diese ausgewählt werden.

# dependencies
- Python 3.6
- Python packages: (pip.exe install)
  - xlsxwriter
  - PyQt5
  - pandas
  - xlrd
  - pyinstaller
