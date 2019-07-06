#Automatisierte Analyse Schritte	
1. Rohdaten von Stick in neuen Ordner legen 	Aminosäuren → 20190523_13:29 → Rohdatei und Auswertung und PDF
2. Die Rohdaten sortieren:	
  - Sigma und Phe werden nicht beachtet und können über das json file konfiguriert werden
  - Ko I werden alle nach oben genommen	Ko I (61)
  - Ko II kommt darunter	Ko II (RV-62)
  - dann alle Patientennummern aufsteigend	71420026, 71420027, 71420028, … Von jeder Zeile alle Spalten kopieren)	
   
3. Kontrollen kontrollieren	
  - alle AS einer Kontrolle mit der Grenzwert-Datei vergleichen	Ala: 54, untere grenze: 53, obere Grenze 59
  - wenn zu hoch = rot markieren, zu niedirg = blau, okay = grün
   - anzeigen, wie viele von Ko I drin sind und wie viele von Ko 2 drin/zu hoch/niedrig sind
   - wenn eine Kontrolle nicht im Grenzbereich ist, schauen ob eine andere mit dem gleichen Namen drin ist 	Thy oder Thy_ph oder Thy_hp4
   - wenn es eine bessere gibt → austauschen	
   - am ende eine Kontrolle auswählen, die am besten ist	Ko I (61) oder Ko II (62)
    
4. Daten übertragen	
   - Patientennummern, die Werte und der Name der ausgetauschen AS in neues Sheet kopieren und transponieren
   - Immer 4 Patienten, eine Leerspalte, wieder 4 Patienten auf eine DINA4 Seite ablegen	
   - Immer 3 Zeilen zusammenlegen, dann eine Leerzeil (letzte hat nur 2 AS)	(Formatierung)
   - Die grau markierten AS bleiben grau	(Formatierung)
   - Die anderen werden mit den Normwerten abgeglichen: zu hoch: rot, zu niedrig: blau	(Formatierung)
    
5. Excel wird exportiert
   - Wenn eine AS nicht rein geht → alle Werte grau markieren	


#dependencies
- Python 3.6
pip.exe install following modules for Python 3.6
- xlsxwriter
- PyQt5
- pandas
- xlrd
- pyinstaller
