cd scripts
::C:\Users\Blitzliner\AppData\Local\Programs\Python\Python36\Scripts\pyinstaller --onedir gui.py
C:\Users\Carme\Anaconda3\Scripts\pyinstaller --onedir gui.py --noconfirm
cd ..
copy scripts\config.json  scripts\dist\gui\config.json
pause