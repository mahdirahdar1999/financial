@echo off
python -m pip install -r requirements.txt
python -m pip install pyinstaller
pyinstaller --noconsole --onefile --add-data "finsuite\resources;finsuite\resources" run.py -n FinSuitePro
pause
