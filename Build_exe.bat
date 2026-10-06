@echo off
rem Run once: builds dist\Excel_Calc.exe (program 1) and dist\Excel_Summary.exe (program 2). No Python needed to use them afterwards.
python -m pip install --upgrade pyinstaller openpyxl xlrd xlutils xlwt
python -m PyInstaller --onefile --noconsole --name Excel_Calc "%~dp0Excel_Accounting.py"
python -m PyInstaller --onefile --noconsole --name Excel_Summary "%~dp0Excel_Summary.py"
echo.
echo Done. Programs: %~dp0dist\Excel_Calc.exe and %~dp0dist\Excel_Summary.exe - drag Excel files onto them.
echo Put reference.txt next to Excel_Calc.exe if you want to edit the category list.
pause
