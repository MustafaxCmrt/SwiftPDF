@echo off
REM
cd /d "%~dp0"
call .venv\Scripts\activate

pyinstaller ^
  --onefile ^
  --windowed ^
  --name "SwiftPDF" ^
  --icon assets\icon2.ico ^
  --add-data "assets;assets" ^
  main.py

echo.
echo Cikti: dist\SwiftPDF.exe
