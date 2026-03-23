@echo off
REM SwiftPDF Windows build — onedir modu (_MEI temp hatasini onler)
cd /d "%~dp0"
call .venv\Scripts\activate

pyinstaller ^
  --onedir ^
  --windowed ^
  --name "SwiftPDF" ^
  --icon assets\icon2.ico ^
  --add-data "assets;assets" ^
  main.py

echo.
echo Cikti: dist\SwiftPDF\ klasoru
