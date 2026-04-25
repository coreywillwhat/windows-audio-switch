@echo off
setlocal
cd /d "%~dp0"

python -m pip install -r requirements.txt
if errorlevel 1 exit /b %errorlevel%

python -m PyInstaller --noconfirm --clean --onefile --windowed --name "audio-switch" audio-switch.py
if errorlevel 1 exit /b %errorlevel%

echo.
echo Built dist\audio-switch.exe
