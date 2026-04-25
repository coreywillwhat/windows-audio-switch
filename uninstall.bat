@echo off
setlocal

reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "Windows Audio Switch" /f >nul 2>nul

set "CONFIG_DIR=%LOCALAPPDATA%\Windows Audio Switch"
if exist "%CONFIG_DIR%" (
    choice /M "Delete saved Windows Audio Switch configuration"
    if errorlevel 2 goto done
    rmdir /S /Q "%CONFIG_DIR%"
)

:done
echo Windows Audio Switch startup entry removed.
