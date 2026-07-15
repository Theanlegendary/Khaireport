@echo off
title DataPusher Setup & Run
echo ===================================================
echo           DataPusher Setup & Run Script
echo ===================================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    py --version >nul 2>&1
    if %errorlevel% neq 0 (
        goto :INSTALL_PYTHON
    ) else (
        set PYTHON_CMD=py
    )
) else (
    set PYTHON_CMD=python
)

:INSTALL_DEPS
echo [1/2] Installing requirements...
%PYTHON_CMD% -m pip install --upgrade pip
%PYTHON_CMD% -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)
echo [SUCCESS] Dependencies installed successfully.
echo.

:RUN_BOT
echo [2/2] Starting bot.py...
echo.
%PYTHON_CMD% bot.py
pause
exit /b 0

:INSTALL_PYTHON
echo [WARNING] Python is not installed on this system!
echo Downloading Python 3.11.8 Installer...
echo.
powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe' -OutFile 'python_installer.exe'"
if not exist python_installer.exe (
    echo [ERROR] Failed to download Python installer. Please install Python manually from python.org.
    pause
    exit /b 1
)

echo Installing Python... (Please wait, a progress bar will appear)
python_installer.exe /passive PrependPath=1 Include_test=0
del python_installer.exe

echo.
echo ===================================================
echo [SUCCESS] Python installation completed!
echo.
echo IMPORTANT: Please CLOSE this window and open this script 
echo again to allow Windows to load the new Python path.
echo ===================================================
pause
exit /b 0
