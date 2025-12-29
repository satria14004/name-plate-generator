@echo off

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    msg * "ERROR: Python is not installed. Please install Python from https://www.python.org/"
    exit /b 1
)

REM Check if required packages are installed
python -c "import flask, flask_cors, pptx" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing required packages...
    start /wait cmd /c "pip install flask flask-cors python-pptx && echo. && echo Installation complete! && timeout /t 3"
)

REM Check for optional tray icon support
python -c "import pystray, PIL" >nul 2>&1
if %errorlevel% neq 0 (
    REM Install tray icon support (optional)
    pip install pystray pillow >nul 2>&1
)

REM Check which server file to use
if exist "server_tray.pyw" (
    REM Use hidden server with tray icon
    start /B pythonw server_tray.pyw
) else (
    REM Use regular server hidden
    start /B pythonw server.py
    timeout /t 2 /nobreak >nul
    start "" nameplate.html
)

exit