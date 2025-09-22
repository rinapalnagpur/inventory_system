@echo off
title Multi-Shop Inventory Management System
echo ==========================================
echo   Multi-Shop Inventory Management System
echo ==========================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH!
    echo.
    echo Please install Python from: https://python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo Python found: 
python --version
echo.

REM Check if we're in the right directory
if not exist "app.py" (
    echo ERROR: app.py not found!
    echo Please make sure you're running this script from the inventory_system directory
    echo.
    pause
    exit /b 1
)

REM Check if virtual environment exists
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo ERROR: Failed to create virtual environment!
        echo.
        pause
        exit /b 1
    )
    echo Virtual environment created successfully!
    echo.
)

REM Activate virtual environment
echo Activating virtual environment...
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
) else (
    echo ERROR: Virtual environment activation script not found!
    pause
    exit /b 1
)

REM Check if pip is working
python -m pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: pip is not working in virtual environment!
    pause
    exit /b 1
)

REM Install/upgrade requirements
echo Checking and installing requirements...
python -m pip install --upgrade pip --quiet
python -m pip install Flask==2.3.3 --quiet
python -m pip install Flask-CORS==4.0.0 --quiet
python -m pip install pandas --quiet
python -m pip install openpyxl --quiet  
python -m pip install xlsxwriter --quiet
python -m pip install werkzeug==2.3.7 --quiet

echo All packages installed successfully!
echo.

REM Create uploads folder if it doesn't exist
if not exist "uploads" (
    mkdir uploads
    echo Created uploads folder
)

echo.
echo ==========================================
echo   Starting Inventory Management System
echo ==========================================
echo.
echo 🌐 System will be available at: http://localhost:5000
echo 🛑 Press Ctrl+C to stop the server
echo.
echo 📁 Make sure your Excel files are ready:
echo    - Sales data file (item names, sales qty, current stock)
echo    - Multi-location stock file (warehouse and shop quantities)
echo.

REM Start the Flask application
python app.py

echo.
echo Server stopped.
pause
