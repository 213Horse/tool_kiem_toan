@echo off
title Kiem Kho App
color 0A

echo ========================================
echo    KIEM KHO - AUTO INSTALL AND RUN
echo ========================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not installed!
    echo.
    echo Please download and install Python from: https://www.python.org/downloads/
    echo Remember to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo [OK] Python found
python --version
echo.

REM Check pip
pip --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] pip not installed!
    pause
    exit /b 1
)

echo [OK] pip found
echo.

REM Install libraries
echo [INFO] Installing required libraries...
echo.
pip install --quiet --upgrade pip

if exist "requirements.txt" (
    echo [INFO] Installing from requirements.txt...
    pip install --quiet -r requirements.txt
) else (
    echo [INFO] Installing libraries directly...
    pip install --quiet pandas openpyxl xlrd
)

if errorlevel 1 (
    echo [ERROR] Cannot install libraries!
    pause
    exit /b 1
)

echo [OK] Libraries installed successfully
echo.

REM Check Excel file
if not exist "DuLieuDauVao.xlsx" (
    echo [WARNING] DuLieuDauVao.xlsx not found!
    echo Please make sure this file is in the same folder!
    echo.
)

REM Run application
echo ========================================
echo    STARTING APPLICATION...
echo ========================================
echo.
python kiem_kho_app.py

if errorlevel 1 (
    echo.
    echo [ERROR] Application encountered an error!
    pause
    exit /b 1
)

pause


