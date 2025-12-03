@echo off
title Build EXE File
color 0B

echo ========================================
echo    BUILD EXE FILE
echo ========================================
echo.
echo This will build .exe file with all libraries
echo to fix "No module named pandas" error
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not installed!
    echo Please install Python from: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [OK] Python found
python --version
echo.

REM Install required libraries
echo [INFO] Installing required libraries...
pip install --quiet --upgrade pip

if exist "requirements.txt" (
    echo [INFO] Installing from requirements.txt...
    pip install --quiet -r requirements.txt
) else (
    echo [INFO] Installing libraries directly...
    pip install --quiet pandas openpyxl xlrd pyinstaller
)

if errorlevel 1 (
    echo [ERROR] Cannot install libraries!
    pause
    exit /b 1
)

echo [OK] Libraries installed successfully
echo.

REM Delete old build folders
if exist "build" (
    echo [INFO] Deleting old build folder...
    rmdir /s /q build
)
if exist "dist" (
    echo [INFO] Deleting old dist folder...
    rmdir /s /q dist
)
if exist "KiemKhoApp.spec" (
    echo [INFO] Deleting old spec file...
    del /q KiemKhoApp.spec
)

echo.

REM Build .exe file
echo ========================================
echo    BUILDING EXE FILE...
echo ========================================
echo.
echo Building, please wait (may take a few minutes)...
echo.

pyinstaller --onefile --windowed --name "KiemKhoApp" ^
    --add-data "DuLieuDauVao.xlsx;." ^
    --add-data "Kiemke_template.xlsx;." ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --hidden-import xlrd ^
    --hidden-import tkinter ^
    --hidden-import tkinter.ttk ^
    --hidden-import tkinter.messagebox ^
    --hidden-import tkinter.filedialog ^
    --hidden-import json ^
    --hidden-import pathlib ^
    --hidden-import shutil ^
    --collect-all pandas ^
    --collect-all openpyxl ^
    --collect-submodules pandas ^
    --collect-submodules openpyxl ^
    kiem_kho_app.py

if errorlevel 1 (
    echo.
    echo [ERROR] Cannot create .exe file!
    echo Please check the error messages above.
    pause
    exit /b 1
)

REM Copy Excel file to dist folder
if exist "DuLieuDauVao.xlsx" (
    if not exist "dist" mkdir dist
    copy "DuLieuDauVao.xlsx" "dist\" >nul
    echo [OK] Copied Excel file to dist folder
)

echo.
echo ========================================
echo    COMPLETE!
echo ========================================
echo.
echo Executable file created at: dist\KiemKhoApp.exe
echo.
echo This file has been built with all required libraries.
echo You can copy the .exe file to another Windows machine and run it.
echo.
pause

