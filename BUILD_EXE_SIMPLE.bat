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

REM Tạo file config tạm thời TRƯỚC KHI build (sẽ được cập nhật sau)
echo # File này được tạo tự động khi build exe > dist_path_config.py
echo # Chứa đường dẫn thư mục dist gốc >> dist_path_config.py
echo DIST_PATH = None  # Sẽ được cập nhật sau khi build >> dist_path_config.py

pyinstaller --onefile --windowed --name "KiemKhoApp" ^
    --add-data "Kiemke_template.xlsx;." ^
    --add-data "dist_path_config.py;." ^
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

REM Cập nhật file config với đường dẫn chính xác sau khi build
REM File này đã được đóng gói vào exe, không cần copy vào thư mục dist
if exist "dist" (
    for %%I in (dist) do set DIST_PATH_FINAL=%%~fI
    echo # File này được tạo tự động khi build exe > dist_path_config.py
    echo # Chứa đường dẫn thư mục dist gốc >> dist_path_config.py
    echo DIST_PATH = r"%DIST_PATH_FINAL%" >> dist_path_config.py
    echo [OK] Updated dist_path_config.py with path: %DIST_PATH_FINAL%
    echo [INFO] Original dist path saved in exe: %DIST_PATH_FINAL%
    echo [INFO] Application will automatically find this path even if moved to another location
    echo [INFO] dist folder now contains only: KiemKhoApp.exe and DuLieuDauVao.xlsx
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

