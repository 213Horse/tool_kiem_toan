@echo off
chcp 65001 >nul
echo Building Windows executable...
echo.

REM Kiểm tra và cài đặt thư viện nếu cần
if exist "requirements.txt" (
    echo Installing required packages...
    pip install --quiet -r requirements.txt
) else (
    echo Installing packages...
    pip install --quiet pandas openpyxl xlrd pyinstaller
)

echo.
echo Building executable with all dependencies...
echo.

REM Xóa build cũ nếu có
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "KiemKhoApp.spec" del /q KiemKhoApp.spec

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

if %errorlevel% neq 0 (
    echo.
    echo Build failed! Check the error messages above.
    pause
    exit /b 1
)

REM Copy file Excel vào thư mục dist
if exist "DuLieuDauVao.xlsx" (
    if not exist "dist" mkdir dist
    copy "DuLieuDauVao.xlsx" "dist\" >nul
)

echo.
echo Build complete! Executable is in the 'dist' folder.
echo Copy DuLieuDauVao.xlsx to the same folder as KiemKhoApp.exe
echo.
pause

