@echo off
chcp 65001 >nul
title Sửa lỗi build file .exe
color 0E

echo ========================================
echo    SỬA LỖI BUILD FILE .EXE
echo ========================================
echo.
echo File này sẽ build lại file .exe với đầy đủ thư viện
echo để khắc phục lỗi "No module named 'pandas'"
echo.

REM Kiểm tra Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python chưa được cài đặt!
    echo Vui lòng cài đặt Python từ: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [OK] Đã tìm thấy Python
python --version
echo.

REM Cài đặt các thư viện cần thiết
echo [INFO] Đang cài đặt các thư viện cần thiết...
pip install --quiet --upgrade pip

if exist "requirements.txt" (
    echo [INFO] Đang cài đặt từ requirements.txt...
    pip install --quiet -r requirements.txt
    set INSTALL_ERR=%errorlevel%
) else (
    echo [INFO] Đang cài đặt các thư viện trực tiếp...
    pip install --quiet pandas openpyxl xlrd pyinstaller
    set INSTALL_ERR=%errorlevel%
)

if %INSTALL_ERR% neq 0 (
    echo [ERROR] Không thể cài đặt thư viện!
    pause
    exit /b 1
)

echo [OK] Đã cài đặt xong các thư viện
echo.

REM Xóa thư mục build và dist cũ
if exist "build" (
    echo [INFO] Đang xóa thư mục build cũ...
    rmdir /s /q build
)
if exist "dist" (
    echo [INFO] Đang xóa thư mục dist cũ...
    rmdir /s /q dist
)
if exist "KiemKhoApp.spec" (
    echo [INFO] Đang xóa file spec cũ...
    del /q KiemKhoApp.spec
)

echo.

REM Tạo file .exe với đầy đủ thư viện
echo ========================================
echo    ĐANG BUILD FILE .EXE...
echo ========================================
echo.
echo Đang build, vui lòng đợi (có thể mất vài phút)...
echo.

pyinstaller --onefile --windowed --name "KiemKhoApp" ^
    --add-data "DuLieuDauVao.xlsx;." ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --hidden-import xlrd ^
    --hidden-import tkinter ^
    --hidden-import tkinter.ttk ^
    --hidden-import tkinter.messagebox ^
    --hidden-import tkinter.filedialog ^
    --hidden-import json ^
    --hidden-import pathlib ^
    --collect-all pandas ^
    --collect-all openpyxl ^
    --collect-submodules pandas ^
    --collect-submodules openpyxl ^
    kiem_kho_app.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Không thể tạo file .exe!
    echo Vui lòng kiểm tra lỗi ở trên.
    pause
    exit /b 1
)

REM Copy file Excel vào thư mục dist
if exist "DuLieuDauVao.xlsx" (
    if not exist "dist" mkdir dist
    copy "DuLieuDauVao.xlsx" "dist\" >nul
    echo [OK] Đã copy file Excel vào thư mục dist
)

echo.
echo ========================================
echo    HOÀN TẤT!
echo ========================================
echo.
echo File thực thi đã được tạo tại: dist\KiemKhoApp.exe
echo.
echo File này đã được build với đầy đủ thư viện.
echo Bạn có thể copy file .exe sang máy Windows khác và chạy.
echo.
pause

