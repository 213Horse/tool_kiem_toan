@echo off
chcp 65001 >nul
title Tạo file thực thi (.exe) cho Windows
color 0B

echo ========================================
echo    TẠO FILE THỰC THI (.EXE)
echo ========================================
echo.
echo File này sẽ tạo file .exe để chạy không cần Python
echo.

REM Kiểm tra Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python chưa được cài đặt!
    echo Vui lòng cài đặt Python trước.
    pause
    exit /b 1
)

echo [OK] Đã tìm thấy Python
echo.

REM Cài đặt các thư viện cần thiết trước khi build
echo [INFO] Đang cài đặt các thư viện cần thiết...
pip install --quiet --upgrade pip

REM Cài đặt từ requirements.txt nếu có
if exist "requirements.txt" (
    echo [INFO] Đang cài đặt từ requirements.txt...
    pip install --quiet -r requirements.txt
) else (
    echo [INFO] Đang cài đặt các thư viện trực tiếp...
    pip install --quiet pandas openpyxl xlrd pyinstaller
)

if %errorlevel% neq 0 (
    echo [ERROR] Không thể cài đặt thư viện!
    pause
    exit /b 1
)

echo [OK] Đã cài đặt xong các thư viện
echo.

REM Kiểm tra file Excel
if not exist "DuLieuDauVao.xlsx" (
    echo [WARNING] Không tìm thấy file DuLieuDauVao.xlsx
    echo File .exe vẫn sẽ được tạo nhưng cần file Excel để chạy!
    echo.
)

REM Xóa thư mục build và dist cũ (nếu có) để build lại sạch
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

REM Tạo file .exe
echo ========================================
echo    ĐANG TẠO FILE .EXE...
echo ========================================
echo.
echo Đang build, vui lòng đợi (có thể mất vài phút)...
echo.

pyinstaller --onefile --windowed --name "KiemKhoApp" ^
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
echo Để sử dụng trên máy Windows khác:
echo 1. Copy file: dist\KiemKhoApp.exe
echo 2. Copy file: dist\DuLieuDauVao.xlsx
echo 3. Đặt cả 2 file vào cùng một thư mục
echo 4. Double-click vào KiemKhoApp.exe để chạy
echo.
pause

