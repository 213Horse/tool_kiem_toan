@echo off
chcp 65001 >nul 2>&1
title Kiem Kho - Tu dong cai dat va chay
color 0A

echo ========================================
echo    KIEM KHO - TU DONG CAI DAT
echo ========================================
echo.

REM Kiem tra Python da cai chua
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python chua duoc cai dat!
    echo.
    echo Vui long tai va cai dat Python tu: https://www.python.org/downloads/
    echo Khi cai dat, nho chon "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

echo [OK] Da tim thay Python
python --version
echo.

REM Kiem tra pip
pip --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] pip chua duoc cai dat!
    pause
    exit /b 1
)

echo [OK] Da tim thay pip
echo.

REM Cai dat cac thu vien can thiet
echo [INFO] Dang cai dat cac thu vien can thiet...
echo.
pip install --quiet --upgrade pip
if errorlevel 1 (
    echo [ERROR] Khong the nang cap pip!
    pause
    exit /b 1
)

REM Cai dat tu requirements.txt neu co, neu khong thi cai truc tiep
if exist "requirements.txt" (
    echo [INFO] Dang cai dat tu requirements.txt...
    pip install --quiet -r requirements.txt
    set INSTALL_ERROR=%errorlevel%
) else (
    echo [INFO] Dang cai dat cac thu vien truc tiep...
    pip install --quiet pandas openpyxl xlrd
    set INSTALL_ERROR=%errorlevel%
)

REM Kiem tra tkinter
python -c "import tkinter" >nul 2>&1
if errorlevel 1 (
    echo [WARNING] Tkinter chua duoc cai dat. Dang thu cai dat...
    pip install --quiet tk
)

if %INSTALL_ERROR% neq 0 (
    echo [ERROR] Khong the cai dat thu vien!
    pause
    exit /b 1
)

echo [OK] Da cai dat xong cac thu vien
echo.

REM Kiem tra file Excel
if not exist "DuLieuDauVao.xlsx" (
    echo [WARNING] Khong tim thay file DuLieuDauVao.xlsx
    echo Vui long dam bao file nay co trong cung thu muc!
    echo.
)

REM Chay ung dung
echo ========================================
echo    DANG KHOI DONG UNG DUNG...
echo ========================================
echo.
python kiem_kho_app.py

if errorlevel 1 (
    echo.
    echo [ERROR] Ung dung gap loi khi chay!
    pause
    exit /b 1
)

pause

