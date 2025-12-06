@echo off
chcp 65001 >nul
title Copy file de build .exe
color 0B

echo ========================================
echo    COPY FILE DE BUILD .EXE
echo ========================================
echo.
echo File nay se tao thu muc chua cac file can thiet
echo de build file .exe tren may Windows co Python
echo.

REM Tao thu muc de copy
set "COPY_FOLDER=COPY_TO_BUILD_EXE"
if exist "%COPY_FOLDER%" (
    echo [INFO] Dang xoa thu muc cu...
    rmdir /s /q "%COPY_FOLDER%"
)

mkdir "%COPY_FOLDER%"
echo [OK] Da tao thu muc: %COPY_FOLDER%
echo.

REM Copy cac file can thiet
echo [INFO] Dang copy cac file can thiet...

REM File chinh
if exist "kiem_kho_app.py" (
    copy "kiem_kho_app.py" "%COPY_FOLDER%\" >nul
    echo [OK] Da copy: kiem_kho_app.py
) else (
    echo [ERROR] Khong tim thay: kiem_kho_app.py
)

REM File Excel
if exist "DuLieuDauVao.xlsx" (
    copy "DuLieuDauVao.xlsx" "%COPY_FOLDER%\" >nul
    echo [OK] Da copy: DuLieuDauVao.xlsx
) else (
    echo [WARNING] Khong tim thay: DuLieuDauVao.xlsx
)

REM File requirements
if exist "requirements.txt" (
    copy "requirements.txt" "%COPY_FOLDER%\" >nul
    echo [OK] Da copy: requirements.txt
) else (
    echo [WARNING] Khong tim thay: requirements.txt
)

REM File de build .exe
if exist "SUA_LOI_BUILD.bat" (
    copy "SUA_LOI_BUILD.bat" "%COPY_FOLDER%\" >nul
    echo [OK] Da copy: SUA_LOI_BUILD.bat
)

if exist "TAO_FILE_EXE_WINDOWS.bat" (
    copy "TAO_FILE_EXE_WINDOWS.bat" "%COPY_FOLDER%\" >nul
    echo [OK] Da copy: TAO_FILE_EXE_WINDOWS.bat
)

if exist "build_windows.bat" (
    copy "build_windows.bat" "%COPY_FOLDER%\" >nul
    echo [OK] Da copy: build_windows.bat
)

REM Copy file huong dan
if exist "HUONG_DAN_BUILD_EXE_WINDOWS.txt" (
    copy "HUONG_DAN_BUILD_EXE_WINDOWS.txt" "%COPY_FOLDER%\" >nul
    echo [OK] Da copy: HUONG_DAN_BUILD_EXE_WINDOWS.txt
)

echo.
echo ========================================
echo    HOAN TAT!
echo ========================================
echo.
echo Da tao thu muc: %COPY_FOLDER%
echo.
echo Ban co the copy toan bo thu muc nay sang may Windows co Python.
echo Sau do chay SUA_LOI_BUILD.bat de build file .exe
echo.
echo Sau khi build xong, copy 2 file sau sang may Windows khac:
echo - dist\KiemKhoApp.exe
echo - dist\DuLieuDauVao.xlsx
echo.
pause





