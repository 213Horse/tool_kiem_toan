@echo off
chcp 65001 >nul
title Copy file cần thiết cho Windows
color 0B

echo ========================================
echo    COPY FILE CẦN THIẾT CHO WINDOWS
echo ========================================
echo.
echo File này sẽ tạo thư mục chứa các file cần thiết
echo để copy sang máy Windows khác
echo.

REM Tạo thư mục để copy
set "COPY_FOLDER=COPY_TO_WINDOWS"
if exist "%COPY_FOLDER%" (
    echo [INFO] Đang xóa thư mục cũ...
    rmdir /s /q "%COPY_FOLDER%"
)

mkdir "%COPY_FOLDER%"
echo [OK] Đã tạo thư mục: %COPY_FOLDER%
echo.

REM Copy các file cần thiết
echo [INFO] Đang copy các file cần thiết...

REM File chính
if exist "kiem_kho_app.py" (
    copy "kiem_kho_app.py" "%COPY_FOLDER%\" >nul
    echo [OK] Đã copy: kiem_kho_app.py
) else (
    echo [ERROR] Không tìm thấy: kiem_kho_app.py
)

REM File Excel
if exist "DuLieuDauVao.xlsx" (
    copy "DuLieuDauVao.xlsx" "%COPY_FOLDER%\" >nul
    echo [OK] Đã copy: DuLieuDauVao.xlsx
) else (
    echo [WARNING] Không tìm thấy: DuLieuDauVao.xlsx
)

REM File batch để chạy
if exist "CHAY_TREN_WINDOWS.bat" (
    copy "CHAY_TREN_WINDOWS.bat" "%COPY_FOLDER%\" >nul
    echo [OK] Đã copy: CHAY_TREN_WINDOWS.bat
) else (
    echo [WARNING] Không tìm thấy: CHAY_TREN_WINDOWS.bat
)

REM File requirements
if exist "requirements.txt" (
    copy "requirements.txt" "%COPY_FOLDER%\" >nul
    echo [OK] Đã copy: requirements.txt
) else (
    echo [WARNING] Không tìm thấy: requirements.txt
)

REM File để build .exe
if exist "SUA_LOI_BUILD.bat" (
    copy "SUA_LOI_BUILD.bat" "%COPY_FOLDER%\" >nul
    echo [OK] Đã copy: SUA_LOI_BUILD.bat
)

if exist "TAO_FILE_EXE_WINDOWS.bat" (
    copy "TAO_FILE_EXE_WINDOWS.bat" "%COPY_FOLDER%\" >nul
    echo [OK] Đã copy: TAO_FILE_EXE_WINDOWS.bat
)

if exist "build_windows.bat" (
    copy "build_windows.bat" "%COPY_FOLDER%\" >nul
    echo [OK] Đã copy: build_windows.bat
)

REM Copy file hướng dẫn
if exist "DANH_SACH_FILE_COPY_WINDOWS.txt" (
    copy "DANH_SACH_FILE_COPY_WINDOWS.txt" "%COPY_FOLDER%\" >nul
    echo [OK] Đã copy: DANH_SACH_FILE_COPY_WINDOWS.txt
)

echo.
echo ========================================
echo    HOÀN TẤT!
echo ========================================
echo.
echo Đã tạo thư mục: %COPY_FOLDER%
echo.
echo Bạn có thể copy toàn bộ thư mục này sang máy Windows.
echo.
echo Trong thư mục có:
echo - Các file để chạy trực tiếp với Python
echo - Các file để build file .exe
echo - File hướng dẫn chi tiết
echo.
pause




