@echo off
chcp 65001 >nul
echo ========================================
echo    XOA CAC FILE KHONG CAN THIET
echo    Sau khi build file .exe thanh cong
echo ========================================
echo.

REM Kiểm tra xem file .exe đã được build chưa
if not exist "dist\KiemKhoApp.exe" (
    echo [LOI] Khong tim thay file dist\KiemKhoApp.exe
    echo Vui long build file .exe truoc khi chay script nay!
    echo.
    pause
    exit /b 1
)

echo [INFO] File .exe da duoc tao thanh cong!
echo.
echo Cac file se bi xoa:
echo - kiem_kho_app.py
echo - requirements.txt
echo - BUILD_EXE_SIMPLE.bat
echo - BUILD_EXE_DEBUG.bat
echo - SUA_LOI_BUILD.bat
echo - TAO_FILE_EXE_WINDOWS.bat
echo - build_windows.bat
echo - CHAY_TREN_WINDOWS.bat
echo - CHAY_TREN_WINDOWS_SIMPLE.bat
echo - HUONG_DAN_BUILD_EXE_WINDOWS.txt
echo - HUONG_DAN_SUA_LOI.txt
echo - HUONG_DAN_SUA_TRANG_MAN_HINH.txt
echo - DANH_SACH_FILE_COPY_WINDOWS.txt
echo - HUONG_DAN_PORTABLE.txt
echo - README.txt (neu co)
echo.
echo Cac file se GIU LAI:
echo - dist\KiemKhoApp.exe (file ung dung chinh)
echo - DuLieuDauVao.xlsx (file du lieu - CAN THIET)
echo - Kiemke_template.xlsx (file template - CAN THIET)
echo - kiem_kho_config.json (file cau hinh - se duoc tao khi chay app)
echo.
set /p confirm="Ban co muon xoa cac file khong can thiet? (Y/N): "

if /i "%confirm%" neq "Y" (
    echo Da huy!
    pause
    exit /b 0
)

echo.
echo [INFO] Dang xoa cac file...

REM Xóa các file Python source
if exist "kiem_kho_app.py" (
    del /q "kiem_kho_app.py"
    echo [OK] Da xoa: kiem_kho_app.py
)

REM Xóa file requirements
if exist "requirements.txt" (
    del /q "requirements.txt"
    echo [OK] Da xoa: requirements.txt
)

REM Xóa các file batch script build
if exist "BUILD_EXE_SIMPLE.bat" (
    del /q "BUILD_EXE_SIMPLE.bat"
    echo [OK] Da xoa: BUILD_EXE_SIMPLE.bat
)

if exist "BUILD_EXE_DEBUG.bat" (
    del /q "BUILD_EXE_DEBUG.bat"
    echo [OK] Da xoa: BUILD_EXE_DEBUG.bat
)

if exist "SUA_LOI_BUILD.bat" (
    del /q "SUA_LOI_BUILD.bat"
    echo [OK] Da xoa: SUA_LOI_BUILD.bat
)

if exist "TAO_FILE_EXE_WINDOWS.bat" (
    del /q "TAO_FILE_EXE_WINDOWS.bat"
    echo [OK] Da xoa: TAO_FILE_EXE_WINDOWS.bat
)

if exist "build_windows.bat" (
    del /q "build_windows.bat"
    echo [OK] Da xoa: build_windows.bat
)

REM Xóa các file batch script chạy
if exist "CHAY_TREN_WINDOWS.bat" (
    del /q "CHAY_TREN_WINDOWS.bat"
    echo [OK] Da xoa: CHAY_TREN_WINDOWS.bat
)

if exist "CHAY_TREN_WINDOWS_SIMPLE.bat" (
    del /q "CHAY_TREN_WINDOWS_SIMPLE.bat"
    echo [OK] Da xoa: CHAY_TREN_WINDOWS_SIMPLE.bat
)

REM Xóa các file hướng dẫn
if exist "HUONG_DAN_BUILD_EXE_WINDOWS.txt" (
    del /q "HUONG_DAN_BUILD_EXE_WINDOWS.txt"
    echo [OK] Da xoa: HUONG_DAN_BUILD_EXE_WINDOWS.txt
)

if exist "HUONG_DAN_SUA_LOI.txt" (
    del /q "HUONG_DAN_SUA_LOI.txt"
    echo [OK] Da xoa: HUONG_DAN_SUA_LOI.txt
)

if exist "HUONG_DAN_SUA_TRANG_MAN_HINH.txt" (
    del /q "HUONG_DAN_SUA_TRANG_MAN_HINH.txt"
    echo [OK] Da xoa: HUONG_DAN_SUA_TRANG_MAN_HINH.txt
)

if exist "DANH_SACH_FILE_COPY_WINDOWS.txt" (
    del /q "DANH_SACH_FILE_COPY_WINDOWS.txt"
    echo [OK] Da xoa: DANH_SACH_FILE_COPY_WINDOWS.txt
)

if exist "HUONG_DAN_PORTABLE.txt" (
    del /q "HUONG_DAN_PORTABLE.txt"
    echo [OK] Da xoa: HUONG_DAN_PORTABLE.txt
)

if exist "README.txt" (
    del /q "README.txt"
    echo [OK] Da xoa: README.txt
)

REM Xóa thư mục build và spec nếu có
if exist "build" (
    rmdir /s /q "build"
    echo [OK] Da xoa thu muc: build
)

if exist "KiemKhoApp.spec" (
    del /q "KiemKhoApp.spec"
    echo [OK] Da xoa: KiemKhoApp.spec
)

echo.
echo ========================================
echo    HOAN TAT!
echo ========================================
echo.
echo Cac file con lai:
echo - dist\KiemKhoApp.exe (file ung dung)
echo - DuLieuDauVao.xlsx (file du lieu)
echo - Kiemke_template.xlsx (file template)
echo.
echo Ban co the copy file .exe ra bat ky dau de su dung!
echo.
pause


