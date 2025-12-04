#!/bin/bash

echo "========================================"
echo "   TAO FILE ZIP DE COPY SANG WINDOWS"
echo "========================================"
echo ""

# Tên file zip
ZIP_NAME="KIEM_KHO_FOR_WINDOWS.zip"
TEMP_DIR="KIEM_KHO_TEMP"

# Xóa file zip cũ nếu có
if [ -f "$ZIP_NAME" ]; then
    echo "[INFO] Xoa file zip cu..."
    rm -f "$ZIP_NAME"
fi

# Xóa thư mục temp cũ nếu có
if [ -d "$TEMP_DIR" ]; then
    echo "[INFO] Xoa thu muc temp cu..."
    rm -rf "$TEMP_DIR"
fi

# Tạo thư mục temp
mkdir -p "$TEMP_DIR"
echo "[OK] Da tao thu muc temp"
echo ""

# Copy các file cần thiết
echo "[INFO] Dang copy cac file can thiet..."

# File chính
if [ -f "kiem_kho_app.py" ]; then
    cp "kiem_kho_app.py" "$TEMP_DIR/"
    echo "[OK] Da copy: kiem_kho_app.py"
else
    echo "[ERROR] Khong tim thay: kiem_kho_app.py"
    exit 1
fi

# File Excel
if [ -f "DuLieuDauVao.xlsx" ]; then
    cp "DuLieuDauVao.xlsx" "$TEMP_DIR/"
    echo "[OK] Da copy: DuLieuDauVao.xlsx"
else
    echo "[WARNING] Khong tim thay: DuLieuDauVao.xlsx"
fi

# File template Excel
if [ -f "Kiemke_template.xlsx" ]; then
    cp "Kiemke_template.xlsx" "$TEMP_DIR/"
    echo "[OK] Da copy: Kiemke_template.xlsx"
else
    echo "[WARNING] Khong tim thay: Kiemke_template.xlsx"
fi

# File requirements
if [ -f "requirements.txt" ]; then
    cp "requirements.txt" "$TEMP_DIR/"
    echo "[OK] Da copy: requirements.txt"
else
    echo "[WARNING] Khong tim thay: requirements.txt"
fi

# File để chạy trực tiếp
if [ -f "CHAY_TREN_WINDOWS_SIMPLE.bat" ]; then
    cp "CHAY_TREN_WINDOWS_SIMPLE.bat" "$TEMP_DIR/"
    echo "[OK] Da copy: CHAY_TREN_WINDOWS_SIMPLE.bat"
fi

if [ -f "CHAY_TREN_WINDOWS.bat" ]; then
    cp "CHAY_TREN_WINDOWS.bat" "$TEMP_DIR/"
    echo "[OK] Da copy: CHAY_TREN_WINDOWS.bat"
fi

# File để build .exe
if [ -f "BUILD_EXE_SIMPLE.bat" ]; then
    cp "BUILD_EXE_SIMPLE.bat" "$TEMP_DIR/"
    echo "[OK] Da copy: BUILD_EXE_SIMPLE.bat"
fi

if [ -f "BUILD_EXE_DEBUG.bat" ]; then
    cp "BUILD_EXE_DEBUG.bat" "$TEMP_DIR/"
    echo "[OK] Da copy: BUILD_EXE_DEBUG.bat"
fi

if [ -f "SUA_LOI_BUILD.bat" ]; then
    cp "SUA_LOI_BUILD.bat" "$TEMP_DIR/"
    echo "[OK] Da copy: SUA_LOI_BUILD.bat"
fi

if [ -f "TAO_FILE_EXE_WINDOWS.bat" ]; then
    cp "TAO_FILE_EXE_WINDOWS.bat" "$TEMP_DIR/"
    echo "[OK] Da copy: TAO_FILE_EXE_WINDOWS.bat"
fi

if [ -f "build_windows.bat" ]; then
    cp "build_windows.bat" "$TEMP_DIR/"
    echo "[OK] Da copy: build_windows.bat"
fi

# File hướng dẫn
if [ -f "HUONG_DAN_BUILD_EXE_WINDOWS.txt" ]; then
    cp "HUONG_DAN_BUILD_EXE_WINDOWS.txt" "$TEMP_DIR/"
    echo "[OK] Da copy: HUONG_DAN_BUILD_EXE_WINDOWS.txt"
fi

if [ -f "DANH_SACH_FILE_COPY_WINDOWS.txt" ]; then
    cp "DANH_SACH_FILE_COPY_WINDOWS.txt" "$TEMP_DIR/"
    echo "[OK] Da copy: DANH_SACH_FILE_COPY_WINDOWS.txt"
fi

if [ -f "HUONG_DAN_SUA_LOI.txt" ]; then
    cp "HUONG_DAN_SUA_LOI.txt" "$TEMP_DIR/"
    echo "[OK] Da copy: HUONG_DAN_SUA_LOI.txt"
fi

if [ -f "HUONG_DAN_SUA_TRANG_MAN_HINH.txt" ]; then
    cp "HUONG_DAN_SUA_TRANG_MAN_HINH.txt" "$TEMP_DIR/"
    echo "[OK] Da copy: HUONG_DAN_SUA_TRANG_MAN_HINH.txt"
fi

if [ -f "HUONG_DAN_PORTABLE.txt" ]; then
    cp "HUONG_DAN_PORTABLE.txt" "$TEMP_DIR/"
    echo "[OK] Da copy: HUONG_DAN_PORTABLE.txt"
fi

if [ -f "XOA_FILE_SAU_BUILD.bat" ]; then
    cp "XOA_FILE_SAU_BUILD.bat" "$TEMP_DIR/"
    echo "[OK] Da copy: XOA_FILE_SAU_BUILD.bat"
fi

if [ -f "HUONG_DAN_XOA_FILE.txt" ]; then
    cp "HUONG_DAN_XOA_FILE.txt" "$TEMP_DIR/"
    echo "[OK] Da copy: HUONG_DAN_XOA_FILE.txt"
fi

# Tạo file README trong zip
cat > "$TEMP_DIR/README.txt" << 'EOF'
═══════════════════════════════════════════════════════════════
   HUONG DAN SU DUNG TREN WINDOWS
═══════════════════════════════════════════════════════════════

CACH 1: CHAY TRUC TIEP (Can Python)
───────────────────────────────────────────────────────────────
1. Giai nen file zip nay
2. Double-click vao: CHAY_TREN_WINDOWS_SIMPLE.bat
3. Ung dung se tu dong cai dat va chay


CACH 2: TAO FILE .EXE (Can Python de build)
───────────────────────────────────────────────────────────────
1. Giai nen file zip nay
2. Double-click vao: BUILD_EXE_SIMPLE.bat
3. Doi qua trinh build hoan tat (co the mat vai phut)
4. File .exe se nam trong: dist\KiemKhoApp.exe
5. Copy file .exe sang may Windows khac de chay


CAC FILE CAN THIET
───────────────────────────────────────────────────────────────
✓ kiem_kho_app.py          - File chinh cua ung dung
✓ DuLieuDauVao.xlsx        - File du lieu Excel (BAT BUOC)
✓ Kiemke_template.xlsx      - File template Excel (de copy khi save)
✓ requirements.txt         - Danh sach thu vien can thiet
✓ BUILD_EXE_SIMPLE.bat     - De build file .exe (khuyen nghi)
✓ CHAY_TREN_WINDOWS_SIMPLE.bat - De chay truc tiep (khuyen nghi)


LUU Y
───────────────────────────────────────────────────────────────
- Neu chay truc tiep: Can cai Python truoc
- Neu build .exe: Can Python de build, sau do khong can Python de chay
- File DuLieuDauVao.xlsx phai cung thu muc voi file .bat hoac .exe
- File Kiemke_template.xlsx se duoc copy va doi ten khi bam SAVE

═══════════════════════════════════════════════════════════════
EOF

echo "[OK] Da tao file README.txt"
echo ""

# Tạo file zip
echo "[INFO] Dang tao file zip..."
cd "$TEMP_DIR"
zip -r "../$ZIP_NAME" . > /dev/null
cd ..

if [ $? -eq 0 ]; then
    echo "[OK] Da tao file zip thanh cong!"
else
    echo "[ERROR] Khong the tao file zip!"
    rm -rf "$TEMP_DIR"
    exit 1
fi

# Xóa thư mục temp
rm -rf "$TEMP_DIR"

echo ""
echo "========================================"
echo "    HOAN TAT!"
echo "========================================"
echo ""
echo "File zip da duoc tao: $ZIP_NAME"
echo ""
echo "Ban co the copy file zip nay sang may Windows va giai nen."
echo "Trong file zip co day du cac file can thiet de:"
echo "- Chay truc tiep ung dung"
echo "- Build file .exe"
echo ""
echo "Xem file README.txt trong zip de biet cach su dung chi tiet."
echo ""

