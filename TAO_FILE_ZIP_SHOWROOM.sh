#!/bin/bash

echo "========================================"
echo "   TAO FILE ZIP SHOWROOM CHO WINDOWS"
echo "========================================"
echo ""

# Tên file zip
ZIP_NAME="KIEM_KHO_SHOWROOM_FOR_WINDOWS.zip"
TEMP_DIR="KIEM_KHO_SHOWROOM_TEMP"

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
if [ -f "kiem_kho_showroom.py" ]; then
    cp "kiem_kho_showroom.py" "$TEMP_DIR/"
    echo "[OK] Da copy: kiem_kho_showroom.py"
else
    echo "[ERROR] Khong tim thay: kiem_kho_showroom.py"
    exit 1
fi

# File Excel Showroom
if [ -f "DuLieuDauVaoShowroom.xlsx" ]; then
    cp "DuLieuDauVaoShowroom.xlsx" "$TEMP_DIR/"
    echo "[OK] Da copy: DuLieuDauVaoShowroom.xlsx"
else
    echo "[WARNING] Khong tim thay: DuLieuDauVaoShowroom.xlsx"
fi

# File template Excel (nếu có)
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
if [ -f "CHAY_TREN_WINDOWS_SHOWROOM_SIMPLE.bat" ]; then
    cp "CHAY_TREN_WINDOWS_SHOWROOM_SIMPLE.bat" "$TEMP_DIR/"
    echo "[OK] Da copy: CHAY_TREN_WINDOWS_SHOWROOM_SIMPLE.bat"
fi

if [ -f "CHAY_TREN_WINDOWS_SHOWROOM.bat" ]; then
    cp "CHAY_TREN_WINDOWS_SHOWROOM.bat" "$TEMP_DIR/"
    echo "[OK] Da copy: CHAY_TREN_WINDOWS_SHOWROOM.bat"
fi

# File để build .exe
if [ -f "BUILD_EXE_SHOWROOM_SIMPLE.bat" ]; then
    cp "BUILD_EXE_SHOWROOM_SIMPLE.bat" "$TEMP_DIR/"
    echo "[OK] Da copy: BUILD_EXE_SHOWROOM_SIMPLE.bat"
fi

if [ -f "BUILD_EXE_SHOWROOM_DEBUG.bat" ]; then
    cp "BUILD_EXE_SHOWROOM_DEBUG.bat" "$TEMP_DIR/"
    echo "[OK] Da copy: BUILD_EXE_SHOWROOM_DEBUG.bat"
fi

if [ -f "build_windows_showroom.bat" ]; then
    cp "build_windows_showroom.bat" "$TEMP_DIR/"
    echo "[OK] Da copy: build_windows_showroom.bat"
fi

# File hướng dẫn
if [ -f "HUONG_DAN_SHOWROOM.txt" ]; then
    cp "HUONG_DAN_SHOWROOM.txt" "$TEMP_DIR/"
    echo "[OK] Da copy: HUONG_DAN_SHOWROOM.txt"
fi

# Tạo file README trong zip
cat > "$TEMP_DIR/README.txt" << 'EOF'
═══════════════════════════════════════════════════════════════
   HUONG DAN SU DUNG KIEM KHO SHOWROOM TREN WINDOWS
═══════════════════════════════════════════════════════════════

CACH 1: CHAY TRUC TIEP (Can Python)
───────────────────────────────────────────────────────────────
1. Giai nen file zip nay
2. Double-click vao: CHAY_TREN_WINDOWS_SHOWROOM_SIMPLE.bat
3. Ung dung se tu dong cai dat va chay


CACH 2: TAO FILE .EXE (Can Python de build)
───────────────────────────────────────────────────────────────
1. Giai nen file zip nay
2. Double-click vao: BUILD_EXE_SHOWROOM_SIMPLE.bat
3. Doi qua trinh build hoan tat (co the mat vai phut)
4. File .exe se nam trong: dist\KiemKhoShowroomApp.exe
5. Copy file .exe sang may Windows khac de chay


CAC FILE CAN THIET
───────────────────────────────────────────────────────────────
✓ kiem_kho_showroom.py          - File chinh cua ung dung Showroom
✓ DuLieuDauVaoShowroom.xlsx     - File du lieu Excel Showroom (BAT BUOC)
✓ Kiemke_template.xlsx          - File template Excel (de copy khi save)
✓ requirements.txt              - Danh sach thu vien can thiet
✓ BUILD_EXE_SHOWROOM_SIMPLE.bat - De build file .exe (khuyen nghi)
✓ CHAY_TREN_WINDOWS_SHOWROOM_SIMPLE.bat - De chay truc tiep (khuyen nghi)


LUU Y
───────────────────────────────────────────────────────────────
- Neu chay truc tiep: Can cai Python truoc
- Neu build .exe: Can Python de build, sau do khong can Python de chay
- File DuLieuDauVaoShowroom.xlsx phai cung thu muc voi file .bat hoac .exe
- File Kiemke_template.xlsx se duoc copy va doi ten khi bam SAVE
- Day la phien ban Showroom, khac voi phien ban chinh

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
echo "- Chay truc tiep ung dung Showroom"
echo "- Build file .exe"
echo ""
echo "Xem file README.txt trong zip de biet cach su dung chi tiet."
echo ""


