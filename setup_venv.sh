#!/bin/bash
# Script setup với virtual environment (khuyến nghị cho macOS)

echo "═══════════════════════════════════════════════════════════════"
echo "        SETUP ỨNG DỤNG KIỂM KHO VỚI VIRTUAL ENVIRONMENT"
echo "═══════════════════════════════════════════════════════════════"
echo ""

# Kiểm tra Python
if ! command -v python3 &> /dev/null; then
    echo "✗ Không tìm thấy Python3. Vui lòng cài đặt Python trước."
    exit 1
fi

echo "✓ Tìm thấy Python: $(python3 --version)"
echo ""

# Kiểm tra tkinter
echo "🔍 Đang kiểm tra tkinter..."
python3 -c "import tkinter" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "⚠ Không tìm thấy tkinter!"
    echo ""
    echo "Đang cài đặt python-tk..."
    if command -v brew &> /dev/null; then
        brew install python-tk
        echo "✓ Đã cài đặt python-tk"
    else
        echo "✗ Không tìm thấy Homebrew. Vui lòng cài đặt python-tk thủ công:"
        echo "   brew install python-tk"
        exit 1
    fi
    echo ""
fi
echo "✓ Tkinter đã sẵn sàng"
echo ""

# Tạo virtual environment
echo "📦 Đang tạo virtual environment..."
python3 -m venv venv

if [ $? -ne 0 ]; then
    echo "✗ Lỗi khi tạo virtual environment!"
    exit 1
fi

echo "✓ Đã tạo virtual environment"
echo ""

# Kích hoạt virtual environment và cài đặt
echo "📥 Đang cài đặt các thư viện..."
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

if [ $? -ne 0 ]; then
    echo "✗ Lỗi khi cài đặt thư viện!"
    deactivate
    exit 1
fi

echo ""
echo "✓ Đã cài đặt thành công!"
echo ""
echo "═══════════════════════════════════════════════════════════════"
echo "🚀 ĐỂ CHẠY ỨNG DỤNG:"
echo ""
echo "   source venv/bin/activate"
echo "   python kiem_kho_app.py"
echo ""
echo "Hoặc chạy trực tiếp:"
echo "   ./venv/bin/python kiem_kho_app.py"
echo ""
echo "═══════════════════════════════════════════════════════════════"

