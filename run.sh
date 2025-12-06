#!/bin/bash
# Script chạy ứng dụng (tự động kích hoạt venv nếu có)

if [ -d "venv" ]; then
    echo "Đang kích hoạt virtual environment..."
    source venv/bin/activate
    python kiem_kho_app.py
else
    echo "Không tìm thấy virtual environment. Chạy trực tiếp với Python..."
    python3 kiem_kho_app.py
fi





