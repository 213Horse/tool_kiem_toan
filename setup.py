#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script setup để cài đặt và build ứng dụng Kiểm Kho
"""

import subprocess
import sys
import os
from pathlib import Path

def install_requirements():
    """Cài đặt các thư viện cần thiết"""
    print("Đang cài đặt các thư viện cần thiết...")
    
    # Thử các cách cài đặt khác nhau
    methods = [
        # Cách 1: Cài vào user directory (an toàn nhất)
        [sys.executable, "-m", "pip", "install", "--user", "-r", "requirements.txt"],
        # Cách 2: Cài với break-system-packages (nếu cách 1 không được)
        [sys.executable, "-m", "pip", "install", "--break-system-packages", "-r", "requirements.txt"],
        # Cách 3: Cài bình thường
        [sys.executable, "-m", "pip", "install", "-r", "requirements.txt"],
    ]
    
    for method in methods:
        try:
            subprocess.check_call(method)
            print("✓ Đã cài đặt thành công các thư viện!")
            return True
        except subprocess.CalledProcessError:
            continue
    
    print("✗ Lỗi khi cài đặt thư viện!")
    print("\n💡 Gợi ý: Thử tạo virtual environment:")
    print("   python3 -m venv venv")
    print("   source venv/bin/activate")
    print("   pip install -r requirements.txt")
    return False

def build_executable():
    """Build file thực thi"""
    print("\nĐang build file thực thi...")
    try:
        if sys.platform == "win32":
            # Windows
            subprocess.check_call(["pyinstaller", "--onefile", "--windowed", 
                                 "--name", "KiemKhoApp", 
                                 "--add-data", "DuLieuDauVao.xlsx;.", 
                                 "kiem_kho_app.py"])
        else:
            # macOS/Linux
            subprocess.check_call(["pyinstaller", "--onefile", "--windowed", 
                                 "--name", "KiemKhoApp", 
                                 "--add-data", "DuLieuDauVao.xlsx:.", 
                                 "kiem_kho_app.py"])
        
        # Copy file Excel vào thư mục dist
        excel_file = Path("DuLieuDauVao.xlsx")
        dist_folder = Path("dist")
        if excel_file.exists() and dist_folder.exists():
            import shutil
            shutil.copy2(excel_file, dist_folder / excel_file.name)
            print(f"✓ Đã copy {excel_file.name} vào thư mục dist")
        
        print("\n✓ Build thành công!")
        print(f"File thực thi nằm trong thư mục: {dist_folder.absolute()}")
        return True
    except subprocess.CalledProcessError:
        print("✗ Lỗi khi build!")
        return False
    except FileNotFoundError:
        print("✗ Không tìm thấy pyinstaller. Vui lòng cài đặt: pip install pyinstaller")
        return False

def main():
    print("=" * 50)
    print("SETUP ỨNG DỤNG KIỂM KHO")
    print("=" * 50)
    
    # Kiểm tra file Excel
    if not Path("DuLieuDauVao.xlsx").exists():
        print("⚠ Cảnh báo: Không tìm thấy file DuLieuDauVao.xlsx")
        print("Vui lòng đảm bảo file này có trong thư mục hiện tại.")
        response = input("Tiếp tục? (y/n): ")
        if response.lower() != 'y':
            return
    
    # Cài đặt requirements
    if not install_requirements():
        return
    
    # Hỏi có muốn build không
    print("\n" + "=" * 50)
    response = input("Bạn có muốn build file thực thi ngay bây giờ? (y/n): ")
    if response.lower() == 'y':
        build_executable()
    else:
        print("\nĐể build sau, chạy:")
        if sys.platform == "win32":
            print("  build_windows.bat")
        else:
            print("  ./build_macos.sh")
    
    print("\n" + "=" * 50)
    print("Hoàn tất!")
    print("=" * 50)

if __name__ == "__main__":
    main()

