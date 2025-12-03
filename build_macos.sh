#!/bin/bash
echo "Building macOS executable..."
echo ""
echo "Make sure you have installed PyInstaller: pip install pyinstaller"
echo ""
pyinstaller --onefile --windowed --name "KiemKhoApp" --add-data "DuLieuDauVao.xlsx:." kiem_kho_app.py
echo ""
echo "Build complete! Executable is in the 'dist' folder."
echo "Copy DuLieuDauVao.xlsx to the same folder as KiemKhoApp"
echo "You can double-click 'KiemKhoApp' to run the application."

