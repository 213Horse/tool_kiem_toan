#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ứng dụng Kiểm Kho - Quét mã vạch để kiểm tra tồn kho thực tế
Chạy trên Windows và macOS
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
from pathlib import Path
import sys

class KiemKhoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kiểm Kho - Quét Mã Vạch")
        self.root.geometry("1200x700")
        # Màu nền nhẹ nhàng hơn
        self.root.configure(bg='#F5F5F5')
        
        # Biến lưu trữ dữ liệu
        self.df = None
        self.current_box_data = None
        self.current_box_number = None
        self.scanned_items = {}  # Lưu các item đã quét: {isbn: {tua, ton_thuc_te, so_thung, ton_trong_thung, ghi_chu}}
        self.edit_entry = None  # Entry widget để chỉnh sửa trực tiếp
        self.editing_item = None  # Item đang được chỉnh sửa
        
        # Load dữ liệu từ Excel
        self.load_data()
        
        # Tạo giao diện
        self.create_ui()
        
        # Bind Enter key để hỗ trợ quét mã vạch
        self.root.bind('<Return>', self.on_enter_pressed)
        
    def load_data(self):
        """Load dữ liệu từ file Excel"""
        try:
            # Kiểm tra nếu đang chạy từ executable (PyInstaller)
            if getattr(sys, 'frozen', False):
                # Chạy từ executable
                base_path = Path(sys._MEIPASS)
                excel_path = base_path / "DuLieuDauVao.xlsx"
                # Nếu không có trong bundle, tìm ở thư mục chứa executable
                if not excel_path.exists():
                    excel_path = Path(sys.executable).parent / "DuLieuDauVao.xlsx"
            else:
                # Chạy từ source code
                excel_path = Path(__file__).parent / "DuLieuDauVao.xlsx"
            
            # Danh sách file thay thế nếu file chính không đọc được
            xls_alternatives = [
                Path(__file__).parent / "KIEM KE Năm -2025 - BP ONLINE.xls",
                Path(__file__).parent / "KIEM KE Năm -2025 - BP ONLINE copy.xls",
                Path(__file__).parent / "DuLieuDauVao.xls",
            ]
            
            # Kiểm tra file .xlsx có hợp lệ không
            if excel_path.exists() and excel_path.suffix.lower() == '.xlsx':
                try:
                    import openpyxl
                    wb = openpyxl.load_workbook(excel_path, read_only=True)
                    if len(wb.sheetnames) == 0:
                        # File .xlsx không có worksheet, tìm file .xls thay thế
                        wb.close()
                        for alt_path in xls_alternatives:
                            if alt_path.exists():
                                excel_path = alt_path
                                break
                        else:
                            raise ValueError("File Excel không có worksheet nào và không tìm thấy file thay thế!")
                    else:
                        wb.close()
                except Exception as e:
                    # Nếu file .xlsx lỗi, thử tìm file .xls thay thế
                    for alt_path in xls_alternatives:
                        if alt_path.exists():
                            excel_path = alt_path
                            break
            
            # Nếu vẫn không tìm thấy file hợp lệ
            if not excel_path.exists():
                # Thử tìm file .xls
                for alt_path in xls_alternatives:
                    if alt_path.exists():
                        excel_path = alt_path
                        break
                
                # Nếu vẫn không có, cho phép người dùng chọn file
                if not excel_path.exists():
                    excel_path = filedialog.askopenfilename(
                        title="Chọn file dữ liệu Excel",
                        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
                    )
                    if not excel_path:
                        messagebox.showerror("Lỗi", "Không tìm thấy file dữ liệu!")
                        sys.exit(1)
                    excel_path = Path(excel_path)
            
            # Đọc file Excel (hỗ trợ cả .xls và .xlsx)
            # Ưu tiên thử đọc với header=0 (dòng 1) trước
            header_row = None
            
            if excel_path.suffix.lower() == '.xls':
                # Thử đọc với header=0 trước (dòng 1 trong Excel)
                try:
                    temp_df = pd.read_excel(excel_path, engine='xlrd', header=0, nrows=2)
                    temp_df.columns = temp_df.columns.str.strip()
                    # Kiểm tra xem có đủ cột cần thiết không
                    cols_lower = [str(c).lower() if pd.notna(c) else '' for c in temp_df.columns]
                    cols_str = ' '.join(cols_lower)
                    keywords_found = sum([
                        'isbn' in cols_str,
                        'số thùng' in cols_str or 'so thung' in cols_str or 'thùng' in cols_str,
                        'tựa' in cols_str or 'tua' in cols_str or 'titles' in cols_str,
                        'tồn' in cols_str or 'ton' in cols_str or 'qty' in cols_str
                    ])
                    if keywords_found >= 3:  # Nếu tìm thấy ít nhất 3 từ khóa, dùng header=0
                        header_row = 0
                except:
                    pass
                
                # Nếu header=0 không hợp lệ, tìm header tự động
                if header_row is None:
                    try:
                        temp_df = pd.read_excel(excel_path, engine='xlrd', header=None, nrows=30)
                        for idx, row in temp_df.iterrows():
                            row_values = [str(cell).lower() if pd.notna(cell) else '' for cell in row.values]
                            row_str = ' '.join(row_values)
                            if 'isbn' in row_str:
                                keywords_found = sum([
                                    'isbn' in row_str,
                                    'số thùng' in row_str or 'so thung' in row_str or 'thùng' in row_str,
                                    'tựa' in row_str or 'tua' in row_str or 'titles' in row_str,
                                    'tồn' in row_str or 'ton' in row_str or 'qty' in row_str
                                ])
                                if keywords_found >= 2:
                                    header_row = idx
                                    break
                    except:
                        pass
                
                # Đọc file với header row đã tìm được
                try:
                    if header_row is not None:
                        self.df = pd.read_excel(excel_path, engine='xlrd', header=header_row)
                    else:
                        # Mặc định dùng header=0
                        self.df = pd.read_excel(excel_path, engine='xlrd', header=0)
                except Exception as e2:
                    raise ValueError(f"Không thể đọc file .xls: {str(e2)}")
            else:
                # File .xlsx - thử đọc với header=0 trước
                try:
                    self.df = pd.read_excel(excel_path, engine='openpyxl', header=0)
                except Exception as e2:
                    raise ValueError(f"Không thể đọc file .xlsx: {str(e2)}")
            
            # Loại bỏ các dòng rỗng hoàn toàn
            self.df = self.df.dropna(how='all')
            
            # Kiểm tra DataFrame có rỗng không (trước khi filter ISBN)
            if self.df.empty:
                raise ValueError("File Excel không có dữ liệu!")
            
            # Loại bỏ các dòng có ISBN rỗng hoặc không hợp lệ (chỉ khi đã có cột ISBN)
            if 'isbn' in self.df.columns or 'ISBN' in self.df.columns:
                isbn_col = 'isbn' if 'isbn' in self.df.columns else 'ISBN'
                original_count = len(self.df)
                
                # Loại bỏ các dòng có ISBN rỗng
                self.df = self.df[self.df[isbn_col].notna()]
                
                # Loại bỏ các dòng có ISBN là số thứ tự đơn giản (1, 2, 3...) hoặc số nhỏ hơn 4 chữ số
                if len(self.df) > 0:
                    # Chuyển ISBN sang string để kiểm tra
                    isbn_str = self.df[isbn_col].astype(str)
                    # Loại bỏ các ISBN chỉ là số đơn giản (1.0, 2.0, ...) hoặc có ít hơn 4 ký tự số
                    mask = ~isbn_str.str.match(r'^\d+\.0?$')  # Không phải chỉ số đơn giản
                    # Giữ lại các ISBN có độ dài hợp lý (ít nhất 4 ký tự sau khi loại bỏ .0)
                    mask = mask & (isbn_str.str.replace('.0', '').str.len() >= 4)
                    self.df = self.df[mask]
                
                # Kiểm tra lại sau khi filter
                if self.df.empty:
                    raise ValueError(f"File Excel không có dữ liệu hợp lệ! Đã loại bỏ {original_count} dòng không hợp lệ.")
            
            # Kiểm tra lại DataFrame có rỗng không (sau khi filter)
            if self.df.empty:
                raise ValueError("File Excel không có dữ liệu sau khi lọc!")
            
            # Xử lý DataFrame
            self._process_dataframe()
            
        except Exception as e:
            error_msg = f"Không thể đọc file Excel: {str(e)}\n\n"
            error_msg += "Vui lòng kiểm tra:\n"
            error_msg += "1. File Excel có đúng định dạng không (.xlsx hoặc .xls)\n"
            error_msg += "2. File có chứa dữ liệu không\n"
            error_msg += "3. File không bị hỏng\n\n"
            error_msg += "Bạn có muốn chọn file khác không?"
            
            result = messagebox.askyesno("Lỗi", error_msg)
            if result:
                # Cho phép chọn file khác
                excel_path = filedialog.askopenfilename(
                    title="Chọn file dữ liệu Excel",
                    filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
                )
                if excel_path:
                    try:
                        # Thử đọc lại với file mới
                        if Path(excel_path).suffix.lower() == '.xls':
                            self.df = pd.read_excel(excel_path, engine='xlrd')
                        else:
                            self.df = pd.read_excel(excel_path, engine='openpyxl')
                        # Nếu thành công, tiếp tục xử lý
                        self._process_dataframe()
                        return
                    except Exception as e2:
                        messagebox.showerror("Lỗi", f"Vẫn không thể đọc file: {str(e2)}")
                        sys.exit(1)
                else:
                    messagebox.showerror("Lỗi", "Không có file nào được chọn!")
                    sys.exit(1)
            else:
                sys.exit(1)
    
    def _process_dataframe(self):
        """Xử lý DataFrame sau khi đọc thành công"""
        # Chuẩn hóa tên cột (loại bỏ khoảng trắng thừa, chuyển sang lowercase)
        self.df.columns = self.df.columns.str.strip()
        
        # Tìm các cột cần thiết (case-insensitive)
        col_mapping = {}
        for col in self.df.columns:
            if pd.isna(col):
                continue
            col_str = str(col).strip()
            col_lower = col_str.lower()
            
            # Số thùng / Thùng
            if ('số thùng' in col_lower or 'so thung' in col_lower or 
                col_lower == 'thùng' or col_lower == 'thung'):
                if 'so_thung' not in col_mapping:
                    col_mapping['so_thung'] = col
            
            # ISBN
            if 'isbn' in col_lower:
                if 'isbn' not in col_mapping:
                    col_mapping['isbn'] = col
            
            # Tựa/Tên sách / Titles
            if ('tựa' in col_lower or 'tua' in col_lower or 'tên' in col_lower or 
                'titles' in col_lower or 'title' in col_lower):
                if 'tua' not in col_mapping:
                    col_mapping['tua'] = col
            
            # Tồn từng tựa / Qty tựa trong thùng / Qty
            # Ưu tiên các cột có tên đầy đủ trước
            if 'qty tựa trong thùng' in col_lower or 'qty tua trong thung' in col_lower:
                if 'ton_tung_tua' not in col_mapping:
                    col_mapping['ton_tung_tua'] = col
            elif (('tồn' in col_lower and 'tựa' in col_lower) or 
                  ('ton' in col_lower and 'tua' in col_lower)):
                if 'ton_tung_tua' not in col_mapping:
                    col_mapping['ton_tung_tua'] = col
            elif col_lower == 'qty' and 'ton_tung_tua' not in col_mapping:
                # Chỉ dùng Qty nếu chưa tìm thấy cột nào khác
                col_mapping['ton_tung_tua'] = col
        
        # Kiểm tra xem có đủ cột không
        if len(col_mapping) < 4:
            messagebox.showwarning("Cảnh báo", 
                f"Không tìm thấy đủ các cột cần thiết. Tìm thấy: {list(col_mapping.keys())}\n"
                f"Các cột trong file: {list(self.df.columns)}")
        
        # Đổi tên cột để dễ sử dụng
        if col_mapping:
            self.df = self.df.rename(columns=col_mapping)
        
        # Làm sạch dữ liệu
        if 'isbn' in self.df.columns:
            self.df['isbn'] = self.df['isbn'].astype(str).str.strip()
    
    def create_ui(self):
        """Tạo giao diện người dùng"""
        # Màu sắc
        bg_color = '#F5F5F5'  # Nền chính nhẹ nhàng
        input_bg = '#FFFFFF'  # Nền input trắng
        input_bg_yellow = '#FFF9C4'  # Nền vàng nhẹ nhàng hơn
        label_fg = '#333333'  # Màu chữ đen nhẹ
        label_required_fg = '#C62828'  # Đỏ đậm cho label bắt buộc
        button_bg = '#E3F2FD'  # Nền button xanh nhẹ
        
        # Frame chính
        main_frame = tk.Frame(self.root, bg=bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # === PHẦN NHẬP THÔNG TIN THÙNG ===
        info_frame = tk.Frame(main_frame, bg=bg_color)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Màu chữ trong input - đen đậm để dễ nhìn
        input_fg = '#000000'  # Đen đậm
        
        # Nhập/Xuất
        tk.Label(info_frame, text="Nhập/Xuất:", bg=bg_color, fg=label_fg, font=('Arial', 11)).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.nhap_xuat_var = tk.StringVar(value="KIEM KE Năm 2025 - BP Online")
        entry1 = tk.Entry(info_frame, textvariable=self.nhap_xuat_var, width=40, bg=input_bg_yellow, 
                          fg=input_fg, font=('Arial', 10), relief=tk.SOLID, bd=1, insertbackground='#000000')
        entry1.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky='ew')
        
        # Thùng / vị trí mới
        tk.Label(info_frame, text="Thùng / vị trí mới:", bg=bg_color, fg=label_fg, font=('Arial', 11)).grid(row=0, column=3, sticky='w', padx=5, pady=5)
        self.vi_tri_moi_var = tk.StringVar()
        entry2 = tk.Entry(info_frame, textvariable=self.vi_tri_moi_var, width=20, bg=input_bg_yellow,
                          fg=input_fg, font=('Arial', 10), relief=tk.SOLID, bd=1, insertbackground='#000000')
        entry2.grid(row=0, column=4, padx=5, pady=5, sticky='ew')
        
        # Số thùng (*)
        tk.Label(info_frame, text="Số thùng (*):", bg=bg_color, fg=label_required_fg, font=('Arial', 11, 'bold')).grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.so_thung_var = tk.StringVar()
        self.so_thung_entry = tk.Entry(info_frame, textvariable=self.so_thung_var, width=40, bg=input_bg_yellow,
                                       fg=input_fg, font=('Arial', 11), relief=tk.SOLID, bd=1, insertbackground='#000000')
        self.so_thung_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        self.so_thung_entry.bind('<Return>', lambda e: self.load_box_data())
        load_btn = tk.Button(info_frame, text="Load", command=self.load_box_data, bg=button_bg, 
                            fg='#1976D2', font=('Arial', 10, 'bold'), width=10, relief=tk.RAISED, bd=1)
        load_btn.grid(row=1, column=2, padx=5, pady=5)
        
        # Ngày (*)
        tk.Label(info_frame, text="Ngày (*):", bg=bg_color, fg=label_required_fg, font=('Arial', 11, 'bold')).grid(row=2, column=0, sticky='w', padx=5, pady=5)
        from datetime import datetime
        self.ngay_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%y"))
        entry3 = tk.Entry(info_frame, textvariable=self.ngay_var, width=40, bg=input_bg_yellow,
                         fg=input_fg, font=('Arial', 10), relief=tk.SOLID, bd=1, insertbackground='#000000')
        entry3.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky='ew')
        
        # Tổ + Nhân Viên
        tk.Label(info_frame, text="Tổ + Nhân Viên:", bg=bg_color, fg=label_fg, font=('Arial', 11)).grid(row=3, column=0, sticky='w', padx=5, pady=5)
        self.nhan_vien_var = tk.StringVar()
        entry4 = tk.Entry(info_frame, textvariable=self.nhan_vien_var, width=40, bg=input_bg_yellow,
                          fg=input_fg, font=('Arial', 10), relief=tk.SOLID, bd=1, insertbackground='#000000')
        entry4.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky='ew')
        
        # Note thùng
        tk.Label(info_frame, text="Note thùng:", bg=bg_color, fg=label_required_fg, font=('Arial', 11, 'bold')).grid(row=4, column=0, sticky='w', padx=5, pady=5)
        self.note_thung_var = tk.StringVar()
        entry5 = tk.Entry(info_frame, textvariable=self.note_thung_var, width=40, bg=input_bg_yellow,
                         fg=input_fg, font=('Arial', 10), relief=tk.SOLID, bd=1, insertbackground='#000000')
        entry5.grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky='ew')
        
        # Nút SAVE
        save_btn = tk.Button(info_frame, text="SAVE", command=self.save_data, 
                            bg='#4CAF50', fg='white', font=('Arial', 12, 'bold'), 
                            width=15, height=2, relief=tk.RAISED, bd=2, cursor='hand2')
        save_btn.grid(row=1, column=3, rowspan=2, padx=20, pady=5, sticky='n')
        
        info_frame.columnconfigure(1, weight=1)
        
        # === PHẦN HIỂN THỊ SỐ TỰA ===
        count_frame = tk.Frame(main_frame, bg=bg_color)
        count_frame.pack(fill=tk.X, pady=(0, 5))
        
        tk.Label(count_frame, text="Số tựa:", bg=bg_color, fg=label_required_fg, font=('Arial', 11, 'bold')).grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.so_tua_var = tk.StringVar(value="0")
        tk.Label(count_frame, textvariable=self.so_tua_var, bg=bg_color, fg='#1976D2', font=('Arial', 14, 'bold')).grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        # === BẢNG DỮ LIỆU ===
        table_frame = tk.Frame(main_frame, bg=bg_color)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tạo Treeview với scrollbar
        scrollbar_y = tk.Scrollbar(table_frame, orient=tk.VERTICAL, bg='#E0E0E0', troughcolor=bg_color)
        scrollbar_x = tk.Scrollbar(table_frame, orient=tk.HORIZONTAL, bg='#E0E0E0', troughcolor=bg_color)
        
        columns = ('ISBN', 'Tựa', 'Tồn thực tế', 'Số thùng', 'Tồn tựa trong thùng', 'Ghi chú')
        
        # Tạo style cho Treeview
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Treeview', background='#FFFFFF', foreground='#333333', 
                       fieldbackground='#FFFFFF', font=('Arial', 10), rowheight=25)
        style.configure('Treeview.Heading', background='#2196F3', foreground='white', 
                       font=('Arial', 10, 'bold'), relief=tk.FLAT)
        style.map('Treeview.Heading', background=[('active', '#1976D2')])
        
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', 
                                 yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set,
                                 height=15, style='Treeview')
        
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)
        
        # Định nghĩa các cột
        self.tree.heading('ISBN', text='ISBN')
        self.tree.heading('Tựa', text='Tựa')
        self.tree.heading('Tồn thực tế', text='Tồn thực tế')
        self.tree.heading('Số thùng', text='Số thùng')
        self.tree.heading('Tồn tựa trong thùng', text='Tồn tựa trong thùng')
        self.tree.heading('Ghi chú', text='Ghi chú')
        
        self.tree.column('ISBN', width=150, anchor='w')
        self.tree.column('Tựa', width=300, anchor='w')
        self.tree.column('Tồn thực tế', width=120, anchor='center')
        self.tree.column('Số thùng', width=100, anchor='center')
        self.tree.column('Tồn tựa trong thùng', width=150, anchor='center')
        self.tree.column('Ghi chú', width=200, anchor='w')
        
        # Tag để highlight màu đỏ (nhẹ nhàng hơn)
        self.tree.tag_configure('error', background='#FFEBEE', foreground='#C62828')
        
        self.tree.grid(row=0, column=0, sticky='nsew')
        scrollbar_y.grid(row=0, column=1, sticky='ns')
        scrollbar_x.grid(row=1, column=0, sticky='ew')
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Bind click để edit trực tiếp tất cả các cột có thể edit
        self.tree.bind('<Button-1>', self.on_item_click)
        self.tree.bind('<Double-1>', self.on_item_double_click)  # Double click cho Ghi chú (để nhập dài)
        
        # === PHẦN NHẬP ISBN (QUÉT MÃ VẠCH) ===
        scan_frame = tk.Frame(main_frame, bg=bg_color)
        scan_frame.pack(fill=tk.X, pady=(10, 0))
        
        tk.Label(scan_frame, text="Quét/Nhập ISBN:", bg=bg_color, fg=label_fg, 
                font=('Arial', 11, 'bold')).grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.isbn_entry = tk.Entry(scan_frame, font=('Arial', 12), width=30, 
                                   bg='#FFFFFF', fg='#000000', relief=tk.SOLID, bd=2, insertbackground='#000000')
        self.isbn_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self.isbn_entry.bind('<Return>', self.on_isbn_entered)
        self.isbn_entry.focus()
        
        scan_frame.columnconfigure(1, weight=1)
    
    def load_box_data(self):
        """Load dữ liệu của thùng được nhập"""
        so_thung = self.so_thung_var.get().strip()
        if not so_thung:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập số thùng!")
            return
        
        try:
            # Tìm cột số thùng (có thể là 'so_thung' sau khi mapping hoặc tên gốc)
            so_thung_col = None
            for col in self.df.columns:
                col_lower = str(col).lower().strip()
                if 'số thùng' in col_lower or 'so thung' in col_lower or col_lower == 'thùng' or col_lower == 'so_thung':
                    so_thung_col = col
                    break
            
            if so_thung_col is None:
                messagebox.showerror("Lỗi", f"Không tìm thấy cột 'Số thùng' trong file Excel!\nCác cột có sẵn: {list(self.df.columns)}")
                return
            
            # Chuyển đổi sang string để so sánh
            self.df[so_thung_col] = self.df[so_thung_col].astype(str).str.strip()
            self.current_box_data = self.df[self.df[so_thung_col] == so_thung].copy()
            
            if self.current_box_data.empty:
                messagebox.showinfo("Thông báo", f"Không tìm thấy dữ liệu cho thùng số {so_thung}")
                self.current_box_number = None
                self.so_tua_var.set("0")
                self.clear_table()
                return
            
            self.current_box_number = so_thung
            self.scanned_items = {}  # Reset danh sách đã quét
            
            # Hiển thị số tựa
            self.so_tua_var.set(str(len(self.current_box_data)))
            
            # Clear bảng
            self.clear_table()
            
            # Thông báo thành công
            messagebox.showinfo("Thành công", f"Đã load {len(self.current_box_data)} tựa cho thùng số {so_thung}")
            
            # Focus vào ô nhập ISBN để sẵn sàng quét
            self.isbn_entry.focus()
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể load dữ liệu thùng: {str(e)}")
    
    def on_isbn_entered(self, event=None):
        """Xử lý khi nhập/quét ISBN"""
        isbn = self.isbn_entry.get().strip()
        if not isbn:
            return
        
        if self.current_box_data is None or self.current_box_data.empty:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập số thùng và load dữ liệu trước!")
            self.isbn_entry.delete(0, tk.END)
            return
        
        # Tìm tựa trong dữ liệu thùng hiện tại
        if 'isbn' in self.current_box_data.columns:
            # Tìm ISBN (có thể không khớp hoàn toàn do format)
            isbn_clean = str(isbn).strip()
            matched_row = None
            
            # Tìm ISBN với nhiều cách khớp
            for idx, row in self.current_box_data.iterrows():
                row_isbn = str(row.get('isbn', '')).strip()
                # Loại bỏ các ký tự đặc biệt để so sánh
                row_isbn_clean = ''.join(filter(str.isdigit, row_isbn))
                isbn_clean_digits = ''.join(filter(str.isdigit, isbn_clean))
                
                # Khớp nếu:
                # 1. Hoàn toàn giống nhau
                # 2. Một trong hai kết thúc bằng cái kia
                # 3. Chỉ số (digits) giống nhau
                if (row_isbn == isbn_clean or 
                    row_isbn.endswith(isbn_clean) or 
                    isbn_clean.endswith(row_isbn) or
                    (row_isbn_clean and isbn_clean_digits and row_isbn_clean == isbn_clean_digits)):
                    matched_row = row
                    break
            
            if matched_row is None:
                messagebox.showwarning("Cảnh báo", f"Không tìm thấy ISBN {isbn} trong thùng số {self.current_box_number}")
                self.isbn_entry.delete(0, tk.END)
                return
            
            # Lấy thông tin từ matched_row
            # Tìm cột 'tua' (có thể là 'tua' sau khi mapping hoặc tên gốc)
            tua = ''
            for col in matched_row.index:
                col_lower = str(col).lower().strip()
                if 'tựa' in col_lower or 'tua' in col_lower or 'tên' in col_lower or 'titles' in col_lower:
                    tua = str(matched_row[col]) if pd.notna(matched_row[col]) else ''
                    break
            
            # Tìm cột 'ton_tung_tua' (có thể là 'ton_tung_tua' sau khi mapping hoặc tên gốc)
            ton_trong_thung = 0
            for col in matched_row.index:
                col_lower = str(col).lower().strip()
                if ('tồn' in col_lower and 'tựa' in col_lower) or ('ton' in col_lower and 'tua' in col_lower) or 'qty tựa trong thùng' in col_lower or 'qty tua trong thung' in col_lower:
                    ton_trong_thung = matched_row[col] if pd.notna(matched_row[col]) else 0
                    try:
                        ton_trong_thung = int(float(ton_trong_thung))  # Chuyển thành số nguyên
                    except:
                        ton_trong_thung = 0
                    break
            
            so_thung = self.current_box_number
            
            # Kiểm tra xem đã quét chưa
            if isbn_clean in self.scanned_items:
                # Update item đã tồn tại
                item_id = self.scanned_items[isbn_clean]['item_id']
                self.tree.delete(item_id)
                del self.scanned_items[isbn_clean]
            
            # Thêm vào bảng với đầy đủ thông tin
            # Chuyển ton_trong_thung thành số nguyên để hiển thị
            ton_trong_thung_display = int(ton_trong_thung) if ton_trong_thung else 0
            
            item_id = self.tree.insert('', tk.END, values=(
                isbn_clean,
                tua,  # Tựa - tự động điền từ dữ liệu
                '',  # Tồn thực tế - để trống để người dùng nhập
                so_thung,
                ton_trong_thung_display,  # Tồn tựa trong thùng - hiển thị số nguyên
                ''   # Ghi chú - để trống
            ), tags=('',))
            
            # Lưu thông tin
            self.scanned_items[isbn_clean] = {
                'item_id': item_id,
                'tua': tua,
                'ton_thuc_te': '',
                'so_thung': so_thung,
                'ton_trong_thung': ton_trong_thung,
                'ghi_chu': ''
            }
            
            # Focus vào cột "Tồn thực tế" để người dùng nhập
            self.tree.selection_set(item_id)
            self.tree.focus(item_id)
            self.tree.see(item_id)
            
            # Tự động focus vào ô "Tồn thực tế" để sẵn sàng nhập
            self.root.after(100, lambda: self.auto_edit_ton_thuc_te(item_id))
            
        else:
            messagebox.showerror("Lỗi", "Không tìm thấy cột 'ISBN' trong dữ liệu!")
        
        # Clear ô nhập ISBN để sẵn sàng quét tiếp
        self.isbn_entry.delete(0, tk.END)
        self.isbn_entry.focus()
    
    def on_item_click(self, event):
        """Xử lý click để edit trực tiếp các cột có thể chỉnh sửa"""
        # Hủy edit cũ nếu có
        if self.edit_entry:
            self.finish_edit()
        
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1
        
        # Cho phép edit: Tồn thực tế (2), Số thùng (3), Tồn tựa trong thùng (4), Ghi chú (5)
        # Không cho edit: ISBN (0), Tựa (1) - chỉ đọc
        if column_index not in [2, 3, 4, 5]:
            return
        
        if not item:
            return
        
        # Lấy giá trị hiện tại
        values = list(self.tree.item(item, 'values'))
        current_value = values[column_index] if column_index < len(values) else ''
        
        # Lấy vị trí của cell
        bbox = self.tree.bbox(item, column)
        if not bbox:
            return
        
        x, y, width, height = bbox
        
        # Tạo Entry widget để edit trực tiếp
        self.edit_entry = tk.Entry(self.tree, font=('Arial', 10), 
                                   relief=tk.FLAT, bd=0, bg='#FFFFFF', fg='#000000')
        self.edit_entry.insert(0, str(current_value))
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        self.edit_entry.focus()
        self.editing_item = item
        self.editing_column = column_index
        
        def finish_on_enter(event):
            self.finish_edit()
        
        def finish_on_focus_out(event):
            # Delay một chút để tránh conflict với click events
            self.root.after(100, self.finish_edit)
        
        self.edit_entry.bind('<Return>', finish_on_enter)
        self.edit_entry.bind('<FocusOut>', finish_on_focus_out)
        self.edit_entry.bind('<Escape>', lambda e: self.cancel_edit())
    
    def finish_edit(self):
        """Hoàn tất việc chỉnh sửa"""
        if not self.edit_entry or not self.editing_item:
            return
        
        new_value = self.edit_entry.get().strip()
        item = self.editing_item
        column_index = self.editing_column
        
        # Lấy giá trị hiện tại
        values = list(self.tree.item(item, 'values'))
        isbn = values[0]
        
        # Xử lý theo từng cột
        if column_index == 2:  # Tồn thực tế
            values[column_index] = new_value
            
            # Kiểm tra và highlight nếu khác nhau
            if isbn in self.scanned_items:
                self.scanned_items[isbn]['ton_thuc_te'] = new_value
                ton_trong_thung = self.scanned_items[isbn]['ton_trong_thung']
                
                try:
                    ton_thuc_te_num = float(new_value) if new_value else 0
                    ton_trong_thung_num = float(ton_trong_thung) if ton_trong_thung else 0
                    
                    # Kiểm tra lệch
                    if abs(ton_thuc_te_num - ton_trong_thung_num) > 0.01:
                        # Tô đỏ và tự động ghi chú lỗi
                        self.tree.item(item, tags=('error',))
                        
                        # Tự động ghi chú lỗi vào cột Ghi chú
                        if ton_thuc_te_num < ton_trong_thung_num:
                            error_note = f"Thiếu {int(ton_trong_thung_num - ton_thuc_te_num)} cuốn"
                        else:
                            error_note = f"Thừa {int(ton_thuc_te_num - ton_trong_thung_num)} cuốn"
                        
                        # Cập nhật Ghi chú
                        if len(values) > 5:
                            values[5] = error_note
                        else:
                            values.append(error_note)
                        
                        # Cập nhật scanned_items
                        self.scanned_items[isbn]['ghi_chu'] = error_note
                        
                        # Hiển thị cảnh báo
                        messagebox.showwarning("Cảnh báo", 
                            f"Tồn thực tế ({int(ton_thuc_te_num)}) khác với tồn trong thùng ({int(ton_trong_thung_num)})!\n"
                            f"Đã tự động ghi chú: {error_note}")
                    else:
                        self.tree.item(item, tags=('',))
                        # Xóa ghi chú lỗi nếu đã khớp
                        if len(values) > 5 and ('Thiếu' in str(values[5]) or 'Thừa' in str(values[5])):
                            values[5] = ''
                            self.scanned_items[isbn]['ghi_chu'] = ''
                except (ValueError, TypeError):
                    # Nếu không phải số, không highlight
                    self.tree.item(item, tags=('',))
        
        elif column_index == 3:  # Số thùng
            values[column_index] = new_value
            if isbn in self.scanned_items:
                self.scanned_items[isbn]['so_thung'] = new_value
        
        elif column_index == 4:  # Tồn tựa trong thùng
            # Chuyển thành số nguyên
            try:
                new_value_int = int(float(new_value)) if new_value else 0
                values[column_index] = new_value_int
                if isbn in self.scanned_items:
                    self.scanned_items[isbn]['ton_trong_thung'] = new_value_int
            except:
                values[column_index] = new_value
        
        elif column_index == 5:  # Ghi chú
            values[column_index] = new_value
            if isbn in self.scanned_items:
                self.scanned_items[isbn]['ghi_chu'] = new_value
        
        # Cập nhật tree với giá trị mới
        self.tree.item(item, values=values)
        
        # Xóa Entry widget
        self.edit_entry.destroy()
        self.edit_entry = None
        self.editing_item = None
        self.isbn_entry.focus()
    
    def cancel_edit(self):
        """Hủy việc chỉnh sửa"""
        if self.edit_entry:
            self.edit_entry.destroy()
            self.edit_entry = None
            self.editing_item = None
        self.isbn_entry.focus()
    
    def auto_edit_ton_thuc_te(self, item_id):
        """Tự động mở edit cho cột 'Tồn thực tế' sau khi thêm item mới"""
        if not item_id:
            return
        
        # Lấy vị trí của cell "Tồn thực tế" (column index 2)
        column = '#3'  # Column index 2 (0-indexed) + 1
        bbox = self.tree.bbox(item_id, column)
        if not bbox:
            return
        
        x, y, width, height = bbox
        
        # Lấy giá trị hiện tại
        values = list(self.tree.item(item_id, 'values'))
        current_value = values[2] if len(values) > 2 else ''
        
        # Tạo Entry widget để edit trực tiếp
        self.edit_entry = tk.Entry(self.tree, font=('Arial', 10), 
                                   relief=tk.FLAT, bd=0, bg='#FFFFFF', fg='#000000')
        self.edit_entry.insert(0, str(current_value))
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        self.edit_entry.focus()
        self.editing_item = item_id
        self.editing_column = 2
        
        def finish_on_enter(event):
            self.finish_edit()
        
        def finish_on_focus_out(event):
            self.root.after(100, self.finish_edit)
        
        self.edit_entry.bind('<Return>', finish_on_enter)
        self.edit_entry.bind('<FocusOut>', finish_on_focus_out)
        self.edit_entry.bind('<Escape>', lambda e: self.cancel_edit())
    
    def on_item_double_click(self, event):
        """Xử lý double click để edit cột 'Ghi chú'"""
        # Hủy edit cũ nếu có
        if self.edit_entry:
            self.finish_edit()
        
        item = self.tree.selection()[0] if self.tree.selection() else None
        if not item:
            return
        
        column = self.tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1
        
        # Chỉ cho phép edit cột "Ghi chú" (index 5) bằng double click
        if column_index != 5:
            return
        
        # Lấy giá trị hiện tại
        values = list(self.tree.item(item, 'values'))
        current_value = values[column_index] if column_index < len(values) else ''
        
        # Tạo popup để edit
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Chỉnh sửa Ghi chú")
        edit_window.geometry("400x120")
        edit_window.transient(self.root)
        edit_window.grab_set()
        
        tk.Label(edit_window, text="Nhập Ghi chú:", font=('Arial', 11)).pack(pady=10)
        
        entry_var = tk.StringVar(value=str(current_value))
        entry = tk.Entry(edit_window, textvariable=entry_var, width=40, font=('Arial', 11))
        entry.pack(pady=5)
        entry.focus()
        entry.select_range(0, tk.END)
        
        def save_edit():
            new_value = entry_var.get().strip()
            values[column_index] = new_value
            
            # Cập nhật tree
            self.tree.item(item, values=values)
            
            # Cập nhật scanned_items
            isbn = values[0]
            if isbn in self.scanned_items:
                self.scanned_items[isbn]['ghi_chu'] = new_value
            
            edit_window.destroy()
            self.isbn_entry.focus()
        
        entry.bind('<Return>', lambda e: save_edit())
        tk.Button(edit_window, text="Lưu", command=save_edit, width=10).pack(pady=5)
    
    def clear_table(self):
        """Xóa tất cả items trong bảng"""
        for item in self.tree.get_children():
            self.tree.delete(item)
    
    def on_enter_pressed(self, event):
        """Xử lý phím Enter"""
        widget = self.root.focus_get()
        if widget == self.isbn_entry:
            self.on_isbn_entered()
    
    def save_data(self):
        """Lưu dữ liệu đã kiểm tra"""
        if not self.scanned_items:
            messagebox.showwarning("Cảnh báo", "Chưa có dữ liệu nào để lưu!")
            return
        
        # Tạo DataFrame từ scanned_items
        save_data = []
        for isbn, info in self.scanned_items.items():
            save_data.append({
                'ISBN': isbn,
                'Tựa': info['tua'],
                'Tồn thực tế': info['ton_thuc_te'],
                'Số thùng': info['so_thung'],
                'Tồn tựa trong thùng': info['ton_trong_thung'],
                'Ghi chú': info['ghi_chu']
            })
        
        df_save = pd.DataFrame(save_data)
        
        # Cho phép người dùng chọn nơi lưu
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"KiemKho_Thung_{self.current_box_number}_{self.ngay_var.get().replace('/', '_')}.xlsx"
        )
        
        if filename:
            try:
                df_save.to_excel(filename, index=False)
                messagebox.showinfo("Thành công", f"Đã lưu dữ liệu vào {filename}")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể lưu file: {str(e)}")

def main():
    root = tk.Tk()
    app = KiemKhoApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

