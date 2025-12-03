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
import json
import shutil

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
        self.error_highlights = {}  # Lưu các highlight widgets: {item_id: [entry1, entry2]}
        self.template_file_path = None  # Đường dẫn file Excel cố định (template)
        self.auto_save_folder = None  # Thư mục tự động lưu file Excel 2 (Kiemkecuoinam)
        self.config_file = self.get_config_file_path()  # Đường dẫn file config
        self.tong_hop_data = []  # Lưu tổng hợp các data đã kiểm kê
        self.notebook = None  # Notebook widget để chứa các tab
        self.tong_hop_tree = None  # Treeview trong tab Tổng hợp
        
        # Load cấu hình từ file (nếu có)
        saved_config = self.load_config()
        need_setup = False
        
        if saved_config:
            template_path = saved_config.get('template_file_path')
            auto_folder = saved_config.get('auto_save_folder')
            
            # Kiểm tra cả 2 đường dẫn có tồn tại không
            if template_path and auto_folder:
                if os.path.isfile(template_path) and os.path.isdir(auto_folder):
                    # Cả 2 đường dẫn đều hợp lệ
                    self.template_file_path = template_path
                    self.auto_save_folder = auto_folder
                else:
                    # Một trong hai hoặc cả hai đường dẫn không tồn tại
                    need_setup = True
            else:
                # Thiếu một trong hai đường dẫn
                need_setup = True
        else:
            # Không có cấu hình
            need_setup = True
        
        # Nếu cần cấu hình, hiển thị dialog
        if need_setup:
            # Đảm bảo root window được hiển thị trước khi tạo dialog
            self.root.deiconify()
            self.root.update()
            self.setup_paths()
            
            # Sau khi setup xong, kiểm tra lại xem có cấu hình hợp lệ không
            if not self.template_file_path or not self.auto_save_folder:
                # Nếu vẫn không có cấu hình hợp lệ, đóng app
                self.root.quit()
                return
        
        # Load dữ liệu từ Excel
        self.load_data()
        
        # Tạo giao diện
        self.create_ui()
        
        # Bind Enter key để hỗ trợ quét mã vạch
        self.root.bind('<Return>', self.on_enter_pressed)
    
    def get_config_file_path(self):
        """Lấy đường dẫn file config"""
        if getattr(sys, 'frozen', False):
            # Chạy từ executable
            config_dir = Path(sys.executable).parent
        else:
            # Chạy từ source code
            config_dir = Path(__file__).parent
        
        return config_dir / "kiem_kho_config.json"
    
    def load_config(self):
        """Load cấu hình từ file"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # Format mới: có cả template_file_path và auto_save_folder
                    if 'template_file_path' in config and 'auto_save_folder' in config:
                        return {
                            'template_file_path': config['template_file_path'],
                            'auto_save_folder': config['auto_save_folder']
                        }
                    # Format cũ - hỗ trợ backward compatibility
                    elif 'auto_save_folder' in config:
                        return {
                            'template_file_path': None,
                            'auto_save_folder': config['auto_save_folder']
                        }
        except Exception as e:
            print(f"Lỗi khi đọc config: {str(e)}")
        return None
    
    def save_config(self, template_path, folder_path):
        """Lưu cấu hình vào file"""
        try:
            config = {
                'template_file_path': template_path,
                'auto_save_folder': folder_path
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Lỗi khi lưu config: {str(e)}")
    
    def setup_paths(self):
        """Hiển thị dialog để cấu hình 2 đường dẫn"""
        try:
            # Đảm bảo root window được hiển thị và update trước
            self.root.deiconify()
            self.root.update()
            self.root.update_idletasks()
        except Exception as e:
            print(f"Lỗi khi hiển thị root window: {str(e)}")
        
        try:
            dialog = tk.Toplevel(self.root)
            dialog.title("Cấu hình đường dẫn")
            dialog.geometry("700x380")
            dialog.configure(bg='#F5F5F5')
            dialog.transient(self.root)
            dialog.grab_set()  # Modal dialog
            
            # Đặt cửa sổ ở giữa màn hình
            dialog.update_idletasks()
            x = (dialog.winfo_screenwidth() // 2) - (700 // 2)
            y = (dialog.winfo_screenheight() // 2) - (380 // 2)
            dialog.geometry(f"700x380+{x}+{y}")
            
            # Đảm bảo dialog được focus và hiển thị trên cùng
            dialog.lift()
            dialog.focus_force()
            dialog.attributes('-topmost', True)  # Luôn hiển thị trên cùng
            dialog.update()
        except Exception as e:
            print(f"Lỗi khi tạo dialog: {str(e)}")
            import traceback
            traceback.print_exc()
            # Nếu không thể tạo dialog, hiển thị messagebox thay thế
            try:
                messagebox.showerror("Lỗi", f"Không thể hiển thị dialog cấu hình:\n{str(e)}")
            except:
                pass
            return
        
        # Label hướng dẫn
        label_text = "Cấu hình 2 đường dẫn:\n1. File Excel cố định (để copy khi SAVE)\n2. Thư mục lưu file bí mật (Kiemkecuoinam)"
        tk.Label(dialog, text=label_text, bg='#F5F5F5', fg='#000000', 
                font=('Arial', 11), justify=tk.LEFT, wraplength=650).pack(pady=15, padx=20)
        
        # Load giá trị đã lưu nếu có
        saved_config = self.load_config()
        
        # Đường dẫn 1: File Excel cố định
        tk.Label(dialog, text="1. File Excel cố định (template):", bg='#F5F5F5', fg='#000000', 
                font=('Arial', 10, 'bold'), anchor='w').pack(pady=(10, 5), padx=20, fill=tk.X)
        
        input_frame1 = tk.Frame(dialog, bg='#F5F5F5')
        input_frame1.pack(pady=5, padx=20, fill=tk.X)
        
        path1_var = tk.StringVar()
        if saved_config and saved_config.get('template_file_path'):
            path1_var.set(saved_config['template_file_path'])
        
        path1_entry = tk.Entry(input_frame1, textvariable=path1_var, width=50, 
                              font=('Arial', 10), relief=tk.SOLID, bd=1,
                              bg='#FFFFFF', fg='#000000', insertbackground='#000000')
        path1_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        def browse_file1():
            file_path = filedialog.askopenfilename(
                title="Chọn file Excel cố định",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            if file_path:
                path1_var.set(file_path)
        
        browse1_btn = tk.Button(input_frame1, text="Chọn file", command=browse_file1,
                               bg='#C8E6C9', fg='#000000', font=('Arial', 9, 'bold'), 
                               relief=tk.RAISED, bd=2, padx=12, pady=4,
                               activebackground='#A5D6A7', activeforeground='#000000',
                               cursor='hand2')
        browse1_btn.pack(side=tk.RIGHT)
        
        # Đường dẫn 2: Thư mục lưu file bí mật
        tk.Label(dialog, text="2. Thư mục lưu file bí mật (Kiemkecuoinam):", bg='#F5F5F5', fg='#000000', 
                font=('Arial', 10, 'bold'), anchor='w').pack(pady=(15, 5), padx=20, fill=tk.X)
        
        input_frame2 = tk.Frame(dialog, bg='#F5F5F5')
        input_frame2.pack(pady=5, padx=20, fill=tk.X)
        
        path2_var = tk.StringVar()
        if saved_config and saved_config.get('auto_save_folder'):
            path2_var.set(saved_config['auto_save_folder'])
        
        path2_entry = tk.Entry(input_frame2, textvariable=path2_var, width=50, 
                              font=('Arial', 10), relief=tk.SOLID, bd=1,
                              bg='#FFFFFF', fg='#000000', insertbackground='#000000')
        path2_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        def browse_folder2():
            folder = filedialog.askdirectory(title="Chọn thư mục lưu file bí mật")
            if folder:
                path2_var.set(folder)
        
        browse2_btn = tk.Button(input_frame2, text="Chọn thư mục", command=browse_folder2,
                               bg='#C8E6C9', fg='#000000', font=('Arial', 9, 'bold'), 
                               relief=tk.RAISED, bd=2, padx=12, pady=4,
                               activebackground='#A5D6A7', activeforeground='#000000',
                               cursor='hand2')
        browse2_btn.pack(side=tk.RIGHT)
        
        # Button OK và Cancel
        button_frame = tk.Frame(dialog, bg='#F5F5F5')
        button_frame.pack(pady=20)
        
        def on_ok():
            template_path = path1_var.get().strip()
            folder_path = path2_var.get().strip()
            
            if not template_path or not folder_path:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập đầy đủ 2 đường dẫn!")
                return
            
            # Kiểm tra file Excel có tồn tại không
            if not os.path.isfile(template_path):
                messagebox.showerror("Lỗi", f"File Excel không tồn tại!\n{template_path}")
                return
            
            # Kiểm tra thư mục có tồn tại không
            if not os.path.isdir(folder_path):
                messagebox.showerror("Lỗi", f"Thư mục không tồn tại!\n{folder_path}")
                return
            
            # Kiểm tra quyền ghi vào thư mục
            try:
                test_file = os.path.join(folder_path, '.kiem_kho_test')
                with open(test_file, 'w') as f:
                    f.write('test')
                os.remove(test_file)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không có quyền ghi vào thư mục!\n{str(e)}")
                return
            
            # Lưu cấu hình
            self.template_file_path = template_path
            self.auto_save_folder = folder_path
            self.save_config(template_path, folder_path)
            dialog.destroy()
        
        def on_cancel():
            # Nếu không nhập đầy đủ và bấm "Bỏ qua", tắt phần mềm
            template_path = path1_var.get().strip()
            folder_path = path2_var.get().strip()
            if not template_path or not folder_path:
                messagebox.showinfo("Thông báo", "Phần mềm sẽ đóng vì chưa cấu hình đầy đủ đường dẫn.")
                self.root.quit()
                sys.exit(0)
            else:
                dialog.destroy()
                self.root.quit()
                sys.exit(0)
        
        ok_btn = tk.Button(button_frame, text="OK", command=on_ok,
                          bg='#BBDEFB', fg='#000000', font=('Arial', 11, 'bold'),
                          relief=tk.RAISED, bd=2, padx=25, pady=8,
                          activebackground='#90CAF9', activeforeground='#000000',
                          cursor='hand2')
        ok_btn.pack(side=tk.LEFT, padx=8)
        
        cancel_btn = tk.Button(button_frame, text="Bỏ qua", command=on_cancel,
                               bg='#FFCDD2', fg='#000000', font=('Arial', 11, 'bold'),
                               relief=tk.RAISED, bd=2, padx=25, pady=8,
                               activebackground='#EF9A9A', activeforeground='#000000',
                               cursor='hand2')
        cancel_btn.pack(side=tk.LEFT, padx=8)
        
        # Bind Enter key
        path1_entry.bind('<Return>', lambda e: path2_entry.focus())
        path2_entry.bind('<Return>', lambda e: on_ok())
        
        # Focus vào ô đầu tiên
        path1_entry.focus()
        path1_entry.select_range(0, tk.END)
        
        # Đợi dialog đóng
        dialog.wait_window()
        
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
                
                # Xác định thư mục để tìm file thay thế (ưu tiên thư mục chứa executable)
                search_dir = Path(sys.executable).parent
                # Danh sách file thay thế nếu file chính không đọc được
                xls_alternatives = [
                    search_dir / "KIEM KE Năm -2025 - BP ONLINE.xls",
                    search_dir / "KIEM KE Năm -2025 - BP ONLINE copy.xls",
                    search_dir / "DuLieuDauVao.xls",
                    # Cũng thử trong bundle
                    base_path / "KIEM KE Năm -2025 - BP ONLINE.xls",
                    base_path / "KIEM KE Năm -2025 - BP ONLINE copy.xls",
                    base_path / "DuLieuDauVao.xls",
                ]
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
        
        # Tạo Notebook để chứa các tab
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # === TAB 1: KIỂM KÊ ===
        tab_kiemke = tk.Frame(self.notebook, bg=bg_color)
        self.notebook.add(tab_kiemke, text="Kiểm kê")
        
        # Frame chính cho tab Kiểm kê
        main_frame = tk.Frame(tab_kiemke, bg=bg_color)
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
        self.vi_tri_moi_entry = tk.Entry(info_frame, textvariable=self.vi_tri_moi_var, width=20, bg=input_bg_yellow,
                          fg=input_fg, font=('Arial', 10), relief=tk.SOLID, bd=1, insertbackground='#000000')
        self.vi_tri_moi_entry.grid(row=0, column=4, padx=5, pady=5, sticky='ew')
        # Thêm validation khi người dùng nhập xong
        self.vi_tri_moi_entry.bind('<FocusOut>', lambda e: self.validate_vi_tri_moi())
        self.vi_tri_moi_entry.bind('<Return>', lambda e: self.validate_vi_tri_moi())
        
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
        
        # Tổ (*) - Bắt buộc
        tk.Label(info_frame, text="Tổ (*):", bg=bg_color, fg=label_required_fg, font=('Arial', 11, 'bold')).grid(row=3, column=0, sticky='w', padx=5, pady=5)
        self.to_var = tk.StringVar()
        self.to_entry = tk.Entry(info_frame, textvariable=self.to_var, width=40, bg=input_bg_yellow,
                          fg=input_fg, font=('Arial', 10), relief=tk.SOLID, bd=1, insertbackground='#000000')
        self.to_entry.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky='ew')
        
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
        
        # Định nghĩa thứ tự cột cố định
        columns = ('ISBN', 'Tựa', 'Tồn thực tế', 'Số thùng', 'Tồn tựa trong thùng', 'Ghi chú')
        # Thứ tự: 0=ISBN, 1=Tựa, 2=Tồn thực tế, 3=Số thùng, 4=Tồn tựa trong thùng, 5=Ghi chú
        
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
        
        # Scrollbar với callback để cập nhật highlights
        def yview_scroll(*args):
            self.tree.yview(*args)
            self.update_all_highlights()
        
        def xview_scroll(*args):
            self.tree.xview(*args)
            self.update_all_highlights()
        
        scrollbar_y.config(command=yview_scroll)
        scrollbar_x.config(command=xview_scroll)
        
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
        
        # Tag để highlight màu đỏ cho cả row khi có lỗi (không dùng nữa, chỉ highlight 2 cell)
        # self.tree.tag_configure('error', background='#FFEBEE', foreground='#C62828')
        
        self.tree.grid(row=0, column=0, sticky='nsew')
        scrollbar_y.grid(row=0, column=1, sticky='ns')
        scrollbar_x.grid(row=1, column=0, sticky='ew')
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Bind click để edit trực tiếp tất cả các cột có thể edit
        self.tree.bind('<Button-1>', self.on_item_click)
        # Double click cũng mở edit (để dễ sử dụng hơn)
        self.tree.bind('<Double-1>', self.on_item_click)
        
        # Bind events để cập nhật highlight khi scroll hoặc resize
        self.tree.bind('<Configure>', lambda e: self.update_all_highlights())
        
        # Bind khi window resize
        self.root.bind('<Configure>', lambda e: self.update_all_highlights())
        
        # Bind khi mouse wheel scroll
        self.tree.bind('<MouseWheel>', lambda e: self.update_all_highlights())
        self.tree.bind('<Button-4>', lambda e: self.update_all_highlights())
        self.tree.bind('<Button-5>', lambda e: self.update_all_highlights())
        
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
        
        # === TAB 2: TỔNG HỢP ===
        tab_tonghop = tk.Frame(self.notebook, bg=bg_color)
        self.notebook.add(tab_tonghop, text="Tổng hợp")
        
        # Frame chính cho tab Tổng hợp
        tonghop_main_frame = tk.Frame(tab_tonghop, bg=bg_color)
        tonghop_main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tiêu đề
        title_frame = tk.Frame(tonghop_main_frame, bg=bg_color)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        title_label = tk.Label(title_frame, text="TỔNG HỢP DỮ LIỆU NHẬP", 
                              bg=bg_color, fg='#000000', font=('Arial', 16, 'bold'))
        title_label.pack()
        
        # Nút Tải file excel tổng hợp
        button_frame = tk.Frame(tonghop_main_frame, bg=bg_color)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        export_btn = tk.Button(button_frame, text="Tải file excel tổng hợp", 
                              command=self.export_tong_hop_excel,
                              bg='#4CAF50', fg='white', font=('Arial', 12, 'bold'), 
                              width=25, height=2, relief=tk.RAISED, bd=2, cursor='hand2')
        export_btn.pack(side=tk.RIGHT, padx=10)
        
        # Bảng tổng hợp
        tonghop_table_frame = tk.Frame(tonghop_main_frame, bg=bg_color)
        tonghop_table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar cho bảng tổng hợp
        tonghop_scrollbar_y = tk.Scrollbar(tonghop_table_frame, orient=tk.VERTICAL, bg='#E0E0E0', troughcolor=bg_color)
        tonghop_scrollbar_x = tk.Scrollbar(tonghop_table_frame, orient=tk.HORIZONTAL, bg='#E0E0E0', troughcolor=bg_color)
        
        # Cột cho bảng tổng hợp
        tonghop_columns = ('N/X', 'Số phiếu', 'Ngày', 'Vị trí mới', 'ISBN', 'Tựa', 'Tồn thực tế', 'Số thùng', 'Ghi chú')
        
        # Tạo Treeview cho tổng hợp
        self.tong_hop_tree = ttk.Treeview(tonghop_table_frame, columns=tonghop_columns, show='headings', 
                                          yscrollcommand=tonghop_scrollbar_y.set, xscrollcommand=tonghop_scrollbar_x.set,
                                          height=20, style='Treeview')
        
        # Cấu hình các cột
        column_widths = {'N/X': 250, 'Số phiếu': 120, 'Ngày': 100, 'Vị trí mới': 120, 'ISBN': 150, 
                        'Tựa': 250, 'Tồn thực tế': 100, 'Số thùng': 120, 'Ghi chú': 200}
        
        for col in tonghop_columns:
            self.tong_hop_tree.heading(col, text=col)
            self.tong_hop_tree.column(col, width=column_widths.get(col, 100), anchor='w')
        
        tonghop_scrollbar_y.config(command=self.tong_hop_tree.yview)
        tonghop_scrollbar_x.config(command=self.tong_hop_tree.xview)
        
        self.tong_hop_tree.grid(row=0, column=0, sticky='nsew')
        tonghop_scrollbar_y.grid(row=0, column=1, sticky='ns')
        tonghop_scrollbar_x.grid(row=1, column=0, sticky='ew')
        
        tonghop_table_frame.grid_rowconfigure(0, weight=1)
        tonghop_table_frame.grid_columnconfigure(0, weight=1)
    
    def get_all_box_numbers(self):
        """Lấy danh sách tất cả mã thùng từ dữ liệu đầu vào"""
        if self.df is None or self.df.empty:
            return set()
        
        # Tìm cột số thùng
        box_col = None
        for col in self.df.columns:
            col_lower = str(col).lower().strip()
            if ('số thùng' in col_lower or 'so thung' in col_lower or 
                col_lower == 'thùng' or col_lower == 'thung'):
                box_col = col
                break
        
        if box_col is None:
            return set()
        
        # Lấy tất cả giá trị số thùng, loại bỏ NaN và chuyển thành string
        box_numbers = self.df[box_col].dropna().astype(str).str.strip()
        # Loại bỏ các giá trị rỗng
        box_numbers = box_numbers[box_numbers != '']
        return set(box_numbers.unique())
    
    def validate_vi_tri_moi(self):
        """Kiểm tra mã thùng mới có trùng với dữ liệu đầu vào không"""
        vi_tri_moi = self.vi_tri_moi_var.get().strip()
        
        # Nếu rỗng thì không cần kiểm tra
        if not vi_tri_moi:
            return True
        
        # Lấy danh sách tất cả mã thùng từ dữ liệu đầu vào
        existing_box_numbers = self.get_all_box_numbers()
        
        # Kiểm tra xem mã thùng mới có trùng với mã thùng nào trong dữ liệu không
        if vi_tri_moi in existing_box_numbers:
            messagebox.showerror(
                "Lỗi", 
                f"Mã thùng mới '{vi_tri_moi}' đã tồn tại trong dữ liệu đầu vào!\n\n"
                f"Vui lòng nhập mã thùng khác với các mã thùng hiện có.\n\n"
                f"Các mã thùng hiện có: {', '.join(sorted(existing_box_numbers)[:10])}"
                + (f" và {len(existing_box_numbers) - 10} mã khác..." if len(existing_box_numbers) > 10 else "")
            )
            # Xóa giá trị và focus lại vào ô nhập
            self.vi_tri_moi_var.set('')
            self.vi_tri_moi_entry.focus()
            return False
        
        return True
    
    def load_box_data(self):
        """Load dữ liệu của thùng được nhập"""
        so_thung = self.so_thung_var.get().strip()
        if not so_thung:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập số thùng!")
            return
        
        # Kiểm tra nếu đang có dữ liệu đã quét
        if self.scanned_items and len(self.scanned_items) > 0:
            # Đếm số item có tồn thực tế đã nhập
            items_with_data = sum(1 for item in self.scanned_items.values() if item.get('ton_thuc_te', '').strip())
            
            if items_with_data > 0:
                # Hiển thị dialog hỏi người dùng
                result = messagebox.askyesnocancel(
                    "Cảnh báo", 
                    f"Bạn đang có {len(self.scanned_items)} item đã quét, trong đó {items_with_data} item đã nhập tồn thực tế.\n\n"
                    "Bạn có muốn lưu dữ liệu trước khi load thùng mới không?\n\n"
                    "• Có: Lưu dữ liệu hiện tại\n"
                    "• Không: Xóa dữ liệu và load thùng mới\n"
                    "• Hủy: Không làm gì"
                )
                
                if result is True:  # Người dùng chọn "Có" - Lưu
                    self.save_data()
                    # Sau khi lưu, tiếp tục load thùng mới
                elif result is False:  # Người dùng chọn "Không" - Reset
                    # Reset dữ liệu
                    self.scanned_items = {}
                    self.clear_table()
                    self.so_tua_var.set("0")
                    # Tiếp tục load thùng mới
                else:  # Người dùng chọn "Hủy" - Không làm gì
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
            
            # Kiểm tra nếu đã quét đủ số tựa trong thùng này
            so_tua_trong_thung = len(self.current_box_data)
            so_tua_da_quet = len(self.scanned_items)
            
            # Nếu đang load lại cùng một thùng và đã quét đủ
            if self.current_box_number == so_thung and so_tua_da_quet >= so_tua_trong_thung:
                messagebox.showwarning(
                    "Cảnh báo", 
                    f"Thùng {so_thung} đã được kiểm kê đủ {so_tua_trong_thung} tựa!\n\n"
                    f"Bạn đã quét {so_tua_da_quet} tựa.\n\n"
                    "Vui lòng load thùng khác hoặc lưu dữ liệu trước khi tiếp tục."
                )
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
        
        # Kiểm tra nếu đã quét đủ số tựa trong thùng
        so_tua_trong_thung = len(self.current_box_data)
        so_tua_da_quet = len(self.scanned_items)
        
        if so_tua_da_quet >= so_tua_trong_thung:
            messagebox.showerror(
                "Lỗi", 
                f"Đã quét đủ số tựa trong thùng!\n\n"
                f"Thùng {self.current_box_number} có {so_tua_trong_thung} tựa.\n"
                f"Bạn đã quét {so_tua_da_quet} tựa.\n\n"
                "Vui lòng lưu dữ liệu hoặc load thùng khác để tiếp tục."
            )
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
                # Kiểm tra xem ISBN có tồn tại trong thùng khác không
                isbn_clean = str(isbn).strip()
                isbn_clean_digits = ''.join(filter(str.isdigit, isbn_clean))
                found_in_other_box = False
                other_box_number = None
                
                # Tìm trong toàn bộ dữ liệu
                if 'isbn' in self.df.columns:
                    # Tìm cột số thùng
                    so_thung_col = None
                    for col in self.df.columns:
                        col_lower = str(col).lower().strip()
                        if 'số thùng' in col_lower or 'so thung' in col_lower or col_lower == 'thùng' or col_lower == 'so_thung':
                            so_thung_col = col
                            break
                    
                    if so_thung_col:
                        # Tìm ISBN trong toàn bộ dữ liệu
                        for idx, row in self.df.iterrows():
                            row_isbn = str(row.get('isbn', '')).strip()
                            row_isbn_clean = ''.join(filter(str.isdigit, row_isbn))
                            
                            # Kiểm tra khớp ISBN
                            if (row_isbn == isbn_clean or 
                                row_isbn.endswith(isbn_clean) or 
                                isbn_clean.endswith(row_isbn) or
                                (row_isbn_clean and isbn_clean_digits and row_isbn_clean == isbn_clean_digits)):
                                # Tìm thấy ISBN, lấy số thùng
                                other_box_number = str(row.get(so_thung_col, '')).strip()
                                if other_box_number and other_box_number != str(self.current_box_number):
                                    found_in_other_box = True
                                    break
                
                # Báo lỗi tương ứng
                if found_in_other_box:
                    messagebox.showerror(
                        "Lỗi", 
                        f"ISBN {isbn} không thuộc thùng đang kiểm kê!\n\n"
                        f"ISBN này thuộc thùng: {other_box_number}\n"
                        f"Thùng đang kiểm kê: {self.current_box_number}\n\n"
                        f"Vui lòng quét đúng ISBN của thùng {self.current_box_number}."
                    )
                else:
                    messagebox.showwarning(
                        "Cảnh báo", 
                        f"Không tìm thấy ISBN {isbn} trong dữ liệu!\n\n"
                        f"Vui lòng kiểm tra lại mã ISBN hoặc thùng số {self.current_box_number}."
                    )
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
            
            # Kiểm tra nếu có "Thùng / vị trí mới" thì dùng giá trị đó, không thì dùng số thùng hiện tại
            vi_tri_moi = self.vi_tri_moi_var.get().strip()
            
            # Validation: Mã thùng mới phải khác với tất cả mã thùng trong dữ liệu đầu vào
            if vi_tri_moi:
                # Kiểm tra lại một lần nữa khi quét ISBN (để đảm bảo)
                existing_box_numbers = self.get_all_box_numbers()
                if vi_tri_moi in existing_box_numbers:
                    messagebox.showerror(
                        "Lỗi", 
                        f"Mã thùng mới '{vi_tri_moi}' đã tồn tại trong dữ liệu đầu vào!\n\n"
                        f"Vui lòng nhập mã thùng khác với các mã thùng hiện có.\n\n"
                        f"Các mã thùng hiện có: {', '.join(sorted(existing_box_numbers)[:10])}"
                        + (f" và {len(existing_box_numbers) - 10} mã khác..." if len(existing_box_numbers) > 10 else "")
                    )
                    self.isbn_entry.delete(0, tk.END)
                    return
                
                so_thung_hien_thi = vi_tri_moi
            else:
                so_thung_hien_thi = so_thung
            
            # Đảm bảo thứ tự đúng với columns: ISBN, Tựa, Tồn thực tế, Số thùng, Tồn tựa trong thùng, Ghi chú
            item_id = self.tree.insert('', tk.END, values=(
                str(isbn_clean),           # 0: ISBN
                str(tua),                  # 1: Tựa
                '',                        # 2: Tồn thực tế - để trống để người dùng nhập
                str(so_thung_hien_thi),    # 3: Số thùng (dùng vị trí mới nếu có)
                str(ton_trong_thung_display),  # 4: Tồn tựa trong thùng
                ''                         # 5: Ghi chú - để trống
            ), tags=('',))
            
            # Lưu thông tin
            # Lưu cả số thùng gốc và số thùng hiển thị (vị trí mới nếu có)
            # Lưu cả vi_tri_moi để có thể sử dụng khi lưu (nếu người dùng chỉnh sửa trực tiếp)
            vi_tri_moi_value = self.vi_tri_moi_var.get().strip()
            self.scanned_items[isbn_clean] = {
                'item_id': item_id,
                'tua': tua,
                'ton_thuc_te': '',
                'so_thung': so_thung_hien_thi,  # Lưu số thùng hiển thị (có thể là vị trí mới)
                'so_thung_goc': so_thung,  # Lưu số thùng gốc từ dữ liệu
                'vi_tri_moi': vi_tri_moi_value,  # Lưu giá trị từ ô "Thùng / vị trí mới" khi quét
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
        
        # Lấy giá trị hiện tại và đảm bảo có đủ 6 cột
        values = list(self.tree.item(item, 'values'))
        while len(values) < 6:
            values.append('')
        
        isbn = values[0] if len(values) > 0 else ''
        
        # Xử lý theo từng cột
        if column_index == 2:  # Tồn thực tế
            values[2] = new_value  # Đảm bảo đúng index
            
            # Kiểm tra và highlight nếu khác nhau
            if isbn in self.scanned_items:
                self.scanned_items[isbn]['ton_thuc_te'] = new_value
                ton_trong_thung = self.scanned_items[isbn]['ton_trong_thung']
                
                try:
                    ton_thuc_te_num = float(new_value) if new_value else 0
                    ton_trong_thung_num = float(ton_trong_thung) if ton_trong_thung else 0
                    
                    # Kiểm tra lệch
                    if abs(ton_thuc_te_num - ton_trong_thung_num) > 0.01:
                        # Tự động ghi chú lỗi vào cột Ghi chú
                        if ton_thuc_te_num < ton_trong_thung_num:
                            error_note = f"Thiếu {int(ton_trong_thung_num - ton_thuc_te_num)} cuốn"
                        else:
                            error_note = f"Dư {int(ton_thuc_te_num - ton_trong_thung_num)} cuốn"
                        
                        # Đảm bảo có đủ 6 cột và đúng thứ tự: ISBN, Tựa, Tồn thực tế, Số thùng, Tồn tựa trong thùng, Ghi chú
                        while len(values) < 6:
                            values.append('')
                        
                        # Đảm bảo thứ tự đúng: values[0]=ISBN, values[1]=Tựa, values[2]=Tồn thực tế, 
                        # values[3]=Số thùng, values[4]=Tồn tựa trong thùng, values[5]=Ghi chú
                        if len(values) >= 6:
                            values[5] = error_note  # Ghi chú ở cột cuối cùng (index 5)
                        else:
                            values.append(error_note)
                        
                        # Cập nhật scanned_items
                        self.scanned_items[isbn]['ghi_chu'] = error_note
                        
                        # Cập nhật tree
                        self.tree.item(item, values=values)
                        
                        # Tô đỏ 2 ô: Tồn thực tế (cột 3) và Ghi chú (cột 6)
                        self.highlight_error_cells(item)
                        
                        # Hiển thị cảnh báo
                        messagebox.showwarning("Cảnh báo", 
                            f"Tồn thực tế ({int(ton_thuc_te_num)}) khác với tồn trong thùng ({int(ton_trong_thung_num)})!\n"
                            f"Đã tự động ghi chú: {error_note}")
                    else:
                        # Không có lỗi - xóa highlight và ghi chú lỗi
                        # Đảm bảo có đủ 6 cột
                        while len(values) < 6:
                            values.append('')
                        
                        # Xóa ghi chú lỗi nếu có
                        if len(values) > 5 and ('Thiếu' in str(values[5]) or 'Dư' in str(values[5])):
                            values[5] = ''
                            self.scanned_items[isbn]['ghi_chu'] = ''
                        
                        # Xóa highlight
                        self.remove_error_highlights(item)
                        
                        # Cập nhật tree
                        self.tree.item(item, values=values)
                except (ValueError, TypeError):
                    # Nếu không phải số, không highlight
                    self.tree.item(item, tags=('',))
        
        elif column_index == 3:  # Số thùng
            # Validation: Mã thùng mới phải khác với tất cả mã thùng trong dữ liệu đầu vào
            if new_value.strip():
                existing_box_numbers = self.get_all_box_numbers()
                if new_value.strip() in existing_box_numbers:
                    messagebox.showerror(
                        "Lỗi", 
                        f"Mã thùng '{new_value.strip()}' đã tồn tại trong dữ liệu đầu vào!\n\n"
                        f"Vui lòng nhập mã thùng khác với các mã thùng hiện có.\n\n"
                        f"Các mã thùng hiện có: {', '.join(sorted(existing_box_numbers)[:10])}"
                        + (f" và {len(existing_box_numbers) - 10} mã khác..." if len(existing_box_numbers) > 10 else "")
                    )
                    # Khôi phục giá trị cũ
                    values[3] = self.scanned_items[isbn]['so_thung'] if isbn in self.scanned_items else ''
                    self.tree.item(item, values=values)
                    return
            
            values[3] = new_value  # Đảm bảo đúng index
            if isbn in self.scanned_items:
                # Khi chỉnh sửa trực tiếp, cập nhật số thùng hiển thị
                # Nhưng giữ nguyên số thùng gốc (từ dữ liệu đầu vào) - KHÔNG BAO GIỜ thay đổi
                new_value_clean = new_value.strip()
                self.scanned_items[isbn]['so_thung'] = new_value_clean
                
                # Đảm bảo so_thung_goc luôn được giữ nguyên (không thay đổi khi chỉnh sửa)
                # Chỉ set so_thung_goc nếu chưa có (lần đầu quét)
                if 'so_thung_goc' not in self.scanned_items[isbn] or not self.scanned_items[isbn].get('so_thung_goc'):
                    # Nếu chưa có, lấy từ current_box_number (số thùng đã load)
                    if self.current_box_number:
                        self.scanned_items[isbn]['so_thung_goc'] = self.current_box_number
                    else:
                        # Nếu không có current_box_number, dùng giá trị hiện tại (fallback)
                        self.scanned_items[isbn]['so_thung_goc'] = new_value_clean
                # Nếu đã có so_thung_goc, KHÔNG BAO GIỜ thay đổi nó
                
                # QUAN TRỌNG: Giữ nguyên vi_tri_moi khi chỉnh sửa trực tiếp
                # Nếu chưa có vi_tri_moi, lấy từ ô input hiện tại
                if 'vi_tri_moi' not in self.scanned_items[isbn] or not self.scanned_items[isbn].get('vi_tri_moi'):
                    vi_tri_moi_current = self.vi_tri_moi_var.get().strip()
                    if vi_tri_moi_current:
                        self.scanned_items[isbn]['vi_tri_moi'] = vi_tri_moi_current
        
        elif column_index == 4:  # Tồn tựa trong thùng
            # Chuyển thành số nguyên
            try:
                new_value_int = int(float(new_value)) if new_value else 0
                values[4] = str(new_value_int)  # Đảm bảo đúng index và chuyển thành string
                if isbn in self.scanned_items:
                    self.scanned_items[isbn]['ton_trong_thung'] = new_value_int
            except:
                values[4] = new_value  # Đảm bảo đúng index
        
        elif column_index == 5:  # Ghi chú
            values[5] = new_value  # Đảm bảo đúng index
            if isbn in self.scanned_items:
                self.scanned_items[isbn]['ghi_chu'] = new_value
                # Nếu người dùng tự chỉnh sửa ghi chú, không tự động ghi đè lại
                # Chỉ cập nhật highlight nếu vẫn còn lệch
                ton_trong_thung = self.scanned_items[isbn]['ton_trong_thung']
                ton_thuc_te = self.scanned_items[isbn].get('ton_thuc_te', '')
                try:
                    ton_thuc_te_num = float(ton_thuc_te) if ton_thuc_te else 0
                    ton_trong_thung_num = float(ton_trong_thung) if ton_trong_thung else 0
                    if abs(ton_thuc_te_num - ton_trong_thung_num) > 0.01:
                        # Vẫn còn lệch, giữ highlight
                        self.highlight_error_cells(item)
                    else:
                        # Đã khớp, xóa highlight
                        self.remove_error_highlights(item)
                except:
                    pass
        
        # Cập nhật tree với giá trị mới
        self.tree.item(item, values=values)
        
        # Nếu là cột Tồn thực tế, kiểm tra lại và cập nhật highlight
        if column_index == 2:
            if isbn in self.scanned_items:
                ton_trong_thung = self.scanned_items[isbn]['ton_trong_thung']
                try:
                    ton_thuc_te_num = float(new_value) if new_value else 0
                    ton_trong_thung_num = float(ton_trong_thung) if ton_trong_thung else 0
                    if abs(ton_thuc_te_num - ton_trong_thung_num) > 0.01:
                        # Vẫn còn lệch - tự động cập nhật ghi chú lỗi nếu chưa có ghi chú tùy chỉnh
                        current_ghi_chu = self.scanned_items[isbn].get('ghi_chu', '')
                        # Chỉ tự động ghi chú nếu ghi chú hiện tại là ghi chú lỗi tự động hoặc rỗng
                        if not current_ghi_chu or ('Thiếu' in current_ghi_chu or 'Dư' in current_ghi_chu):
                            if ton_thuc_te_num < ton_trong_thung_num:
                                error_note = f"Thiếu {int(ton_trong_thung_num - ton_thuc_te_num)} cuốn"
                            else:
                                error_note = f"Dư {int(ton_thuc_te_num - ton_trong_thung_num)} cuốn"
                            values[5] = error_note
                            self.scanned_items[isbn]['ghi_chu'] = error_note
                            self.tree.item(item, values=values)
                        self.highlight_error_cells(item)
                    else:
                        # Đã khớp - chỉ xóa ghi chú lỗi tự động, giữ lại ghi chú tùy chỉnh
                        current_ghi_chu = self.scanned_items[isbn].get('ghi_chu', '')
                        if current_ghi_chu and ('Thiếu' in current_ghi_chu or 'Dư' in current_ghi_chu):
                            # Chỉ xóa nếu là ghi chú lỗi tự động
                            values[5] = ''
                            self.scanned_items[isbn]['ghi_chu'] = ''
                            self.tree.item(item, values=values)
                        self.remove_error_highlights(item)
                except:
                    self.remove_error_highlights(item)
        
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
    
    def highlight_error_cells(self, item_id):
        """Tô đỏ 2 ô: Tồn thực tế và Ghi chú"""
        # Xóa highlight cũ nếu có
        self.remove_error_highlights(item_id)
        
        # Lấy giá trị từ tree
        values = list(self.tree.item(item_id, 'values'))
        
        # Tô đỏ ô "Tồn thực tế" (cột 3, index 2)
        bbox_ton = self.tree.bbox(item_id, '#3')
        if bbox_ton:
            x, y, width, height = bbox_ton
            ton_value = values[2] if len(values) > 2 else ''
            highlight1 = tk.Label(self.tree, bg='#FFCDD2', fg='#C62828', 
                                  text=str(ton_value), font=('Arial', 10, 'bold'), 
                                  relief=tk.FLAT, anchor='center')
            highlight1.place(x=x, y=y, width=width, height=height)
            # Cho phép click qua để edit
            highlight1.bind('<Button-1>', lambda e, item=item_id, col='#3': self.on_highlight_click(e, item, col))
        
        # Tô đỏ ô "Ghi chú" (cột 6, index 5)
        bbox_ghi_chu = self.tree.bbox(item_id, '#6')
        if bbox_ghi_chu:
            x, y, width, height = bbox_ghi_chu
            ghi_chu_value = values[5] if len(values) > 5 else ''
            highlight2 = tk.Label(self.tree, bg='#FFCDD2', fg='#C62828', 
                                 text=str(ghi_chu_value), font=('Arial', 10, 'bold'), 
                                 relief=tk.FLAT, anchor='w', padx=5)
            highlight2.place(x=x, y=y, width=width, height=height)
            # Cho phép click qua để edit
            highlight2.bind('<Button-1>', lambda e, item=item_id, col='#6': self.on_highlight_click(e, item, col))
        
        # Lưu các highlight widgets
        if item_id not in self.error_highlights:
            self.error_highlights[item_id] = []
        if bbox_ton:
            self.error_highlights[item_id].append(highlight1)
        if bbox_ghi_chu:
            self.error_highlights[item_id].append(highlight2)
        
        # Cập nhật lại highlight khi scroll hoặc resize
        self.root.after(100, lambda: self.update_error_highlights(item_id))
    
    def on_highlight_click(self, event, item_id, column):
        """Xử lý click vào highlight để edit cell"""
        # Hủy edit cũ nếu có
        if self.edit_entry:
            self.finish_edit()
        
        column_index = int(column.replace('#', '')) - 1
        
        # Lấy giá trị hiện tại
        values = list(self.tree.item(item_id, 'values'))
        current_value = values[column_index] if column_index < len(values) else ''
        
        # Lấy vị trí của cell
        bbox = self.tree.bbox(item_id, column)
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
        self.editing_item = item_id
        self.editing_column = column_index
        
        def finish_on_enter(event):
            self.finish_edit()
        
        def finish_on_focus_out(event):
            self.root.after(100, self.finish_edit)
        
        self.edit_entry.bind('<Return>', finish_on_enter)
        self.edit_entry.bind('<FocusOut>', finish_on_focus_out)
        self.edit_entry.bind('<Escape>', lambda e: self.cancel_edit())
    
    def update_error_highlights(self, item_id):
        """Cập nhật lại vị trí highlight khi scroll"""
        if item_id not in self.error_highlights:
            return
        
        # Kiểm tra xem item còn tồn tại không
        try:
            if item_id not in self.tree.get_children(''):
                return
        except:
            return
        
        # Lấy giá trị từ tree
        values = list(self.tree.item(item_id, 'values'))
        
        widgets = self.error_highlights[item_id]
        if len(widgets) >= 2:
            # Cập nhật ô Tồn thực tế
            bbox_ton = self.tree.bbox(item_id, '#3')
            if bbox_ton and widgets[0].winfo_exists():
                x, y, width, height = bbox_ton
                ton_value = values[2] if len(values) > 2 else ''
                widgets[0].config(text=str(ton_value))
                widgets[0].place(x=x, y=y, width=width, height=height)
            
            # Cập nhật ô Ghi chú
            bbox_ghi_chu = self.tree.bbox(item_id, '#6')
            if bbox_ghi_chu and len(widgets) > 1 and widgets[1].winfo_exists():
                x, y, width, height = bbox_ghi_chu
                ghi_chu_value = values[5] if len(values) > 5 else ''
                widgets[1].config(text=str(ghi_chu_value))
                widgets[1].place(x=x, y=y, width=width, height=height)
    
    def update_all_highlights(self):
        """Cập nhật tất cả các highlight khi scroll hoặc resize"""
        # Delay một chút để đảm bảo tree đã cập nhật
        self.root.after(10, self._do_update_all_highlights)
    
    def _do_update_all_highlights(self):
        """Thực hiện cập nhật tất cả highlights"""
        for item_id in list(self.error_highlights.keys()):
            try:
                # Kiểm tra item còn tồn tại không
                if item_id in self.tree.get_children(''):
                    self.update_error_highlights(item_id)
                else:
                    # Xóa highlight nếu item không còn tồn tại
                    self.remove_error_highlights(item_id)
            except:
                pass
    
    def remove_error_highlights(self, item_id):
        """Xóa highlight của các ô lỗi"""
        if item_id in self.error_highlights:
            for widget in self.error_highlights[item_id]:
                try:
                    widget.destroy()
                except:
                    pass
            del self.error_highlights[item_id]
    
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
    
    
    def clear_table(self):
        """Xóa tất cả items trong bảng"""
        # Xóa tất cả highlights trước
        for item_id in list(self.error_highlights.keys()):
            self.remove_error_highlights(item_id)
        
        # Xóa tất cả items
        for item in self.tree.get_children():
            self.tree.delete(item)
    
    def on_enter_pressed(self, event):
        """Xử lý phím Enter"""
        widget = self.root.focus_get()
        if widget == self.isbn_entry:
            self.on_isbn_entered()
    
    def save_data(self):
        """Lưu dữ liệu đã kiểm tra vào tab Tổng hợp"""
        if not self.scanned_items:
            messagebox.showwarning("Cảnh báo", "Chưa có dữ liệu để lưu!")
            return
        
        # Kiểm tra ràng buộc: Tổ là bắt buộc
        to_value = self.to_var.get().strip() if hasattr(self, 'to_var') and self.to_var.get() else ''
        if not to_value:
            messagebox.showerror("Lỗi", "Vui lòng nhập 'Tổ' trước khi lưu!")
            # Focus vào ô input Tổ
            if hasattr(self, 'to_entry'):
                self.to_entry.focus()
                self.to_entry.select_range(0, tk.END)
            return
        
        # Lấy giá trị từ các input
        vi_tri_moi_global = self.vi_tri_moi_var.get().strip() if hasattr(self, 'vi_tri_moi_var') else ''
        ngay_value = self.ngay_var.get().strip() if hasattr(self, 'ngay_var') else ''
        nhap_xuat_value = self.nhap_xuat_var.get().strip() if hasattr(self, 'nhap_xuat_var') else ''
        
        # Tạo số phiếu theo format P-DD/MM/YYYY
        from datetime import datetime
        try:
            if ngay_value:
                # Parse ngày từ format DD/MM/YY hoặc DD/MM/YYYY
                if len(ngay_value.split('/')) == 3:
                    parts = ngay_value.split('/')
                    if len(parts[2]) == 2:
                        # Format DD/MM/YY -> DD/MM/YYYY
                        parts[2] = '20' + parts[2]
                    ngay_parsed = datetime.strptime('/'.join(parts), "%d/%m/%Y")
                else:
                    ngay_parsed = datetime.now()
            else:
                ngay_parsed = datetime.now()
        except:
            ngay_parsed = datetime.now()
        
        so_phieu = f"P-{ngay_parsed.strftime('%d/%m/%Y')}"
        
        # Thêm dữ liệu vào tổng hợp
        for isbn, info in self.scanned_items.items():
            # Lấy số thùng gốc
            so_thung_goc = info.get('so_thung_goc', '')
            if not so_thung_goc:
                so_thung_goc = self.current_box_number if self.current_box_number else info.get('so_thung', '')
            so_thung_goc_clean = str(so_thung_goc).strip()
            
            # Xác định số thùng mới
            so_thung_moi = ''
            if vi_tri_moi_global and vi_tri_moi_global != so_thung_goc_clean:
                so_thung_moi = vi_tri_moi_global
            else:
                vi_tri_moi_saved = info.get('vi_tri_moi', '').strip()
                if vi_tri_moi_saved and vi_tri_moi_saved != so_thung_goc_clean:
                    so_thung_moi = vi_tri_moi_saved
                else:
                    so_thung_hien_thi = info.get('so_thung', so_thung_goc)
                    so_thung_hien_thi_clean = str(so_thung_hien_thi).strip()
                    if so_thung_hien_thi_clean != so_thung_goc_clean and so_thung_hien_thi_clean:
                        so_thung_moi = so_thung_hien_thi_clean
            
            # Lấy giá trị từ input "Nhập/Xuất" ở tab Kiểm kê và hiển thị trực tiếp vào cột N/X
            nx_value = nhap_xuat_value.strip() if nhap_xuat_value else ""
            
            # Thêm vào tổng hợp
            self.tong_hop_data.append({
                'N/X': nx_value,
                'Số phiếu': so_phieu,
                'Ngày': ngay_value,
                'Vị trí mới': so_thung_moi if so_thung_moi else so_thung_goc_clean,
                'ISBN': isbn,
                'Tựa': info['tua'],
                'Tồn thực tế': info['ton_thuc_te'],
                'Số thùng': so_thung_goc_clean,
                'Ghi chú': info['ghi_chu']
            })
        
        # Cập nhật bảng tổng hợp
        self.update_tong_hop_table()
        
        # Xóa dữ liệu đã quét
        self.scanned_items.clear()
        self.clear_table()
        self.so_tua_var.set("0")
        
        # Reset các input: Số thùng và Thùng / vị trí mới
        if hasattr(self, 'so_thung_var'):
            self.so_thung_var.set("")
        if hasattr(self, 'vi_tri_moi_var'):
            self.vi_tri_moi_var.set("")
        
        # Reset current_box_number và current_box_data
        self.current_box_number = None
        self.current_box_data = None
        
        # Chuyển sang tab Tổng hợp
        self.notebook.select(1)
        
        messagebox.showinfo("Thành công", f"Đã lưu {len(self.tong_hop_data)} dòng vào Tổng hợp!")
    
    def update_tong_hop_table(self):
        """Cập nhật bảng tổng hợp"""
        # Xóa tất cả items cũ
        for item in self.tong_hop_tree.get_children():
            self.tong_hop_tree.delete(item)
        
        # Thêm dữ liệu mới
        for data in self.tong_hop_data:
            self.tong_hop_tree.insert('', 'end', values=(
                data.get('N/X', ''),
                data.get('Số phiếu', ''),
                data.get('Ngày', ''),
                data.get('Vị trí mới', ''),
                data.get('ISBN', ''),
                data.get('Tựa', ''),
                data.get('Tồn thực tế', ''),
                data.get('Số thùng', ''),
                data.get('Ghi chú', '')
            ))
    
    def export_tong_hop_excel(self):
        """Xuất file Excel tổng hợp (logic giống save_data cũ)"""
        if not self.tong_hop_data:
            messagebox.showwarning("Cảnh báo", "Chưa có dữ liệu tổng hợp để xuất!")
            return
        
        # Tạo DataFrame từ tổng hợp data
        df_save = pd.DataFrame(self.tong_hop_data)
        
        # Tạo tên file theo format
        from datetime import datetime
        ngay_hien_tai = datetime.now().strftime("%d/%m/%Y")
        to_value = self.to_var.get().strip() if hasattr(self, 'to_var') and self.to_var.get() else ""
        
        # Thay thế ký tự không hợp lệ trong tên file
        to_safe = to_value.replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
        ngay_safe = ngay_hien_tai.replace('/', '_')
        
        # File 1: Kiemke_dd/mm/yyyy_Tổ.xlsx (file chính)
        ten_file_1 = f"Kiemke_{ngay_safe}_{to_safe}.xlsx" if to_safe else f"Kiemke_{ngay_safe}.xlsx"
        
        # File 2: Kiemkecuoinam_dd/mm/yyyy_Tổ.xlsx (file ngầm)
        ten_file_2 = f"Kiemkecuoinam_{ngay_safe}_{to_safe}.xlsx" if to_safe else f"Kiemkecuoinam_{ngay_safe}.xlsx"
        
        # Kiểm tra có cấu hình đầy đủ không
        if not self.template_file_path or not os.path.exists(self.template_file_path):
            messagebox.showerror("Lỗi", "Chưa cấu hình file Excel cố định! Vui lòng cấu hình lại.")
            return
        
        # Cho phép người dùng chọn đường dẫn để lưu file
        filename = filedialog.asksaveasfilename(
            title="Chọn đường dẫn để lưu file Excel tổng hợp",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=ten_file_1
        )
        
        if filename:
            try:
                # File 1: Chỉ copy file template từ đường dẫn đã cấu hình và đổi tên (KHÔNG ghi đè data)
                shutil.copy2(self.template_file_path, filename)
                # KHÔNG ghi đè dữ liệu vào file 1 - giữ nguyên data từ template
                
                # File 2: Tự động lưu vào thư mục đã cấu hình (nếu có)
                if self.auto_save_folder:
                    try:
                        # Tạo đường dẫn file 2 với tên file đúng format
                        file2_path = os.path.join(self.auto_save_folder, ten_file_2)
                        
                        # Lưu file 2 với data tổng hợp (ngầm, không hiển thị thông báo)
                        df_save.to_excel(file2_path, index=False)
                        
                        # Hiển thị thông báo thành công (chỉ đề cập file 1)
                        messagebox.showinfo("Thành công", f"Đã lưu file tổng hợp vào {filename}")
                    except Exception as e2:
                        # Nếu lỗi khi lưu file 2, chỉ log lỗi nhưng không hiển thị cho người dùng
                        print(f"Lỗi khi lưu file tự động: {str(e2)}")
                        # Vẫn hiển thị thành công cho file 1
                        messagebox.showinfo("Thành công", f"Đã lưu file tổng hợp vào {filename}")
                else:
                    # Không có thư mục tự động, chỉ lưu file 1
                    messagebox.showinfo("Thành công", f"Đã lưu file tổng hợp vào {filename}")
                    
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể lưu file: {str(e)}")

def main():
    try:
        root = tk.Tk()
        # Đảm bảo root window được hiển thị ngay từ đầu
        root.deiconify()
        root.update()
        
        app = KiemKhoApp(root)
        
        # Kiểm tra xem app đã được khởi tạo thành công chưa
        if not hasattr(app, 'template_file_path') or not app.template_file_path:
            print("App không được khởi tạo đúng cách")
            root.quit()
            return
        
        root.mainloop()
    except Exception as e:
        # Hiển thị lỗi nếu có
        import traceback
        error_msg = f"Lỗi khi khởi động ứng dụng:\n{str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)
        try:
            # Tạo một root window mới để hiển thị lỗi
            error_root = tk.Tk()
            error_root.withdraw()  # Ẩn window chính
            messagebox.showerror("Lỗi", f"Lỗi khi khởi động ứng dụng:\n{str(e)}")
            error_root.destroy()
        except:
            # Nếu không thể hiển thị messagebox, in ra console
            print("Không thể hiển thị dialog lỗi")
            print(error_msg)

if __name__ == "__main__":
    main()

