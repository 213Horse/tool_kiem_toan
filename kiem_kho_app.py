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
import time
import traceback

class KiemKhoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kiểm Kho - Quét Mã Vạch")
        # Tăng chiều cao để hiển thị đủ tất cả các phần tử
        self.root.geometry("1200x800")
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
        self.is_processing_edit = False  # Flag để tránh xử lý edit 2 lần
        self.template_file_path = None  # Đường dẫn file Excel cố định (template)
        self.auto_save_folder = None  # Thư mục tự động lưu file Excel 2 (Kiemkecuoinam)
        self.config_folder = None  # Thư mục lưu file config (do người dùng chọn)
        self.config_file = self.get_config_file_path()  # Đường dẫn file config
        self.tong_hop_data = []  # Lưu tổng hợp các data đã kiểm kê
        self.notebook = None  # Notebook widget để chứa các tab
        self.tong_hop_tree = None  # Treeview trong tab Tổng hợp
        self.so_tua_da_quet_var = None  # Biến để hiển thị số tựa đã quét
        
        # Load cấu hình từ file (nếu có)
        saved_config = self.load_config()
        need_setup = False
        
        if saved_config:
            template_path = saved_config.get('template_file_path')
            auto_folder = saved_config.get('auto_save_folder')
            config_folder = saved_config.get('config_folder')
            
            # Kiểm tra cả 3 đường dẫn có tồn tại và hợp lệ không
            if template_path and auto_folder and config_folder:
                # Windows-safe: Normalize paths và kiểm tra tồn tại
                try:
                    template_path_normalized = str(Path(template_path).resolve())
                    auto_folder_normalized = str(Path(auto_folder).resolve())
                    config_folder_normalized = str(Path(config_folder).resolve())
                    
                    # Kiểm tra tồn tại với normalized paths
                    if (os.path.isfile(template_path_normalized) and 
                        os.path.isdir(auto_folder_normalized) and 
                        os.path.isdir(config_folder_normalized)):
                        # Cả 3 đường dẫn đều hợp lệ - không cần setup lại
                        self.template_file_path = template_path_normalized
                        self.auto_save_folder = auto_folder_normalized
                        self.config_folder = config_folder_normalized
                        # Cập nhật config_file để dùng đúng thư mục
                        self.config_file = Path(self.config_folder) / "kiem_kho_config.json"
                    else:
                        # Một hoặc nhiều đường dẫn không tồn tại - cần setup lại
                        need_setup = True
                except Exception as e:
                    # Nếu có lỗi khi normalize (ví dụ: đường dẫn không tồn tại), cần setup lại
                    print(f"Lỗi khi kiểm tra đường dẫn: {str(e)}")
                    need_setup = True
            else:
                # Thiếu một hoặc nhiều đường dẫn - cần setup lại
                need_setup = True
        else:
            # Không có cấu hình - cần setup lần đầu
            need_setup = True
        
        # Nếu cần cấu hình, hiển thị dialog (chỉ lần đầu hoặc khi config không hợp lệ)
        if need_setup:
            # Đảm bảo root window được hiển thị trước khi tạo dialog
            self.root.deiconify()
            self.root.update()
            self.setup_paths()
            
            # Sau khi setup xong, kiểm tra lại xem có cấu hình hợp lệ không (cả 3 đường dẫn)
            if not self.template_file_path or not self.auto_save_folder or not self.config_folder:
                # Nếu vẫn không có cấu hình hợp lệ, đóng app
                self.root.quit()
                return
        
        # Load dữ liệu từ Excel
        self.load_data()
        
        # Tạo giao diện
        self.create_ui()
        
        # Tự động điều chỉnh kích thước cửa sổ để hiển thị đủ tất cả các phần tử
        self.root.update_idletasks()
        # Lấy chiều cao yêu cầu của cửa sổ
        req_height = self.root.winfo_reqheight()
        req_width = self.root.winfo_reqwidth()
        # Đảm bảo chiều cao đủ để hiển thị tất cả (thêm một chút padding)
        current_height = self.root.winfo_height()
        if req_height > current_height:
            # Điều chỉnh geometry để hiển thị đủ
            self.root.geometry(f"{req_width}x{req_height + 50}")
        
        # Bind Enter key để hỗ trợ quét mã vạch
        self.root.bind('<Return>', self.on_enter_pressed)
    
    def get_config_file_path(self):
        """Lấy đường dẫn file config - tìm ở nhiều vị trí để đảm bảo tìm được file đã lưu"""
        # Nếu đã có config_folder do người dùng chọn, dùng nó (ưu tiên cao nhất)
        if self.config_folder:
            return Path(self.config_folder) / "kiem_kho_config.json"
        
        # Tìm file config ở nhiều vị trí để lấy config_folder
        search_locations = []
        
        if getattr(sys, 'frozen', False):
            # Chạy từ executable
            exe_dir = Path(sys.executable).parent
            search_locations.extend([
                exe_dir / "kiem_kho_config.json",
                exe_dir.parent / "kiem_kho_config.json",
            ])
            
            user_home = Path.home()
            search_locations.extend([
                user_home / "Desktop" / "kiem_kho_config.json",
                user_home / "Documents" / "kiem_kho_config.json",
                user_home / "kiem_kho_config.json",
            ])
        else:
            # Chạy từ source code
            search_locations.append(Path(__file__).parent / "kiem_kho_config.json")
        
        # Tìm file config ở các vị trí và đọc config_folder từ đó
        for config_file_path in search_locations:
            if config_file_path.exists():
                try:
                    with open(config_file_path, 'r', encoding='utf-8') as f:
                        old_config = json.load(f)
                        # Nếu file này có config_folder, dùng nó
                        if 'config_folder' in old_config and old_config['config_folder']:
                            config_folder_path = Path(old_config['config_folder'])
                            if config_folder_path.exists() and config_folder_path.is_dir():
                                self.config_folder = str(config_folder_path.resolve())
                                # Trả về đường dẫn file config trong config_folder (vị trí thực sự)
                                return Path(self.config_folder) / "kiem_kho_config.json"
                        # Nếu không có config_folder nhưng có đầy đủ thông tin, dùng file này
                        elif 'template_file_path' in old_config and 'auto_save_folder' in old_config:
                            # File này có thể là file config hợp lệ, nhưng chưa có config_folder
                            # Trả về file này để load
                            return config_file_path
                except Exception as e:
                    # Bỏ qua lỗi và tiếp tục tìm
                    continue
        
        # Nếu không tìm thấy, dùng vị trí mặc định
        if getattr(sys, 'frozen', False):
            return Path(sys.executable).parent / "kiem_kho_config.json"
        else:
            return Path(__file__).parent / "kiem_kho_config.json"
    
    def load_config(self):
        """Load cấu hình từ file - Windows-safe, tìm ở nhiều vị trí"""
        try:
            # Danh sách các vị trí tìm file config (theo thứ tự ưu tiên)
            search_locations = []
            
            # Ưu tiên 1: File config trong config_folder (nếu đã biết)
            if self.config_folder:
                search_locations.append(Path(self.config_folder) / "kiem_kho_config.json")
            
            # Ưu tiên 2: File config hiện tại
            if self.config_file:
                search_locations.append(self.config_file)
            
            # Ưu tiên 3: Tìm ở các vị trí khác
            if getattr(sys, 'frozen', False):
                exe_dir = Path(sys.executable).parent
                search_locations.extend([
                    exe_dir / "kiem_kho_config.json",
                    exe_dir.parent / "kiem_kho_config.json",
                ])
                
                user_home = Path.home()
                search_locations.extend([
                    user_home / "Desktop" / "kiem_kho_config.json",
                    user_home / "Documents" / "kiem_kho_config.json",
                    user_home / "kiem_kho_config.json",
                ])
            else:
                search_locations.append(Path(__file__).parent / "kiem_kho_config.json")
            
            # Loại bỏ trùng lặp và chỉ giữ các file tồn tại
            # Windows-safe: Normalize paths để so sánh (case-insensitive trên Windows)
            unique_locations = []
            seen = set()
            for loc in search_locations:
                try:
                    if loc.exists():
                        # Windows-safe: Normalize path và lowercase để so sánh
                        if sys.platform == 'win32':
                            loc_str = str(loc.resolve()).lower().replace('\\', '/')
                        else:
                            loc_str = str(loc.resolve())
                        
                        if loc_str not in seen:
                            seen.add(loc_str)
                            unique_locations.append(loc)
                except Exception:
                    # Bỏ qua nếu không thể resolve
                    continue
            
            # Đọc từ các vị trí theo thứ tự ưu tiên
            for config_file_path in unique_locations:
                max_retries = 3
                retry_count = 0
                while retry_count < max_retries:
                    try:
                        with open(config_file_path, 'r', encoding='utf-8') as f:
                            config = json.load(f)
                        
                        # Format mới: có cả template_file_path, auto_save_folder và config_folder
                        if 'template_file_path' in config and 'auto_save_folder' in config:
                            # Normalize paths cho Windows (quan trọng: phải normalize đúng cách)
                            template_path = None
                            auto_folder = None
                            config_folder = None
                            
                            try:
                                if config['template_file_path']:
                                    template_path = str(Path(config['template_file_path']).resolve())
                            except:
                                pass
                            
                            try:
                                if config['auto_save_folder']:
                                    auto_folder = str(Path(config['auto_save_folder']).resolve())
                            except:
                                pass
                            
                            try:
                                if config.get('config_folder'):
                                    config_folder = str(Path(config['config_folder']).resolve())
                            except:
                                pass
                            
                            # QUAN TRỌNG: Nếu file này có config_folder, đọc file config từ đó (vị trí thực sự)
                            if config_folder:
                                try:
                                    actual_config_file = Path(config_folder) / "kiem_kho_config.json"
                                    # Windows-safe: So sánh đường dẫn case-insensitive
                                    if actual_config_file.exists():
                                        current_path_str = str(config_file_path.resolve())
                                        actual_path_str = str(actual_config_file.resolve())
                                        
                                        # So sánh case-insensitive trên Windows
                                        if sys.platform == 'win32':
                                            if current_path_str.lower() != actual_path_str.lower():
                                                # Đọc lại từ file config thực sự
                                                with open(actual_config_file, 'r', encoding='utf-8') as f2:
                                                    config = json.load(f2)
                                                # Normalize lại paths
                                                if config.get('template_file_path'):
                                                    template_path = str(Path(config['template_file_path']).resolve())
                                                if config.get('auto_save_folder'):
                                                    auto_folder = str(Path(config['auto_save_folder']).resolve())
                                                if config.get('config_folder'):
                                                    config_folder = str(Path(config['config_folder']).resolve())
                                        else:
                                            if current_path_str != actual_path_str:
                                                # Đọc lại từ file config thực sự
                                                with open(actual_config_file, 'r', encoding='utf-8') as f2:
                                                    config = json.load(f2)
                                                # Normalize lại paths
                                                if config.get('template_file_path'):
                                                    template_path = str(Path(config['template_file_path']).resolve())
                                                if config.get('auto_save_folder'):
                                                    auto_folder = str(Path(config['auto_save_folder']).resolve())
                                                if config.get('config_folder'):
                                                    config_folder = str(Path(config['config_folder']).resolve())
                                except Exception as e:
                                    print(f"Lỗi khi đọc file config thực sự: {str(e)}")
                                    # Nếu không đọc được, dùng config từ file đầu tiên
                            
                            # Cập nhật config_folder và config_file
                            if config_folder:
                                self.config_folder = config_folder
                                self.config_file = Path(self.config_folder) / "kiem_kho_config.json"
                            
                            return {
                                'template_file_path': template_path,
                                'auto_save_folder': auto_folder,
                                'config_folder': config_folder
                            }
                        
                        # Format cũ - hỗ trợ backward compatibility
                        elif 'auto_save_folder' in config:
                            auto_folder = str(Path(config['auto_save_folder']).resolve()) if config['auto_save_folder'] else None
                            return {
                                'template_file_path': None,
                                'auto_save_folder': auto_folder,
                                'config_folder': None
                            }
                        
                        break  # Thành công, thoát khỏi loop
                    except PermissionError:
                        retry_count += 1
                        if retry_count < max_retries:
                            time.sleep(0.2)
                        else:
                            # Chuyển sang file tiếp theo
                            break
                    except Exception as e:
                        # Chuyển sang file tiếp theo
                        break
        except Exception as e:
            print(f"Lỗi khi đọc config: {str(e)}")
            traceback.print_exc()
        return None
    
    def save_config(self, template_path, folder_path, config_folder=None):
        """Lưu cấu hình vào file - Windows-safe"""
        try:
            # Normalize paths cho Windows
            template_path_normalized = str(Path(template_path).resolve())
            folder_path_normalized = str(Path(folder_path).resolve())
            
            # Nếu có config_folder, normalize và cập nhật
            if config_folder:
                config_folder_normalized = str(Path(config_folder).resolve())
                self.config_folder = config_folder_normalized
                # Cập nhật lại config_file path
                self.config_file = Path(self.config_folder) / "kiem_kho_config.json"
            
            config = {
                'template_file_path': template_path_normalized,
                'auto_save_folder': folder_path_normalized,
                'config_folder': str(Path(self.config_folder).resolve()) if self.config_folder else None
            }
            
            # Windows-specific: Retry nếu file bị lock
            max_retries = 3
            retry_count = 0
            save_success = False
            
            while retry_count < max_retries and not save_success:
                try:
                    with open(self.config_file, 'w', encoding='utf-8') as f:
                        json.dump(config, f, ensure_ascii=False, indent=2)
                    save_success = True
                except PermissionError:
                    retry_count += 1
                    if retry_count < max_retries:
                        time.sleep(0.2)
                    else:
                        print(f"Lỗi khi lưu config: File bị lock sau {max_retries} lần thử")
                except Exception as e:
                    print(f"Lỗi khi lưu config: {str(e)}")
                    traceback.print_exc()
                    break
        except Exception as e:
            print(f"Lỗi khi lưu config: {str(e)}")
            traceback.print_exc()
    
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
            dialog.geometry("700x480")
            dialog.configure(bg='#F5F5F5')
            dialog.transient(self.root)
            dialog.grab_set()  # Modal dialog
            
            # Đặt cửa sổ ở giữa màn hình
            dialog.update_idletasks()
            x = (dialog.winfo_screenwidth() // 2) - (700 // 2)
            y = (dialog.winfo_screenheight() // 2) - (480 // 2)
            dialog.geometry(f"700x480+{x}+{y}")
            
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
        label_text = "Cấu hình 3 đường dẫn:\n1. File Excel cố định (để copy khi SAVE)\n2. Thư mục lưu file bí mật (Kiemkecuoinam)\n3. Thư mục lưu file cấu hình (kiem_kho_config.json)"
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
            # Tạm thời hạ dialog cấu hình xuống để dialog chọn file có thể hiển thị
            dialog.attributes('-topmost', False)
            dialog.grab_release()
            dialog.update()
            
            file_path = filedialog.askopenfilename(
                title="Chọn file Excel cố định",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            # Khôi phục lại dialog cấu hình lên trên cùng
            dialog.grab_set()
            dialog.attributes('-topmost', True)
            dialog.lift()
            dialog.focus_force()
            dialog.update()
            
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
            # Tạm thời hạ dialog cấu hình xuống để dialog chọn thư mục có thể hiển thị
            dialog.attributes('-topmost', False)
            dialog.grab_release()
            dialog.update()
            
            folder = filedialog.askdirectory(title="Chọn thư mục lưu file bí mật")
            
            # Khôi phục lại dialog cấu hình lên trên cùng
            dialog.grab_set()
            dialog.attributes('-topmost', True)
            dialog.lift()
            dialog.focus_force()
            dialog.update()
            
            if folder:
                path2_var.set(folder)
        
        browse2_btn = tk.Button(input_frame2, text="Chọn thư mục", command=browse_folder2,
                               bg='#C8E6C9', fg='#000000', font=('Arial', 9, 'bold'), 
                               relief=tk.RAISED, bd=2, padx=12, pady=4,
                               activebackground='#A5D6A7', activeforeground='#000000',
                               cursor='hand2')
        browse2_btn.pack(side=tk.RIGHT)
        
        # Đường dẫn 3: Thư mục lưu file config
        tk.Label(dialog, text="3. Thư mục lưu file cấu hình (kiem_kho_config.json):", bg='#F5F5F5', fg='#000000', 
                font=('Arial', 10, 'bold'), anchor='w').pack(pady=(15, 5), padx=20, fill=tk.X)
        
        input_frame3 = tk.Frame(dialog, bg='#F5F5F5')
        input_frame3.pack(pady=5, padx=20, fill=tk.X)
        
        path3_var = tk.StringVar()
        if saved_config and saved_config.get('config_folder'):
            path3_var.set(saved_config['config_folder'])
        elif self.config_folder:
            path3_var.set(self.config_folder)
        else:
            # Mặc định: thư mục chứa DuLieuDauVao.xlsx
            default_config_dir = self.get_config_file_path().parent
            path3_var.set(str(default_config_dir))
        
        path3_entry = tk.Entry(input_frame3, textvariable=path3_var, width=50, 
                              font=('Arial', 10), relief=tk.SOLID, bd=1,
                              bg='#FFFFFF', fg='#000000', insertbackground='#000000')
        path3_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        def browse_folder3():
            # Tạm thời hạ dialog cấu hình xuống để dialog chọn thư mục có thể hiển thị
            dialog.attributes('-topmost', False)
            dialog.grab_release()
            dialog.update()
            
            folder = filedialog.askdirectory(title="Chọn thư mục lưu file cấu hình")
            
            # Khôi phục lại dialog cấu hình lên trên cùng
            dialog.grab_set()
            dialog.attributes('-topmost', True)
            dialog.lift()
            dialog.focus_force()
            dialog.update()
            
            if folder:
                path3_var.set(folder)
        
        browse3_btn = tk.Button(input_frame3, text="Chọn thư mục", command=browse_folder3,
                               bg='#C8E6C9', fg='#000000', font=('Arial', 9, 'bold'), 
                               relief=tk.RAISED, bd=2, padx=12, pady=4,
                               activebackground='#A5D6A7', activeforeground='#000000',
                               cursor='hand2')
        browse3_btn.pack(side=tk.RIGHT)
        
        # Button OK và Cancel
        button_frame = tk.Frame(dialog, bg='#F5F5F5')
        button_frame.pack(pady=20)
        
        def on_ok():
            template_path = path1_var.get().strip()
            folder_path = path2_var.get().strip()
            config_folder_path = path3_var.get().strip()
            
            if not template_path or not folder_path or not config_folder_path:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập đầy đủ 3 đường dẫn!")
                return
            
            # Kiểm tra file Excel có tồn tại không
            if not os.path.isfile(template_path):
                messagebox.showerror("Lỗi", f"File Excel không tồn tại!\n{template_path}")
                return
            
            # Kiểm tra thư mục lưu file bí mật có tồn tại không
            if not os.path.isdir(folder_path):
                messagebox.showerror("Lỗi", f"Thư mục lưu file bí mật không tồn tại!\n{folder_path}")
                return
            
            # Kiểm tra thư mục lưu file config có tồn tại không
            if not os.path.isdir(config_folder_path):
                messagebox.showerror("Lỗi", f"Thư mục lưu file cấu hình không tồn tại!\n{config_folder_path}")
                return
            
            # Kiểm tra quyền ghi vào thư mục lưu file bí mật - Windows-safe
            try:
                test_file = Path(folder_path) / '.kiem_kho_test'
                # Windows-specific: Normalize path
                test_file = test_file.resolve()
                with open(test_file, 'w', encoding='utf-8') as f:
                    f.write('test')
                # Windows-specific: Retry nếu file bị lock
                max_retries = 3
                retry_count = 0
                delete_success = False
                while retry_count < max_retries and not delete_success:
                    try:
                        test_file.unlink()
                        delete_success = True
                    except PermissionError:
                        retry_count += 1
                        if retry_count < max_retries:
                            time.sleep(0.2)
                        else:
                            # File có thể bị lock, nhưng không quan trọng vì chỉ là file test
                            pass
            except Exception as e:
                error_msg = str(e)
                traceback.print_exc()
                messagebox.showerror("Lỗi", f"Không có quyền ghi vào thư mục lưu file bí mật!\n{error_msg}")
                return
            
            # Kiểm tra quyền ghi vào thư mục lưu file config - Windows-safe
            try:
                test_file_config = Path(config_folder_path) / '.kiem_kho_test_config'
                # Windows-specific: Normalize path
                test_file_config = test_file_config.resolve()
                with open(test_file_config, 'w', encoding='utf-8') as f:
                    f.write('test')
                # Windows-specific: Retry nếu file bị lock
                max_retries = 3
                retry_count = 0
                delete_success = False
                while retry_count < max_retries and not delete_success:
                    try:
                        test_file_config.unlink()
                        delete_success = True
                    except PermissionError:
                        retry_count += 1
                        if retry_count < max_retries:
                            time.sleep(0.2)
                        else:
                            # File có thể bị lock, nhưng không quan trọng vì chỉ là file test
                            pass
            except Exception as e:
                error_msg = str(e)
                traceback.print_exc()
                messagebox.showerror("Lỗi", f"Không có quyền ghi vào thư mục lưu file cấu hình!\n{error_msg}")
                return
            
            # Lưu cấu hình (lưu vào thư mục config mới)
            self.save_config(template_path, folder_path, config_folder_path)
            
            # Cập nhật biến (quan trọng: phải cập nhật sau khi save để lần sau load được)
            self.template_file_path = template_path
            self.auto_save_folder = folder_path
            self.config_folder = config_folder_path
            # Cập nhật config_file để trỏ đến đúng vị trí
            self.config_file = Path(self.config_folder) / "kiem_kho_config.json"
            
            # Đảm bảo config đã được lưu thành công
            if self.config_file.exists():
                print(f"[OK] Đã lưu cấu hình vào: {self.config_file}")
            
            dialog.destroy()
        
        def on_cancel():
            # Nếu không nhập đầy đủ và bấm "Bỏ qua", tắt phần mềm
            template_path = path1_var.get().strip()
            folder_path = path2_var.get().strip()
            config_folder_path = path3_var.get().strip()
            if not template_path or not folder_path or not config_folder_path:
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
        path2_entry.bind('<Return>', lambda e: path3_entry.focus())
        path3_entry.bind('<Return>', lambda e: on_ok())
        
        # Focus vào ô đầu tiên
        path1_entry.focus()
        path1_entry.select_range(0, tk.END)
        
        # Đợi dialog đóng
        dialog.wait_window()
        
    def get_original_dist_path(self):
        """Lấy đường dẫn thư mục dist gốc từ file config (được đóng gói trong exe)"""
        if not getattr(sys, 'frozen', False):
            return None
        
        try:
            # Ưu tiên 1: Tìm file config trong thư mục chứa exe (nếu copy cùng với exe)
            exe_dir = Path(sys.executable).parent
            config_file_py = exe_dir / "dist_path_config.py"
            
            # Thử đọc từ file Python trong thư mục exe trước
            if config_file_py.exists():
                try:
                    import importlib.util
                    spec = importlib.util.spec_from_file_location("dist_path_config", config_file_py)
                    if spec and spec.loader:
                        config_module = importlib.util.module_from_spec(spec)
                        spec.loader.exec_module(config_module)
                        if hasattr(config_module, 'DIST_PATH') and config_module.DIST_PATH:
                            dist_path = Path(config_module.DIST_PATH)
                            if dist_path.exists():
                                return dist_path
                except Exception as e:
                    print(f"Lỗi khi đọc dist_path_config.py từ exe dir: {str(e)}")
            
            # Ưu tiên 2: Đọc từ file Python đã đóng gói trong exe (bundle)
            try:
                if getattr(sys, '_MEIPASS', None):
                    config_path = Path(sys._MEIPASS) / "dist_path_config.py"
                    if config_path.exists():
                        import importlib.util
                        spec = importlib.util.spec_from_file_location("dist_path_config", config_path)
                        if spec and spec.loader:
                            config_module = importlib.util.module_from_spec(spec)
                            spec.loader.exec_module(config_module)
                            if hasattr(config_module, 'DIST_PATH') and config_module.DIST_PATH:
                                dist_path = Path(config_module.DIST_PATH)
                                if dist_path.exists():
                                    return dist_path
            except Exception as e:
                print(f"Lỗi khi đọc dist_path_config.py từ bundle: {str(e)}")
            
            # Ưu tiên 3: Tìm file txt trong thư mục chứa exe
            config_file = exe_dir / "dist_path.txt"
            
            # Ưu tiên 3: Tìm trong thư mục cha
            if not config_file.exists():
                parent_dir = exe_dir.parent
                config_file = parent_dir / "dist_path.txt"
            
            # Ưu tiên 4: Tìm trong các thư mục phổ biến
            if not config_file.exists():
                user_home = Path.home()
                possible_locations = [
                    user_home / "Desktop" / "dist" / "dist_path.txt",
                    user_home / "Documents" / "dist" / "dist_path.txt",
                    user_home / "Downloads" / "dist" / "dist_path.txt",
                ]
                for loc in possible_locations:
                    if loc.exists():
                        config_file = loc
                        break
            
            # Đọc đường dẫn từ file txt (backup)
            if config_file.exists():
                with open(config_file, 'r', encoding='utf-8') as f:
                    dist_path = f.read().strip()
                    if dist_path and Path(dist_path).exists():
                        return Path(dist_path)
        except Exception as e:
            print(f"Lỗi khi đọc dist_path config: {str(e)}")
        
        return None
    
    def load_data(self):
        """Load dữ liệu từ file Excel"""
        try:
            # Kiểm tra nếu đang chạy từ executable (PyInstaller)
            if getattr(sys, 'frozen', False):
                # Chạy từ executable
                # ƯU TIÊN: Tìm file từ thư mục dist gốc (nơi build ban đầu)
                original_dist_path = self.get_original_dist_path()
                
                if original_dist_path:
                    # Tìm file từ thư mục dist gốc
                    excel_path = original_dist_path / "DuLieuDauVao.xlsx"
                    search_dir = original_dist_path
                else:
                    # Fallback: Tìm trong thư mục chứa executable
                    exe_dir = Path(sys.executable).parent
                    excel_path = exe_dir / "DuLieuDauVao.xlsx"
                    search_dir = exe_dir
                
                # Nếu không có trong thư mục dist gốc hoặc thư mục exe, thử tìm trong bundle
                if not excel_path.exists():
                    base_path = Path(sys._MEIPASS)
                    excel_path = base_path / "DuLieuDauVao.xlsx"
                
                # Danh sách file thay thế nếu file chính không đọc được
                xls_alternatives = [
                    search_dir / "KIEM KE Năm -2025 - BP ONLINE.xls",
                    search_dir / "KIEM KE Năm -2025 - BP ONLINE copy.xls",
                    search_dir / "DuLieuDauVao.xls",
                ]
                # Fallback: thử trong bundle nếu không tìm thấy
                if getattr(sys, '_MEIPASS', None):
                    base_path = Path(sys._MEIPASS)
                    xls_alternatives.extend([
                        base_path / "KIEM KE Năm -2025 - BP ONLINE.xls",
                        base_path / "KIEM KE Năm -2025 - BP ONLINE copy.xls",
                        base_path / "DuLieuDauVao.xls",
                    ])
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
        
        # Nút RESET
        reset_btn = tk.Button(info_frame, text="RESET", command=self.reset_scanned_data, 
                            bg='#FF9800', fg='white', font=('Arial', 12, 'bold'), 
                            width=15, height=2, relief=tk.RAISED, bd=2, cursor='hand2')
        reset_btn.grid(row=3, column=3, rowspan=2, padx=20, pady=5, sticky='n')
        
        info_frame.columnconfigure(1, weight=1)
        
        # === PHẦN HIỂN THỊ SỐ TỰA ===
        count_frame = tk.Frame(main_frame, bg=bg_color)
        count_frame.pack(fill=tk.X, pady=(0, 5))
        
        tk.Label(count_frame, text="Số tựa:", bg=bg_color, fg=label_required_fg, font=('Arial', 11, 'bold')).grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.so_tua_var = tk.StringVar(value="0")
        tk.Label(count_frame, textvariable=self.so_tua_var, bg=bg_color, fg='#1976D2', font=('Arial', 14, 'bold')).grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        # Hiển thị số tựa đã quét
        tk.Label(count_frame, text="Đã quét:", bg=bg_color, fg=label_required_fg, font=('Arial', 11, 'bold')).grid(row=0, column=2, padx=(20, 5), pady=5, sticky='w')
        self.so_tua_da_quet_var = tk.StringVar(value="0")
        tk.Label(count_frame, textvariable=self.so_tua_da_quet_var, bg=bg_color, fg='#4CAF50', font=('Arial', 14, 'bold')).grid(row=0, column=3, padx=5, pady=5, sticky='w')
        
        # === BẢNG DỮ LIỆU ===
        table_frame = tk.Frame(main_frame, bg=bg_color)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tạo Treeview với scrollbar
        scrollbar_y = tk.Scrollbar(table_frame, orient=tk.VERTICAL, bg='#E0E0E0', troughcolor=bg_color)
        scrollbar_x = tk.Scrollbar(table_frame, orient=tk.HORIZONTAL, bg='#E0E0E0', troughcolor=bg_color)
        
        # Định nghĩa thứ tự cột cố định
        columns = ('ISBN', 'Tựa', 'Tồn thực tế', 'Số thùng', 'Tồn tựa trong thùng', 'Tình trạng', 'Ghi chú')
        # Thứ tự: 0=ISBN, 1=Tựa, 2=Tồn thực tế, 3=Số thùng, 4=Tồn tựa trong thùng, 5=Tình trạng, 6=Ghi chú
        
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
        self.tree.heading('Tình trạng', text='Tình trạng')
        self.tree.heading('Ghi chú', text='Ghi chú')
        
        self.tree.column('ISBN', width=150, anchor='w')
        self.tree.column('Tựa', width=300, anchor='w')
        self.tree.column('Tồn thực tế', width=120, anchor='center')
        self.tree.column('Số thùng', width=100, anchor='center')
        self.tree.column('Tồn tựa trong thùng', width=150, anchor='center')
        self.tree.column('Tình trạng', width=100, anchor='center')
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
                font=('Arial', 13, 'bold')).grid(row=0, column=0, padx=5, pady=8, sticky='w')
        self.isbn_entry = tk.Entry(scan_frame, font=('Arial', 16, 'bold'), width=30, 
                                   bg='#FFFFFF', fg='#000000', relief=tk.SOLID, bd=3, insertbackground='#000000')
        # Tăng chiều cao bằng cách thêm padding
        self.isbn_entry.grid(row=0, column=1, padx=5, pady=8, sticky='ew', ipady=8)
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
        tonghop_columns = ('N/X', 'Số phiếu', 'Ngày', 'Vị trí mới', 'ISBN', 'Tựa', 'Tồn thực tế', 'Số thùng', 'Tình trạng', 'Ghi chú')
        
        # Tạo Treeview cho tổng hợp với tối ưu hiệu suất
        self.tong_hop_tree = ttk.Treeview(tonghop_table_frame, columns=tonghop_columns, show='headings', 
                                          yscrollcommand=tonghop_scrollbar_y.set, xscrollcommand=tonghop_scrollbar_x.set,
                                          height=20, style='Treeview')
        
        # Tối ưu hiệu suất: tắt một số tính năng không cần thiết
        # Không cần selectmode vì không có selection logic phức tạp
        
        # Cấu hình các cột
        column_widths = {'N/X': 250, 'Số phiếu': 120, 'Ngày': 100, 'Vị trí mới': 120, 'ISBN': 150, 
                        'Tựa': 250, 'Tồn thực tế': 100, 'Số thùng': 120, 'Tình trạng': 100, 'Ghi chú': 200}
        
        for col in tonghop_columns:
            self.tong_hop_tree.heading(col, text=col)
            # Tối ưu: không stretch, không minwidth để tăng tốc độ render và scroll
            self.tong_hop_tree.column(col, width=column_widths.get(col, 100), anchor='w', 
                                     stretch=False, minwidth=0)
        
        # Tối ưu scrollbar để scroll mượt hơn - không có delay
        # Sử dụng trực tiếp method của treeview để giảm overhead
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
            
            # Đếm số tựa đã quét từ tab Tổng hợp cho thùng này
            so_tua_da_quet = self.count_scanned_titles_for_box(so_thung)
            
            # Hiển thị số tựa tổng và số tựa đã quét
            self.so_tua_var.set(str(len(self.current_box_data)))
            if hasattr(self, 'so_tua_da_quet_var') and self.so_tua_da_quet_var:
                self.so_tua_da_quet_var.set(str(so_tua_da_quet))
            
            # Clear bảng
            self.clear_table()
            
            # Thông báo thành công với thông tin số tựa đã quét
            if so_tua_da_quet > 0:
                messagebox.showinfo("Thành công", 
                    f"Đã load {len(self.current_box_data)} tựa cho thùng số {so_thung}\n\n"
                    f"Đã quét: {so_tua_da_quet} tựa (đã lưu trong Tổng hợp)\n"
                    f"Còn lại: {len(self.current_box_data) - so_tua_da_quet} tựa")
            else:
                messagebox.showinfo("Thành công", f"Đã load {len(self.current_box_data)} tựa cho thùng số {so_thung}")
            
            # Focus vào ô nhập ISBN để sẵn sàng quét
            self.isbn_entry.focus()
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể load dữ liệu thùng: {str(e)}")
    
    def count_scanned_titles_for_box(self, so_thung):
        """Đếm số tựa đã quét từ tab Tổng hợp cho một thùng cụ thể"""
        if not so_thung or not self.tong_hop_data:
            return 0
        
        so_thung_clean = str(so_thung).strip()
        count = 0
        
        for data in self.tong_hop_data:
            # So sánh số thùng (có thể là 'Số thùng' hoặc 'Vị trí mới')
            so_thung_in_data = str(data.get('Số thùng', '')).strip()
            vi_tri_moi_in_data = str(data.get('Vị trí mới', '')).strip()
            
            # Đếm nếu số thùng khớp
            if so_thung_in_data == so_thung_clean or vi_tri_moi_in_data == so_thung_clean:
                count += 1
        
        return count
    
    def is_isbn_already_scanned(self, isbn, so_thung):
        """Kiểm tra xem ISBN đã được quét và lưu trong tab Tổng hợp cho thùng này chưa - tối ưu"""
        if not isbn or not so_thung or not self.tong_hop_data:
            return False
        
        try:
            isbn_clean = str(isbn).strip()
            isbn_clean_digits = ''.join(filter(str.isdigit, isbn_clean))
            so_thung_clean = str(so_thung).strip()
            
            # Tối ưu: chỉ kiểm tra các item có số thùng khớp trước
            for data in self.tong_hop_data:
                # Kiểm tra số thùng trước (nhanh hơn)
                so_thung_in_data = str(data.get('Số thùng', '')).strip()
                vi_tri_moi_in_data = str(data.get('Vị trí mới', '')).strip()
                
                # Nếu số thùng không khớp, bỏ qua ngay
                if so_thung_in_data != so_thung_clean and vi_tri_moi_in_data != so_thung_clean:
                    continue
                
                # Kiểm tra ISBN chỉ khi số thùng khớp
                isbn_in_data = str(data.get('ISBN', '')).strip()
                
                # So sánh nhanh: khớp chính xác trước
                if isbn_in_data == isbn_clean:
                    return True
                
                # So sánh với digits
                isbn_in_data_digits = ''.join(filter(str.isdigit, isbn_in_data))
                if isbn_in_data_digits and isbn_clean_digits and isbn_in_data_digits == isbn_clean_digits:
                    return True
                
                # So sánh với endswith/startswith (chậm hơn, chỉ khi cần)
                if (isbn_in_data.endswith(isbn_clean) or isbn_clean.endswith(isbn_in_data)):
                    return True
            
            return False
        except Exception as e:
            # Nếu có lỗi, trả về False để không block quét
            print(f"Lỗi khi kiểm tra ISBN đã quét: {str(e)}")
            return False
    
    def on_isbn_entered(self, event=None):
        """Xử lý khi nhập/quét ISBN - tối ưu để tránh freeze"""
        try:
            isbn = self.isbn_entry.get().strip()
            if not isbn:
                return
            
            if self.current_box_data is None or self.current_box_data.empty:
                # Sử dụng after để không block UI
                self.root.after(10, lambda: messagebox.showwarning("Cảnh báo", "Vui lòng nhập số thùng và load dữ liệu trước!"))
                self.isbn_entry.delete(0, tk.END)
                return
            
            # Kiểm tra nếu đã quét đủ số tựa trong thùng
            so_tua_trong_thung = len(self.current_box_data)
            so_tua_da_quet = len(self.scanned_items)
            
            if so_tua_da_quet >= so_tua_trong_thung:
                # Fix closure issue: capture giá trị vào biến local
                box_num = self.current_box_number
                self.root.after(10, lambda t=so_tua_trong_thung, d=so_tua_da_quet, b=box_num: messagebox.showerror(
                    "Lỗi", 
                    f"Đã quét đủ số tựa trong thùng!\n\n"
                    f"Thùng {b} có {t} tựa.\n"
                    f"Bạn đã quét {d} tựa.\n\n"
                    "Vui lòng lưu dữ liệu hoặc load thùng khác để tiếp tục."
                ))
                self.isbn_entry.delete(0, tk.END)
                return
            
            # Tìm tựa trong dữ liệu thùng hiện tại - tối ưu với vectorization
            if 'isbn' in self.current_box_data.columns:
                isbn_clean = str(isbn).strip()
                isbn_clean_digits = ''.join(filter(str.isdigit, isbn_clean))
                matched_row = None
                
                # Tối ưu: sử dụng vectorization thay vì iterrows() (nhanh hơn nhiều)
                try:
                    # Chuyển đổi ISBN sang số để so sánh nhanh hơn
                    isbn_col = self.current_box_data['isbn'].astype(str).str.strip()
                    
                    # Tìm khớp chính xác trước (nhanh nhất)
                    exact_match = isbn_col == isbn_clean
                    if exact_match.any():
                        matched_row = self.current_box_data[exact_match].iloc[0]
                    else:
                        # Tìm khớp với digits
                        isbn_col_digits = isbn_col.str.replace(r'\D', '', regex=True)
                        digit_match = isbn_col_digits == isbn_clean_digits
                        if digit_match.any():
                            matched_row = self.current_box_data[digit_match].iloc[0]
                        else:
                            # Tìm khớp với endswith/startswith (chậm hơn nhưng cần thiết)
                            for idx, row in self.current_box_data.iterrows():
                                row_isbn = str(row.get('isbn', '')).strip()
                                row_isbn_clean = ''.join(filter(str.isdigit, row_isbn))
                                
                                if (row_isbn == isbn_clean or 
                                    row_isbn.endswith(isbn_clean) or 
                                    isbn_clean.endswith(row_isbn) or
                                    (row_isbn_clean and isbn_clean_digits and row_isbn_clean == isbn_clean_digits)):
                                    matched_row = row
                                    break
                except Exception as e:
                    # Fallback về cách cũ nếu vectorization lỗi
                    print(f"Lỗi khi tìm ISBN với vectorization: {str(e)}")
                    for idx, row in self.current_box_data.iterrows():
                        row_isbn = str(row.get('isbn', '')).strip()
                        row_isbn_clean = ''.join(filter(str.isdigit, row_isbn))
                        
                        if (row_isbn == isbn_clean or 
                            row_isbn.endswith(isbn_clean) or 
                            isbn_clean.endswith(row_isbn) or
                            (row_isbn_clean and isbn_clean_digits and row_isbn_clean == isbn_clean_digits)):
                            matched_row = row
                            break
                
                # Kiểm tra xem ISBN này đã được quét và lưu trong tab Tổng hợp chưa
                if matched_row is not None:
                    if self.is_isbn_already_scanned(isbn_clean, self.current_box_number):
                        # Fix closure issue: capture giá trị vào biến local
                        isbn_msg = isbn_clean
                        box_msg = self.current_box_number
                        self.root.after(10, lambda i=isbn_msg, b=box_msg: messagebox.showwarning(
                            "Cảnh báo",
                            f"ISBN {i} đã được quét và lưu trong tab Tổng hợp cho thùng {b}!\n\n"
                            "Vui lòng không quét lại ISBN đã được lưu."
                        ))
                        self.isbn_entry.delete(0, tk.END)
                        return
                
                if matched_row is None:
                    # Kiểm tra xem ISBN có tồn tại trong thùng khác không - tối ưu
                    found_in_other_box = False
                    other_box_number = None
                    
                    # Tìm trong toàn bộ dữ liệu - chỉ tìm nếu cần thiết
                    if 'isbn' in self.df.columns:
                        # Tìm cột số thùng (cache nếu có thể)
                        so_thung_col = None
                        for col in self.df.columns:
                            col_lower = str(col).lower().strip()
                            if 'số thùng' in col_lower or 'so thung' in col_lower or col_lower == 'thùng' or col_lower == 'so_thung':
                                so_thung_col = col
                                break
                        
                        if so_thung_col:
                            # Tối ưu: chỉ tìm trong một số dòng đầu tiên hoặc sử dụng vectorization
                            try:
                                # Tìm với vectorization (nhanh hơn)
                                isbn_col_df = self.df['isbn'].astype(str).str.strip()
                                exact_match_df = isbn_col_df == isbn_clean
                                if exact_match_df.any():
                                    matched_idx = exact_match_df.idxmax()
                                    other_box_number = str(self.df.loc[matched_idx, so_thung_col]).strip()
                                    if other_box_number and other_box_number != str(self.current_box_number):
                                        found_in_other_box = True
                                else:
                                    # Tìm với digits
                                    isbn_col_digits_df = isbn_col_df.str.replace(r'\D', '', regex=True)
                                    digit_match_df = isbn_col_digits_df == isbn_clean_digits
                                    if digit_match_df.any():
                                        matched_idx = digit_match_df.idxmax()
                                        other_box_number = str(self.df.loc[matched_idx, so_thung_col]).strip()
                                        if other_box_number and other_box_number != str(self.current_box_number):
                                            found_in_other_box = True
                            except Exception as e:
                                # Fallback: chỉ tìm trong 1000 dòng đầu để tránh chậm
                                print(f"Lỗi khi tìm ISBN trong df: {str(e)}")
                                max_search = min(1000, len(self.df))
                                for idx in range(max_search):
                                    row = self.df.iloc[idx]
                                    row_isbn = str(row.get('isbn', '')).strip()
                                    row_isbn_clean = ''.join(filter(str.isdigit, row_isbn))
                                    
                                    if (row_isbn == isbn_clean or 
                                        row_isbn.endswith(isbn_clean) or 
                                        isbn_clean.endswith(row_isbn) or
                                        (row_isbn_clean and isbn_clean_digits and row_isbn_clean == isbn_clean_digits)):
                                        other_box_number = str(row.get(so_thung_col, '')).strip()
                                        if other_box_number and other_box_number != str(self.current_box_number):
                                            found_in_other_box = True
                                        break
                    
                    # Báo lỗi tương ứng - sử dụng after để không block
                    # Fix closure issue: capture giá trị vào biến local
                    isbn_msg = isbn
                    other_box = other_box_number
                    current_box = self.current_box_number
                    if found_in_other_box:
                        self.root.after(10, lambda i=isbn_msg, o=other_box, c=current_box: messagebox.showerror(
                            "Lỗi", 
                            f"ISBN {i} không thuộc thùng đang kiểm kê!\n\n"
                            f"ISBN này thuộc thùng: {o}\n"
                            f"Thùng đang kiểm kê: {c}\n\n"
                            f"Vui lòng quét đúng ISBN của thùng {c}."
                        ))
                    else:
                        self.root.after(10, lambda i=isbn_msg, c=current_box: messagebox.showwarning(
                            "Cảnh báo", 
                            f"Không tìm thấy ISBN {i} trong dữ liệu!\n\n"
                            f"Vui lòng kiểm tra lại mã ISBN hoặc thùng số {c}."
                        ))
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
                        # Fix closure issue: capture giá trị vào biến local
                        vi_tri_msg = vi_tri_moi
                        existing_list = ', '.join(sorted(existing_box_numbers)[:10])
                        existing_count = len(existing_box_numbers)
                        existing_suffix = f" và {existing_count - 10} mã khác..." if existing_count > 10 else ""
                        self.root.after(10, lambda v=vi_tri_msg, e=existing_list, s=existing_suffix: messagebox.showerror(
                            "Lỗi", 
                            f"Mã thùng mới '{v}' đã tồn tại trong dữ liệu đầu vào!\n\n"
                            f"Vui lòng nhập mã thùng khác với các mã thùng hiện có.\n\n"
                            f"Các mã thùng hiện có: {e}{s}"
                        ))
                        self.isbn_entry.delete(0, tk.END)
                        return
                    
                    so_thung_hien_thi = vi_tri_moi
                else:
                    so_thung_hien_thi = so_thung
                
                # Đảm bảo thứ tự đúng với columns: ISBN, Tựa, Tồn thực tế, Số thùng, Tồn tựa trong thùng, Tình trạng, Ghi chú
                item_id = self.tree.insert('', tk.END, values=(
                    str(isbn_clean),           # 0: ISBN
                    str(tua),                  # 1: Tựa
                    '',                        # 2: Tồn thực tế - để trống để người dùng nhập
                    str(so_thung_hien_thi),    # 3: Số thùng (dùng vị trí mới nếu có)
                    str(ton_trong_thung_display),  # 4: Tồn tựa trong thùng
                    '',                        # 5: Tình trạng - sẽ tự động điền khi có lệch
                    ''                         # 6: Ghi chú - để trống để người dùng tự nhập
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
                    'tinh_trang': '',  # Tình trạng sẽ tự động điền khi có lệch
                    'ghi_chu': ''  # Ghi chú để người dùng tự nhập
                }
                
                # Cập nhật số tựa đã quét (chỉ hiển thị số tựa đã lưu trong Tổng hợp)
                if hasattr(self, 'so_tua_da_quet_var') and self.so_tua_da_quet_var and self.current_box_number:
                    so_tua_da_quet = self.count_scanned_titles_for_box(self.current_box_number)
                    self.so_tua_da_quet_var.set(str(so_tua_da_quet))
                
                # Focus vào cột "Tồn thực tế" để người dùng nhập
                self.tree.selection_set(item_id)
                self.tree.focus(item_id)
                self.tree.see(item_id)
                
                # Tự động focus vào ô "Tồn thực tế" để sẵn sàng nhập
                # Fix closure issue: capture item_id vào biến local
                item_id_to_edit = item_id
                self.root.after(100, lambda i=item_id_to_edit: self.auto_edit_ton_thuc_te(i))
                
            else:
                self.root.after(10, lambda: messagebox.showerror("Lỗi", "Không tìm thấy cột 'ISBN' trong dữ liệu!"))
            
        except Exception as e:
            # Xử lý lỗi để tránh crash
            error_msg = str(e)
            print(f"Lỗi khi quét ISBN: {error_msg}")
            import traceback
            traceback.print_exc()
            # Fix closure issue: capture error_msg vào biến local
            self.root.after(10, lambda e=error_msg: messagebox.showerror("Lỗi", f"Lỗi khi quét ISBN: {e}"))
        
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
        
        # Cho phép edit: Tồn thực tế (2), Số thùng (3), Tồn tựa trong thùng (4), Ghi chú (6)
        # Không cho edit: ISBN (0), Tựa (1), Tình trạng (5) - chỉ đọc
        if column_index not in [2, 3, 4, 6]:
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
        # Tránh xử lý 2 lần nếu đang trong quá trình xử lý
        if self.is_processing_edit:
            return
        
        if not self.edit_entry or not self.editing_item:
            return
        
        # Đặt flag để tránh xử lý lại
        self.is_processing_edit = True
        
        try:
            new_value = self.edit_entry.get().strip()
            item = self.editing_item
            column_index = self.editing_column
            
            # Lấy giá trị hiện tại và đảm bảo có đủ 7 cột
            values = list(self.tree.item(item, 'values'))
            while len(values) < 7:
                values.append('')
            
            isbn = values[0] if len(values) > 0 else ''
            
            # Nếu đang edit cột "Ghi chú", đảm bảo lấy giá trị từ scanned_items để không mất dữ liệu
            if column_index == 6 and isbn in self.scanned_items:
                saved_ghi_chu = self.scanned_items[isbn].get('ghi_chu', '')
                # Nếu có giá trị đã lưu, đảm bảo values[6] có giá trị đúng
                if saved_ghi_chu:
                    if len(values) <= 6:
                        while len(values) < 7:
                            values.append('')
                    values[6] = saved_ghi_chu
            
            # Xử lý theo từng cột
            if column_index == 2:  # Tồn thực tế
                values[2] = new_value  # Đảm bảo đúng index
                
                # Kiểm tra và highlight nếu khác nhau - CHỈ chạy cho cột Tồn thực tế
                if isbn in self.scanned_items:
                    self.scanned_items[isbn]['ton_thuc_te'] = new_value
                    ton_trong_thung = self.scanned_items[isbn]['ton_trong_thung']
                    
                    try:
                        ton_thuc_te_num = float(new_value) if new_value else 0
                        ton_trong_thung_num = float(ton_trong_thung) if ton_trong_thung else 0
                        
                        # Kiểm tra lệch
                        if abs(ton_thuc_te_num - ton_trong_thung_num) > 0.01:
                            # Tự động điền "Thiếu" hoặc "Dư" vào cột Tình trạng (index 5)
                            if ton_thuc_te_num < ton_trong_thung_num:
                                tinh_trang = "Thiếu"
                            else:
                                tinh_trang = "Dư"
                            
                            # Đảm bảo có đủ 7 cột và đúng thứ tự: ISBN, Tựa, Tồn thực tế, Số thùng, Tồn tựa trong thùng, Tình trạng, Ghi chú
                            while len(values) < 7:
                                values.append('')
                            
                            # Đảm bảo thứ tự đúng: values[0]=ISBN, values[1]=Tựa, values[2]=Tồn thực tế, 
                            # values[3]=Số thùng, values[4]=Tồn tựa trong thùng, values[5]=Tình trạng, values[6]=Ghi chú
                            values[5] = tinh_trang  # Tình trạng ở index 5
                            # Không tự động điền vào cột Ghi chú (index 6) - để người dùng tự nhập
                            
                            # Cập nhật scanned_items
                            self.scanned_items[isbn]['tinh_trang'] = tinh_trang
                            
                            # Cập nhật tree
                            self.tree.item(item, values=values)
                            
                            # Tô đỏ 2 ô: Tồn thực tế (cột 2) và Tình trạng (cột 5)
                            self.highlight_error_cells(item)
                            
                            # Không hiển thị cảnh báo nữa vì đã có cột "Tình trạng" để hiển thị
                        else:
                            # Không có lỗi - xóa highlight và tình trạng
                            # Đảm bảo có đủ 7 cột
                            while len(values) < 7:
                                values.append('')
                            
                            # Xóa tình trạng nếu có
                            if len(values) > 5:
                                values[5] = ''
                                if 'tinh_trang' in self.scanned_items[isbn]:
                                    del self.scanned_items[isbn]['tinh_trang']
                            
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
                        # Reset flag và cleanup trước khi return
                        self.is_processing_edit = False
                        if self.edit_entry:
                            self.edit_entry.destroy()
                            self.edit_entry = None
                            self.editing_item = None
                        return
                
                values[3] = new_value  # Đảm bảo đúng index
                if isbn in self.scanned_items:
                    # Khi chỉnh sửa trực tiếp, cập nhật số thùng hiển thị
                    # Nhưng giữ nguyên số thùng gốc (từ dữ liệu đầu vào) - KHÔNG BAO GIỜ thay đổi
                    new_value_clean = new_value.strip()
                    self.scanned_items[isbn]['so_thung'] = new_value_clean
                else:
                    # Nếu ISBN không có trong scanned_items, không làm gì
                    pass
                
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
            
            elif column_index == 6:  # Ghi chú
                # Đảm bảo có đủ 7 cột
                while len(values) < 7:
                    values.append('')
                # Chỉ lưu giá trị mới vào values và scanned_items - không có logic phức tạp
                values[6] = new_value
                if isbn in self.scanned_items:
                    self.scanned_items[isbn]['ghi_chu'] = new_value
                # Cập nhật tree với giá trị mới
                self.tree.item(item, values=values)
                # Đảm bảo không có highlight nào che mất nội dung cột "Ghi chú"
                # (highlight chỉ được tạo cho cột "Tồn thực tế" và "Tình trạng", không phải "Ghi chú")
                # Return ngay để không chạy phần cập nhật tree chung bên dưới
                return
            
            # Cập nhật tree với giá trị mới (chỉ cho các cột khác, không phải Ghi chú)
            self.tree.item(item, values=values)
            
            # Nếu là cột Tồn thực tế, kiểm tra lại và cập nhật highlight
            if column_index == 2:
                if isbn in self.scanned_items:
                    ton_trong_thung = self.scanned_items[isbn]['ton_trong_thung']
                    try:
                        ton_thuc_te_num = float(new_value) if new_value else 0
                        ton_trong_thung_num = float(ton_trong_thung) if ton_trong_thung else 0
                        if abs(ton_thuc_te_num - ton_trong_thung_num) > 0.01:
                            # Vẫn còn lệch - tự động cập nhật tình trạng
                            if ton_thuc_te_num < ton_trong_thung_num:
                                tinh_trang = "Thiếu"
                            else:
                                tinh_trang = "Dư"
                            # Đảm bảo có đủ 7 cột và giữ nguyên giá trị Ghi chú (index 6)
                            while len(values) < 7:
                                values.append('')
                            ghi_chu_backup = values[6] if len(values) > 6 else ''  # Backup giá trị Ghi chú
                            values[5] = tinh_trang  # Tình trạng ở index 5
                            values[6] = ghi_chu_backup  # Giữ nguyên giá trị Ghi chú
                            self.scanned_items[isbn]['tinh_trang'] = tinh_trang
                            self.tree.item(item, values=values)
                            self.highlight_error_cells(item)
                        else:
                            # Đã khớp - xóa tình trạng nhưng giữ nguyên Ghi chú
                            while len(values) < 7:
                                values.append('')
                            ghi_chu_backup = values[6] if len(values) > 6 else ''  # Backup giá trị Ghi chú
                            values[5] = ''  # Xóa tình trạng
                            values[6] = ghi_chu_backup  # Giữ nguyên giá trị Ghi chú
                            if 'tinh_trang' in self.scanned_items[isbn]:
                                del self.scanned_items[isbn]['tinh_trang']
                            self.tree.item(item, values=values)
                            self.remove_error_highlights(item)
                    except:
                        self.remove_error_highlights(item)
        finally:
            # Reset flag và cleanup
            self.is_processing_edit = False
            self._finish_edit_scheduled = False  # Reset flag finish_edit
            # Xóa Entry widget
            if self.edit_entry:
                try:
                    self.edit_entry.destroy()
                except:
                    pass  # Widget có thể đã bị destroy
                self.edit_entry = None
                self.editing_item = None
            # Focus lại vào ISBN entry
            try:
                self.isbn_entry.focus()
            except:
                pass  # Widget có thể không tồn tại
    
    def cancel_edit(self):
        """Hủy việc chỉnh sửa"""
        if self.edit_entry:
            try:
                self.edit_entry.destroy()
            except:
                pass  # Widget có thể đã bị destroy
            self.edit_entry = None
            self.editing_item = None
        self.is_processing_edit = False  # Reset flag khi hủy
        self._finish_edit_scheduled = False  # Reset flag finish_edit
        try:
            self.isbn_entry.focus()
        except:
            pass  # Widget có thể không tồn tại
    
    def highlight_error_cells(self, item_id):
        """Tô đỏ 2 ô: Tồn thực tế và Tình trạng"""
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
            # Fix closure issue: capture item_id và column vào biến local
            item_id_click = item_id
            highlight1.bind('<Button-1>', lambda e, i=item_id_click, c='#3': self.on_highlight_click(e, i, c))
        
        # Tô đỏ ô "Tình trạng" (cột 6, index 5)
        bbox_tinh_trang = self.tree.bbox(item_id, '#6')
        if bbox_tinh_trang:
            x, y, width, height = bbox_tinh_trang
            tinh_trang_value = values[5] if len(values) > 5 else ''
            highlight2 = tk.Label(self.tree, bg='#FFCDD2', fg='#C62828', 
                                 text=str(tinh_trang_value), font=('Arial', 10, 'bold'), 
                                 relief=tk.FLAT, anchor='center')
            highlight2.place(x=x, y=y, width=width, height=height)
            # Không cho phép edit cột Tình trạng (chỉ đọc)
        
        # Lưu các highlight widgets
        if item_id not in self.error_highlights:
            self.error_highlights[item_id] = []
        if bbox_ton:
            self.error_highlights[item_id].append(highlight1)
        if bbox_tinh_trang:
            self.error_highlights[item_id].append(highlight2)
        
        # Cập nhật lại highlight khi scroll hoặc resize
        # Fix closure issue: capture item_id vào biến local
        item_id_to_update = item_id
        self.root.after(100, lambda i=item_id_to_update: self.update_error_highlights(i))
    
    def on_highlight_click(self, event, item_id, column):
        """Xử lý click vào highlight để edit cell"""
        # Hủy edit cũ nếu có
        if self.edit_entry:
            self.finish_edit()
        
        column_index = int(column.replace('#', '')) - 1
        
        # Không cho phép edit cột Tình trạng (5) - chỉ đọc
        if column_index == 5:
            return
        
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
            
            # Cập nhật ô Tình trạng
            bbox_tinh_trang = self.tree.bbox(item_id, '#6')
            if bbox_tinh_trang and len(widgets) > 1 and widgets[1].winfo_exists():
                x, y, width, height = bbox_tinh_trang
                tinh_trang_value = values[5] if len(values) > 5 else ''
                widgets[1].config(text=str(tinh_trang_value))
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
        
        # Reset flag finish_edit khi bắt đầu edit mới
        self._finish_edit_scheduled = False
        
        def finish_on_enter(event):
            if not self._finish_edit_scheduled:
                self._finish_edit_scheduled = True
                self.finish_edit()
        
        def finish_on_focus_out(event):
            if not self._finish_edit_scheduled:
                self._finish_edit_scheduled = True
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
    
    def reset_scanned_data(self):
        """Reset lại tất cả dữ liệu đã quét và nhập để làm lại"""
        # Xác nhận với người dùng
        result = messagebox.askyesno(
            "Xác nhận Reset", 
            "Bạn có chắc chắn muốn reset lại tất cả dữ liệu đã quét?\n\n"
            "Tất cả các ISBN đã quét và dữ liệu nhập sẽ bị xóa.\n"
            "Các thông tin như Số thùng, Ngày, Tổ sẽ được giữ lại."
        )
        
        if not result:
            return
        
        # Xóa tất cả items trong bảng
        self.clear_table()
        
        # Xóa tất cả scanned items
        self.scanned_items = {}
        
        # Reset số tựa về 0
        if hasattr(self, 'so_tua_var'):
            self.so_tua_var.set("0")
        
        # Xóa input ISBN
        if hasattr(self, 'isbn_entry'):
            self.isbn_entry.delete(0, tk.END)
            self.isbn_entry.focus()
        
        # Hủy edit nếu đang edit
        if self.edit_entry:
            self.cancel_edit()
        
        # Thông báo thành công
        messagebox.showinfo("Thành công", "Đã reset lại tất cả dữ liệu đã quét.\nBạn có thể bắt đầu quét lại từ đầu.")
    
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
        
        # Thêm dữ liệu vào tổng hợp - tối ưu cho dữ liệu lớn
        items_to_add = []
        for isbn, info in self.scanned_items.items():
            try:
                # Kiểm tra xem item có tồn thực tế chưa
                if not info.get('ton_thuc_te', '').strip():
                    continue  # Bỏ qua item chưa nhập tồn thực tế
                
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
                
                # Thêm vào danh sách để append sau (hiệu quả hơn)
                items_to_add.append({
                    'N/X': nx_value,
                    'Số phiếu': so_phieu,
                    'Ngày': ngay_value,
                    'Vị trí mới': so_thung_moi if so_thung_moi else so_thung_goc_clean,
                    'ISBN': isbn,
                    'Tựa': info.get('tua', ''),
                    'Tồn thực tế': info.get('ton_thuc_te', ''),
                    'Số thùng': so_thung_goc_clean,
                    'Tình trạng': info.get('tinh_trang', ''),  # Lấy từ scanned_items
                    'Ghi chú': info.get('ghi_chu', '')  # Ghi chú do người dùng tự nhập
                })
            except Exception as e:
                # Bỏ qua item lỗi và tiếp tục
                print(f"Lỗi khi xử lý item {isbn}: {str(e)}")
                continue
        
        # Thêm tất cả items vào tổng hợp cùng lúc (hiệu quả hơn append từng cái)
        self.tong_hop_data.extend(items_to_add)
        
        # Lưu số thùng hiện tại trước khi reset để cập nhật số tựa đã quét
        saved_box_number = self.current_box_number
        
        # Cập nhật bảng tổng hợp (có xử lý lỗi và batch processing bên trong)
        self.update_tong_hop_table()
        
        # Cập nhật số tựa đã quét sau khi lưu (trước khi reset current_box_number)
        if saved_box_number and hasattr(self, 'so_tua_da_quet_var') and self.so_tua_da_quet_var:
            so_tua_da_quet = self.count_scanned_titles_for_box(saved_box_number)
            self.so_tua_da_quet_var.set(str(so_tua_da_quet))
        
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
        
        messagebox.showinfo("Thành công", f"Đã lưu {len(items_to_add)} dòng mới vào Tổng hợp!\nTổng cộng: {len(self.tong_hop_data)} dòng")
    
    def update_tong_hop_table(self):
        """Cập nhật bảng tổng hợp - tối ưu cho dữ liệu lớn"""
        if not self.tong_hop_tree:
            return
        
        try:
            # Tắt cập nhật UI để tăng tốc độ
            self.tong_hop_tree.config(cursor='wait')
            self.root.config(cursor='wait')
            
            # Xóa tất cả items cũ - sử dụng cách nhanh hơn
            children = self.tong_hop_tree.get_children()
            if children:
                self.tong_hop_tree.delete(*children)
            
            # Tắt update trong quá trình insert để tăng tốc độ
            self.root.update_idletasks()
            
            # Kiểm tra số lượng dữ liệu
            total_items = len(self.tong_hop_data)
            
            # Với dữ liệu lớn (>1000 items), sử dụng batch insert để tránh freeze UI
            if total_items > 1000:
                # Batch insert với update_idletasks mỗi 100 items để không freeze UI
                batch_size = 100
                for i in range(0, total_items, batch_size):
                    batch = self.tong_hop_data[i:i + batch_size]
                    for data in batch:
                        try:
                            self.tong_hop_tree.insert('', 'end', values=(
                                data.get('N/X', ''),
                                data.get('Số phiếu', ''),
                                data.get('Ngày', ''),
                                data.get('Vị trí mới', ''),
                                data.get('ISBN', ''),
                                data.get('Tựa', ''),
                                data.get('Tồn thực tế', ''),
                                data.get('Số thùng', ''),
                                data.get('Tình trạng', ''),
                                data.get('Ghi chú', '')
                            ))
                        except Exception as e:
                            # Bỏ qua item lỗi và tiếp tục
                            print(f"Lỗi khi insert item: {str(e)}")
                            continue
                    
                    # Update UI mỗi batch để không freeze
                    if i + batch_size < total_items:
                        self.root.update_idletasks()
            else:
                # Với dữ liệu nhỏ, insert trực tiếp
                for data in self.tong_hop_data:
                    try:
                        self.tong_hop_tree.insert('', 'end', values=(
                            data.get('N/X', ''),
                            data.get('Số phiếu', ''),
                            data.get('Ngày', ''),
                            data.get('Vị trí mới', ''),
                            data.get('ISBN', ''),
                            data.get('Tựa', ''),
                            data.get('Tồn thực tế', ''),
                            data.get('Số thùng', ''),
                            data.get('Tình trạng', ''),
                            data.get('Ghi chú', '')
                        ))
                    except Exception as e:
                        # Bỏ qua item lỗi và tiếp tục
                        print(f"Lỗi khi insert item: {str(e)}")
                        continue
            
            # Bật lại cursor bình thường
            self.tong_hop_tree.config(cursor='')
            self.root.config(cursor='')
            self.root.update_idletasks()
            
        except Exception as e:
            # Xử lý lỗi để tránh crash
            self.tong_hop_tree.config(cursor='')
            self.root.config(cursor='')
            messagebox.showerror("Lỗi", f"Không thể cập nhật bảng tổng hợp: {str(e)}\n\nSố lượng dữ liệu: {len(self.tong_hop_data)}")
            print(f"Lỗi khi update_tong_hop_table: {str(e)}")
    
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
        # KHÔNG set initialdir để không hiển thị đường dẫn đã chọn trước đó
        # Mở dialog ở thư mục mặc định của hệ thống (Documents hoặc Desktop)
        user_home = Path.home()
        if sys.platform == 'win32':
            # Windows: mở ở Documents (không phải đường dẫn đã chọn trước đó)
            default_dir = str(user_home / "Documents")
        else:
            # macOS/Linux: mở ở Home
            default_dir = str(user_home)
        
        filename = filedialog.asksaveasfilename(
            title="Chọn đường dẫn để lưu file Excel tổng hợp",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=ten_file_1,
            initialdir=default_dir  # Luôn mở ở thư mục mặc định, không phải đường dẫn cũ
        )
        
        if not filename:
            # Người dùng đã hủy
            return
        
        # Lưu file vào đường dẫn đã chọn
        try:
            # Normalize path cho Windows
            filename = str(Path(filename).resolve())
            template_path_normalized = str(Path(self.template_file_path).resolve())
            
            # File 1: Chỉ copy file template từ đường dẫn đã cấu hình và đổi tên (KHÔNG ghi đè data)
            # Windows-specific: Retry nếu file bị lock
            max_retries = 5
            retry_count = 0
            copy_success = False
            
            while retry_count < max_retries and not copy_success:
                try:
                    shutil.copy2(template_path_normalized, filename)
                    copy_success = True
                except PermissionError as pe:
                    retry_count += 1
                    if retry_count < max_retries:
                        time.sleep(0.5)  # Đợi 0.5 giây trước khi thử lại
                    else:
                        messagebox.showerror("Lỗi", 
                            f"Không thể copy file template!\n\n"
                            f"File có thể đang được mở trong Excel hoặc chương trình khác.\n"
                            f"Vui lòng đóng file và thử lại.\n\n"
                            f"Lỗi: {str(pe)}")
                        return
                except Exception as e:
                    messagebox.showerror("Lỗi", f"Không thể copy file template: {str(e)}")
                    traceback.print_exc()
                    return
            
            # KHÔNG ghi đè dữ liệu vào file 1 - giữ nguyên data từ template
            
            # File 2: Tự động lưu vào thư mục đã cấu hình (nếu có)
            if self.auto_save_folder:
                try:
                    # Tạo đường dẫn file 2 với tên file đúng format - Windows-safe
                    file2_path = str(Path(self.auto_save_folder) / ten_file_2)
                    
                    # Windows-specific: Retry nếu file bị lock
                    max_retries = 5
                    retry_count = 0
                    save_success = False
                    
                    while retry_count < max_retries and not save_success:
                        try:
                            # Lưu file 2 với data tổng hợp (ngầm, không hiển thị thông báo)
                            df_save.to_excel(file2_path, index=False, engine='openpyxl')
                            save_success = True
                        except PermissionError as pe:
                            retry_count += 1
                            if retry_count < max_retries:
                                time.sleep(0.5)  # Đợi 0.5 giây trước khi thử lại
                            else:
                                # Log lỗi nhưng không hiển thị cho người dùng
                                print(f"Lỗi khi lưu file tự động (file có thể đang mở): {str(pe)}")
                                traceback.print_exc()
                                break
                        except Exception as e2:
                            # Log lỗi nhưng không hiển thị cho người dùng
                            print(f"Lỗi khi lưu file tự động: {str(e2)}")
                            traceback.print_exc()
                            break
                    
                    # Hiển thị thông báo thành công
                    messagebox.showinfo("Thành công", f"Đã lưu file tổng hợp thành công!")
                except Exception as e2:
                    # Nếu lỗi khi lưu file 2, chỉ log lỗi nhưng không hiển thị cho người dùng
                    print(f"Lỗi khi lưu file tự động: {str(e2)}")
                    traceback.print_exc()
                    # Vẫn hiển thị thành công cho file 1
                    messagebox.showinfo("Thành công", f"Đã lưu file tổng hợp thành công!")
            else:
                # Không có thư mục tự động, chỉ lưu file 1
                messagebox.showinfo("Thành công", f"Đã lưu file tổng hợp thành công!")
                
        except Exception as e:
            error_msg = str(e)
            traceback.print_exc()
            messagebox.showerror("Lỗi", f"Không thể lưu file: {error_msg}")

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

