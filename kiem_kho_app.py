#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ứng dụng Kiểm Kho - Quét mã vạch để kiểm tra tồn kho thực tế
Chạy trên Windows và macOS
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
# Lazy import pandas - chỉ import khi cần thiết để tăng tốc độ khởi động
# import pandas as pd  # Đã chuyển sang lazy import
import os
from pathlib import Path
import sys
import json
import shutil
import time
import traceback
import base64
import signal
import atexit

class KiemKhoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kiểm Kho - Quét Mã Vạch")
        # Tăng chiều cao để hiển thị đủ tất cả các phần tử
        self.root.geometry("1200x800")
        # Màu nền nhẹ nhàng hơn
        self.root.configure(bg='#F5F5F5')
        
        # Import pandas ngay sau khi UI được tạo (để có thể dùng trong toàn bộ class)
        # Import sớm hơn một chút để tránh lỗi "name 'pd' is not defined"
        try:
            import pandas as pd
            self.pd = pd  # Lưu vào instance để dùng trong các method khác
        except ImportError:
            self.pd = None
            # Sẽ báo lỗi khi load_data được gọi
        
        # Biến lưu trữ dữ liệu
        self.df = None
        self.current_box_data = None
        self.current_box_number = None
        self.scanned_items = {}  # Lưu các item đã quét: {isbn: {tua, ton_thuc_te, so_thung, ton_trong_thung, ghi_chu}}
        self.edit_entry = None  # Entry widget để chỉnh sửa trực tiếp
        self.editing_item = None  # Item đang được chỉnh sửa
        self.error_highlights = {}  # Lưu các highlight widgets: {item_id: [entry1, entry2]}
        self.is_processing_edit = False  # Flag để tránh xử lý edit 2 lần
        self.template_file_path = None  # Đường dẫn CHỌN ĐƯỜNG DẪN FILE TỔNG HỢP MẶC ĐỊNH
        self.auto_save_folder = None  # Thư mục tự động lưu file Excel 2 (Kiemkecuoinam)
        self.config_folder = None  # Thư mục lưu file config (do người dùng chọn)
        self.config_file = self.get_config_file_path()  # Đường dẫn file config
        self.tong_hop_data = []  # Lưu tổng hợp các data đã kiểm kê
        self.notebook = None  # Notebook widget để chứa các tab
        self.tong_hop_tree = None  # Treeview trong tab Tổng hợp
        self.so_tua_da_quet_var = None  # Biến để hiển thị số tựa đã quét
        self.tong_hop_edit_entry = None  # Entry widget để chỉnh sửa trong tab Tổng hợp
        self.tong_hop_editing_item = None  # Item đang được chỉnh sửa trong tab Tổng hợp
        self.tong_hop_editing_column = None  # Cột đang được chỉnh sửa trong tab Tổng hợp
        self._tong_hop_finish_scheduled = None  # ID của scheduled finish_tong_hop_edit call
        self.is_processing_tong_hop_edit = False  # Flag để tránh xử lý 2 lần
        
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
        
        # Tạo giao diện TRƯỚC để hiển thị nhanh hơn (tối ưu tốc độ khởi động)
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
        
        # Bind sự kiện đóng cửa sổ để kiểm tra dữ liệu chưa lưu
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Kiểm tra và khôi phục dữ liệu backup nếu có
        self.check_and_restore_backup()
        
        # Load dữ liệu SAU KHI UI đã hiển thị (defer loading để tăng tốc độ khởi động)
        # Sử dụng after() để load dữ liệu sau khi UI đã render xong (100ms delay)
        self.root.after(100, self.load_data_deferred)
        
        # Bắt đầu auto-save định kỳ (mỗi 30 giây)
        self.start_auto_save()
        
        # Đăng ký xử lý signal để lưu backup khi shutdown (cúp điện, tắt máy)
        self.setup_signal_handlers()
    
    def get_config_location_file(self):
        """Lấy đường dẫn file pointer trỏ đến vị trí config thực sự"""
        if getattr(sys, 'frozen', False):
            # Chạy từ executable - lưu file pointer cùng thư mục với exe
            exe_dir = Path(sys.executable).parent
            return exe_dir / ".set_up"
        else:
            # Chạy từ source code
            return Path(__file__).parent / ".set_up"
    
    def _set_file_hidden(self, file_path):
        """Đặt thuộc tính hidden/system cho file trên Windows để chặn người dùng mở"""
        if sys.platform == 'win32':
            try:
                import ctypes
                # Đặt thuộc tính hidden và system
                # FILE_ATTRIBUTE_HIDDEN = 2
                # FILE_ATTRIBUTE_SYSTEM = 4
                # FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM = 6
                ctypes.windll.kernel32.SetFileAttributesW(str(file_path), 6)
            except Exception as e:
                # Nếu không thể đặt thuộc tính, không quan trọng lắm
                pass
    
    def _encode_path(self, path_str):
        """Mã hóa đường dẫn để lưu vào file pointer"""
        try:
            # Sử dụng base64 với một key đơn giản để mã hóa
            key = "KiemKho2025SecretKey"
            # XOR với key (đơn giản nhưng hiệu quả)
            encoded_bytes = []
            key_bytes = key.encode('utf-8')
            path_bytes = path_str.encode('utf-8')
            for i, byte in enumerate(path_bytes):
                encoded_bytes.append(byte ^ key_bytes[i % len(key_bytes)])
            # Encode base64 để có thể lưu dạng text
            encoded = base64.b64encode(bytes(encoded_bytes)).decode('utf-8')
            return encoded
        except Exception as e:
            print(f"Lỗi khi mã hóa đường dẫn: {str(e)}")
            return None
    
    def _decode_path(self, encoded_str):
        """Giải mã đường dẫn từ file pointer"""
        try:
            # Decode base64
            encoded_bytes = base64.b64decode(encoded_str.encode('utf-8'))
            # XOR với key để giải mã
            key = "KiemKho2025SecretKey"
            decoded_bytes = []
            key_bytes = key.encode('utf-8')
            for i, byte in enumerate(encoded_bytes):
                decoded_bytes.append(byte ^ key_bytes[i % len(key_bytes)])
            # Decode về string
            decoded = bytes(decoded_bytes).decode('utf-8')
            return decoded
        except Exception as e:
            print(f"Lỗi khi giải mã đường dẫn: {str(e)}")
            return None
    
    def get_config_file_path(self):
        """Lấy đường dẫn file config - đọc từ file pointer hoặc tìm ở nhiều vị trí"""
        # Nếu đã có config_folder do người dùng chọn, dùng nó (ưu tiên cao nhất)
        if self.config_folder:
            return Path(self.config_folder) / "kiem_kho_config.json"
        
        # Ưu tiên 1: Đọc từ file pointer (trỏ đến vị trí config thực sự)
        location_file = self.get_config_location_file()
        if location_file.exists():
            try:
                with open(location_file, 'r', encoding='utf-8') as f:
                    encoded_path = f.read().strip()
                    if encoded_path:
                        # Giải mã đường dẫn
                        config_path = self._decode_path(encoded_path)
                        if config_path:
                            config_file_path = Path(config_path)
                            if config_file_path.exists():
                                # Đọc config_folder từ file config này
                                try:
                                    with open(config_file_path, 'r', encoding='utf-8') as cf:
                                        config = json.load(cf)
                                        if 'config_folder' in config and config['config_folder']:
                                            config_folder_path = Path(config['config_folder'])
                                            if config_folder_path.exists():
                                                self.config_folder = str(config_folder_path.resolve())
                                                return Path(self.config_folder) / "kiem_kho_config.json"
                                        # Nếu không có config_folder, dùng file này luôn
                                        return config_file_path
                                except:
                                    return config_file_path
            except Exception as e:
                print(f"Lỗi khi đọc file pointer: {str(e)}")
        
        # Ưu tiên 2: Tìm file config ở nhiều vị trí để lấy config_folder
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
            
            # Ưu tiên 1: Đọc từ file pointer (trỏ đến vị trí config thực sự)
            location_file = self.get_config_location_file()
            if location_file.exists():
                try:
                    with open(location_file, 'r', encoding='utf-8') as f:
                        encoded_path = f.read().strip()
                        if encoded_path:
                            # Giải mã đường dẫn
                            config_path = self._decode_path(encoded_path)
                            if config_path:
                                config_file_path = Path(config_path)
                                if config_file_path.exists():
                                    search_locations.insert(0, config_file_path)  # Ưu tiên cao nhất
                except Exception as e:
                    print(f"Lỗi khi đọc file pointer: {str(e)}")
            
            # Ưu tiên 2: File config trong config_folder (nếu đã biết)
            if self.config_folder:
                search_locations.append(Path(self.config_folder) / "kiem_kho_config.json")
            
            # Ưu tiên 3: File config hiện tại
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
        label_text = "Cấu hình 3 đường dẫn:\n1. CHỌN ĐƯỜNG DẪN FILE TỔNG HỢP MẶC ĐỊNH (để copy khi SAVE)\n2. CHỌN ĐƯỜNG DẪN FILE THEO DÕI CHI TIẾT (Kiemkecuoinam)\n3. Thư mục lưu file cấu hình (kiem_kho_config.json)"
        tk.Label(dialog, text=label_text, bg='#F5F5F5', fg='#000000', 
                font=('Arial', 11), justify=tk.LEFT, wraplength=650).pack(pady=15, padx=20)
        
        # Load giá trị đã lưu nếu có
        saved_config = self.load_config()
        
        # Đường dẫn 1: CHỌN ĐƯỜNG DẪN FILE TỔNG HỢP MẶC ĐỊNH
        tk.Label(dialog, text="1. CHỌN ĐƯỜNG DẪN FILE TỔNG HỢP MẶC ĐỊNH:", bg='#F5F5F5', fg='#000000', 
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
                title="Chọn CHỌN ĐƯỜNG DẪN FILE TỔNG HỢP MẶC ĐỊNH",
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
        
        # Đường dẫn 2: CHỌN ĐƯỜNG DẪN FILE THEO DÕI CHI TIẾT
        tk.Label(dialog, text="2. CHỌN ĐƯỜNG DẪN FILE THEO DÕI CHI TIẾT (Kiemkecuoinam):", bg='#F5F5F5', fg='#000000', 
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
            
            folder = filedialog.askdirectory(title="Chọn CHỌN ĐƯỜNG DẪN FILE THEO DÕI CHI TIẾT")
            
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
            
            # Kiểm tra CHỌN ĐƯỜNG DẪN FILE THEO DÕI CHI TIẾT có tồn tại không
            if not os.path.isdir(folder_path):
                messagebox.showerror("Lỗi", f"CHỌN ĐƯỜNG DẪN FILE THEO DÕI CHI TIẾT không tồn tại!\n{folder_path}")
                return
            
            # Kiểm tra thư mục lưu file config có tồn tại không
            if not os.path.isdir(config_folder_path):
                messagebox.showerror("Lỗi", f"Thư mục lưu file cấu hình không tồn tại!\n{config_folder_path}")
                return
            
            # Kiểm tra quyền ghi vào CHỌN ĐƯỜNG DẪN FILE THEO DÕI CHI TIẾT - Windows-safe
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
                messagebox.showerror("Lỗi", f"Không có quyền ghi vào CHỌN ĐƯỜNG DẪN FILE THEO DÕI CHI TIẾT!\n{error_msg}")
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
            
            # QUAN TRỌNG: Lưu file pointer trỏ đến vị trí config thực sự (đã mã hóa)
            # File này sẽ được lưu ở vị trí cố định (cùng thư mục với exe) để dễ tìm
            try:
                location_file = self.get_config_location_file()
                # Mã hóa đường dẫn trước khi lưu
                encoded_path = self._encode_path(str(self.config_file.resolve()))
                if encoded_path:
                    with open(location_file, 'w', encoding='utf-8') as f:
                        f.write(encoded_path)
                    # Đặt thuộc tính hidden/system trên Windows để chặn người dùng mở
                    self._set_file_hidden(location_file)
                    print(f"[OK] Đã lưu file pointer (đã mã hóa và ẩn) vào: {location_file}")
                    print(f"[OK] File pointer trỏ đến: {self.config_file}")
                else:
                    print(f"[WARNING] Không thể mã hóa đường dẫn")
            except Exception as e:
                print(f"[WARNING] Không thể lưu file pointer: {str(e)}")
                # Không quan trọng lắm, vẫn có thể tìm bằng cách khác
            
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
        # Sử dụng pandas đã được import trong __init__
        if self.pd is None:
            try:
                import pandas as pd
                self.pd = pd
            except ImportError:
                messagebox.showerror("Lỗi", "Không thể import pandas! Vui lòng cài đặt: pip install pandas")
                sys.exit(1)
        
        pd = self.pd  # Alias để dùng trong hàm này
        
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
                        # Đảm bảo pandas đã được import
                        if self.pd is None:
                            import pandas as pd
                            self.pd = pd
                        pd = self.pd
                        
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
    
    def load_data_deferred(self):
        """Load dữ liệu sau khi UI đã hiển thị (deferred loading để tăng tốc độ khởi động)"""
        try:
            # Cập nhật cursor để hiển thị đang load
            self.root.config(cursor='wait')
            self.root.update_idletasks()
            
            # Load dữ liệu
            self.load_data()
        finally:
            # Khôi phục cursor
            self.root.config(cursor='')
    
    def _process_dataframe(self):
        """Xử lý DataFrame sau khi đọc thành công"""
        # Đảm bảo pandas đã được import
        if self.pd is None:
            try:
                import pandas as pd
                self.pd = pd
            except ImportError:
                messagebox.showerror("Lỗi", "Không thể import pandas! Vui lòng cài đặt: pip install pandas")
                return
        
        pd = self.pd  # Alias để dùng trong hàm này
        
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
        columns = ('Số thứ tự', 'ISBN', 'Tựa', 'Tồn thực tế', 'Số thùng', 'Tồn tựa trong thùng', 'Tình trạng', 'Ghi chú')
        # Thứ tự: 0=Số thứ tự, 1=ISBN, 2=Tựa, 3=Tồn thực tế, 4=Số thùng, 5=Tồn tựa trong thùng, 6=Tình trạng, 7=Ghi chú
        
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
        self.tree.heading('Số thứ tự', text='STT')
        self.tree.heading('ISBN', text='ISBN')
        self.tree.heading('Tựa', text='Tựa')
        self.tree.heading('Tồn thực tế', text='Tồn thực tế')
        self.tree.heading('Số thùng', text='Số thùng')
        self.tree.heading('Tồn tựa trong thùng', text='Tồn tựa trong thùng')
        self.tree.heading('Tình trạng', text='Tình trạng')
        self.tree.heading('Ghi chú', text='Ghi chú')
        
        self.tree.column('Số thứ tự', width=80, anchor='center')
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
        tonghop_columns = ('N/X', 'Số phiếu', 'Ngày', 'Vị trí mới', 'ISBN', 'Tựa', 'Tồn thực tế', 'Số thùng', 'Tình trạng', 'Ghi chú', 'Note thùng')
        
        # Tạo Treeview cho tổng hợp với tối ưu hiệu suất
        self.tong_hop_tree = ttk.Treeview(tonghop_table_frame, columns=tonghop_columns, show='headings', 
                                          yscrollcommand=tonghop_scrollbar_y.set, xscrollcommand=tonghop_scrollbar_x.set,
                                          height=20, style='Treeview')
        
        # Tối ưu hiệu suất: tắt một số tính năng không cần thiết
        # Không cần selectmode vì không có selection logic phức tạp
        
        # Cấu hình các cột
        column_widths = {'N/X': 250, 'Số phiếu': 120, 'Ngày': 100, 'Vị trí mới': 120, 'ISBN': 150, 
                        'Tựa': 250, 'Tồn thực tế': 100, 'Số thùng': 120, 'Tình trạng': 100, 'Ghi chú': 200, 'Note thùng': 200}
        
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
        
        # Bind events để cho phép chỉnh sửa và xóa dòng trong tab Tổng hợp
        self.tong_hop_tree.bind('<Double-1>', self.on_tong_hop_item_click)
        self.tong_hop_tree.bind('<Button-1>', self.on_tong_hop_item_click)
        self.tong_hop_tree.bind('<Delete>', self.on_tong_hop_delete)
        self.tong_hop_tree.bind('<Key-Delete>', self.on_tong_hop_delete)
    
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
            
            # Chuyển đổi sang string để so sánh (không phân biệt chữ hoa/thường)
            self.df[so_thung_col] = self.df[so_thung_col].astype(str).str.strip()
            # So sánh không phân biệt chữ hoa/thường
            so_thung_lower = so_thung.lower()
            self.current_box_data = self.df[self.df[so_thung_col].str.lower() == so_thung_lower].copy()
            
            if self.current_box_data.empty:
                messagebox.showinfo("Thông báo", f"Không tìm thấy dữ liệu cho thùng số {so_thung}")
                self.current_box_number = None
                self.so_tua_var.set("0")
                self.clear_table()
                return
            
            # Kiểm tra nếu đã quét đủ số tựa trong thùng này
            so_tua_trong_thung = len(self.current_box_data)
            so_tua_da_quet = len(self.scanned_items)
            
            # Nếu đang load lại cùng một thùng và đã quét đủ (so sánh không phân biệt chữ hoa/thường)
            current_box_lower = str(self.current_box_number).lower() if self.current_box_number else ''
            if current_box_lower == so_thung.lower() and so_tua_da_quet >= so_tua_trong_thung:
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
            # So sánh số thùng (có thể là 'Số thùng' hoặc 'Vị trí mới') - không phân biệt chữ hoa/thường
            so_thung_in_data = str(data.get('Số thùng', '')).strip()
            vi_tri_moi_in_data = str(data.get('Vị trí mới', '')).strip()
            
            # Đếm nếu số thùng khớp (so sánh không phân biệt chữ hoa/thường)
            if (so_thung_in_data.lower() == so_thung_clean.lower() or 
                vi_tri_moi_in_data.lower() == so_thung_clean.lower()):
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
            
            # Tối ưu: chỉ kiểm tra các item có số thùng khớp trước (không phân biệt chữ hoa/thường)
            so_thung_clean_lower = so_thung_clean.lower()
            for data in self.tong_hop_data:
                # Kiểm tra số thùng trước (nhanh hơn)
                so_thung_in_data = str(data.get('Số thùng', '')).strip()
                vi_tri_moi_in_data = str(data.get('Vị trí mới', '')).strip()
                
                # Nếu số thùng không khớp, bỏ qua ngay (so sánh không phân biệt chữ hoa/thường)
                if (so_thung_in_data.lower() != so_thung_clean_lower and 
                    vi_tri_moi_in_data.lower() != so_thung_clean_lower):
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
                
                # Đảm bảo pandas đã được import
                if self.pd is None:
                    try:
                        import pandas as pd
                        self.pd = pd
                    except ImportError:
                        messagebox.showerror("Lỗi", "Không thể import pandas!")
                        return
                
                pd = self.pd  # Alias để dùng trong hàm này
                
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
                
                # Kiểm tra xem đã quét chưa để tự động tăng số lượng
                ton_thuc_te_value = '1'  # Mặc định là 1 khi quét lần đầu
                ghi_chu_value = ''
                tinh_trang_value = ''
                is_existing_item = False
                
                if isbn_clean in self.scanned_items:
                    # ISBN đã tồn tại - tăng số lượng lên 1
                    is_existing_item = True
                    item_id_old = self.scanned_items[isbn_clean]['item_id']
                    
                    # Lấy giá trị "Tồn thực tế" từ cả scanned_items và tree để đảm bảo chính xác
                    old_ton_thuc_te_from_items = self.scanned_items[isbn_clean].get('ton_thuc_te', '')
                    old_values = list(self.tree.item(item_id_old, 'values'))
                    old_ton_thuc_te_from_tree = old_values[3] if len(old_values) > 3 else ''  # Tồn thực tế ở index 3
                    
                    # Chuyển đổi sang string và strip để so sánh
                    # QUAN TRỌNG: Kiểm tra cả None và empty string, nhưng không bỏ qua giá trị '0'
                    old_ton_thuc_te_from_items_str = ''
                    if old_ton_thuc_te_from_items is not None:
                        old_ton_thuc_te_from_items_str = str(old_ton_thuc_te_from_items).strip()
                    
                    old_ton_thuc_te_from_tree_str = ''
                    if old_ton_thuc_te_from_tree is not None:
                        old_ton_thuc_te_from_tree_str = str(old_ton_thuc_te_from_tree).strip()
                    
                    # Ưu tiên lấy từ scanned_items (chính xác nhất), nếu không có hoặc rỗng thì lấy từ tree
                    # Kiểm tra bằng cách thử parse, không chỉ kiểm tra empty string
                    old_ton_thuc_te = ''
                    
                    # Thử parse từ scanned_items trước (ưu tiên)
                    if old_ton_thuc_te_from_items_str != '':
                        try:
                            test_num = float(old_ton_thuc_te_from_items_str)
                            old_ton_thuc_te = old_ton_thuc_te_from_items_str
                        except (ValueError, TypeError):
                            pass
                    
                    # Nếu không có từ scanned_items, thử từ tree
                    if not old_ton_thuc_te and old_ton_thuc_te_from_tree_str != '':
                        try:
                            test_num = float(old_ton_thuc_te_from_tree_str)
                            old_ton_thuc_te = old_ton_thuc_te_from_tree_str
                        except (ValueError, TypeError):
                            pass
                    
                    try:
                        # Tăng số lượng lên 1
                        if old_ton_thuc_te:
                            # Parse giá trị và tăng lên 1
                            old_ton_thuc_te_num = int(float(old_ton_thuc_te))
                            ton_thuc_te_value = str(old_ton_thuc_te_num + 1)
                        else:
                            # Nếu không có giá trị, mặc định là 1
                            ton_thuc_te_value = '1'
                    except (ValueError, TypeError) as e:
                        # Nếu không parse được, mặc định là 1
                        ton_thuc_te_value = '1'
                    
                    # Giữ lại giá trị Ghi chú và Tình trạng cũ
                    ghi_chu_value = old_values[7] if len(old_values) > 7 else ''  # Ghi chú ở index 7
                    tinh_trang_value = old_values[6] if len(old_values) > 6 else ''  # Tình trạng ở index 6
                    
                    # Xóa highlight cũ trước khi xóa item để tránh lỗi khi click vào highlight
                    try:
                        self.remove_error_highlights(item_id_old)
                    except:
                        pass
                    
                    # Xóa item cũ
                    try:
                        self.tree.delete(item_id_old)
                    except:
                        pass  # Item có thể đã bị xóa rồi
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
                
                # Đảm bảo thứ tự đúng với columns: Số thứ tự, ISBN, Tựa, Tồn thực tế, Số thùng, Tồn tựa trong thùng, Tình trạng, Ghi chú
                # Tính số thứ tự: số dòng hiện tại + 1
                so_thu_tu = len(self.tree.get_children()) + 1
                
                # Đảm bảo ton_thuc_te_value là string và không rỗng
                if not ton_thuc_te_value or str(ton_thuc_te_value).strip() == '':
                    ton_thuc_te_value = '1'
                else:
                    ton_thuc_te_value = str(ton_thuc_te_value).strip()
                
                # Đảm bảo ISBN được format đúng và hiển thị
                isbn_display = str(isbn_clean).strip() if isbn_clean else ''
                if not isbn_display:
                    isbn_display = str(isbn).strip() if isbn else ''
                
                item_id = self.tree.insert('', tk.END, values=(
                    str(so_thu_tu),            # 0: Số thứ tự
                    isbn_display,               # 1: ISBN - đảm bảo hiển thị đúng
                    str(tua) if tua else '',   # 2: Tựa
                    ton_thuc_te_value,         # 3: Tồn thực tế - tự động điền 1 hoặc tăng lên
                    str(so_thung_hien_thi) if so_thung_hien_thi else '',    # 4: Số thùng (dùng vị trí mới nếu có)
                    str(ton_trong_thung_display),  # 5: Tồn tựa trong thùng
                    tinh_trang_value if tinh_trang_value else '',          # 6: Tình trạng - giữ lại nếu đã có
                    ghi_chu_value if ghi_chu_value else ''              # 7: Ghi chú - giữ lại nếu đã có
                ), tags=('',))
                
                # Đảm bảo giá trị ISBN được hiển thị đúng ngay sau khi insert
                current_values_check = list(self.tree.item(item_id, 'values'))
                if len(current_values_check) >= 2 and current_values_check[1] != isbn_display:
                    current_values_check[1] = isbn_display
                    self.tree.item(item_id, values=current_values_check)
                
                # Đảm bảo giá trị được hiển thị đúng ngay sau khi insert
                # Cập nhật lại để đảm bảo đồng bộ
                current_values_check = list(self.tree.item(item_id, 'values'))
                if len(current_values_check) >= 4 and current_values_check[3] != ton_thuc_te_value:
                    current_values_check[3] = ton_thuc_te_value
                    self.tree.item(item_id, values=current_values_check)
                
                # Lưu thông tin
                # Lưu cả số thùng gốc và số thùng hiển thị (vị trí mới nếu có)
                # Lưu cả vi_tri_moi để có thể sử dụng khi lưu (nếu người dùng chỉnh sửa trực tiếp)
                vi_tri_moi_value = self.vi_tri_moi_var.get().strip()
                self.scanned_items[isbn_clean] = {
                    'item_id': item_id,
                    'tua': tua,
                    'ton_thuc_te': ton_thuc_te_value,  # Lưu giá trị đã tự động điền hoặc đã tăng
                    'so_thung': so_thung_hien_thi,  # Lưu số thùng hiển thị (có thể là vị trí mới)
                    'so_thung_goc': so_thung,  # Lưu số thùng gốc từ dữ liệu
                    'vi_tri_moi': vi_tri_moi_value,  # Lưu giá trị từ ô "Thùng / vị trí mới" khi quét
                    'ton_trong_thung': ton_trong_thung,
                    'tinh_trang': tinh_trang_value,  # Giữ lại tình trạng cũ nếu có
                    'ghi_chu': ghi_chu_value  # Giữ lại ghi chú cũ nếu có
                }
                
                # Cập nhật số tựa đã quét (chỉ hiển thị số tựa đã lưu trong Tổng hợp)
                if hasattr(self, 'so_tua_da_quet_var') and self.so_tua_da_quet_var and self.current_box_number:
                    so_tua_da_quet = self.count_scanned_titles_for_box(self.current_box_number)
                    self.so_tua_da_quet_var.set(str(so_tua_da_quet))
                
                # Scroll đến item để người dùng thấy được item vừa thêm/cập nhật
                self.tree.selection_set(item_id)
                self.tree.focus(item_id)
                self.tree.see(item_id)
                
                # Nếu là item mới (quét lần đầu), tự động focus vào ô "Tồn thực tế" để người dùng nhập và Enter
                # Nếu là item đã tồn tại (quét lại), chỉ cộng dồn số lượng, không cần focus
                if not is_existing_item:
                    # Tự động focus vào ô "Tồn thực tế" để người dùng có thể chỉnh sửa và Enter
                    # Fix closure issue: capture item_id vào biến local
                    item_id_to_edit = item_id
                    self.root.after(100, lambda i=item_id_to_edit: self.auto_edit_ton_thuc_te(i))
                else:
                    # Nếu là item đã tồn tại (đã cộng dồn), đảm bảo giá trị được hiển thị đúng
                    # Cập nhật lại giá trị trong tree ngay lập tức để đảm bảo hiển thị đúng
                    current_values = list(self.tree.item(item_id, 'values'))
                    if len(current_values) >= 4:
                        # Đảm bảo giá trị Tồn thực tế đúng
                        if current_values[3] != ton_thuc_te_value:
                            current_values[3] = ton_thuc_te_value
                            self.tree.item(item_id, values=current_values)
                            # Force update để hiển thị ngay
                            self.root.update_idletasks()
                    
                    # Tự động kiểm tra và cập nhật highlight/tình trạng nếu có lệch
                    if isbn_clean in self.scanned_items:
                        # Sau đó mới kiểm tra và cập nhật highlight/tình trạng
                        self.root.after(200, lambda i=item_id, isbn=isbn_clean: self._check_and_update_status_after_increment(i, isbn))
                
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
        
        # Cho phép edit: Tồn thực tế (3), Số thùng (4), Tồn tựa trong thùng (5), Ghi chú (7)
        # Không cho edit: Số thứ tự (0), ISBN (1), Tựa (2), Tình trạng (6) - chỉ đọc
        if column_index not in [3, 4, 5, 7]:
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
            
            # Lấy giá trị hiện tại và đảm bảo có đủ 8 cột
            values = list(self.tree.item(item, 'values'))
            while len(values) < 8:
                values.append('')
            
            isbn = values[1] if len(values) > 1 else ''  # ISBN ở index 1 (sau Số thứ tự)
            
            # Nếu đang edit cột "Ghi chú", đảm bảo lấy giá trị từ scanned_items để không mất dữ liệu
            if column_index == 7 and isbn in self.scanned_items:
                saved_ghi_chu = self.scanned_items[isbn].get('ghi_chu', '')
                # Nếu có giá trị đã lưu, đảm bảo values[7] có giá trị đúng
                if saved_ghi_chu:
                    if len(values) <= 7:
                        while len(values) < 8:
                            values.append('')
                    values[7] = saved_ghi_chu
            
            # Xử lý theo từng cột
            if column_index == 3:  # Tồn thực tế
                values[3] = new_value  # Đảm bảo đúng index
                
                # Kiểm tra và highlight nếu khác nhau - CHỈ chạy cho cột Tồn thực tế
                if isbn in self.scanned_items:
                    self.scanned_items[isbn]['ton_thuc_te'] = new_value
                    # Lưu backup khi có thay đổi
                    self.save_backup_on_change()
                    ton_trong_thung = self.scanned_items[isbn]['ton_trong_thung']
                    
                    try:
                        ton_thuc_te_num = float(new_value) if new_value else 0
                        ton_trong_thung_num = float(ton_trong_thung) if ton_trong_thung else 0
                        
                        # Kiểm tra lệch
                        if abs(ton_thuc_te_num - ton_trong_thung_num) > 0.01:
                            # Tự động điền "Thiếu" hoặc "Dư" vào cột Tình trạng (index 5)
                            if ton_thuc_te_num < ton_trong_thung_num:
                                tinh_trang = "Thiếu"
                                so_luong_lech = int(ton_trong_thung_num - ton_thuc_te_num)
                                ghi_chu_auto = f"Thiếu {so_luong_lech} cuốn"
                            else:
                                tinh_trang = "Dư"
                                so_luong_lech = int(ton_thuc_te_num - ton_trong_thung_num)
                                ghi_chu_auto = f"Dư {so_luong_lech} cuốn"
                            
                            # Đảm bảo có đủ 8 cột và đúng thứ tự: Số thứ tự, ISBN, Tựa, Tồn thực tế, Số thùng, Tồn tựa trong thùng, Tình trạng, Ghi chú
                            while len(values) < 8:
                                values.append('')
                            
                            # Lấy giá trị Ghi chú hiện tại từ scanned_items (để giữ lại phần người dùng nhập)
                            ghi_chu_hien_tai = self.scanned_items[isbn].get('ghi_chu', '')
                            if not ghi_chu_hien_tai:
                                # Nếu không có trong scanned_items, lấy từ values
                                ghi_chu_hien_tai = values[7] if len(values) > 7 else ''
                            
                            # Tự động điền số lượng thiếu/dư vào ô Ghi chú
                            # QUAN TRỌNG: Xóa TẤT CẢ các pattern lỗi cũ trước khi thêm lỗi mới
                            import re
                            if ghi_chu_hien_tai and ghi_chu_hien_tai.strip():
                                # Xóa TẤT CẢ các pattern "Thiếu X cuốn" hoặc "Dư X cuốn" và các dấu câu/khoảng trắng sau đó
                                # Pattern: "Thiếu X cuốn" hoặc "Dư X cuốn" có thể có dấu phẩy, dấu chấm, khoảng trắng sau đó
                                ghi_chu_cleaned = re.sub(r'(Thiếu \d+ cuốn|Dư \d+ cuốn)[,\.\s]*', '', ghi_chu_hien_tai, flags=re.IGNORECASE)
                                ghi_chu_cleaned = re.sub(r'\.\s*\.', '.', ghi_chu_cleaned)  # Xóa dấu chấm kép
                                ghi_chu_cleaned = ghi_chu_cleaned.strip()
                                
                                # Thêm thông tin thiếu/dư mới vào đầu
                                if ghi_chu_cleaned:
                                    values[7] = f"{ghi_chu_auto}. {ghi_chu_cleaned}"
                                else:
                                    values[7] = ghi_chu_auto
                            else:
                                # Chưa có nội dung, điền mới
                                values[7] = ghi_chu_auto
                            
                            values[6] = tinh_trang  # Tình trạng ở index 6
                            
                            # Cập nhật scanned_items
                            self.scanned_items[isbn]['tinh_trang'] = tinh_trang
                            self.scanned_items[isbn]['ghi_chu'] = values[7]  # Cập nhật ghi chú
                            
                            # Cập nhật tree
                            self.tree.item(item, values=values)
                            
                            # Tô đỏ 2 ô: Tồn thực tế (cột 2) và Tình trạng (cột 5)
                            self.highlight_error_cells(item)
                            
                            # Không hiển thị cảnh báo nữa vì đã có cột "Tình trạng" để hiển thị
                        else:
                            # Không có lỗi - xóa highlight và tình trạng
                            # Xóa thông tin thiếu/dư khỏi Ghi chú (nếu có) nhưng giữ lại phần người dùng nhập
                            while len(values) < 8:
                                values.append('')
                            
                            ghi_chu_hien_tai = values[7] if len(values) > 7 else ''
                            
                            # Xóa thông tin thiếu/dư tự động khỏi Ghi chú (nếu có)
                            if ghi_chu_hien_tai:
                                import re
                                # Loại bỏ các pattern như "Thiếu X cuốn" hoặc "Dư X cuốn" và các dấu câu/khoảng trắng sau đó
                                ghi_chu_cleaned = re.sub(r'^(Thiếu \d+ cuốn|Dư \d+ cuốn)[,\.\s]*', '', ghi_chu_hien_tai, flags=re.IGNORECASE)
                                ghi_chu_cleaned = re.sub(r'[,\.\s]*(Thiếu \d+ cuốn|Dư \d+ cuốn)[,\.\s]*', '', ghi_chu_cleaned, flags=re.IGNORECASE)
                                values[7] = ghi_chu_cleaned.strip()
                            else:
                                values[7] = ''
                            
                            # Xóa tình trạng nếu có
                            if len(values) > 6:
                                values[6] = ''
                                if 'tinh_trang' in self.scanned_items[isbn]:
                                    del self.scanned_items[isbn]['tinh_trang']
                            
                            # Cập nhật ghi chú đã làm sạch
                            self.scanned_items[isbn]['ghi_chu'] = values[7]
                            
                            # Lưu backup khi có thay đổi
                            self.save_backup_on_change()
                            
                            # Cập nhật tree
                            self.tree.item(item, values=values)
                            
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
                        values[4] = self.scanned_items[isbn]['so_thung'] if isbn in self.scanned_items else ''  # Số thùng ở index 4
                        self.tree.item(item, values=values)
                        # Reset flag và cleanup trước khi return
                        self.is_processing_edit = False
                        if self.edit_entry:
                            self.edit_entry.destroy()
                            self.edit_entry = None
                            self.editing_item = None
                        return
                
                values[4] = new_value  # Đảm bảo đúng index (Số thùng ở index 4)
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
        
            elif column_index == 5:  # Tồn tựa trong thùng
                # Chuyển thành số nguyên
                try:
                    new_value_int = int(float(new_value)) if new_value else 0
                    values[5] = str(new_value_int)  # Đảm bảo đúng index và chuyển thành string (Tồn tựa trong thùng ở index 5)
                    if isbn in self.scanned_items:
                        self.scanned_items[isbn]['ton_trong_thung'] = new_value_int
                except:
                    values[5] = new_value  # Đảm bảo đúng index
            
            elif column_index == 7:  # Ghi chú
                # Đảm bảo có đủ 8 cột
                while len(values) < 8:
                    values.append('')
                # Chỉ lưu giá trị mới vào values và scanned_items - không có logic phức tạp
                values[7] = new_value
                if isbn in self.scanned_items:
                    self.scanned_items[isbn]['ghi_chu'] = new_value
                    # Lưu backup khi có thay đổi
                    self.save_backup_on_change()
                # Cập nhật tree với giá trị mới
                self.tree.item(item, values=values)
                # Đảm bảo không có highlight nào che mất nội dung cột "Ghi chú"
                # (highlight chỉ được tạo cho cột "Tồn thực tế" và "Tình trạng", không phải "Ghi chú")
                # Return ngay để không chạy phần cập nhật tree chung bên dưới
                return
            
            # Cập nhật tree với giá trị mới (chỉ cho các cột khác, không phải Ghi chú)
            self.tree.item(item, values=values)
            
            # Nếu là cột Tồn thực tế, kiểm tra lại và cập nhật highlight
            if column_index == 3:
                if isbn in self.scanned_items:
                    ton_trong_thung = self.scanned_items[isbn]['ton_trong_thung']
                    try:
                        ton_thuc_te_num = float(new_value) if new_value else 0
                        ton_trong_thung_num = float(ton_trong_thung) if ton_trong_thung else 0
                        if abs(ton_thuc_te_num - ton_trong_thung_num) > 0.01:
                            # Vẫn còn lệch - tự động cập nhật tình trạng và số lượng thiếu/dư vào Ghi chú
                            if ton_thuc_te_num < ton_trong_thung_num:
                                tinh_trang = "Thiếu"
                                so_luong_lech = int(ton_trong_thung_num - ton_thuc_te_num)
                                ghi_chu_auto = f"Thiếu {so_luong_lech} cuốn"
                            else:
                                tinh_trang = "Dư"
                                so_luong_lech = int(ton_thuc_te_num - ton_trong_thung_num)
                                ghi_chu_auto = f"Dư {so_luong_lech} cuốn"
                            
                            # Đảm bảo có đủ 8 cột
                            while len(values) < 8:
                                values.append('')
                            
                            # Lấy giá trị Ghi chú hiện tại từ scanned_items (để giữ lại phần người dùng nhập)
                            ghi_chu_hien_tai = self.scanned_items[isbn].get('ghi_chu', '')
                            if not ghi_chu_hien_tai:
                                # Nếu không có trong scanned_items, lấy từ values
                                ghi_chu_hien_tai = values[7] if len(values) > 7 else ''
                            
                            # Tự động điền số lượng thiếu/dư vào ô Ghi chú
                            # QUAN TRỌNG: Xóa TẤT CẢ các pattern lỗi cũ trước khi thêm lỗi mới
                            import re
                            if ghi_chu_hien_tai and ghi_chu_hien_tai.strip():
                                # Xóa TẤT CẢ các pattern "Thiếu X cuốn" hoặc "Dư X cuốn" và các dấu câu/khoảng trắng sau đó
                                # Pattern: "Thiếu X cuốn" hoặc "Dư X cuốn" có thể có dấu phẩy, dấu chấm, khoảng trắng sau đó
                                ghi_chu_cleaned = re.sub(r'(Thiếu \d+ cuốn|Dư \d+ cuốn)[,\.\s]*', '', ghi_chu_hien_tai, flags=re.IGNORECASE)
                                ghi_chu_cleaned = re.sub(r'\.\s*\.', '.', ghi_chu_cleaned)  # Xóa dấu chấm kép
                                ghi_chu_cleaned = ghi_chu_cleaned.strip()
                                
                                # Thêm thông tin thiếu/dư mới vào đầu
                                if ghi_chu_cleaned:
                                    values[7] = f"{ghi_chu_auto}. {ghi_chu_cleaned}"
                                else:
                                    values[7] = ghi_chu_auto
                            else:
                                # Chưa có nội dung, điền mới
                                values[7] = ghi_chu_auto
                            
                            values[6] = tinh_trang  # Tình trạng ở index 6
                            
                            # Cập nhật scanned_items
                            self.scanned_items[isbn]['tinh_trang'] = tinh_trang
                            self.scanned_items[isbn]['ghi_chu'] = values[7]  # Cập nhật ghi chú
                            
                            self.tree.item(item, values=values)
                            self.highlight_error_cells(item)
                        else:
                            # Đã khớp - xóa tình trạng
                            # Xóa thông tin thiếu/dư khỏi Ghi chú (nếu có) nhưng giữ lại phần người dùng nhập
                            while len(values) < 8:
                                values.append('')
                            
                            # Lấy giá trị Ghi chú hiện tại từ scanned_items (để giữ lại phần người dùng nhập)
                            ghi_chu_hien_tai = self.scanned_items[isbn].get('ghi_chu', '')
                            if not ghi_chu_hien_tai:
                                # Nếu không có trong scanned_items, lấy từ values
                                ghi_chu_hien_tai = values[7] if len(values) > 7 else ''
                            
                            # Xóa thông tin thiếu/dư tự động khỏi Ghi chú (nếu có)
                            if ghi_chu_hien_tai:
                                # Loại bỏ các pattern như "Thiếu X cuốn" hoặc "Dư X cuốn" và các dấu câu/khoảng trắng sau đó
                                import re
                                ghi_chu_cleaned = re.sub(r'^(Thiếu \d+ cuốn|Dư \d+ cuốn)[,\.\s]*', '', ghi_chu_hien_tai, flags=re.IGNORECASE)
                                ghi_chu_cleaned = re.sub(r'[,\.\s]*(Thiếu \d+ cuốn|Dư \d+ cuốn)[,\.\s]*', '', ghi_chu_cleaned, flags=re.IGNORECASE)
                                values[7] = ghi_chu_cleaned.strip()
                            else:
                                values[7] = ''
                            
                            values[6] = ''  # Xóa tình trạng
                            
                            # Cập nhật scanned_items
                            if 'tinh_trang' in self.scanned_items[isbn]:
                                del self.scanned_items[isbn]['tinh_trang']
                            self.scanned_items[isbn]['ghi_chu'] = values[7]  # Cập nhật ghi chú đã làm sạch
                            
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
    
    def _check_and_update_status_after_increment(self, item_id, isbn):
        """Kiểm tra và cập nhật tình trạng sau khi tăng số lượng"""
        if isbn not in self.scanned_items:
            return
        
        try:
            values = list(self.tree.item(item_id, 'values'))
            if len(values) < 8:
                return
            
            ton_thuc_te_str = values[3] if len(values) > 3 else ''  # Tồn thực tế ở index 3
            ton_trong_thung = self.scanned_items[isbn]['ton_trong_thung']
            
            try:
                ton_thuc_te_num = float(ton_thuc_te_str) if ton_thuc_te_str else 0
                ton_trong_thung_num = float(ton_trong_thung) if ton_trong_thung else 0
                
                if abs(ton_thuc_te_num - ton_trong_thung_num) > 0.01:
                    # Có lệch - tự động cập nhật tình trạng và số lượng thiếu/dư vào Ghi chú
                    if ton_thuc_te_num < ton_trong_thung_num:
                        tinh_trang = "Thiếu"
                        so_luong_lech = int(ton_trong_thung_num - ton_thuc_te_num)
                        ghi_chu_auto = f"Thiếu {so_luong_lech} cuốn"
                    else:
                        tinh_trang = "Dư"
                        so_luong_lech = int(ton_thuc_te_num - ton_trong_thung_num)
                        ghi_chu_auto = f"Dư {so_luong_lech} cuốn"
                    
                    # Đảm bảo có đủ 8 cột
                    while len(values) < 8:
                        values.append('')
                    
                    # Lấy giá trị Ghi chú hiện tại
                    ghi_chu_hien_tai = values[7] if len(values) > 7 else ''
                    
                    # QUAN TRỌNG: Xóa các thông tin thiếu/dư cũ trong Ghi chú trước khi thêm mới
                    # Để tránh tích lũy thông tin cũ (ví dụ: "Dư 2 cuốn. Dư 1 cuốn")
                    # QUAN TRỌNG: Xóa TẤT CẢ các pattern lỗi cũ trước khi thêm lỗi mới
                    import re
                    ghi_chu_cleaned = ''
                    if ghi_chu_hien_tai and ghi_chu_hien_tai.strip():
                        # Xóa TẤT CẢ các pattern "Thiếu X cuốn" hoặc "Dư X cuốn" và các dấu câu/khoảng trắng sau đó
                        ghi_chu_cleaned = re.sub(r'(Thiếu \d+ cuốn|Dư \d+ cuốn)[,\.\s]*', '', ghi_chu_hien_tai, flags=re.IGNORECASE)
                        ghi_chu_cleaned = re.sub(r'\.\s*\.', '.', ghi_chu_cleaned)  # Xóa dấu chấm kép
                        ghi_chu_cleaned = ghi_chu_cleaned.strip()
                    
                    # Thêm thông tin thiếu/dư mới vào đầu
                    if ghi_chu_cleaned:
                        values[7] = f"{ghi_chu_auto}. {ghi_chu_cleaned}"
                    else:
                        values[7] = ghi_chu_auto
                    
                    values[6] = tinh_trang
                    
                    # Cập nhật scanned_items
                    self.scanned_items[isbn]['tinh_trang'] = tinh_trang
                    self.scanned_items[isbn]['ghi_chu'] = values[7]
                    
                    self.tree.item(item_id, values=values)
                    self.highlight_error_cells(item_id)
                else:
                    # Đã khớp - xóa tình trạng
                    while len(values) < 7:
                        values.append('')
                    
                    ghi_chu_hien_tai = values[7] if len(values) > 7 else ''
                    
                    if ghi_chu_hien_tai:
                        import re
                        ghi_chu_cleaned = re.sub(r'^(Thiếu \d+ cuốn|Dư \d+ cuốn)[,\.\s]*', '', ghi_chu_hien_tai, flags=re.IGNORECASE)
                        ghi_chu_cleaned = re.sub(r'[,\.\s]*(Thiếu \d+ cuốn|Dư \d+ cuốn)[,\.\s]*', '', ghi_chu_cleaned, flags=re.IGNORECASE)
                        values[7] = ghi_chu_cleaned.strip()
                    else:
                        values[7] = ''
                    
                    values[6] = ''
                    
                    if 'tinh_trang' in self.scanned_items[isbn]:
                        del self.scanned_items[isbn]['tinh_trang']
                    self.scanned_items[isbn]['ghi_chu'] = values[7]
                    
                    self.tree.item(item_id, values=values)
                    self.remove_error_highlights(item_id)
            except:
                pass
        except:
            pass
    
    def highlight_error_cells(self, item_id):
        """Tô đỏ 2 ô: Tồn thực tế và Tình trạng"""
        # Xóa highlight cũ nếu có
        self.remove_error_highlights(item_id)
        
        # Lấy giá trị từ tree
        values = list(self.tree.item(item_id, 'values'))
        
        # Tô đỏ ô "Tồn thực tế" (cột 4, index 3)
        bbox_ton = self.tree.bbox(item_id, '#4')
        if bbox_ton:
            x, y, width, height = bbox_ton
            ton_value = values[3] if len(values) > 3 else ''
            highlight1 = tk.Label(self.tree, bg='#FFCDD2', fg='#C62828', 
                                  text=str(ton_value), font=('Arial', 10, 'bold'), 
                                  relief=tk.FLAT, anchor='center')
            highlight1.place(x=x, y=y, width=width, height=height)
            # Cho phép click qua để edit
            # Fix closure issue: capture item_id và column vào biến local
            item_id_click = item_id
            highlight1.bind('<Button-1>', lambda e, i=item_id_click, c='#4': self.on_highlight_click(e, i, c))
        
        # Tô đỏ ô "Tình trạng" (cột 7, index 6)
        bbox_tinh_trang = self.tree.bbox(item_id, '#7')
        if bbox_tinh_trang:
            x, y, width, height = bbox_tinh_trang
            tinh_trang_value = values[6] if len(values) > 6 else ''
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
        
        # Không cho phép edit cột Tình trạng (6) - chỉ đọc
        if column_index == 6:
            return
        
        # Kiểm tra xem item có tồn tại không
        try:
            # Kiểm tra item có tồn tại bằng cách lấy children
            all_items = self.tree.get_children()
            if item_id not in all_items:
                return  # Item không tồn tại, không làm gì cả
        except:
            return  # Lỗi khi kiểm tra, không làm gì cả
        
        # Lấy giá trị hiện tại
        try:
            values = list(self.tree.item(item_id, 'values'))
            current_value = values[column_index] if column_index < len(values) else ''
        except Exception as e:
            # Item có thể đã bị xóa, không làm gì cả
            return
        
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
            bbox_ton = self.tree.bbox(item_id, '#4')
            if bbox_ton and widgets[0].winfo_exists():
                x, y, width, height = bbox_ton
                ton_value = values[3] if len(values) > 3 else ''
                widgets[0].config(text=str(ton_value))
                widgets[0].place(x=x, y=y, width=width, height=height)
            
            # Cập nhật ô Tình trạng
            bbox_tinh_trang = self.tree.bbox(item_id, '#7')
            if bbox_tinh_trang and len(widgets) > 1 and widgets[1].winfo_exists():
                x, y, width, height = bbox_tinh_trang
                tinh_trang_value = values[6] if len(values) > 6 else ''
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
        
        # Lấy vị trí của cell "Tồn thực tế" (column index 3)
        column = '#4'  # Column index 3 (0-indexed) + 1
        bbox = self.tree.bbox(item_id, column)
        if not bbox:
            return
        
        x, y, width, height = bbox
        
        # Lấy giá trị hiện tại
        values = list(self.tree.item(item_id, 'values'))
        current_value = values[3] if len(values) > 3 else ''  # Tồn thực tế ở index 3
        
        # Tạo Entry widget để edit trực tiếp
        self.edit_entry = tk.Entry(self.tree, font=('Arial', 10), 
                                   relief=tk.FLAT, bd=0, bg='#FFFFFF', fg='#000000')
        self.edit_entry.insert(0, str(current_value))
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        self.edit_entry.focus()
        self.editing_item = item_id
        self.editing_column = 3  # Tồn thực tế ở index 3
        
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
        
        # Kiểm tra ràng buộc: Số tựa đã quét phải bằng số tựa trong thùng
        if self.current_box_data is not None and not self.current_box_data.empty and self.current_box_number:
            so_tua_trong_thung = len(self.current_box_data)
            so_tua_da_quet_lan_nay = len(self.scanned_items)
            
            # QUAN TRỌNG: Đếm cả các tựa đã lưu trong tab Tổng hợp từ các lần save trước
            so_tua_da_luu_truoc = self.count_scanned_titles_for_box(self.current_box_number) if self.current_box_number else 0
            
            # Tổng số tựa đã quét = số tựa quét lần này + số tựa đã lưu trước đó
            so_tua_da_quet_tong = so_tua_da_quet_lan_nay + so_tua_da_luu_truoc
            
            if so_tua_da_quet_tong < so_tua_trong_thung:
                # Tạo dialog tùy chỉnh với 2 nút
                dialog = tk.Toplevel(self.root)
                dialog.title("Cảnh báo")
                dialog.geometry("600x400")
                dialog.resizable(False, False)
                dialog.transient(self.root)
                dialog.grab_set()
                dialog.configure(bg='#f5f5f5')
                
                # Đặt dialog ở giữa màn hình
                dialog.update_idletasks()
                x = (dialog.winfo_screenwidth() // 2) - (600 // 2)
                y = (dialog.winfo_screenheight() // 2) - (400 // 2)
                dialog.geometry(f"600x400+{x}+{y}")
                
                # Frame chính với padding đủ để các nút không bị che
                main_frame = tk.Frame(dialog, padx=30, pady=30, bg='#f5f5f5')
                main_frame.pack(fill=tk.BOTH, expand=True)
                
                # Icon cảnh báo
                icon_frame = tk.Frame(main_frame, bg='#f5f5f5')
                icon_frame.pack(pady=(0, 15))
                icon_label = tk.Label(icon_frame, text="⚠️", font=('Arial', 32), bg='#f5f5f5')
                icon_label.pack()
                
                # Thông điệp - hiển thị cả số tựa đã lưu trước đó
                if so_tua_da_luu_truoc > 0:
                    message_text = (
                        f"Chưa quét đủ số tựa trong thùng!\n\n"
                        f"Thùng {self.current_box_number} có {so_tua_trong_thung} tựa.\n"
                        f"Đã quét lần này: {so_tua_da_quet_lan_nay} tựa.\n"
                        f"Đã lưu trước đó: {so_tua_da_luu_truoc} tựa.\n"
                        f"Tổng đã quét: {so_tua_da_quet_tong} tựa.\n"
                        f"Còn thiếu: {so_tua_trong_thung - so_tua_da_quet_tong} tựa.\n\n"
                        f"Bạn muốn tiếp tục lưu hay hủy để quét tiếp?"
                    )
                else:
                    message_text = (
                        f"Chưa quét đủ số tựa trong thùng!\n\n"
                        f"Thùng {self.current_box_number} có {so_tua_trong_thung} tựa.\n"
                        f"Đã quét: {so_tua_da_quet_lan_nay} tựa.\n"
                        f"Còn thiếu: {so_tua_trong_thung - so_tua_da_quet_lan_nay} tựa.\n\n"
                        f"Bạn muốn tiếp tục lưu hay hủy để quét tiếp?"
                    )
                message_frame = tk.Frame(main_frame, bg='#f5f5f5')
                message_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
                message_label = tk.Label(
                    message_frame, 
                    text=message_text,
                    font=('Arial', 12),
                    justify=tk.CENTER,
                    bg='#f5f5f5',
                    fg='#333333',
                    wraplength=540
                )
                message_label.pack()
                
                # Frame chứa các nút với padding đủ để không bị che
                button_frame = tk.Frame(main_frame, bg='#f5f5f5')
                button_frame.pack(pady=(15, 25))
                
                # Biến để lưu kết quả
                result = {'value': None}
                
                def on_continue_save():
                    result['value'] = 'continue'
                    dialog.destroy()
                
                def on_cancel():
                    result['value'] = 'cancel'
                    dialog.destroy()
                
                # Nút "Tiếp tục và lưu"
                continue_btn = tk.Button(
                    button_frame,
                    text="Tiếp tục và lưu",
                    command=on_continue_save,
                    font=('Arial', 12, 'bold'),
                    bg='#4CAF50',
                    fg='white',
                    padx=25,
                    pady=12,
                    relief=tk.FLAT,
                    cursor='hand2',
                    activebackground='#45a049',
                    activeforeground='white',
                    bd=0,
                    highlightthickness=0
                )
                continue_btn.pack(side=tk.LEFT, padx=15)
                
                # Nút "Cancel"
                cancel_btn = tk.Button(
                    button_frame,
                    text="Cancel",
                    command=on_cancel,
                    font=('Arial', 12, 'bold'),
                    bg='#f44336',
                    fg='white',
                    padx=25,
                    pady=12,
                    relief=tk.FLAT,
                    cursor='hand2',
                    activebackground='#da190b',
                    activeforeground='white',
                    bd=0,
                    highlightthickness=0
                )
                cancel_btn.pack(side=tk.LEFT, padx=15)
                
                # Đợi dialog đóng
                dialog.wait_window()
                
                # Xử lý kết quả
                if result['value'] == 'cancel':
                    # Focus vào ô nhập ISBN để tiếp tục quét
                    if hasattr(self, 'isbn_entry'):
                        self.isbn_entry.focus()
                    return
                # Nếu chọn "Tiếp tục và lưu", tiếp tục với logic lưu bên dưới
        
        # Lấy giá trị từ các input
        vi_tri_moi_global = self.vi_tri_moi_var.get().strip() if hasattr(self, 'vi_tri_moi_var') else ''
        ngay_value = self.ngay_var.get().strip() if hasattr(self, 'ngay_var') else ''
        nhap_xuat_value = self.nhap_xuat_var.get().strip() if hasattr(self, 'nhap_xuat_var') else ''
        note_thung_value = self.note_thung_var.get().strip() if hasattr(self, 'note_thung_var') else ''
        
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
                
                # Xác định số thùng mới (vị trí mới)
                # QUAN TRỌNG: "Vị trí mới" chỉ lấy từ vi_tri_moi, KHÔNG fallback về so_thung_goc
                so_thung_moi = ''
                
                # Ưu tiên 1: vi_tri_moi_global (từ ô input "Thùng / vị trí mới")
                if vi_tri_moi_global and vi_tri_moi_global.strip():
                    so_thung_moi = vi_tri_moi_global.strip()
                else:
                    # Ưu tiên 2: vi_tri_moi từ scanned_items (đã lưu khi quét)
                    vi_tri_moi_saved = info.get('vi_tri_moi', '').strip()
                    if vi_tri_moi_saved:
                        so_thung_moi = vi_tri_moi_saved
                    # Nếu không có vi_tri_moi, để trống (không lấy từ so_thung gốc)
                
                # Lấy giá trị từ input "Nhập/Xuất" ở tab Kiểm kê và hiển thị trực tiếp vào cột N/X
                nx_value = nhap_xuat_value.strip() if nhap_xuat_value else ""
                
                # Thêm vào danh sách để append sau (hiệu quả hơn)
                items_to_add.append({
                    'N/X': nx_value,
                    'Số phiếu': so_phieu,
                    'Ngày': ngay_value,
                    'Vị trí mới': so_thung_moi,  # Chỉ lấy từ vi_tri_moi, không fallback về so_thung_goc
                    'ISBN': isbn,
                    'Tựa': info.get('tua', ''),
                    'Tồn thực tế': info.get('ton_thuc_te', ''),
                    'Số thùng': so_thung_goc_clean,
                    'Tình trạng': info.get('tinh_trang', ''),  # Lấy từ scanned_items
                    'Ghi chú': info.get('ghi_chu', ''),  # Ghi chú do người dùng tự nhập
                    'Note thùng': note_thung_value  # Note thùng từ input
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
        
        # Xóa dữ liệu đã quét
        self.scanned_items.clear()
        self.clear_table()
        self.so_tua_var.set("0")
        
        # QUAN TRỌNG: Reset field "Đã quét" về 0 sau khi save thành công
        # Không cập nhật lại từ cache vì sẽ load thùng mới
        if hasattr(self, 'so_tua_da_quet_var') and self.so_tua_da_quet_var:
            self.so_tua_da_quet_var.set("0")
        
        # Reset các input: Số thùng, Thùng / vị trí mới, và Note thùng
        if hasattr(self, 'so_thung_var'):
            self.so_thung_var.set("")
        if hasattr(self, 'vi_tri_moi_var'):
            self.vi_tri_moi_var.set("")
        if hasattr(self, 'note_thung_var'):
            self.note_thung_var.set("")
        
        # Reset current_box_number và current_box_data để đảm bảo không còn cache
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
                                data.get('Ghi chú', ''),
                                data.get('Note thùng', '')
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
                            data.get('Ghi chú', ''),
                            data.get('Note thùng', '')
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
    
    def on_tong_hop_item_click(self, event):
        """Xử lý click để edit trực tiếp các cột có thể chỉnh sửa trong tab Tổng hợp"""
        # Hủy edit cũ nếu có
        if self.tong_hop_edit_entry:
            self.finish_tong_hop_edit()
        
        region = self.tong_hop_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        item = self.tong_hop_tree.identify_row(event.y)
        column = self.tong_hop_tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1
        
        # Cho phép edit: Vị trí mới (3), Tồn thực tế (6), Tình trạng (8), Ghi chú (9), Note thùng (10)
        # Không cho edit: N/X (0), Số phiếu (1), Ngày (2), ISBN (4), Tựa (5), Số thùng (7) - chỉ đọc
        editable_columns = [3, 6, 8, 9, 10]  # Vị trí mới, Tồn thực tế, Tình trạng, Ghi chú, Note thùng
        if column_index not in editable_columns:
            return
        
        if not item:
            return
        
        # Lấy giá trị hiện tại
        values = list(self.tong_hop_tree.item(item, 'values'))
        current_value = values[column_index] if column_index < len(values) else ''
        
        # Lấy vị trí của cell
        bbox = self.tong_hop_tree.bbox(item, column)
        if not bbox:
            return
        
        x, y, width, height = bbox
        
        # Tạo Entry widget để edit trực tiếp
        self.tong_hop_edit_entry = tk.Entry(self.tong_hop_tree, font=('Arial', 10), 
                                           relief=tk.FLAT, bd=0, bg='#FFFFFF', fg='#000000')
        self.tong_hop_edit_entry.insert(0, str(current_value))
        self.tong_hop_edit_entry.select_range(0, tk.END)
        self.tong_hop_edit_entry.place(x=x, y=y, width=width, height=height)
        self.tong_hop_edit_entry.focus()
        self.tong_hop_editing_item = item
        self.tong_hop_editing_column = column_index
        
        def finish_on_enter(event):
            # Hủy scheduled call nếu có
            if self._tong_hop_finish_scheduled is not None:
                try:
                    # Chỉ cancel nếu ID hợp lệ
                    if isinstance(self._tong_hop_finish_scheduled, str) and self._tong_hop_finish_scheduled:
                        self.root.after_cancel(self._tong_hop_finish_scheduled)
                except (ValueError, TypeError):
                    # Bỏ qua lỗi nếu ID không hợp lệ
                    pass
                self._tong_hop_finish_scheduled = None
            self.finish_tong_hop_edit()
        
        def finish_on_focus_out(event):
            # Kiểm tra entry còn tồn tại và chưa đang xử lý
            if not self.tong_hop_edit_entry or self.is_processing_tong_hop_edit:
                return
            
            # Kiểm tra widget còn tồn tại
            try:
                if not self.tong_hop_edit_entry.winfo_exists():
                    return
            except:
                return
            
            # Hủy scheduled call cũ nếu có (nếu có)
            if self._tong_hop_finish_scheduled is not None:
                try:
                    if isinstance(self._tong_hop_finish_scheduled, str) and self._tong_hop_finish_scheduled:
                        self.root.after_cancel(self._tong_hop_finish_scheduled)
                except (ValueError, TypeError, tk.TclError):
                    pass
                self._tong_hop_finish_scheduled = None
            
            # Gọi trực tiếp - KHÔNG dùng after() hoặc after_idle() để tránh lỗi
            # Sử dụng try-except để bắt mọi lỗi có thể xảy ra
            try:
                self.finish_tong_hop_edit()
            except Exception as e:
                # Nếu có lỗi, chỉ log và cleanup
                print(f"Error in finish_on_focus_out: {e}")
                try:
                    if self.tong_hop_edit_entry:
                        self.tong_hop_edit_entry.destroy()
                except:
                    pass
                self.tong_hop_edit_entry = None
                self.tong_hop_editing_item = None
                self.tong_hop_editing_column = None
                self.is_processing_tong_hop_edit = False
        
        self.tong_hop_edit_entry.bind('<Return>', finish_on_enter)
        self.tong_hop_edit_entry.bind('<FocusOut>', finish_on_focus_out)
        self.tong_hop_edit_entry.bind('<Escape>', lambda e: self.cancel_tong_hop_edit())
    
    def _do_finish_tong_hop_edit(self):
        """Wrapper để reset scheduled flag trước khi gọi finish_tong_hop_edit"""
        self._tong_hop_finish_scheduled = None
        # Gọi finish trực tiếp
        self.finish_tong_hop_edit()
    
    def finish_tong_hop_edit(self):
        """Hoàn tất việc chỉnh sửa trong tab Tổng hợp"""
        # Tránh xử lý 2 lần nếu đang trong quá trình xử lý
        if self.is_processing_tong_hop_edit:
            return
        
        # Hủy scheduled call nếu có
        if self._tong_hop_finish_scheduled is not None:
            try:
                if isinstance(self._tong_hop_finish_scheduled, str) and self._tong_hop_finish_scheduled:
                    self.root.after_cancel(self._tong_hop_finish_scheduled)
            except (ValueError, TypeError, tk.TclError):
                # Bỏ qua lỗi nếu ID không hợp lệ hoặc đã bị cancel
                pass
            self._tong_hop_finish_scheduled = None
        
        # Kiểm tra entry và item còn tồn tại
        if not self.tong_hop_edit_entry or not self.tong_hop_editing_item:
            return
        
        # Kiểm tra entry widget còn tồn tại trong window
        try:
            if not self.tong_hop_edit_entry.winfo_exists():
                # Entry đã bị destroy, cleanup và return
                self.tong_hop_edit_entry = None
                self.tong_hop_editing_item = None
                self.tong_hop_editing_column = None
                return
        except:
            # Nếu không thể kiểm tra, giả sử đã bị destroy
            self.tong_hop_edit_entry = None
            self.tong_hop_editing_item = None
            self.tong_hop_editing_column = None
            return
        
        # Đặt flag để tránh xử lý lại
        self.is_processing_tong_hop_edit = True
        
        try:
            # Lấy giá trị mới
            new_value = self.tong_hop_edit_entry.get()
            
            # Lấy item ID và column index
            item_id = self.tong_hop_editing_item
            column_index = self.tong_hop_editing_column
            
            # Lấy giá trị hiện tại của dòng
            values = list(self.tong_hop_tree.item(item_id, 'values'))
            
            # Cập nhật giá trị trong tree
            values[column_index] = new_value
            self.tong_hop_tree.item(item_id, values=values)
            
            # Tìm index của item trong tree để map với tong_hop_data
            all_items = list(self.tong_hop_tree.get_children())
            if item_id in all_items:
                data_index = all_items.index(item_id)
                
                # Cập nhật trong tong_hop_data theo index
                if 0 <= data_index < len(self.tong_hop_data):
                    # Map column index sang tên cột trong data
                    column_mapping = {
                        3: 'Vị trí mới',
                        6: 'Tồn thực tế',
                        8: 'Tình trạng',
                        9: 'Ghi chú',
                        10: 'Note thùng'
                    }
                    
                    column_name = column_mapping.get(column_index)
                    if column_name:
                        self.tong_hop_data[data_index][column_name] = new_value
                        
                        # Nếu là cột "Tồn thực tế" (column_index == 6), tự động check và cập nhật Tình trạng và Ghi chú
                        if column_index == 6:
                            self._check_and_update_tinh_trang_tong_hop(data_index, new_value, values, item_id)
                        
                        # Lưu backup khi chỉnh sửa (gọi trực tiếp, không dùng after)
                        try:
                            self.save_backup()
                        except Exception as backup_error:
                            # Không hiển thị lỗi cho người dùng, chỉ log
                            print(f"Error saving backup: {backup_error}")
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể cập nhật dữ liệu: {str(e)}")
        finally:
            # Reset flag
            self.is_processing_tong_hop_edit = False
            
            # Xóa entry widget
            if self.tong_hop_edit_entry:
                try:
                    self.tong_hop_edit_entry.destroy()
                except:
                    pass
                self.tong_hop_edit_entry = None
            self.tong_hop_editing_item = None
            self.tong_hop_editing_column = None
    
    def _check_and_update_tinh_trang_tong_hop(self, data_index, ton_thuc_te_new, values, item_id):
        """Tự động check và cập nhật Tình trạng và Ghi chú khi sửa Tồn thực tế ở tab Tổng hợp"""
        try:
            if data_index < 0 or data_index >= len(self.tong_hop_data):
                return
            
            data = self.tong_hop_data[data_index]
            isbn = data.get('ISBN', '').strip()
            so_thung = data.get('Số thùng', '').strip()
            
            if not isbn or not so_thung:
                return
            
            # Tìm "Tồn tựa trong thùng" từ dữ liệu gốc (self.df)
            ton_trong_thung = 0
            if self.df is not None and not self.df.empty:
                # Tìm trong df dựa trên ISBN và số thùng
                isbn_clean = str(isbn).strip()
                isbn_clean_digits = ''.join(filter(str.isdigit, isbn_clean))
                
                # Tìm cột số thùng và ISBN trong df
                so_thung_col = None
                isbn_col = None
                ton_tung_tua_col = None
                
                for col in self.df.columns:
                    col_lower = str(col).lower().strip()
                    if ('số thùng' in col_lower or 'so thung' in col_lower or 
                        col_lower == 'thùng' or col_lower == 'thung'):
                        so_thung_col = col
                    elif 'isbn' in col_lower:
                        isbn_col = col
                    elif (('tồn' in col_lower and 'tựa' in col_lower) or 
                          ('ton' in col_lower and 'tua' in col_lower) or 
                          'qty tựa trong thùng' in col_lower or 
                          'qty tua trong thung' in col_lower):
                        ton_tung_tua_col = col
                
                # Tìm dòng khớp với ISBN và số thùng
                if so_thung_col and isbn_col and ton_tung_tua_col:
                    matched_row = None
                    for idx, row in self.df.iterrows():
                        row_isbn = str(row.get(isbn_col, '')).strip()
                        row_so_thung = str(row.get(so_thung_col, '')).strip()
                        
                        # So sánh ISBN (có thể không khớp hoàn toàn)
                        row_isbn_digits = ''.join(filter(str.isdigit, row_isbn))
                        isbn_match = (row_isbn == isbn_clean or 
                                     row_isbn.endswith(isbn_clean) or 
                                     isbn_clean.endswith(row_isbn) or
                                     (row_isbn_digits and isbn_clean_digits and row_isbn_digits == isbn_clean_digits))
                        
                        # So sánh số thùng (không phân biệt chữ hoa/thường)
                        row_so_thung_lower = row_so_thung.lower()
                        so_thung_lower = so_thung.lower()
                        so_thung_match = (row_so_thung_lower == so_thung_lower or 
                                         row_so_thung_lower.endswith(so_thung_lower) or 
                                         so_thung_lower.endswith(row_so_thung_lower))
                        
                        if isbn_match and so_thung_match:
                            matched_row = row
                            break
                    
                    if matched_row is not None:
                        ton_trong_thung = matched_row.get(ton_tung_tua_col, 0)
                        try:
                            ton_trong_thung = int(float(ton_trong_thung)) if ton_trong_thung else 0
                        except:
                            ton_trong_thung = 0
            
            # So sánh Tồn thực tế với Tồn tựa trong thùng
            try:
                ton_thuc_te_num = float(ton_thuc_te_new) if ton_thuc_te_new else 0
                ton_trong_thung_num = float(ton_trong_thung) if ton_trong_thung else 0
                
                # Đảm bảo values có đủ 11 cột (theo tonghop_columns)
                while len(values) < 11:
                    values.append('')
                
                if abs(ton_thuc_te_num - ton_trong_thung_num) > 0.01:
                    # Có lệch - tự động cập nhật tình trạng và số lượng thiếu/dư vào Ghi chú
                    if ton_thuc_te_num < ton_trong_thung_num:
                        tinh_trang = "Thiếu"
                        so_luong_lech = int(ton_trong_thung_num - ton_thuc_te_num)
                        ghi_chu_auto = f"Thiếu {so_luong_lech} cuốn"
                    else:
                        tinh_trang = "Dư"
                        so_luong_lech = int(ton_thuc_te_num - ton_trong_thung_num)
                        ghi_chu_auto = f"Dư {so_luong_lech} cuốn"
                    
                    # Lấy giá trị Ghi chú hiện tại từ tong_hop_data (để giữ lại phần người dùng nhập)
                    ghi_chu_hien_tai = data.get('Ghi chú', '')
                    if not ghi_chu_hien_tai:
                        # Nếu không có trong data, lấy từ values
                        ghi_chu_hien_tai = values[9] if len(values) > 9 else ''
                    
                    # Tự động điền số lượng thiếu/dư vào ô Ghi chú
                    # QUAN TRỌNG: Xóa TẤT CẢ các pattern lỗi cũ trước khi thêm lỗi mới
                    import re
                    if ghi_chu_hien_tai and ghi_chu_hien_tai.strip():
                        # Xóa TẤT CẢ các pattern "Thiếu X cuốn" hoặc "Dư X cuốn" và các dấu câu/khoảng trắng sau đó
                        ghi_chu_cleaned = re.sub(r'(Thiếu \d+ cuốn|Dư \d+ cuốn)[,\.\s]*', '', ghi_chu_hien_tai, flags=re.IGNORECASE)
                        ghi_chu_cleaned = re.sub(r'\.\s*\.', '.', ghi_chu_cleaned)  # Xóa dấu chấm kép
                        ghi_chu_cleaned = ghi_chu_cleaned.strip()
                        
                        # Thêm thông tin thiếu/dư mới vào đầu
                        if ghi_chu_cleaned:
                            ghi_chu_moi = f"{ghi_chu_auto}. {ghi_chu_cleaned}"
                        else:
                            ghi_chu_moi = ghi_chu_auto
                    else:
                        # Chưa có nội dung, điền mới
                        ghi_chu_moi = ghi_chu_auto
                    
                    values[8] = tinh_trang  # Tình trạng ở index 8
                    values[9] = ghi_chu_moi  # Ghi chú mới
                    
                    # Cập nhật tong_hop_data
                    self.tong_hop_data[data_index]['Tình trạng'] = tinh_trang
                    self.tong_hop_data[data_index]['Ghi chú'] = ghi_chu_moi
                    
                    # Cập nhật tree
                    if item_id:
                        self.tong_hop_tree.item(item_id, values=values)
                else:
                    # Đã khớp - xóa tình trạng
                    # Xóa thông tin thiếu/dư khỏi Ghi chú (nếu có) nhưng giữ lại phần người dùng nhập
                    # Lấy giá trị Ghi chú hiện tại từ tong_hop_data (để giữ lại phần người dùng nhập)
                    ghi_chu_hien_tai = data.get('Ghi chú', '')
                    if not ghi_chu_hien_tai:
                        # Nếu không có trong data, lấy từ values
                        ghi_chu_hien_tai = values[9] if len(values) > 9 else ''
                    
                    # Xóa thông tin thiếu/dư tự động khỏi Ghi chú (nếu có)
                    if ghi_chu_hien_tai:
                        # Loại bỏ các pattern như "Thiếu X cuốn" hoặc "Dư X cuốn" và các dấu câu/khoảng trắng sau đó
                        import re
                        ghi_chu_cleaned = re.sub(r'^(Thiếu \d+ cuốn|Dư \d+ cuốn)[,\.\s]*', '', ghi_chu_hien_tai, flags=re.IGNORECASE)
                        ghi_chu_cleaned = re.sub(r'[,\.\s]*(Thiếu \d+ cuốn|Dư \d+ cuốn)[,\.\s]*', '', ghi_chu_cleaned, flags=re.IGNORECASE)
                        ghi_chu_cleaned = ghi_chu_cleaned.strip()
                    else:
                        ghi_chu_cleaned = ''
                    
                    values[8] = ''  # Xóa tình trạng
                    values[9] = ghi_chu_cleaned  # Ghi chú đã làm sạch
                    
                    # Cập nhật tong_hop_data
                    self.tong_hop_data[data_index]['Tình trạng'] = ''
                    self.tong_hop_data[data_index]['Ghi chú'] = ghi_chu_cleaned
                    
                    # Cập nhật tree
                    if item_id:
                        self.tong_hop_tree.item(item_id, values=values)
            except Exception as e:
                # Nếu có lỗi khi so sánh, không làm gì cả
                print(f"Lỗi khi check tình trạng: {str(e)}")
                pass
        except Exception as e:
            # Nếu có lỗi, không làm gì cả
            print(f"Lỗi khi check và update tình trạng: {str(e)}")
            pass
    
    def cancel_tong_hop_edit(self):
        """Hủy việc chỉnh sửa trong tab Tổng hợp"""
        # Hủy scheduled call nếu có
        if self._tong_hop_finish_scheduled is not None:
            try:
                # Chỉ cancel nếu ID hợp lệ
                if isinstance(self._tong_hop_finish_scheduled, str) and self._tong_hop_finish_scheduled:
                    self.root.after_cancel(self._tong_hop_finish_scheduled)
            except (ValueError, TypeError):
                # Bỏ qua lỗi nếu ID không hợp lệ
                pass
            self._tong_hop_finish_scheduled = None
        
        if self.tong_hop_edit_entry:
            try:
                self.tong_hop_edit_entry.destroy()
            except:
                pass
            self.tong_hop_edit_entry = None
        self.tong_hop_editing_item = None
        self.tong_hop_editing_column = None
    
    def on_tong_hop_delete(self, event):
        """Xóa dòng được chọn trong tab Tổng hợp"""
        selected_items = self.tong_hop_tree.selection()
        if not selected_items:
            return
        
        # Xác nhận xóa
        result = messagebox.askyesno("Xác nhận", f"Bạn có chắc chắn muốn xóa {len(selected_items)} dòng đã chọn?")
        if not result:
            return
        
        try:
            # Lấy tất cả items trong tree theo thứ tự
            all_items = list(self.tong_hop_tree.get_children())
            
            # Tìm index của các items được chọn
            selected_indices = []
            for item_id in selected_items:
                if item_id in all_items:
                    selected_indices.append(all_items.index(item_id))
            
            # Sắp xếp theo thứ tự ngược lại để xóa từ cuối lên (tránh lỗi index)
            selected_indices.sort(reverse=True)
            
            # Xóa từ tong_hop_data trước (theo index)
            for idx in selected_indices:
                if 0 <= idx < len(self.tong_hop_data):
                    del self.tong_hop_data[idx]
            
            # Xóa khỏi tree
            for item_id in selected_items:
                self.tong_hop_tree.delete(item_id)
            
            # Lưu backup sau khi xóa
            self.save_backup()
            
            messagebox.showinfo("Thành công", f"Đã xóa {len(selected_items)} dòng!")
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xóa dòng: {str(e)}")
            traceback.print_exc()
    
    def export_tong_hop_excel(self):
        """Xuất file Excel tổng hợp (logic giống save_data cũ)"""
        if not self.tong_hop_data:
            messagebox.showwarning("Cảnh báo", "Chưa có dữ liệu tổng hợp để xuất!")
            return
        
        # Sử dụng pandas đã được import trong __init__
        if self.pd is None:
            try:
                import pandas as pd
                self.pd = pd
            except ImportError:
                messagebox.showerror("Lỗi", "Không thể import pandas! Vui lòng cài đặt: pip install pandas")
                return
        
        pd = self.pd  # Alias để dùng trong hàm này
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
            messagebox.showerror("Lỗi", "Chưa cấu hình CHỌN ĐƯỜNG DẪN FILE TỔNG HỢP MẶC ĐỊNH! Vui lòng cấu hình lại.")
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
    
    def get_backup_file_path(self):
        """Lấy đường dẫn file backup - luôn lưu cùng thư mục với file application"""
        if getattr(sys, 'frozen', False):
            # Chạy từ executable - lưu cùng thư mục với file .exe
            return Path(sys.executable).parent / "kiem_kho_backup.json"
        else:
            # Chạy từ source code - lưu cùng thư mục với file .py
            return Path(__file__).parent / "kiem_kho_backup.json"
    
    def save_backup(self):
        """Tự động lưu backup dữ liệu (scanned_items và tong_hop_data)"""
        try:
            backup_file = self.get_backup_file_path()
            backup_data = {
                'scanned_items': self.scanned_items,
                'tong_hop_data': self.tong_hop_data,
                'current_box_number': self.current_box_number,
                'timestamp': time.time()
            }
            
            # Lưu vào file tạm trước, sau đó rename để tránh mất dữ liệu khi crash
            temp_file = backup_file.with_suffix('.tmp')
            with open(temp_file, 'w', encoding='utf-8') as f:
                json.dump(backup_data, f, ensure_ascii=False, indent=2)
            
            # Rename file tạm thành file chính (atomic operation)
            if backup_file.exists():
                backup_file.unlink()
            temp_file.rename(backup_file)
            
        except Exception as e:
            # Không hiển thị lỗi cho người dùng vì đây là auto-save
            print(f"Lỗi khi lưu backup: {str(e)}")
    
    def check_and_restore_backup(self):
        """Kiểm tra và khôi phục dữ liệu backup nếu có"""
        try:
            backup_file = self.get_backup_file_path()
            if not backup_file.exists():
                return  # Không có backup
            
            # Đọc backup
            with open(backup_file, 'r', encoding='utf-8') as f:
                backup_data = json.load(f)
            
            scanned_items_backup = backup_data.get('scanned_items', {})
            tong_hop_data_backup = backup_data.get('tong_hop_data', [])
            current_box_number_backup = backup_data.get('current_box_number')
            timestamp = backup_data.get('timestamp', 0)
            
            # Kiểm tra xem có dữ liệu để khôi phục không
            has_scanned_items = scanned_items_backup and len(scanned_items_backup) > 0
            has_tong_hop_data = tong_hop_data_backup and len(tong_hop_data_backup) > 0
            
            if not has_scanned_items and not has_tong_hop_data:
                return  # Không có dữ liệu để khôi phục
            
            # Hiển thị dialog cho phép người dùng chọn khôi phục
            dialog = tk.Toplevel(self.root)
            dialog.title("Khôi phục dữ liệu")
            dialog.geometry("600x400")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()
            dialog.configure(bg='#f5f5f5')
            
            # Đặt dialog ở phía trên cửa sổ chính để không che các nút SAVE/RESET
            dialog.update_idletasks()
            # Lấy vị trí của cửa sổ chính
            root_x = self.root.winfo_x()
            root_y = self.root.winfo_y()
            root_width = self.root.winfo_width()
            # Đặt dialog ở giữa theo chiều ngang của cửa sổ chính, nhưng ở phía trên
            x = root_x + (root_width // 2) - (600 // 2)
            y = root_y + 100  # Đặt cách đỉnh cửa sổ chính 100px để không che các nút
            dialog.geometry(f"600x400+{x}+{y}")
            
            # Frame chính với padding đủ để các nút không bị che
            main_frame = tk.Frame(dialog, padx=30, pady=30, bg='#f5f5f5')
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # Icon thông tin
            icon_frame = tk.Frame(main_frame, bg='#f5f5f5')
            icon_frame.pack(pady=(0, 15))
            icon_label = tk.Label(icon_frame, text="💾", font=('Arial', 32), bg='#f5f5f5')
            icon_label.pack()
            
            # Thông điệp
            from datetime import datetime
            backup_time = datetime.fromtimestamp(timestamp).strftime("%d/%m/%Y %H:%M:%S") if timestamp else "Không xác định"
            
            message_text = (
                f"Phát hiện dữ liệu backup từ lần chạy trước!\n\n"
                f"Thời gian backup: {backup_time}\n\n"
            )
            
            if has_scanned_items:
                message_text += f"• Dữ liệu đang quét: {len(scanned_items_backup)} tựa\n"
            if has_tong_hop_data:
                message_text += f"• Dữ liệu tổng hợp: {len(tong_hop_data_backup)} dòng\n"
            if current_box_number_backup:
                message_text += f"• Thùng đang kiểm kê: {current_box_number_backup}\n"
            
            message_text += "\nBạn có muốn khôi phục dữ liệu này không?"
            
            message_frame = tk.Frame(main_frame, bg='#f5f5f5')
            message_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
            message_label = tk.Label(
                message_frame, 
                text=message_text,
                font=('Arial', 12),
                justify=tk.CENTER,
                bg='#f5f5f5',
                fg='#333333',
                wraplength=540
            )
            message_label.pack()
            
            # Frame chứa các nút với padding đủ để không bị che
            button_frame = tk.Frame(main_frame, bg='#f5f5f5')
            button_frame.pack(pady=(15, 25))
            
            result = {'value': None}
            
            def on_restore():
                result['value'] = 'restore'
                dialog.destroy()
            
            def on_discard():
                result['value'] = 'discard'
                dialog.destroy()
            
            # Nút "Khôi phục"
            restore_btn = tk.Button(
                button_frame,
                text="Khôi phục",
                command=on_restore,
                font=('Arial', 12, 'bold'),
                bg='#4CAF50',
                fg='white',
                padx=25,
                pady=12,
                relief=tk.FLAT,
                cursor='hand2',
                activebackground='#45a049',
                activeforeground='white',
                bd=0,
                highlightthickness=0
            )
            restore_btn.pack(side=tk.LEFT, padx=15)
            
            # Nút "Bỏ qua"
            discard_btn = tk.Button(
                button_frame,
                text="Bỏ qua",
                command=on_discard,
                font=('Arial', 12, 'bold'),
                bg='#757575',
                fg='white',
                padx=25,
                pady=12,
                relief=tk.FLAT,
                cursor='hand2',
                activebackground='#616161',
                activeforeground='white',
                bd=0,
                highlightthickness=0
            )
            discard_btn.pack(side=tk.LEFT, padx=15)
            
            # Đợi dialog đóng
            dialog.wait_window()
            
            # Xử lý kết quả
            if result['value'] == 'restore':
                # Khôi phục dữ liệu
                self.scanned_items = scanned_items_backup
                self.tong_hop_data = tong_hop_data_backup
                self.current_box_number = current_box_number_backup
                
                # Cập nhật UI sau khi khôi phục
                if hasattr(self, 'tong_hop_tree') and self.tong_hop_tree:
                    self.update_tong_hop_table()
                
                # Nếu có dữ liệu đang quét, hiển thị lại trong bảng
                if self.scanned_items:
                    # Hiển thị lại dữ liệu đã quét trong bảng
                    self.clear_table()
                    for isbn, info in self.scanned_items.items():
                        # Tạo lại item trong tree
                        item_id = self.tree.insert('', tk.END, values=(
                            len(self.tree.get_children()) + 1,  # STT
                            isbn,
                            info.get('tua', ''),
                            info.get('ton_thuc_te', ''),
                            info.get('so_thung', ''),
                            info.get('ton_trong_thung', ''),
                            info.get('tinh_trang', ''),
                            info.get('ghi_chu', '')
                        ))
                        # Cập nhật item_id trong scanned_items
                        info['item_id'] = item_id
                    
                    # Cập nhật số tựa đã quét
                    if hasattr(self, 'so_tua_var'):
                        self.so_tua_var.set(str(len(self.scanned_items)))
                    if hasattr(self, 'so_tua_da_quet_var') and self.current_box_number:
                        so_tua_da_quet = self.count_scanned_titles_for_box(self.current_box_number)
                        self.so_tua_da_quet_var.set(str(so_tua_da_quet))
                    
                    # Cập nhật số thùng nếu có
                    if hasattr(self, 'so_thung_var') and self.current_box_number:
                        self.so_thung_var.set(self.current_box_number)
                
                messagebox.showinfo("Thành công", 
                    f"Đã khôi phục dữ liệu!\n\n"
                    f"Dữ liệu đang quét: {len(self.scanned_items)} tựa\n"
                    f"Dữ liệu tổng hợp: {len(self.tong_hop_data)} dòng")
            else:
                # Người dùng không muốn khôi phục - giữ nguyên file backup (không xóa)
                pass
        
        except Exception as e:
            # Không hiển thị lỗi cho người dùng, chỉ log
            print(f"Lỗi khi khôi phục backup: {str(e)}")
    
    def start_auto_save(self):
        """Bắt đầu auto-save định kỳ (mỗi 30 giây)"""
        def auto_save_periodic():
            # Chỉ lưu nếu có dữ liệu
            if self.scanned_items or self.tong_hop_data:
                self.save_backup()
            # Lên lịch lại sau 30 giây
            self.root.after(30000, auto_save_periodic)
        
        # Bắt đầu auto-save ngay lập tức, sau đó mỗi 30 giây
        auto_save_periodic()  # Chạy ngay lần đầu
    
    def save_backup_on_change(self):
        """Lưu backup ngay lập tức khi có thay đổi dữ liệu"""
        # Delay một chút để tránh lưu quá nhiều lần
        if hasattr(self, '_backup_scheduled') and self._backup_scheduled is not None:
            try:
                # Chỉ cancel nếu ID hợp lệ
                if isinstance(self._backup_scheduled, str) and self._backup_scheduled:
                    self.root.after_cancel(self._backup_scheduled)
            except (ValueError, TypeError, tk.TclError):
                # Bỏ qua lỗi nếu ID không hợp lệ
                pass
            self._backup_scheduled = None
        
        def do_save():
            try:
                self.save_backup()
            except Exception as e:
                # Không hiển thị lỗi cho người dùng
                print(f"Error saving backup: {e}")
            finally:
                self._backup_scheduled = None
        
        try:
            after_id = self.root.after(2000, do_save)  # Lưu sau 2 giây
            # Chỉ lưu nếu after() trả về giá trị hợp lệ
            if after_id is not None:
                self._backup_scheduled = after_id
            else:
                # Nếu after() trả về None, lưu ngay
                self.save_backup()
        except Exception as e:
            # Nếu có lỗi với after(), lưu ngay
            print(f"Error scheduling backup: {e}")
            try:
                self.save_backup()
            except:
                pass
    
    def setup_signal_handlers(self):
        """Đăng ký xử lý signal để lưu backup khi shutdown (cúp điện, tắt máy)"""
        def signal_handler(signum, frame):
            """Xử lý signal shutdown - lưu backup ngay lập tức"""
            try:
                # Hủy scheduled backup nếu có để tránh conflict
                if hasattr(self, '_backup_scheduled') and self._backup_scheduled:
                    try:
                        self.root.after_cancel(self._backup_scheduled)
                    except:
                        pass
                
                # Lưu backup ngay lập tức
                if self.scanned_items or self.tong_hop_data:
                    self.save_backup()
                    print(f"Đã lưu backup khi nhận signal {signum}")
            except Exception as e:
                print(f"Lỗi khi lưu backup trong signal handler: {str(e)}")
        
        def atexit_handler():
            """Xử lý khi exit - lưu backup"""
            try:
                if self.scanned_items or self.tong_hop_data:
                    self.save_backup()
                    print("Đã lưu backup khi exit")
            except Exception as e:
                print(f"Lỗi khi lưu backup trong atexit handler: {str(e)}")
        
        # Đăng ký signal handlers (chỉ trên Unix/Linux/macOS, Windows không hỗ trợ tốt)
        if hasattr(signal, 'SIGTERM'):
            try:
                signal.signal(signal.SIGTERM, signal_handler)
            except:
                pass
        
        if hasattr(signal, 'SIGINT'):
            try:
                signal.signal(signal.SIGINT, signal_handler)
            except:
                pass
        
        # Đăng ký atexit handler (hoạt động trên cả Windows và Unix)
        atexit.register(atexit_handler)
    
    def on_closing(self):
        """Xử lý sự kiện đóng cửa sổ - kiểm tra dữ liệu chưa lưu"""
        # Kiểm tra xem có dữ liệu chưa lưu không (cả scanned_items và tong_hop_data)
        has_scanned_items = self.scanned_items and len(self.scanned_items) > 0
        has_tong_hop_data = self.tong_hop_data and len(self.tong_hop_data) > 0
        
        if has_scanned_items or has_tong_hop_data:
            # Có dữ liệu chưa lưu, hiển thị dialog cảnh báo
            dialog = tk.Toplevel(self.root)
            dialog.title("Cảnh báo")
            dialog.geometry("600x320")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()
            dialog.configure(bg='#f5f5f5')
            
            # Đặt dialog ở giữa màn hình
            dialog.update_idletasks()
            x = (dialog.winfo_screenwidth() // 2) - (600 // 2)
            y = (dialog.winfo_screenheight() // 2) - (320 // 2)
            dialog.geometry(f"600x320+{x}+{y}")
            
            # Frame chính với padding lớn hơn
            main_frame = tk.Frame(dialog, padx=30, pady=25, bg='#f5f5f5')
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # Icon cảnh báo
            icon_frame = tk.Frame(main_frame, bg='#f5f5f5')
            icon_frame.pack(pady=(0, 15))
            icon_label = tk.Label(icon_frame, text="⚠️", font=('Arial', 32), bg='#f5f5f5')
            icon_label.pack()
            
            # Thông điệp
            message_parts = []
            if has_scanned_items:
                message_parts.append(f"• {len(self.scanned_items)} tựa đang quét chưa được lưu vào Tổng hợp")
            if has_tong_hop_data:
                message_parts.append(f"• {len(self.tong_hop_data)} dòng dữ liệu trong tab Tổng hợp chưa được lưu backup")
            
            message_text = (
                f"Bạn có muốn lưu backup dữ liệu trước khi đóng phần mềm không?\n\n"
                + "\n".join(message_parts) + "\n\n"
                f"Bạn muốn làm gì?"
            )
            
            message_frame = tk.Frame(main_frame, bg='#f5f5f5')
            message_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 25))
            message_label = tk.Label(
                message_frame,
                text=message_text,
                font=('Arial', 12),
                justify=tk.CENTER,
                bg='#f5f5f5',
                fg='#333333',
                wraplength=540
            )
            message_label.pack()
            
            # Frame chứa các nút
            button_frame = tk.Frame(main_frame, bg='#f5f5f5')
            button_frame.pack(pady=(10, 0))
            
            result = {'value': None}
            
            def on_save():
                result['value'] = 'save'
                dialog.destroy()
            
            def on_close():
                result['value'] = 'close'
                dialog.destroy()
            
            # Nút "Lưu"
            save_btn = tk.Button(
                button_frame,
                text="Lưu",
                command=on_save,
                font=('Arial', 12, 'bold'),
                bg='#4CAF50',
                fg='white',
                padx=25,
                pady=12,
                relief=tk.FLAT,
                cursor='hand2',
                activebackground='#45a049',
                activeforeground='white',
                bd=0,
                highlightthickness=0
            )
            save_btn.pack(side=tk.LEFT, padx=15)
            
            # Nút "Đóng không lưu"
            close_btn = tk.Button(
                button_frame,
                text="Đóng không lưu",
                command=on_close,
                font=('Arial', 12, 'bold'),
                bg='#757575',
                fg='white',
                padx=25,
                pady=12,
                relief=tk.FLAT,
                cursor='hand2',
                activebackground='#616161',
                activeforeground='white',
                bd=0,
                highlightthickness=0
            )
            close_btn.pack(side=tk.LEFT, padx=15)
            
            # Đợi dialog đóng
            dialog.wait_window()
            
            # Xử lý kết quả
            if result['value'] == 'save':
                # Lưu vào backup file trước khi đóng
                try:
                    self.save_backup()
                    # Sau khi lưu backup xong, đóng phần mềm
                    self.root.quit()
                    self.root.destroy()
                except Exception as e:
                    # Nếu lưu backup lỗi, hỏi lại có muốn đóng không
                    error_result = messagebox.askyesno(
                        "Lỗi",
                        f"Không thể lưu backup: {str(e)}\n\n"
                        f"Bạn có muốn đóng phần mềm mà không lưu không?"
                    )
                    if error_result:
                        self.root.quit()
                        self.root.destroy()
            elif result['value'] == 'close':
                # Đóng phần mềm luôn, không lưu gì cả
                self.root.quit()
                self.root.destroy()
            # Nếu result['value'] là None (người dùng đóng dialog bằng X), không làm gì cả
        else:
            # Không có dữ liệu chưa lưu, đóng phần mềm bình thường
            self.root.quit()
            self.root.destroy()

def main():
    try:
        root = tk.Tk()
        # Hiển thị window ngay lập tức để tăng tốc độ khởi động (perceived speed)
        root.deiconify()
        root.update_idletasks()  # Cập nhật UI ngay lập tức
        
        # Tạo app (sẽ load dữ liệu sau khi UI hiển thị)
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

