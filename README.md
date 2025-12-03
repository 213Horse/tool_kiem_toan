# Ứng dụng Kiểm Kho - Quét Mã Vạch

Ứng dụng desktop chạy trên Windows và macOS để kiểm tra tồn kho thực tế bằng cách quét mã vạch.

## Tính năng

- Đọc dữ liệu từ file Excel `DuLieuDauVao.xlsx`
- Nhập số thùng để load danh sách tựa trong thùng đó
- Quét/nhập ISBN để tự động điền thông tin
- So sánh tồn thực tế với tồn trong thùng và highlight màu đỏ nếu khác nhau
- Lưu kết quả kiểm tra ra file Excel

## Cài đặt và Chạy

### Cách 1: Chạy trực tiếp từ source code (cần Python)

1. Cài đặt Python 3.8 trở lên
2. Cài đặt các thư viện cần thiết:
```bash
pip install -r requirements.txt
```

3. Chạy ứng dụng:
```bash
python kiem_kho_app.py
```

### Cách 2: Tạo file thực thi (không cần Python)

#### Trên Windows:
1. Mở Command Prompt hoặc PowerShell
2. Chạy lệnh:
```bash
build_windows.bat
```
3. File thực thi sẽ nằm trong thư mục `dist/KiemKhoApp.exe`
4. Double-click vào `KiemKhoApp.exe` để chạy

#### Trên macOS:
1. Mở Terminal
2. Chạy lệnh:
```bash
chmod +x build_macos.sh
./build_macos.sh
```
3. File thực thi sẽ nằm trong thư mục `dist/KiemKhoApp`
4. Double-click vào `KiemKhoApp` để chạy

## Hướng dẫn sử dụng

### Bước 1: Khởi động ứng dụng
- Double-click vào file `KiemKhoApp.exe` (Windows) hoặc `KiemKhoApp` (macOS)
- Đảm bảo file `DuLieuDauVao.xlsx` cùng thư mục với file thực thi

### Bước 2: Load dữ liệu thùng
1. Nhập số thùng vào ô **"Số thùng (*)"** (màu vàng)
2. Nhấn **Enter** hoặc click nút **"Load"**
3. Ứng dụng sẽ hiển thị số tựa có trong thùng đó ở trên bảng

### Bước 3: Quét mã vạch và kiểm tra
1. **Quét mã vạch**: 
   - Sử dụng máy quét mã vạch (máy quét sẽ tự động nhập ISBN như gõ phím)
   - Hoặc nhập ISBN thủ công vào ô **"Quét/Nhập ISBN"**
   - Nhấn **Enter** sau khi quét/nhập
   - Thông tin tựa sẽ tự động hiển thị trong bảng với các cột:
     - ISBN
     - Tựa (tên sách)
     - Tồn thực tế (để trống, bạn sẽ nhập)
     - Số thùng
     - Tồn tựa trong thùng (từ dữ liệu Excel)
     - Ghi chú (để trống)

2. **Nhập tồn thực tế**: 
   - Double-click vào ô **"Tồn thực tế"** của tựa vừa quét
   - Nhập số lượng thực tế bạn đếm được
   - Nhấn Enter hoặc click "Lưu"

3. **Kiểm tra lệch**: 
   - Nếu tồn thực tế khác tồn trong thùng, ô **"Tồn thực tế"** sẽ tự động **tô đỏ**
   - Điều này giúp bạn dễ dàng phát hiện các tựa có số lượng không khớp

4. **Nhập ghi chú** (tùy chọn): 
   - Double-click vào ô **"Ghi chú"** để nhập ghi chú cho tựa đó
   - Ví dụ: "Thiếu 2 cuốn", "Thừa 1 cuốn", v.v.

5. **Tiếp tục quét**: 
   - Sau mỗi lần quét, ô ISBN sẽ tự động được clear để sẵn sàng quét tiếp
   - Lặp lại cho đến khi quét hết các tựa trong thùng

### Bước 4: Lưu kết quả
1. Click nút **"SAVE"** (màu vàng, góc phải trên)
2. Chọn nơi lưu file Excel kết quả
3. File sẽ chứa tất cả các tựa đã kiểm tra với thông tin đầy đủ

## Lưu ý quan trọng

- **File Excel**: File `DuLieuDauVao.xlsx` phải cùng thư mục với file thực thi
- **Máy quét mã vạch**: Máy quét sẽ hoạt động như bàn phím, tự động nhập ISBN khi quét
- **ISBN không khớp**: Ứng dụng sẽ tự động tìm ISBN tương tự nếu không khớp hoàn toàn (bỏ qua ký tự đặc biệt)
- **Quét lại**: Nếu quét cùng một ISBN nhiều lần, dữ liệu sẽ được cập nhật (không tạo bản ghi trùng)
- **Highlight đỏ**: Ô sẽ tự động tô đỏ khi tồn thực tế ≠ tồn trong thùng

## Yêu cầu file Excel

File `DuLieuDauVao.xlsx` cần có các cột sau:
- **Số thùng**: Số thùng chứa tựa
- **ISBN**: Mã ISBN của tựa
- **Tựa**: Tên tựa sách
- **Tồn từng tựa**: Số lượng tồn của tựa trong thùng

Lưu ý: Tên cột có thể có khoảng trắng, ứng dụng sẽ tự động nhận diện.

## Lưu ý

- File `DuLieuDauVao.xlsx` phải cùng thư mục với file thực thi
- Máy quét mã vạch sẽ tự động nhập ISBN khi quét (giống như gõ phím)
- Có thể nhập ISBN thủ công nếu không có máy quét

# tool_kiem_toan
