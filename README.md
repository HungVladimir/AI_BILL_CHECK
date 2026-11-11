# 🤖 AI BILL PROCESSOR - HỆ THỐNG XỬ LÝ & ĐỐI SOÁT HÓA ĐƠN AI (OPENXML, MINIO, ORACLE DB)

Hệ thống tự động hóa trích xuất ảnh hóa đơn nguyên bản (RAW) từ Excel, bóc tách đơn giá từng dịch vụ AI (ChatGPT Plus, Claude Pro, Gemini Pro), đối soát KnoxID/Email thành viên, tải ảnh lên MinIO S3 và đồng bộ dữ liệu vào bảng `LOCAL_AI.AI_INVOICE_CHECK` trong Oracle DB.

Đặc biệt, hệ thống hỗ trợ **chạy trực tiếp trên máy cá nhân (Local)** với dữ liệu thật trong thư mục Private, cung cấp chế độ **TEST an toàn** (xuất file Excel báo cáo đối soát và thư mục ảnh bóc tách) trước khi chạy **ĐỒNG BỘ THẬT**.

---

## 🌟 TÍNH NĂNG NỔI BẬT

1. **Bảo toàn 100% chất lượng ảnh gốc (RAW)**:
   - Trích xuất trực tiếp từ các file nén OpenXML (`xl/drawings/` + `xl/media/`).
   - Không qua nén lại, giữ nguyên kích thước pixel gốc, bitrate và metadata của ảnh.
2. **Nhận diện Sheet chuẩn theo quy ước nghiệp vụ**:
   - **File 1**: Luôn ưu tiên tuyệt đối sheet **`Tổng hợp hóa đơn`** (chứa cột `Họ và tên`, `Hóa đơn ChatGPT`, `Hóa đơn Claude`...).
   - **File 2**: Luôn ưu tiên tuyệt đối sheet **`Thanh toán`**.
3. **Đối soát KnoxID, Email và Đơn giá đa tầng (File 2)**:
   - **Phần thông tin tài khoản**: Cột `KnoxID`, `Name`, và các cột công cụ (`ChatGPT Plus`, `Claude Pro`, `Gemini Pro`...) với 2 cột con `Đăng nhập` và `Thanh toán`. Hệ thống lấy username ở cột **`Thanh toán`** này để tự động thêm đuôi `@gmail.com` hoặc giữ nguyên `@yahoo.com.vn`.
   - **Phần đơn giá theo tháng**: Các cụm cột tháng (`Aug-26`, `Sep-26`, `Oct-26`...) với các cột con `ChatGPT Plus`, `Claude Pro`, `Total`. Hệ thống tự động bóc tách đơn giá chuẩn xác của từng người cho từng công cụ theo tháng cần xử lý.
4. **Tách biệt lưu trữ an toàn (Private vs Project)**:
   - **File Excel Báo cáo**: Được lưu vào thư mục Private (`D:\Private_Documents\Doc\AI-Bill\output_reports\`).
   - **Ảnh hóa đơn bóc tách**: Được lưu vào thư mục Dự án (`D:\Project\AI_BILL_CHECK\extracted_images\{tháng}\`), vì thư mục Private không hỗ trợ lưu ảnh. Link ảnh trong file Excel trỏ trực tiếp đến thư mục này.
5. **Xuất Báo cáo Đối soát Excel chuyên nghiệp (`excel_reporter.py`)**:
   - Tự động xuất file Excel 2 sheet:
     - **Sheet 1 (Chi tiết đối soát)**: STT, Mã ID, Tháng, Họ Tên, Knox ID, Email, Dịch vụ AI, Gói cước, Số tiền VNĐ, Trạng thái khớp, Đường dẫn MinIO, Kích thước ảnh, Link file ảnh local.
     - **Sheet 2 (Tổng hợp chi phí & KPI)**: Bảng thống kê tổng hợp số lượng và chi phí từng dịch vụ, tỷ trọng %, tổng cộng.
   - Bóc tách toàn bộ ảnh gốc ra thư mục `extracted_images_{tháng}/` với tên `{ID}_{Họ_Tên}_{Provider}.png` để mở xem ngay.
6. **Sinh khóa định danh chuẩn Oracle 10 ký tự**:
   - Tương thích 100% với hàm `LOCAL_AI.FN_MAKE_ID(10)`.
7. **Đồng bộ MinIO S3 & Oracle DB**:
   - MinIO Path: `aibill/{YYYY-MM}/{ID}.png` (MIME: `image/png`).
   - Oracle Table: `LOCAL_AI.AI_INVOICE_CHECK` (Đầy đủ 22 cột theo chuẩn database công ty).
   - Hỗ trợ cờ `--clear-month` tự động dọn dẹp bản ghi cũ nếu chạy lại cùng một tháng.

---

## 📁 CẤU TRÚC THƯ MỤC DỰ ÁN

```text
MEP_OneAI_System/ai_bill_processor/ (và D:\Project\AI_BILL_CHECK\, D:\Project\ai_bill_processor\)
├── config.py             # Cấu hình dịch vụ AI, đường dẫn thư mục Private, MinIO, Oracle DB
├── id_generator.py       # Thuật toán sinh ID 10 ký tự tương thích Oracle FN_MAKE_ID
├── excel_extractor.py    # Bộ bóc tách ảnh RAW từ OpenXML, hỗ trợ Header đa tầng & Fallback
├── excel_matcher.py      # Bộ đối soát dữ liệu File 2 (KnoxID, Email, Đơn giá an toàn)
├── excel_reporter.py     # Bộ xuất Báo cáo đối soát Excel 2 sheet & lưu ảnh bóc tách
├── minio_client.py       # Quản lý tải ảnh lên MinIO S3 (aibill/{YYYY-MM}/{ID}.png)
├── db_sync.py            # Đồng bộ dữ liệu vào bảng LOCAL_AI.AI_INVOICE_CHECK (hỗ trợ clear-month)
├── main_sync.py          # CLI Runner chính điều phối toàn bộ quy trình
├── test_pipeline.py      # Bộ kiểm thử tự động 7 bước
├── input/                # Thư mục dự phòng nếu không dùng thư mục Private
├── output_debug/         # Thư mục chứa báo cáo CSV, JSON và ảnh preview nội bộ
└── test_data/            # Dữ liệu giả lập cho unit test
```

---

## 📂 CẤU TRÚC 2 FOLDER XỬ LÝ ĐA THÁNG (BATCH PROCESSING)

Mặc định hệ thống trỏ tới: `D:\Private_Documents\Doc\AI-Bill` (hoặc cấu hình tùy ý qua `--folder1`, `--folder2`, `--dir`).

```text
D:\Private_Documents\Doc\AI-Bill\
├── invoices_file1/                                 <-- [FOLDER 1]: Chứa các file Excel 1 của từng tháng
│   ├── (AI Crew) thanh toán hóa đơn tháng 8.xlsx    (Tháng 8)
│   ├── (AI Crew) thanh toán hóa đơn tháng 9.xlsx    (Tháng 9)
│   └── (AI Crew) thanh toán hóa đơn tháng 10.xlsx   (Tháng 10)
├── members_file2/                                  <-- [FOLDER 2]: Chứa file Excel 2 tổng hợp giá các tháng
│   └── Chi tiết thanh toán.xlsx                    (Sheet 'Thanh toán' chứa các cột Aug-26, Sep-26, Oct-26...)
└── output_reports/                                 <-- Nơi xuất file Excel báo cáo đối soát tổng hợp
    └── Bao_cao_tong_hop_doi_soat_hoa_don_AI_2026-08_den_2026-10_TEST.xlsx
```

### 🧠 Logic hoạt động:
1. Tự động quét **tất cả** các file Excel trong **Folder 1** (`invoices_file1/`).
2. Tự động nhận diện tháng của từng file (File tháng 8 so với cột `Aug-26`, file tháng 9 so với cột `Sep-26`, file tháng 10 so với `Oct-26` trong **File 2**).
3. Bóc tách ảnh hóa đơn OpenXML nguyên bản của từng tháng lưu vào project (`extracted_images/{YYYY-MM}/`).
4. **Tổng hợp toàn bộ các tháng lại thành 1 FILE BÁO CÁO DUY NHẤT** trong thư mục `output_reports/`:
   - **Sheet 1 (`AI_INVOICE_CHECK`)**: Toàn bộ bản ghi của tất cả các tháng theo chuẩn 22 cột database + link ảnh local.
   - **Sheet 2 (`Tổng hợp chi phí`)**: Có **Bảng 1** thống kê chi tiết theo từng tháng và **Bảng 2** thống kê theo từng nhà cung cấp AI.

---

## 🚀 HƯỚNG DẪN CHẠY DỰ ÁN

### 1. CHẠY TEST ĐỐI SOÁT TẤT CẢ CÁC THÁNG (MẶC ĐỊNH - AN TOÀN)
> Tự động quét toàn bộ Folder 1 và Folder 2, bóc tách ảnh và đối soát giá từng tháng, xuất 1 file Excel báo cáo tổng hợp.

```bash
python main_sync.py
```
Hoặc truyền đường dẫn thư mục tùy chỉnh:
```bash
python main_sync.py --folder1 "D:\Private\invoices_file1" --folder2 "D:\Private\members_file2"
```

### 2. CHẠY ĐỒNG BỘ THẬT LÊN MINIO & DATABASE ORACLE
> Sau khi kiểm tra file Excel báo cáo đã khớp 100%, chạy lệnh sau để tải ảnh toàn bộ các tháng lên MinIO và lưu vào Oracle DB:

```bash
python main_sync.py --sync
```
Nếu muốn dọn dẹp dữ liệu cũ của các tháng đó trước khi nạp lại:
```bash
python main_sync.py --sync --clear-month
```

### 3. CHẠY FILE ĐƠN LẺ (SINGLE FILE)
Nếu bạn chỉ muốn đối soát riêng 1 file của 1 tháng:
```bash
python main_sync.py --file1 "D:\path\to\file1_thang8.xlsx" --file2 "D:\path\to\file2.xlsx"
```
