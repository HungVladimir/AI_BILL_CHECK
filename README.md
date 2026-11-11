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

## 📂 CẤU TRÚC THƯ MỤC PRIVATE (TRÊN MÁY BẠN)

Mặc định hệ thống trỏ tới: `D:\Private_Documents\Doc\AI-Bill` (có thể đổi trong `config.py` hoặc `.env`).
Hỗ trợ cả 2 cách sắp xếp:

### Cách 1: Để trực tiếp (Tiện nhất, đang dùng)
```text
D:\Private_Documents\Doc\AI-Bill\
├── (AI Crew) thanh toán hóa đơn tháng 2026.08.xlsx   (File 1: Ảnh hóa đơn)
├── [AI Crew] member share.xlsx                      (File 2: KnoxID & Đơn giá)
└── output_reports/                                  (Tự động tạo ra)
    ├── Bao_cao_doi_soat_hoa_don_AI_2026-08_TEST.xlsx
    └── extracted_images_2026-08/
        ├── 7ac3d0ee16_Nguyễn_Thị_Kim_Ngân_ChatGPT.png
        └── ...
```

### Cách 2: Phân loại theo thư mục con
```text
D:\Private_Documents\Doc\AI-Bill\
├── invoices_file1/
│   ├── (AI Crew) thanh toán hóa đơn tháng 2026.08.xlsx
│   └── (AI Crew) thanh toán hóa đơn tháng 2026.10.xlsx
├── members_file2/
│   └── [AI Crew] member share.xlsx
└── output_reports/
    ├── Bao_cao_doi_soat_hoa_don_AI_2026-08_TEST.xlsx
    └── extracted_images_2026-08/
```

---

## 🚀 HƯỚNG DẪN CHẠY DỰ ÁN

### 1. CHẠY TEST ĐỐI SOÁT (AN TOÀN - KHÔNG GHI DB/MINIO)
> Chế độ này đọc file thật từ thư mục Private, bóc tách ảnh thật, đối soát khớp giá, xuất file Excel báo cáo và xuất ảnh bóc tách vào `output_reports/` để bạn mở lên kiểm tra mắt thấy tai nghe.

Chỉ cần mở Terminal tại thư mục dự án và chạy:
```bash
python main_sync.py
```
*(Mặc định không truyền cờ gì là chạy chế độ TEST an toàn)*

Hoặc truyền rõ ràng:
```bash
python main_sync.py --test
```

**Kết quả nhận được:**
* 📊 **File Excel Báo cáo:** `D:\Private_Documents\Doc\AI-Bill\output_reports\Bao_cao_doi_soat_hoa_don_AI_{Tháng}_TEST.xlsx`
* 📸 **Folder ảnh đã bóc tách:** `D:\Private_Documents\Doc\AI-Bill\output_reports\extracted_images_{Tháng}\`

---

### 2. CHẠY ĐỒNG BỘ THẬT (UPLOAD MINIO & LƯU ORACLE DB)
> Sau khi bạn đã mở file Excel báo cáo và kiểm tra toàn bộ thông tin đã khớp chuẩn 100%, hãy chạy lệnh này để chính thức lưu vào hệ thống:

```bash
python main_sync.py --sync
```

Nếu bạn muốn dọn dẹp các bản ghi cũ của tháng đó trước khi nạp lại (tránh trùng lặp dữ liệu):
```bash
python main_sync.py --sync --clear-month
```

**Hệ thống sẽ:**
1. Upload toàn bộ ảnh vào MinIO S3: `aibill/{YYYY-MM}/{ID}.png`.
2. Insert toàn bộ dòng hóa đơn vào Oracle DB: `LOCAL_AI.AI_INVOICE_CHECK`.
3. Xuất file Excel báo cáo hoàn tất: `Bao_cao_doi_soat_hoa_don_AI_{Tháng}_SYNCED.xlsx`.
4. Dữ liệu ngay lập tức hiển thị trên màn hình giao diện web OneAI!

---

### 3. TÙY CHỌN TRUYỀN THAM SỐ (NẾU CẦN CHỈ ĐỊNH ĐẶC BIỆT)

* **Chỉ định tháng thanh toán cụ thể:**
  ```bash
  python main_sync.py --month "2026-08"
  ```
* **Chỉ định thư mục Private tùy ý:**
  ```bash
  python main_sync.py --dir "D:\Thu_Muc_Khac\Doc\AI-Bill"
  ```
* **Chỉ định file lẻ trực tiếp:**
  ```bash
  python main_sync.py --file1 "D:\path\to\file1.xlsx" --file2 "D:\path\to\file2.xlsx"
  ```
