# 🤖 AI BILL PROCESSOR (MODULE XỬ LÝ HÓA ĐƠN AI)

Hệ thống tự động hóa trích xuất ảnh hóa đơn nguyên bản (RAW) từ Excel, bóc tách đơn giá từng dịch vụ AI (ChatGPT Plus, Claude Pro, Gemini Pro), tải ảnh lên MinIO S3 và đồng bộ dữ liệu vào bảng `LOCAL_AI.AI_INVOICE_CHECK` trong Oracle DB.

---

## 🌟 TÍNH NĂNG NỔI BẬT

1. **Bảo toàn 100% chất lượng ảnh gốc (RAW)**:
   - Trích xuất trực tiếp từ các file nén OpenXML (`xl/drawings/` + `xl/media/`).
   - Không qua nén lại, giữ nguyên kích thước pixel gốc, bitrate và metadata của ảnh.
2. **Tự động nhận diện Tháng từ tên file**:
   - `(AI Crew) thanh toán hóa đơn tháng 2026.10.xlsx` ➔ Tự động bóc tách `2026-10` và ánh xạ sang cột `Oct-26` ở File 2.
3. **Phân tách nhiều hóa đơn trên cùng 1 nhân sự**:
   - Nhân sự có cả ChatGPT và Claude (hoặc Gemini) sẽ tự động được phân tách thành các dòng độc lập trong DB, mỗi dòng có ID riêng, đơn giá riêng và link ảnh MinIO riêng.
4. **Chuẩn hóa Email thanh toán theo quy tắc**:
   - Lấy giá trị từ cột `Thanh toán` (sheet *Thanh toán* của File 2).
   - Nếu có đuôi `@yahoo.com.vn` hoặc `@gmail.com` ➔ Giữ nguyên.
   - Nếu chỉ có tên người dùng ➔ Tự động thêm `@gmail.com`.
   - Gán `BILLING_EMAIL = SUBMITTER_EMAIL` và `EMAIL_MATCH = 1`.
5. **Sinh khóa định danh ngẫu nhiên chuẩn Oracle 10 ký tự**:
   - Tương thích 100% với hàm `LOCAL_AI.FN_MAKE_ID(10)`.
6. **Đồng bộ MinIO và Oracle DB**:
   - MinIO Path: `aibill/{YYYY-MM}/{ID}.png` (MIME: `image/png`).
   - Oracle Table: `LOCAL_AI.AI_INVOICE_CHECK` (Đầy đủ 22 cột).
7. **Chế độ kiểm thử an toàn (Dry-Run)**:
   - Cho phép chạy thử nghiệm bóc tách, đối soát và xuất báo cáo CSV/JSON mà không ghi đè vào DB hay MinIO production.

---

## 📁 CẤU TRÚC THƯ MỤC

```text
MEP_OneAI_System/ai_bill_processor/
├── config.py             # Cấu hình dịch vụ (ChatGPT, Claude, Gemini), MinIO, Oracle DB
├── id_generator.py       # Thuật toán sinh ID 10 ký tự tương thích Oracle FN_MAKE_ID
├── excel_extractor.py    # Bộ bóc tách ảnh RAW nguyên bản từ OpenXML và tọa độ Cell
├── excel_matcher.py      # Bộ đối soát dữ liệu File 2 (KnoxID, Email, Đơn giá theo tháng)
├── minio_client.py       # Quản lý tải ảnh lên MinIO S3 (aibill/{YYYY-MM}/{ID}.png)
├── db_sync.py            # Đồng bộ dữ liệu vào bảng LOCAL_AI.AI_INVOICE_CHECK
├── main_sync.py          # CLI Runner chính chạy toàn bộ quy trình
├── test_pipeline.py      # Bộ kiểm thử tự động 7 bước
├── input/                # Thư mục đặt 2 file Excel đầu vào
├── output_debug/         # Thư mục chứa báo cáo CSV, JSON và ảnh preview
└── test_data/            # Dữ liệu mẫu giả lập cho unit test
```

---

## 🚀 HƯỚNG DẪN SỬ DỤNG

### 1. Chạy kiểm thử tự động (Unit Test 7 bước)
```bash
python MEP_OneAI_System/ai_bill_processor/test_pipeline.py
```

### 2. Chạy thử nghiệm an toàn (Dry-Run)
Đặt 2 file Excel vào thư mục `MEP_OneAI_System/ai_bill_processor/input/`, sau đó chạy:
```bash
python MEP_OneAI_System/ai_bill_processor/main_sync.py --dry-run
```
Hoặc chỉ định đường dẫn file trực tiếp:
```bash
python MEP_OneAI_System/ai_bill_processor/main_sync.py --file1 "đường/dẫn/file1.xlsx" --file2 "đường/dẫn/file2.xlsx" --dry-run
```

### 3. Chạy thực tế (Tải lên MinIO & Nạp vào Oracle DB)
```bash
python MEP_OneAI_System/ai_bill_processor/main_sync.py --file1 "đường/dẫn/file1.xlsx" --file2 "đường/dẫn/file2.xlsx"
```
Khi chạy xong, hệ thống sẽ:
1. Upload toàn bộ ảnh vào MinIO bucket `aibill/{YYYY-MM}/{ID}.png`.
2. Insert toàn bộ bản ghi vào bảng `LOCAL_AI.AI_INVOICE_CHECK`.
3. Xuất file báo cáo kiểm toán chi tiết tại `output_debug/audit_report_{YYYY-MM}.csv`.

