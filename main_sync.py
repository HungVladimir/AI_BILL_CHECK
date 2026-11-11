# ==============================================================================
# MODULE: AI BILL PROCESSOR - MAIN RUNNER (CLI)
# File: ai_bill_processor/main_sync.py
# ==============================================================================
r"""
Chương trình chính điều phối toàn bộ luồng xử lý Hóa đơn AI:
  1. Tự động nhận diện File 1 & File 2 từ thư mục Private (D:\Private_Documents\Doc\AI-Bill).
  2. Tự động đọc Tháng thanh toán từ tên File 1 (hoặc tham số --month).
  3. Trích xuất ảnh RAW nguyên bản theo từng ô (ChatGPT, Claude, Gemini).
  4. Đọc File 2, chuẩn hóa email, đối soát KnoxID và bóc tách đơn giá.
  5. Chế độ TEST (Mặc định):
     - KHÔNG upload MinIO, KHÔNG ghi Database.
     - Xuất Báo cáo đối soát chi tiết ra file Excel (.xlsx) kèm thư mục ảnh bóc tách
       lưu thẳng vào thư mục Private (output_reports/).
  6. Chế độ SYNC (--sync):
     - Upload ảnh lên MinIO S3 (aibill/{YYYY-MM}/{ID}.png).
     - Ghi bản ghi vào bảng Oracle DB (LOCAL_AI.AI_INVOICE_CHECK).
     - Xuất Báo cáo đối soát hoàn tất (.xlsx) vào thư mục Private.

Cách chạy:
  python main_sync.py          (Chạy TEST đối soát an toàn, xuất Excel ra Private)
  python main_sync.py --sync   (Đồng bộ THẬT lên MinIO & Database)
"""
import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path
import pandas as pd

# Thiết lập mã hóa UTF-8 cho Terminal Windows
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

try:
    from .config import (
        BASE_DIR,
        AI_BILL_PRIVATE_DIR,
        DIR_EXCEL1,
        DIR_EXCEL2,
        DIR_EXCEL_OUTPUT,
        DIR_PROJECT_IMAGES,
        TABLE_NAME,
        MINIO_BUCKET_NAME,
        MINIO_PREFIX,
        extract_month_from_filename,
        locate_ai_bill_files
    )
    from .excel_extractor import ExcelImageExtractor
    from .excel_matcher import ExcelDataMatcher
    from .excel_reporter import export_reconciliation_excel
    from .minio_client import MinioUploader
    from .db_sync import DatabaseSynchronizer
except ImportError:
    from config import (
        BASE_DIR,
        AI_BILL_PRIVATE_DIR,
        DIR_EXCEL1,
        DIR_EXCEL2,
        DIR_EXCEL_OUTPUT,
        DIR_PROJECT_IMAGES,
        TABLE_NAME,
        MINIO_BUCKET_NAME,
        MINIO_PREFIX,
        extract_month_from_filename,
        locate_ai_bill_files
    )
    from excel_extractor import ExcelImageExtractor
    from excel_matcher import ExcelDataMatcher
    from excel_reporter import export_reconciliation_excel
    from minio_client import MinioUploader
    from db_sync import DatabaseSynchronizer


def parse_args():
    parser = argparse.ArgumentParser(description="AI Bill Processor - Tự động bóc tách hóa đơn AI và đối soát MinIO / DB")
    parser.add_argument("--test", action="store_true", help="Chế độ TEST (Mặc định): Đọc file thật, đối soát, xuất file Excel báo cáo vào thư mục Private, KHÔNG ghi DB/MinIO")
    parser.add_argument("--sync", "--save", dest="sync", action="store_true", help="Chế độ ĐỒNG BỘ THẬT: Tải ảnh lên MinIO S3 và lưu bản ghi vào Oracle DB")
    parser.add_argument("--dry-run", action="store_true", help="Tương đương --test (Không ghi DB, không upload MinIO)")
    parser.add_argument("--dir", type=str, default=None, help="Đường dẫn thư mục Private tùy chỉnh (Mặc định lấy từ config/AI_BILL_PRIVATE_DIR)")
    parser.add_argument("--file1", type=str, default=None, help="Đường dẫn file Excel 1 (chứa ảnh hóa đơn)")
    parser.add_argument("--file2", type=str, default=None, help="Đường dẫn file Excel 2 (chứa đơn giá và KnoxID)")
    parser.add_argument("--month", type=str, default=None, help="Tháng thanh toán YYYY-MM (Mặc định tự đọc từ tên File 1)")
    parser.add_argument("--clear-month", action="store_true", help="Xóa dữ liệu cũ của tháng đó trong DB trước khi nạp mới (tránh trùng lặp khi chạy lại sync)")
    parser.add_argument("--skip-minio", action="store_true", help="Bỏ qua bước upload MinIO khi chạy sync")
    parser.add_argument("--skip-db", action="store_true", help="Bỏ qua bước insert Database khi chạy sync")
    return parser.parse_args()


def main():
    args = parse_args()

    # Xác định chế độ chạy (Mặc định an toàn là TEST nếu không truyền --sync)
    is_test_mode = True
    if args.sync:
        is_test_mode = False
    elif args.test or args.dry_run:
        is_test_mode = True

    print("=" * 75)
    if is_test_mode:
        print("🧪 AI BILL PROCESSOR - CHẾ ĐỘ TEST ĐỐI SOÁT (AN TOÀN - KHÔNG GHI DB/MINIO)")
        print("   (File báo cáo đối soát Excel và ảnh bóc tách sẽ được lưu vào thư mục Private)")
    else:
        print("🚀 AI BILL PROCESSOR - CHẾ ĐỘ ĐỒNG BỘ THẬT (UPLOAD MINIO & LƯU ORACLE DB)")
    print("=" * 75)

    # 1. Xác định thư mục Private & Dò tìm File 1, File 2
    target_private_dir = Path(args.dir.strip().strip('"').strip("'")) if args.dir else AI_BILL_PRIVATE_DIR
    print(f"📁 Thư mục Private cấu hình: {target_private_dir}")

    auto_f1, auto_f2, excel_out_dir, img_out_dir = locate_ai_bill_files(target_private_dir, target_month=args.month)

    # Ưu tiên tham số truyền tay nếu có
    f1_path = Path(args.file1.strip().strip('"').strip("'")) if args.file1 else auto_f1
    f2_path = Path(args.file2.strip().strip('"').strip("'")) if args.file2 else auto_f2

    # Nếu vẫn chưa tìm thấy trong Private, fallback tìm trong thư mục cục bộ input/
    if not f1_path or not f2_path:
        local_input = BASE_DIR / "input"
        loc_f1, loc_f2, _, _ = locate_ai_bill_files(local_input, target_month=args.month)
        f1_path = f1_path or loc_f1
        f2_path = f2_path or loc_f2

    # Chế độ nhập đường dẫn nếu hoàn toàn chưa thấy file
    if not f1_path or not f2_path:
        print("\n💡 CHƯA TÌM THẤY ĐỦ 2 FILE EXCEL TRONG THƯ MỤC PRIVATE:")
        print(f"   Thư mục kiểm tra: {target_private_dir}")
        print("   (Mẹo: Bạn có thể dán đường dẫn hoặc KÉO THẢ file từ Windows Explorer vào đây)\n")
        try:
            if not f1_path:
                inp_f1 = input("👉 Nhập đường dẫn File 1 (chứa ảnh hóa đơn): ").strip().strip('"').strip("'")
                if inp_f1:
                    f1_path = Path(inp_f1)

            if not f2_path:
                inp_f2 = input("👉 Nhập đường dẫn File 2 (chứa đơn giá & KnoxID): ").strip().strip('"').strip("'")
                if inp_f2:
                    f2_path = Path(inp_f2)
        except (KeyboardInterrupt, EOFError):
            print("\nĐã hủy thao tác.")
            sys.exit(0)

    if not f1_path or not f1_path.exists():
        print(f"\n❌ LỖI: Không tìm thấy File 1 tại: {f1_path}")
        print("👉 Vui lòng kiểm tra lại file trong thư mục Private hoặc truyền qua tham số --file1")
        sys.exit(1)

    if not f2_path or not f2_path.exists():
        print(f"\n❌ LỖI: Không tìm thấy File 2 tại: {f2_path}")
        print("👉 Vui lòng kiểm tra lại file trong thư mục Private hoặc truyền qua tham số --file2")
        sys.exit(1)

    print(f"   - File 1 (Hóa đơn ảnh): {f1_path.name}")
    print(f"   - File 2 (Member & Đơn giá): {f2_path.name}")

    # 2. Xác định Tháng thanh toán
    month = args.month or extract_month_from_filename(f1_path.name)
    print(f"📅 Tháng thanh toán xác định: {month}")

    # Đảm bảo thư mục output_reports trong Private tồn tại (để lưu file Excel)
    excel_out_dir.mkdir(parents=True, exist_ok=True)
    # Thư mục lưu ảnh: Luôn nằm trong Project folder
    img_month_dir = DIR_PROJECT_IMAGES / month
    img_month_dir.mkdir(parents=True, exist_ok=True)

    # 3. BƯỚC 1: Trích xuất ảnh RAW từ File 1
    print("\n" + "-" * 75)
    print("BƯỚC 1: TRÍCH XUẤT ẢNH HÓA ĐƠN RAW TỪ FILE 1 (OPENXML)")
    print("-" * 75)
    extractor = ExcelImageExtractor(str(f1_path))
    invoices = extractor.extract_invoices()

    if not invoices:
        print("⚠️ Không tìm thấy ảnh hóa đơn nào trong File 1. Kết thúc.")
        sys.exit(0)

    # 4. BƯỚC 2: Đối chiếu và ghép dữ liệu với File 2
    print("\n" + "-" * 75)
    print("BƯỚC 2: ĐỐI SOÁT VỚI FILE 2 (KNOXID, EMAIL & ĐƠN GIÁ)")
    print("-" * 75)
    matcher = ExcelDataMatcher(str(f2_path), month_str=month)
    records = matcher.pair_and_build_records(invoices)

    # Gán thêm thông tin kích thước ảnh phục vụ xuất báo cáo
    for rec in records:
        for inv in invoices:
            if rec["SUBMITTER_NAME"] == inv["submitter_name"] and rec["PROVIDER"] == inv["provider_name"]:
                rec["_width"] = inv.get("width", 0)
                rec["_height"] = inv.get("height", 0)
                break

    # 5. XỬ LÝ THEO CHẾ ĐỘ (TEST vs SYNC THẬT)
    if is_test_mode:
        print("\n" + "-" * 75)
        print("BƯỚC 3: XỬ LÝ KIỂM THỬ (TEST MODE) - BỎ QUA GHI MINIO & DATABASE")
        print("-" * 75)
        print("   🔒 [An toàn] Chế độ TEST không can thiệp vào MinIO S3 hay cơ sở dữ liệu.")
        print(f"   📊 File báo cáo Excel sẽ được lưu vào thư mục Private: {excel_out_dir}")
        print(f"   📸 Toàn bộ ảnh hóa đơn bóc tách sẽ lưu vào Project folder: {img_month_dir}")
    else:
        # 5.1 Upload MinIO thật
        print("\n" + "-" * 75)
        print(f"BƯỚC 3: TẢI ẢNH LÊN MINIO S3 (Bucket: {MINIO_BUCKET_NAME} | Prefix: {MINIO_PREFIX}/{month})")
        print("-" * 75)
        minio_uploader = MinioUploader(dry_run=args.skip_minio)
        for rec in records:
            minio_uploader.upload_image(
                object_key=rec["FILE_PATH"],
                raw_bytes=rec["_raw_bytes"],
                output_debug_dir=img_month_dir
            )

        # 5.2 Insert Oracle Database thật
        print("\n" + "-" * 75)
        print(f"BƯỚC 4: LƯU VÀO CƠ SỞ DỮ LIỆU ORACLE (Bảng: {TABLE_NAME})")
        print("-" * 75)
        db_syncer = DatabaseSynchronizer(dry_run=args.skip_db)
        inserted_count = db_syncer.insert_records(
            records=records,
            clear_existing_month=args.clear_month,
            month=month
        )
        db_syncer.close()

    # 6. BƯỚC CUỐI: Xuất Báo cáo đối soát Excel ra thư mục Private & Ảnh vào Project
    print("\n" + "-" * 75)
    print("BƯỚC CUỐI: XUẤT BÁO CÁO ĐỐI SOÁT EXCEL VÀO PRIVATE & ẢNH VÀO PROJECT")
    print("-" * 75)
    excel_report_path = export_reconciliation_excel(
        records=records,
        month_str=month,
        excel_output_dir=excel_out_dir,
        image_output_dir=img_month_dir,
        is_test_mode=is_test_mode
    )

    print(f"📊 [Báo cáo Excel] Đã tạo thành công trong thư mục Private:")
    print(f"   👉 {excel_report_path}")
    print(f"📸 [Thư mục ảnh] Đã bóc tách {len(records)} ảnh gốc vào thư mục Dự án:")
    print(f"   👉 {img_month_dir}")

    print("\n" + "=" * 75)
    if is_test_mode:
        print(f"🎉 HOÀN THÀNH KIỂM THỬ ĐỐI SOÁT: {len(records)} HÓA ĐƠN ĐÃ ĐƯỢC PHÂN TÍCH!")
        print("💡 Bạn có thể mở file Excel trên trong thư mục Private để kiểm tra đối chiếu.")
        print("👉 Khi số liệu đã chuẩn xác và bạn muốn đồng bộ THẬT lên MinIO & Database, hãy chạy:")
        print("   python main_sync.py --sync")
    else:
        print(f"🎉 HOÀN THÀNH ĐỒNG BỘ TOÀN DIỆN: {len(records)} HÓA ĐƠN ĐÃ LƯU THÀNH CÔNG VÀO DB & MINIO!")
        print("🌐 Dữ liệu hiện đã sẵn sàng hiển thị trên giao diện màn hình OneAI.")
    print("=" * 75)


if __name__ == "__main__":
    main()
