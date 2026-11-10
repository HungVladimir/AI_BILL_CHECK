# ==============================================================================
# MODULE: AI BILL PROCESSOR - MAIN RUNNER (CLI)
# File: ai_bill_processor/main_sync.py
# ==============================================================================
"""
Chương trình chính điều phối toàn bộ luồng xử lý:
  1. Tiếp nhận 2 file Excel (hoặc tự động quét trong thư mục input/).
  2. Tự động nhận diện Tháng thanh toán từ tên File 1.
  3. Trích xuất ảnh RAW nguyên bản theo từng ô (ChatGPT, Claude, Gemini).
  4. Đọc File 2, chuẩn hóa email, lấy KnoxID và bóc tách đơn giá từng dịch vụ.
  5. Tải ảnh lên MinIO S3 theo chuẩn: aibill/{YYYY-MM}/{ID}.png.
  6. Ghi bản ghi vào bảng LOCAL_AI.AI_INVOICE_CHECK trong Oracle DB.
  7. Xuất file báo cáo kiểm toán CSV/JSON tại output_debug/.

Cách chạy:
  python main_sync.py --file1 "path/to/file1.xlsx" --file2 "path/to/file2.xlsx"
  python main_sync.py --dry-run   (Chạy thử nghiệm an toàn, không ghi DB hay MinIO)
"""
import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path
import pandas as pd

# Thiết lập mã hóa UTF-8 cho Terminal
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

try:
    from .config import (
        BASE_DIR,
        extract_month_from_filename
    )
    from .excel_extractor import ExcelImageExtractor
    from .excel_matcher import ExcelDataMatcher
    from .minio_client import MinioUploader
    from .db_sync import DatabaseSynchronizer
except ImportError:
    from config import (
        BASE_DIR,
        extract_month_from_filename
    )
    from excel_extractor import ExcelImageExtractor
    from excel_matcher import ExcelDataMatcher
    from minio_client import MinioUploader
    from db_sync import DatabaseSynchronizer


def parse_args():
    parser = argparse.ArgumentParser(description="AI Bill Processor - Tự động bóc tách hóa đơn AI và đồng bộ MinIO & DB")
    parser.add_argument("--dir", type=str, default=None, help="Đường dẫn thư mục private bên ngoài chứa 2 file Excel")
    parser.add_argument("--file1", type=str, default=None, help="Đường dẫn file Excel 1 (chứa ảnh hóa đơn)")
    parser.add_argument("--file2", type=str, default=None, help="Đường dẫn file Excel 2 (chứa đơn giá và KnoxID)")
    parser.add_argument("--month", type=str, default=None, help="Tháng thanh toán YYYY-MM (Mặc định tự đọc từ tên File 1)")
    parser.add_argument("--dry-run", action="store_true", help="Chạy thử nghiệm (Không ghi DB, không upload MinIO)")
    parser.add_argument("--skip-minio", action="store_true", help="Bỏ qua bước upload MinIO")
    parser.add_argument("--skip-db", action="store_true", help="Bỏ qua bước insert Database")
    return parser.parse_args()


def auto_detect_files(input_dir: Path):
    """Tự động tìm 2 file Excel trong một thư mục bất kỳ nếu không truyền đường dẫn trực tiếp."""
    if not input_dir.exists():
        return None, None

    excel_files = [f for f in input_dir.glob("*.xlsx") if not f.name.startswith("~$")] + \
                  [f for f in input_dir.glob("*.xlsm") if not f.name.startswith("~$")]
    if not excel_files:
        return None, None

    file1 = None
    file2 = None

    for f in excel_files:
        name_lower = f.name.lower()
        if any(k in name_lower for k in ["thanh toán hóa đơn", "hóa đơn", "crew", "bill"]):
            file1 = f
        elif any(k in name_lower for k in ["chi tiết", "payment", "boq", "knox", "detail"]):
            file2 = f

    if not file1 and len(excel_files) >= 1:
        file1 = excel_files[0]
    if not file2 and len(excel_files) >= 2:
        file2 = excel_files[1]

    return file1, file2


def main():
    args = parse_args()

    print("=" * 70)
    print("🚀 AI BILL PROCESSOR - TỰ ĐỘNG BÓC TÁCH HÓA ĐƠN AI & ĐỒNG BỘ MINIO / DB")
    print("=" * 70)

    input_dir = BASE_DIR / "input"
    debug_dir = BASE_DIR / "output_debug"
    debug_dir.mkdir(parents=True, exist_ok=True)

    # 1. Xác định file đầu vào (Hỗ trợ CLI params -> .env -> --dir -> auto_detect tại input_dir)
    env_dir = (os.getenv("AI_BILL_PRIVATE_DIR") or os.getenv("AI_BILL_DIR") or "").strip() or None
    env_f1 = (os.getenv("AI_BILL_FILE1") or "").strip() or None
    env_f2 = (os.getenv("AI_BILL_FILE2") or "").strip() or None

    f1_path = Path(args.file1.strip().strip('"').strip("'")) if args.file1 else (Path(env_f1) if env_f1 else None)
    f2_path = Path(args.file2.strip().strip('"').strip("'")) if args.file2 else (Path(env_f2) if env_f2 else None)

    # Nếu người dùng truyền --dir hoặc khai báo AI_BILL_PRIVATE_DIR trong .env
    dir_arg = args.dir.strip().strip('"').strip("'") if args.dir else None
    target_scan_dir = Path(dir_arg) if dir_arg else (Path(env_dir) if env_dir else None)
    if target_scan_dir and target_scan_dir.exists() and target_scan_dir.is_dir():
        d_f1, d_f2 = auto_detect_files(target_scan_dir)
        f1_path = f1_path or d_f1
        f2_path = f2_path or d_f2

    # Nếu vẫn chưa có, fallback tìm trong thư mục cục bộ input/
    if not f1_path or not f2_path:
        d_f1, d_f2 = auto_detect_files(input_dir)
        f1_path = f1_path or d_f1
        f2_path = f2_path or d_f2

    # Hỗ trợ chế độ nhập trực tiếp nếu chưa tìm thấy file
    if not f1_path or not f2_path:
        print("\n💡 CHƯA TÌM THẤY FILE EXCEL - CHẾ ĐỘ NHẬP ĐƯỜNG DẪN TRỰC TIẾP:")
        print("   (Mẹo: Bạn có thể dán đường dẫn hoặc KÉO THẢ thư mục/file từ Windows Explorer vào đây rồi gõ Enter)\n")
        try:
            inp_dir = input("👉 Nhập đường dẫn thư mục Private (hoặc Enter để bỏ qua): ").strip().strip('"').strip("'")
            if inp_dir:
                p_dir = Path(inp_dir)
                if p_dir.exists() and p_dir.is_dir():
                    d_f1, d_f2 = auto_detect_files(p_dir)
                    f1_path = f1_path or d_f1
                    f2_path = f2_path or d_f2

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

    if not f1_path or not f2_path:
        print("\n❌ LỖI: Vui lòng cung cấp đường dẫn tới 2 file Excel đầu vào!")
        print("👉 Cách 1 (Điền thư mục vào file .env):")
        print('   Mở MEP_OneAI_System/ai_bill_processor/.env và điền: AI_BILL_PRIVATE_DIR=D:/Thu_Muc_Private')
        print("👉 Cách 2 (Truyền đường dẫn thư mục khi chạy):")
        print('   python main_sync.py --dir "D:/Thu_Muc_Private" --dry-run')
        print("👉 Cách 3 (Truyền đường dẫn từng file):")
        print('   python main_sync.py --file1 "D:/Path/File1.xlsx" --file2 "D:/Path/File2.xlsx" --dry-run')
        sys.exit(1)

    if not f1_path.exists():
        print(f"❌ LỖI: Không tìm thấy File 1 tại: {f1_path}")
        sys.exit(1)
    if not f2_path.exists():
        print(f"❌ LỖI: Không tìm thấy File 2 tại: {f2_path}")
        sys.exit(1)

    print(f"📁 File 1 (Chứa ảnh hóa đơn): {f1_path}")
    print(f"📁 File 2 (Chứa đơn giá & KnoxID): {f2_path}")

    # 2. Xác định Tháng thanh toán
    month = args.month or extract_month_from_filename(f1_path.name)
    print(f"📅 Tháng thanh toán xác định: {month}")

    # 3. Trích xuất ảnh RAW từ File 1
    print("\n" + "-" * 70)
    print("BƯỚC 1: TRÍCH XUẤT ẢNH HÓA ĐƠN RAW TỪ FILE 1 (OPENXML)")
    print("-" * 70)
    extractor = ExcelImageExtractor(str(f1_path))
    invoices = extractor.extract_invoices()

    if not invoices:
        print("⚠️ Không tìm thấy ảnh hóa đơn nào trong File 1. Kết thúc.")
        sys.exit(0)

    # 4. Đối chiếu và ghép dữ liệu với File 2
    print("\n" + "-" * 70)
    print("BƯỚC 2: ĐỐI SOÁT VỚI FILE 2 (KNOXID, EMAIL & ĐƠN GIÁ)")
    print("-" * 70)
    matcher = ExcelDataMatcher(str(f2_path), month_str=month)
    records = matcher.pair_and_build_records(invoices)

    # 5. Upload ảnh lên MinIO
    print("\n" + "-" * 70)
    print(f"BƯỚC 3: TẢI ẢNH LÊN MINIO S3 (aibill/{month}/...)")
    print("-" * 70)
    minio_uploader = MinioUploader(dry_run=args.dry_run or args.skip_minio)
    for rec in records:
        minio_uploader.upload_image(
            object_key=rec["FILE_PATH"],
            raw_bytes=rec["_raw_bytes"],
            output_debug_dir=debug_dir
        )

    # 6. Đồng bộ vào Oracle Database
    print("\n" + "-" * 70)
    print("BƯỚC 4: LƯU VÀO CƠ SỞ DỮ LIỆU (LOCAL_AI.AI_INVOICE_CHECK)")
    print("-" * 70)
    db_syncer = DatabaseSynchronizer(dry_run=args.dry_run or args.skip_db)
    inserted_count = db_syncer.insert_records(records)
    db_syncer.close()

    # 7. Xuất báo cáo tổng kết ra CSV và JSON để kiểm toán
    print("\n" + "-" * 70)
    print("BƯỚC 5: XUẤT BÁO CÁO KIỂM TOÁN TỔNG KẾT")
    print("-" * 70)

    # Lọc bỏ các trường raw bytes trước khi xuất báo cáo
    clean_records = []
    for r in records:
        clean_r = {k: v for k, v in r.items() if not k.startswith("_")}
        clean_records.append(clean_r)

    csv_path = debug_dir / f"audit_report_{month}.csv"
    json_path = debug_dir / f"audit_report_{month}.json"

    df_out = pd.DataFrame(clean_records)
    df_out.to_csv(csv_path, index=False, encoding="utf-8-sig")

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(clean_records, f, ensure_ascii=False, indent=2)

    print(f"📊 Báo cáo CSV: {csv_path}")
    print(f"📊 Báo cáo JSON: {json_path}")
    print(f"📸 Ảnh preview đã trích xuất: {debug_dir / 'aibill' / month}")

    print("\n" + "=" * 70)
    print(f"🎉 HOÀN THÀNH XỬ LÝ TOÀN BỘ: {len(records)} HÓA ĐƠN THÀNH CÔNG!")
    print("=" * 70)


if __name__ == "__main__":
    main()
