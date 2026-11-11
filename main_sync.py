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
        locate_ai_bill_files,
        locate_batch_ai_bill_files
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
        locate_ai_bill_files,
        locate_batch_ai_bill_files
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
    parser.add_argument("--folder1", "--dir1", "--invoices-dir", type=str, default=None, help="Folder 1: chứa các file Excel 1 của từng tháng (Mặc định: invoices_file1/)")
    parser.add_argument("--folder2", "--dir2", "--members-dir", type=str, default=None, help="Folder 2: chứa file Excel 2 tổng hợp giá tiền các tháng (Mặc định: members_file2/)")
    parser.add_argument("--file1", type=str, default=None, help="Chạy đơn lẻ: Đường dẫn file Excel 1 cụ thể")
    parser.add_argument("--file2", type=str, default=None, help="Đường dẫn file Excel 2 cụ thể")
    parser.add_argument("--month", type=str, default=None, help="Lọc tháng thanh toán YYYY-MM cụ thể (Ví dụ: 2026-08)")
    parser.add_argument("--clear-month", action="store_true", help="Xóa dữ liệu cũ của các tháng đó trong DB trước khi nạp mới (tránh trùng lặp khi chạy lại sync)")
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

    print("=" * 80)
    if is_test_mode:
        print("🧪 AI BILL PROCESSOR - CHẾ ĐỘ TEST ĐỐI SOÁT (AN TOÀN - KHÔNG GHI DB/MINIO)")
        print("   (File báo cáo đối soát Excel tổng hợp và ảnh bóc tách sẽ được lưu vào Private & Project)")
    else:
        print("🚀 AI BILL PROCESSOR - CHẾ ĐỘ ĐỒNG BỘ THẬT (UPLOAD MINIO & LƯU ORACLE DB)")
    print("=" * 80)

    # 1. Xác định thư mục Private & Dò tìm File 1 (Folder 1), File 2 (Folder 2)
    target_private_dir = Path(args.dir.strip().strip('"').strip("'")) if args.dir else AI_BILL_PRIVATE_DIR
    print(f"📁 Thư mục Private cấu hình: {target_private_dir}")

    folder1_path = Path(args.folder1.strip().strip('"').strip("'")) if args.folder1 else (target_private_dir / "invoices_file1")
    folder2_path = Path(args.folder2.strip().strip('"').strip("'")) if args.folder2 else (target_private_dir / "members_file2")

    f2_path = Path(args.file2.strip().strip('"').strip("'")) if args.file2 else None
    files_to_process = []

    if args.file1:
        # Chế độ chạy file đơn lẻ theo tham số --file1
        f1_single = Path(args.file1.strip().strip('"').strip("'"))
        m_single = args.month or extract_month_from_filename(f1_single.name)
        files_to_process.append((f1_single, m_single))
        if not f2_path:
            _, auto_f2, excel_out_dir, _ = locate_ai_bill_files(target_private_dir, target_month=m_single)
            f2_path = auto_f2
        else:
            excel_out_dir = target_private_dir / "output_reports"
    else:
        # Chế độ quét theo Folder 1 & Folder 2
        batch_f1_list, auto_f2, excel_out_dir, _ = locate_batch_ai_bill_files(
            private_dir=target_private_dir,
            folder1=folder1_path,
            folder2=folder2_path,
            target_month=args.month
        )
        files_to_process = batch_f1_list
        if not f2_path:
            f2_path = auto_f2

    # Fallback kiểm tra trong thư mục cục bộ input/ nếu Private chưa có file
    if not files_to_process or not f2_path:
        local_input = BASE_DIR / "input"
        local_f1_list, loc_f2, _, _ = locate_batch_ai_bill_files(
            private_dir=local_input,
            folder1=local_input / "invoices_file1",
            folder2=local_input / "members_file2",
            target_month=args.month
        )
        if not files_to_process:
            files_to_process = local_f1_list
        if not f2_path:
            f2_path = loc_f2

    # Chế độ nhập đường dẫn nếu hoàn toàn chưa thấy file
    if not files_to_process or not f2_path:
        print("\n💡 CHƯA TÌM THẤY ĐỦ FILE TRONG CÁC THƯ MỤC CẤU HÌNH:")
        print(f"   📂 Folder 1 (Chứa file Excel 1 từng tháng): {folder1_path}")
        print(f"   📂 Folder 2 (Chứa file Excel 2 tổng hợp):    {folder2_path}")
        print("   (Mẹo: Bạn có thể sao chép các file vào 2 folder trên, hoặc nhập đường dẫn trực tiếp)\n")
        try:
            if not files_to_process:
                inp_f1 = input("👉 Nhập đường dẫn File hoặc Folder 1: ").strip().strip('"').strip("'")
                if inp_f1:
                    p1 = Path(inp_f1)
                    if p1.is_dir():
                        files_to_process, _, _, _ = locate_batch_ai_bill_files(folder1=p1, target_month=args.month)
                    elif p1.exists():
                        files_to_process = [(p1, args.month or extract_month_from_filename(p1.name))]

            if not f2_path:
                inp_f2 = input("👉 Nhập đường dẫn File 2 (chứa đơn giá & KnoxID): ").strip().strip('"').strip("'")
                if inp_f2:
                    p2 = Path(inp_f2)
                    if p2.is_dir():
                        _, f2_path, _, _ = locate_batch_ai_bill_files(folder2=p2)
                    elif p2.exists():
                        f2_path = p2
        except (KeyboardInterrupt, EOFError):
            print("\nĐã hủy thao tác.")
            sys.exit(0)

    if not files_to_process:
        print(f"\n❌ LỖI: Không tìm thấy bất kỳ file Excel 1 nào trong: {folder1_path}")
        print("👉 Vui lòng đặt các file Excel 1 vào Folder 1 (invoices_file1/) hoặc truyền tham số --folder1 / --file1")
        sys.exit(1)

    if not f2_path or not f2_path.exists():
        print(f"\n❌ LỖI: Không tìm thấy File 2 tổng hợp tại: {f2_path or folder2_path}")
        print("👉 Vui lòng đặt File 2 vào Folder 2 (members_file2/) hoặc truyền tham số --file2")
        sys.exit(1)

    print(f"\n📋 THÔNG TIN ĐẦU VÀO ĐƯỢC XÁC ĐỊNH:")
    print(f"   - Folder 1 (Hóa đơn ảnh): {folder1_path.resolve() if folder1_path.exists() else 'Đường dẫn chỉ định'}")
    print(f"     👉 Đã phát hiện {len(files_to_process)} file hóa đơn:")
    for idx, (f_path, m_str) in enumerate(files_to_process, start=1):
        print(f"        {idx}. {f_path.name} -> Tháng đối soát: {m_str}")
    print(f"   - Folder 2 / File tổng hợp: {f2_path.name}")

    # Đảm bảo các thư mục đầu ra tồn tại
    excel_out_dir.mkdir(parents=True, exist_ok=True)
    DIR_PROJECT_IMAGES.mkdir(parents=True, exist_ok=True)

    # 2. Khởi tạo Matcher và nạp File 2 một lần duy nhất
    print("\n" + "=" * 80)
    print("BƯỚC 1: NẠP DỮ LIỆU ĐƠN GIÁ & TIÊU CHUẨN EMAIL TỪ FILE 2 TỔNG HỢP")
    print("=" * 80)
    matcher = ExcelDataMatcher(str(f2_path))
    matcher.load_users_data()

    # 3. Duyệt qua từng file Excel 1 và đối soát theo tháng
    all_records = []
    month_summary = {}

    for idx, (f1_path, target_month) in enumerate(files_to_process, start=1):
        print("\n" + "-" * 80)
        print(f"📂 [{idx}/{len(files_to_process)}] XỬ LÝ FILE 1: {f1_path.name}")
        print(f"   📅 Tháng thanh toán: {target_month}")
        print("-" * 80)

        # 3.1 Trích xuất ảnh raw từ File 1
        extractor = ExcelImageExtractor(str(f1_path))
        invoices = extractor.extract_invoices()

        if not invoices:
            print(f"⚠️ [Cảnh báo] File '{f1_path.name}' không có ảnh hóa đơn nào hợp lệ. Bỏ qua.")
            continue

        # 3.2 Đối soát dữ liệu với tháng target_month trong File 2
        month_records = matcher.pair_and_build_records(invoices, month_override=target_month)

        # Gán kích thước ảnh
        for rec in month_records:
            for inv in invoices:
                if rec["SUBMITTER_NAME"] == inv["submitter_name"] and rec["PROVIDER"] == inv["provider_name"]:
                    rec["_width"] = inv.get("width", 0)
                    rec["_height"] = inv.get("height", 0)
                    break

        # 3.3 Lưu ảnh bóc tách vào project theo từng tháng
        img_month_dir = DIR_PROJECT_IMAGES / target_month
        img_month_dir.mkdir(parents=True, exist_ok=True)
        for rec in month_records:
            raw_b = rec.get("_raw_bytes")
            if raw_b:
                safe_name = rec.get("SUBMITTER_NAME", "User").replace(" ", "_")
                p_name = rec.get("PROVIDER", "AI")
                img_file_name = f"{rec['ID']}_{safe_name}_{p_name}.png"
                img_path = img_month_dir / img_file_name
                try:
                    if not img_path.exists() or img_path.stat().st_size == 0:
                        with open(img_path, "wb") as f_img:
                            f_img.write(raw_b)
                except Exception:
                    pass
                rec["_local_img_path"] = str(img_path.resolve())

        # 3.4 Nếu chế độ SYNC thật -> tải ảnh lên MinIO
        if not is_test_mode and not args.skip_minio:
            minio_uploader = MinioUploader(dry_run=False)
            for rec in month_records:
                minio_uploader.upload_image(
                    object_key=rec["FILE_PATH"],
                    raw_bytes=rec["_raw_bytes"],
                    output_debug_dir=img_month_dir
                )

        all_records.extend(month_records)
        m_amt = sum(r.get("AMOUNT", 0) for r in month_records)
        month_summary[target_month] = {
            "file": f1_path.name,
            "count": len(month_records),
            "amount": m_amt
        }
        print(f"   ✅ Đã bóc tách {len(month_records)} hóa đơn cho tháng {target_month} | Tổng tiền: {m_amt:,.0f} VNĐ")

    if not all_records:
        print("\n⚠️ Không có bản ghi nào được bóc tách từ các file đã quét. Kết thúc.")
        sys.exit(0)

    # 4. Ghi cơ sở dữ liệu nếu ở chế độ SYNC thật
    if not is_test_mode:
        print("\n" + "=" * 80)
        print(f"BƯỚC 4: LƯU TOÀN BỘ BẢN GHI VÀO ORACLE DATABASE (Bảng: {TABLE_NAME})")
        print("=" * 80)
        db_syncer = DatabaseSynchronizer(dry_run=args.skip_db)
        if args.clear_month:
            for m_key in month_summary.keys():
                db_syncer.clear_month_data(m_key)
        inserted_count = db_syncer.insert_records(
            records=all_records,
            clear_existing_month=False
        )
        db_syncer.close()
        print(f"🎉 Đã lưu thành công {inserted_count} bản ghi vào Database.")

    # 5. Xuất 1 file Báo cáo đối soát tổng hợp duy nhất
    print("\n" + "=" * 80)
    print("BƯỚC CUỐI: XUẤT 1 FILE BÁO CÁO ĐỐI SOÁT TỔNG HỢP DUY NHẤT (EXCEL)")
    print("=" * 80)
    all_months_sorted = sorted(list(month_summary.keys()))
    if len(all_months_sorted) > 1:
        rep_month_label = f"{all_months_sorted[0]}_den_{all_months_sorted[-1]}"
    elif all_months_sorted:
        rep_month_label = all_months_sorted[0]
    else:
        rep_month_label = "all"

    excel_report_path = export_reconciliation_excel(
        records=all_records,
        month_str=rep_month_label,
        excel_output_dir=excel_out_dir,
        is_test_mode=is_test_mode
    )

    print(f"📊 [Báo cáo Tổng hợp Excel] Đã tạo thành công trong thư mục Private:")
    print(f"   👉 {excel_report_path}")
    print(f"📸 [Thư mục ảnh] Toàn bộ ảnh hóa đơn các tháng đã lưu trong:")
    print(f"   👉 {DIR_PROJECT_IMAGES}")

    # 6. In bảng tổng kết toàn diện
    print("\n" + "=" * 80)
    print("📈 TỔNG KẾT ĐỐI SOÁT THEO TỪNG THÁNG:")
    print(f"{'Tháng':<12} | {'Tên File Excel 1':<45} | {'Số Hóa đơn':<12} | {'Tổng tiền':<18}")
    print("-" * 95)
    grand_total = 0
    grand_count = 0
    for m_k, s in sorted(month_summary.items()):
        amt_val = s["amount"]
        print(f"{m_k:<12} | {s['file']:<45} | {s['count']:<12} | {amt_val:>14,.0f} VNĐ")
        grand_total += amt_val
        grand_count += s["count"]
    print("-" * 95)
    print(f"{'TỔNG CỘNG':<12} | {f'{len(month_summary)} files':<45} | {grand_count:<12} | {grand_total:>14,.0f} VNĐ")
    print("=" * 80)

    if is_test_mode:
        print(f"\n🎉 HOÀN THÀNH KIỂM THỬ: {len(all_records)} HÓA ĐƠN ({len(month_summary)} THÁNG) ĐÃ ĐƯỢC TỔNG HỢP VÀO 1 FILE BÁO CÁO!")
        print("💡 Hãy mở file Excel báo cáo ở trên trong thư mục Private để kiểm tra đối soát.")
        print("👉 Để đồng bộ THẬT lên MinIO & Database Oracle, chạy:")
        print("   python main_sync.py --sync")
    else:
        print(f"\n🎉 ĐÃ ĐỒNG BỘ THÀNH CÔNG: {len(all_records)} HÓA ĐƠN VÀO MINIO & DATABASE!")


if __name__ == "__main__":
    main()
