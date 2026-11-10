# ==============================================================================
# TEST SCRIPT: AI BILL PROCESSOR PIPELINE VERIFICATION
# ==============================================================================
import sys
import os
import io
from pathlib import Path
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image, ImageDraw

sys.path.insert(0, str(Path(__file__).parent))

from config import extract_month_from_filename, convert_month_to_excel2_header
from excel_extractor import ExcelImageExtractor
from excel_matcher import ExcelDataMatcher
from minio_client import MinioUploader
from db_sync import DatabaseSynchronizer

def test_full_pipeline():
    print("=" * 60)
    print("🧪 BẮT ĐẦU KIỂM THỬ TỔNG THỂ AI BILL PROCESSOR")
    print("=" * 60)

    # 1. Test helper functions
    assert extract_month_from_filename("(AI Crew) thanh toán hóa đơn tháng 2026.10.xlsx") == "2026-10"
    assert extract_month_from_filename("(AI Crew) thanh toán hóa đơn tháng 2026.09.xlsx") == "2026-09"
    assert convert_month_to_excel2_header("2026-10") == "Oct-26"
    assert convert_month_to_excel2_header("2026-09") == "Sep-26"
    print("✅ 1. Kiểm tra trích xuất Tháng & Header Conversion: ĐẠT")

    # 2. Tạo ảnh mẫu nhỏ để nhúng vào Excel
    def make_dummy_img(text, color):
        im = Image.new("RGB", (120, 60), color=color)
        d = ImageDraw.Draw(im)
        d.text((10, 20), text, fill=(255, 255, 255))
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        buf.seek(0)
        return buf

    # 3. Tạo Mock File 1: '(AI Crew) thanh toán hóa đơn tháng 2026.10.xlsx'
    wb1 = openpyxl.Workbook()
    ws1 = wb1.active
    ws1.title = "Sheet1"
    ws1.append(["STT", "Họ và Tên", "GenID", "Team", "Group", "Part", "Hóa đơn ChatGPT", "Hóa đơn Claude", "Số tiền thanh toán", "Tiêu chuẩn thanh toán"])
    ws1.append([1, "Nguyễn Thị Kim Ngân", "24420206", "DTM Team", "DTM Group", "", "", "", 499000, 1200000])
    ws1.append([2, "Chu Thị Thảo Nga", "16425218", "DTM Team", "DTM Group", "", "", "", 0, 1200000])
    ws1.append([3, "Nguyễn Hữu Đồng", "17439982", "EHS Team", "EHS Group", "", "", "", 1082222, 1200000])

    # Nhúng ảnh vào Cell:
    # Row 2 (Kim Ngân), Col 7 (Hóa đơn ChatGPT -> G2)
    img_kn_gpt = OpenpyxlImage(make_dummy_img("GPT_KimNgan", (34, 139, 34)))
    ws1.add_image(img_kn_gpt, "G2")

    # Row 4 (Hữu Đồng), Col 7 (Hóa đơn ChatGPT -> G4)
    img_hd_gpt = OpenpyxlImage(make_dummy_img("GPT_HuuDong", (34, 139, 34)))
    ws1.add_image(img_hd_gpt, "G4")

    # Row 4 (Hữu Đồng), Col 8 (Hóa đơn Claude -> H4)
    img_hd_cld = OpenpyxlImage(make_dummy_img("CLD_HuuDong", (210, 105, 30)))
    ws1.add_image(img_hd_cld, "H4")

    test_dir = Path(__file__).parent / "test_data"
    test_dir.mkdir(parents=True, exist_ok=True)

    test_f1 = test_dir / "(AI Crew) thanh toán hóa đơn tháng 2026.10.xlsx"
    try:
        wb1.save(test_f1)
    except PermissionError:
        print(f"\n⚠️ [CẢNH BÁO KHÓA FILE] File '{test_f1.name}' đang được mở trong Excel!")
        print("👉 Hệ thống tự động chuyển sang file kiểm thử dự phòng để tiếp tục...")
        test_f1 = test_dir / f"(AI Crew) thanh toán hóa đơn tháng 2026.10_test_tmp.xlsx"
        wb1.save(test_f1)
    wb1.close()

    # 4. Tạo Mock File 2: 'Chi tiết thanh toán 2026.10.xlsx'
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = "Thanh toán"
    # Row 1: Header tháng
    ws2.append(["No.", "KnoxID", "Name", "Gemini Pro", "", "Oct-26", "", "", ""])
    # Row 2: Sub-header
    ws2.append(["", "", "", "Đăng nhập", "Thanh toán", "ChatGPT Plus", "Claude Pro", "Gemini Pro", "Total"])
    # Data rows
    ws2.append([1, "kimngan.nguyen", "Nguyễn Thị Kim Ngân", "kimngan", "nguyenthikimngan269", 499000, 0, 0, 499000])
    ws2.append([2, "thaonga.chu", "Chu Thị Thảo Nga", "thaonga", "themeteor_xt@yahoo.com.vn", 0, 0, 0, 0])
    ws2.append([3, "dong.nh", "Nguyễn Hữu Đồng", "dong.nh", "dongnguyen30794", 527778, 554444, 0, 1082222])

    test_f2 = test_dir / "Chi tiết thanh toán 2026.10.xlsx"
    try:
        wb2.save(test_f2)
    except PermissionError:
        print(f"\n⚠️ [CẢNH BÁO KHÓA FILE] File '{test_f2.name}' đang được mở trong Excel!")
        print("👉 Hệ thống tự động chuyển sang file kiểm thử dự phòng để tiếp tục...")
        test_f2 = test_dir / f"Chi tiết thanh toán 2026.10_test_tmp.xlsx"
        wb2.save(test_f2)
    wb2.close()

    # 5. Test trích xuất và đối soát
    ext = ExcelImageExtractor(str(test_f1))
    invs = ext.extract_invoices()
    print(f"📸 Số lượng ảnh trích xuất từ File 1: {len(invs)}")
    assert len(invs) == 3, f"Kỳ vọng 3 ảnh, thực tế {len(invs)}"
    print("✅ 2. Trích xuất ảnh OpenXML: ĐẠT")

    matcher = ExcelDataMatcher(str(test_f2), month_str="2026-10")
    records = matcher.pair_and_build_records(invs)
    print(f"📊 Số lượng bản ghi ghép nối: {len(records)}")
    assert len(records) == 3, f"Kỳ vọng 3 bản ghi, thực tế {len(records)}"
    print("✅ 3. Ghép nối dữ liệu & Tính tiền: ĐẠT")

    # Kiểm tra bản ghi 1 (Kim Ngân):
    r1 = [r for r in records if r["SUBMITTER_NAME"] == "Nguyễn Thị Kim Ngân"][0]
    assert r1["USERID"] == "kimngan.nguyen"
    assert r1["SUBMITTER_EMAIL"] == "nguyenthikimngan269@gmail.com"
    assert r1["AMOUNT"] == 499000
    assert r1["PROVIDER"] == "ChatGPT"
    assert len(r1["ID"]) == 10
    assert r1["FILE_PATH"] == f"aibill/2026-10/{r1['ID']}.png"
    print("✅ 4. Kiểm tra bản ghi đơn lẻ (Kim Ngân): ĐẠT")

    # Kiểm tra bản ghi Hữu Đồng (2 bản ghi tách biệt!):
    hd_recs = [r for r in records if r["SUBMITTER_NAME"] == "Nguyễn Hữu Đồng"]
    assert len(hd_recs) == 2, "Nguyễn Hữu Đồng phải có đúng 2 bản ghi riêng!"
    hd_gpt = [r for r in hd_recs if r["PROVIDER"] == "ChatGPT"][0]
    hd_cld = [r for r in hd_recs if r["PROVIDER"] == "Claude"][0]
    assert hd_gpt["USERID"] == "dong.nh"
    assert hd_gpt["SUBMITTER_EMAIL"] == "dongnguyen30794@gmail.com"
    assert hd_gpt["AMOUNT"] == 527778
    assert hd_cld["AMOUNT"] == 554444
    assert hd_gpt["ID"] != hd_cld["ID"], "Hai hóa đơn phải có ID khác nhau!"
    print("✅ 5. Kiểm tra người có 2 hóa đơn tách 2 dòng độc lập (Hữu Đồng): ĐẠT")

    # 6. Test Dry-run MinIO & DB
    uploader = MinioUploader(dry_run=True)
    debug_out = test_dir / "debug"
    for r in records:
        uploader.upload_image(r["FILE_PATH"], r["_raw_bytes"], output_debug_dir=debug_out)
    print("✅ 6. Kiểm tra MinIO Uploader (Dry-Run): ĐẠT")

    db_sync = DatabaseSynchronizer(dry_run=True)
    ins_cnt = db_sync.insert_records(records)
    assert ins_cnt == 3
    print("✅ 7. Kiểm tra Database Synchronizer (Dry-Run): ĐẠT")

    print("\n" + "=" * 60)
    print("🎉 TOÀN BỘ 7 BƯỚC KIỂM THỬ ĐÃ THÀNH CÔNG 100%!")
    print("=" * 60)

if __name__ == "__main__":
    test_full_pipeline()

