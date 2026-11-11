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
from excel_reporter import export_reconciliation_excel
from minio_client import MinioUploader
from db_sync import DatabaseSynchronizer

def test_full_pipeline():
    print("=" * 60)
    print("🧪 BẮT ĐẦU KIỂM THỬ TỔNG THỂ AI BILL PROCESSOR")
    print("=" * 60)

    # 1. Test helper functions
    assert extract_month_from_filename("(AI Crew) thanh toán hóa đơn tháng 2026.10.xlsx") == "2026-10"
    assert extract_month_from_filename("(AI Crew) thanh toán hóa đơn tháng 2026.09.xlsx") == "2026-09"
    assert extract_month_from_filename("(AI Crew) thanh toán hóa đơn tháng 8.xlsx") == "2026-08"
    assert extract_month_from_filename("thang 8.2026.xlsx") == "2026-08"
    assert extract_month_from_filename("Aug-26.xlsx") == "2026-08"
    assert extract_month_from_filename("Oct-26.xlsx") == "2026-10"
    assert convert_month_to_excel2_header("2026-10") == "Oct-26"
    assert convert_month_to_excel2_header("2026-09") == "Sep-26"
    assert convert_month_to_excel2_header("2026-08") == "Aug-26"
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
    ws1.title = "Tổng hợp hóa đơn"
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

    # 4. Tạo Mock File 2: 'Chi tiết thanh toán 2026.10.xlsx' gồm 2 sheet chuẩn:
    # Sheet 1: 'Tiêu chuẩn email' (KnoxID, Name, AI_PROVIDER, AI_GEN, AI_TIER, EMAIL_PAYMENT)
    # Sheet 2: 'Thanh toán' (KnoxID, Name, Đăng nhập, Thanh toán, và các cột tháng Aug-26, Sep-26, Oct-26...)
    wb2 = openpyxl.Workbook()

    # Tạo Sheet 1: Tiêu chuẩn email
    ws_email_std = wb2.active
    ws_email_std.title = "Tiêu chuẩn email"
    ws_email_std.append(["KnoxID", "Name", "AI_PROVIDER", "AI_GEN", "AI_TIER", "EMAIL_PAYMENT"])
    ws_email_std.append(["kimngan.nguyen", "Nguyễn Thị Kim Ngân", "OpenAI", "ChatGPT", "Pro", "nguyenthikimngan269@gmail.com"])
    ws_email_std.append(["thaonga.chu", "Chu Thị Thảo Nga", "OpenAI", "ChatGPT", "Pro", "themeteor_xt@yahoo.com.vn"])
    ws_email_std.append(["dong.nh", "Nguyễn Hữu Đồng", "OpenAI", "ChatGPT", "Plus", "dongnguyen30794@gmail.com"])
    ws_email_std.append(["dong.nh", "Nguyễn Hữu Đồng", "Anthropic", "Claude", "Pro", "dongnguyen30794@gmail.com"])

    # Tạo Sheet 2: Thanh toán (Mô phỏng 100% cấu trúc thực tế từ ảnh media_1794371395755_944ee58d.jpg)
    ws2 = wb2.create_sheet(title="Thanh toán")
    # Row 1: Header tháng
    ws2.append(["No.", "KnoxID", "Name", "Gemini Pro", "", "Aug-26", "", "", "Sep-26", "", "", "Oct-26", "", "", ""])
    # Row 2: Sub-header
    ws2.append(["", "", "", "Đăng nhập", "Thanh toán", "ChatGPT Plus", "Claude Pro", "Total", "ChatGPT Plus", "Claude Pro", "Total", "ChatGPT Plus", "Claude Pro", "Gemini Pro", "Total"])
    # Data rows
    ws2.append([1, "kimngan.nguyen", "Nguyễn Thị Kim Ngân", "", "", 499000, 0, 499000, 499000, 0, 499000, 499000, 0, 0, 499000])
    ws2.append([2, "thaonga.chu", "Chu Thị Thảo Nga", "", "", 0, 0, 0, 522500, 0, 522500, 522500, 0, 0, 522500])
    ws2.append([3, "dong.nh", "Nguyễn Hữu Đồng", "dongnguyen30794", "dongnguyen30794", 527778, 554444, 1082222, 527778, 554444, 1082222, 527778, 554444, 0, 1082222])

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
    for inv in invs:
        assert inv["plan"] is None, "Plan trích xuất phải là None (null)"
    print("✅ 2. Trích xuất ảnh OpenXML & Kiểm tra PLAN null: ĐẠT")

    matcher = ExcelDataMatcher(str(test_f2), month_str="2026-10")
    records = matcher.pair_and_build_records(invs)
    print(f"📊 Số lượng bản ghi ghép nối: {len(records)}")
    assert len(records) == 3, f"Kỳ vọng 3 bản ghi, thực tế {len(records)}"
    print("✅ 3. Ghép nối dữ liệu từ Sheet 'Tiêu chuẩn email' & Sheet 'Thanh toán': ĐẠT")

    # Kiểm tra bản ghi 1 (Kim Ngân):
    r1 = [r for r in records if r["SUBMITTER_NAME"] == "Nguyễn Thị Kim Ngân"][0]
    assert r1["USERID"] == "kimngan.nguyen"
    assert r1["SUBMITTER_EMAIL"] == "nguyenthikimngan269@gmail.com"
    assert r1["BILLING_EMAIL"] == "nguyenthikimngan269@gmail.com"
    assert r1["AMOUNT"] == 499000
    assert r1["PROVIDER"] == "ChatGPT"
    assert r1["PLAN"] is None, "PLAN phải là null"
    assert r1["STATUS"] == "confirmed", "STATUS phải là confirmed"
    assert r1["EMAIL_MATCH"] == 1, "EMAIL_MATCH phải là 1"
    assert r1["DELETE_TOKEN"] is None
    assert r1["INVOICE_NUMBER"] is None
    assert r1["INVOICE_DATE"] is None
    assert len(r1["ID"]) == 10
    assert r1["FILE_PATH"] == f"aibill/2026-10/{r1['ID']}.png"
    print("✅ 4. Kiểm tra bản ghi đơn lẻ (Kim Ngân) với 22 trường chuẩn DB: ĐẠT")

    # Kiểm tra bản ghi Hữu Đồng (2 bản ghi tách biệt!):
    hd_recs = [r for r in records if r["SUBMITTER_NAME"] == "Nguyễn Hữu Đồng"]
    assert len(hd_recs) == 2, "Nguyễn Hữu Đồng phải có đúng 2 bản ghi riêng!"
    hd_gpt = [r for r in hd_recs if r["PROVIDER"] == "ChatGPT"][0]
    hd_cld = [r for r in hd_recs if r["PROVIDER"] == "Claude"][0]
    assert hd_gpt["USERID"] == "dong.nh"
    assert hd_gpt["SUBMITTER_EMAIL"] == "dongnguyen30794@gmail.com"
    assert hd_gpt["AMOUNT"] == 527778
    assert hd_gpt["STATUS"] == "confirmed"
    assert hd_gpt["EMAIL_MATCH"] == 1
    assert hd_cld["AMOUNT"] == 554444
    assert hd_cld["STATUS"] == "confirmed"
    assert hd_gpt["ID"] != hd_cld["ID"], "Hai hóa đơn phải có ID khác nhau!"
    print("✅ 5. Kiểm tra người có 2 hóa đơn tách 2 dòng độc lập (Hữu Đồng): ĐẠT")

    # 5.1 Kiểm tra đa tháng (Aug-26, Sep-26) và khớp tên không dấu / Unicode NFD
    matcher_aug = ExcelDataMatcher(str(test_f2), month_str="2026-08")
    recs_aug = matcher_aug.pair_and_build_records(invs)
    r_aug_kn = [r for r in recs_aug if "Kim Ngân" in r["SUBMITTER_NAME"]][0]
    assert r_aug_kn["AMOUNT"] == 499000, f"Kỳ vọng 499000, thực tế {r_aug_kn['AMOUNT']}"

    matcher_sep = ExcelDataMatcher(str(test_f2), month_str="2026-09")
    recs_sep = matcher_sep.pair_and_build_records(invs)
    r_sep_kn = [r for r in recs_sep if "Kim Ngân" in r["SUBMITTER_NAME"]][0]
    assert r_sep_kn["AMOUNT"] == 499000, f"Kỳ vọng 499000, thực tế {r_sep_kn['AMOUNT']}"

    # Kiểm tra so khớp tên không dấu (e.g. 'Nguyen Thi Kim Ngan' -> 'Nguyễn Thị Kim Ngân')
    inv_unacc = [{
        "row_idx": 2,
        "submitter_name": "Nguyen Thi Kim Ngan",
        "provider_key": "chatgpt",
        "provider_name": "ChatGPT",
        "plan": None,
        "file_name": "gpt_raw.png",
        "raw_bytes": b"fake_bytes_for_test",
        "width": 100,
        "height": 50,
        "invoice_date": None,
        "invoice_number": None
    }]
    recs_unacc = matcher.pair_and_build_records(inv_unacc)
    assert len(recs_unacc) == 1
    assert recs_unacc[0]["USERID"] == "kimngan.nguyen"
    assert recs_unacc[0]["AMOUNT"] == 499000
    print("✅ 5.1 Kiểm tra đối soát Đa tháng & So khớp tên không dấu: ĐẠT")

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

    # 8. Test Export Báo cáo Excel ra Private & Ảnh ra Project
    mock_private_reports = test_dir / "mock_private_reports"
    mock_project_images = test_dir / "mock_project_images"
    rep_path = export_reconciliation_excel(
        records=records,
        month_str="2026-10",
        excel_output_dir=mock_private_reports,
        image_output_dir=mock_project_images,
        is_test_mode=True
    )
    assert rep_path.exists(), "File báo cáo Excel phải được tạo thành công!"
    assert any(mock_project_images.glob("*.png")), "Ảnh phải được lưu vào thư mục dự án!"

    # Kiểm tra cấu trúc Sheet 1 của file Excel báo cáo:
    wb_out = openpyxl.load_workbook(rep_path)
    assert "AI_INVOICE_CHECK" in wb_out.sheetnames, "Sheet 1 phải là 'AI_INVOICE_CHECK'"
    ws_out1 = wb_out["AI_INVOICE_CHECK"]

    expected_headers = [
        "ID", "CREATED_AT", "MONTH", "SUBMITTER_NAME", "SUBMITTER_EMAIL",
        "FILE_NAME", "FILE_PATH", "MIME", "DELETE_TOKEN", "PROVIDER",
        "INVOICE_NUMBER", "INVOICE_DATE", "BILLING_EMAIL", "AMOUNT", "PLAN",
        "EMAIL_MATCH", "WARNINGS", "AI_RAW", "STATUS", "ADMIN_NOTE",
        "REVIEWED_AT", "USERID", "LOCAL_IMAGE_PATH"
    ]
    actual_headers = [ws_out1.cell(1, c).value for c in range(1, len(expected_headers) + 1)]
    assert actual_headers == expected_headers, f"Tiêu đề cột không khớp:\nKỳ vọng: {expected_headers}\nThực tế: {actual_headers}"

    # Kiểm tra dòng dữ liệu 1 (Dòng 2 trong file)
    row2_status = ws_out1.cell(2, 19).value  # Cột STATUS
    row2_email_match = ws_out1.cell(2, 16).value  # Cột EMAIL_MATCH
    row2_plan = ws_out1.cell(2, 15).value  # Cột PLAN
    row2_del_tok = ws_out1.cell(2, 9).value  # Cột DELETE_TOKEN

    assert row2_status == "confirmed", f"STATUS phải là 'confirmed', thực tế: {row2_status}"
    assert row2_email_match == 1, f"EMAIL_MATCH phải là 1, thực tế: {row2_email_match}"
    assert row2_plan is None, f"PLAN phải là None, thực tế: {row2_plan}"
    assert row2_del_tok is None, f"DELETE_TOKEN phải là None, thực tế: {row2_del_tok}"
    wb_out.close()

    print("✅ 8. Kiểm tra File Excel Báo cáo với 22 Cột chuẩn Database: ĐẠT")

    print("\n" + "=" * 60)
    print("🎉 TOÀN BỘ 8 BƯỚC KIỂM THỬ ĐÃ THÀNH CÔNG 100%!")
    print("=" * 60)

if __name__ == "__main__":
    test_full_pipeline()

