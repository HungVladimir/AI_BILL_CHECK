# ==============================================================================
# MODULE: AI BILL PROCESSOR - EXCEL REPORT GENERATOR
# File: ai_bill_processor/excel_reporter.py
# ==============================================================================
"""
Tự động xuất Báo cáo đối soát Hóa đơn AI ra file Excel chuyên nghiệp (.xlsx)
và lưu trực tiếp vào thư mục Private của người dùng.
Hỗ trợ cả chế độ TEST (Dry-Run) và chế độ SYNC (Đồng bộ thật).
Đồng thời xuất thư mục ảnh đã bóc tách nguyên bản để người dùng mở kiểm tra.
"""
import os
from pathlib import Path
from typing import List, Dict, Any
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


try:
    from .config import DIR_PROJECT_IMAGES, BASE_DIR
except ImportError:
    from config import DIR_PROJECT_IMAGES, BASE_DIR


def export_reconciliation_excel(
    records: List[Dict[str, Any]],
    month_str: str,
    excel_output_dir: Path,
    image_output_dir: Optional[Path] = None,
    is_test_mode: bool = True
) -> Path:
    """
    Xuất báo cáo đối soát ra file Excel và lưu ảnh bóc tách:
      - File Excel báo cáo (.xlsx): Lưu tại excel_output_dir (trong thư mục Private).
      - Ảnh hóa đơn bóc tách (.png): Lưu tại image_output_dir (trong thư mục Dự án / Project vì Private không lưu ảnh).
    """
    # 1. Thư mục lưu ảnh bóc tách: Luôn lưu trong Project folder
    if image_output_dir is None:
        image_output_dir = DIR_PROJECT_IMAGES / month_str
    image_output_dir = Path(image_output_dir)
    image_output_dir.mkdir(parents=True, exist_ok=True)

    # Lưu ảnh ra thư mục dự án
    for r in records:
        raw_b = r.get("_raw_bytes")
        if raw_b:
            safe_name = r.get("SUBMITTER_NAME", "User").replace(" ", "_")
            p_name = r.get("PROVIDER", "AI")
            img_file_name = f"{r['ID']}_{safe_name}_{p_name}.png"
            img_path = image_output_dir / img_file_name
            with open(img_path, "wb") as f:
                f.write(raw_b)
            r["_local_img_path"] = str(img_path.resolve())

    # 2. Thư mục lưu file Excel báo cáo: Thư mục Private (fallback dự phòng nếu private chưa có)
    excel_output_dir = Path(excel_output_dir)
    try:
        excel_output_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        fallback_dir = BASE_DIR / "output_debug"
        fallback_dir.mkdir(parents=True, exist_ok=True)
        excel_output_dir = fallback_dir

    # 3. Khởi tạo Workbook Excel
    wb = openpyxl.Workbook()

    # --- SHEET 1: ĐỊNH DẠNG CHUẨN 22 CỘT DATABASE LOCAL_AI.AI_INVOICE_CHECK ---
    ws_detail = wb.active
    ws_detail.title = "AI_INVOICE_CHECK"
    ws_detail.views.sheetView[0].showGridLines = True

    # Định nghĩa màu sắc & font
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_header_kpi = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    fill_ok = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")      # Xanh nhạt

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    # 22 Cột chuẩn Database + Cột LOCAL_IMAGE_PATH hỗ trợ mở ảnh kiểm tra
    headers = [
        "ID",
        "CREATED_AT",
        "MONTH",
        "SUBMITTER_NAME",
        "SUBMITTER_EMAIL",
        "FILE_NAME",
        "FILE_PATH",
        "MIME",
        "DELETE_TOKEN",
        "PROVIDER",
        "INVOICE_NUMBER",
        "INVOICE_DATE",
        "BILLING_EMAIL",
        "AMOUNT",
        "PLAN",
        "EMAIL_MATCH",
        "WARNINGS",
        "AI_RAW",
        "STATUS",
        "ADMIN_NOTE",
        "REVIEWED_AT",
        "USERID",
        "LOCAL_IMAGE_PATH"
    ]

    for col_idx, h in enumerate(headers, start=1):
        cell = ws_detail.cell(row=1, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws_detail.row_dimensions[1].height = 28

    # Ghi dữ liệu từng dòng từ dòng 2
    for row_idx, r in enumerate(records, start=2):
        amount = r.get("AMOUNT", 0)
        knox = r.get("USERID", "")

        row_data = [
            r.get("ID", ""),
            r.get("CREATED_AT", ""),
            r.get("MONTH", month_str),
            r.get("SUBMITTER_NAME", ""),
            r.get("SUBMITTER_EMAIL", ""),
            r.get("FILE_NAME", ""),
            r.get("FILE_PATH", ""),
            r.get("MIME", "image/png"),
            None,                       # DELETE_TOKEN: null
            r.get("PROVIDER", ""),
            None,                       # INVOICE_NUMBER: null
            None,                       # INVOICE_DATE: null
            r.get("BILLING_EMAIL", ""),
            amount,                     # AMOUNT: giá tiền
            None,                       # PLAN: null
            1,                          # EMAIL_MATCH: 1 cố định
            None,                       # WARNINGS: null
            None,                       # AI_RAW: null
            "confirmed",                # STATUS: confirmed cố định
            None,                       # ADMIN_NOTE: null
            None,                       # REVIEWED_AT: null
            knox,                       # USERID: knoxid user
            r.get("_local_img_path", "")# LOCAL_IMAGE_PATH
        ]

        for col_idx, val in enumerate(row_data, start=1):
            cell = ws_detail.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border

            # Căn lề & định dạng số
            if col_idx == 14:  # AMOUNT
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx == 16: # EMAIL_MATCH
                cell.number_format = '0'
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [1, 2, 3, 8, 10, 19, 22]: # ID, Dates, MIME, Status, KnoxID
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

            # Highlight trạng thái confirmed
            if col_idx == 19:
                cell.fill = fill_ok

    # --- SHEET 2: TỔNG HỢP KPI & CHI PHÍ ---
    ws_kpi = wb.create_sheet(title="Tổng hợp chi phí")
    ws_kpi.views.sheetView[0].showGridLines = True

    font_title = Font(name="Calibri", size=14, bold=True, color="1F497D")
    font_sub = Font(name="Calibri", size=10, italic=True, color="595959")

    ws_kpi["A1"] = f"TỔNG HỢP CHI PHÍ AI CREW - THÁNG {month_str}"
    ws_kpi["A1"].font = font_title
    ws_kpi["A2"] = f"Thời gian lập: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Tổng số hóa đơn: {len(records)}"
    ws_kpi["A2"].font = font_sub

    kpi_headers = ["STT", "Nhà cung cấp", "Số lượng hóa đơn", "Tổng tiền (VNĐ)", "Tỷ trọng"]
    for c_idx, h in enumerate(kpi_headers, start=1):
        cell = ws_kpi.cell(row=4, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_kpi
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_kpi.row_dimensions[4].height = 24

    # Tính toán KPI
    provider_stats = {}
    total_amount = sum(r.get("AMOUNT", 0) for r in records)
    for r in records:
        p = r.get("PROVIDER", "Khác")
        amt = r.get("AMOUNT", 0)
        if p not in provider_stats:
            provider_stats[p] = {"count": 0, "amount": 0}
        provider_stats[p]["count"] += 1
        provider_stats[p]["amount"] += amt

    kpi_row = 5
    for idx, (p_name, stat) in enumerate(provider_stats.items(), start=1):
        amt = stat["amount"]
        ratio = (amt / total_amount) if total_amount > 0 else 0

        ws_kpi.cell(row=kpi_row, column=1, value=idx).alignment = Alignment(horizontal="center")
        ws_kpi.cell(row=kpi_row, column=2, value=p_name).alignment = Alignment(horizontal="left")
        ws_kpi.cell(row=kpi_row, column=3, value=stat["count"]).alignment = Alignment(horizontal="center")

        amt_cell = ws_kpi.cell(row=kpi_row, column=4, value=amt)
        amt_cell.number_format = '#,##0" đ"'
        amt_cell.alignment = Alignment(horizontal="right")

        ratio_cell = ws_kpi.cell(row=kpi_row, column=5, value=ratio)
        ratio_cell.number_format = '0.0%'
        ratio_cell.alignment = Alignment(horizontal="right")

        for c in range(1, 6):
            ws_kpi.cell(row=kpi_row, column=c).border = thin_border
        kpi_row += 1

    # Dòng Tổng cộng
    ws_kpi.cell(row=kpi_row, column=2, value="TỔNG CỘNG").font = Font(name="Calibri", bold=True)
    ws_kpi.cell(row=kpi_row, column=3, value=len(records)).font = Font(name="Calibri", bold=True)
    ws_kpi.cell(row=kpi_row, column=3).alignment = Alignment(horizontal="center")
    total_cell = ws_kpi.cell(row=kpi_row, column=4, value=total_amount)
    total_cell.font = Font(name="Calibri", bold=True)
    total_cell.number_format = '#,##0" đ"'
    total_cell.alignment = Alignment(horizontal="right")

    for c in range(1, 6):
        ws_kpi.cell(row=kpi_row, column=c).border = thin_border

    # Tự động canh độ rộng cột cho cả 2 sheet
    for sheet in [ws_detail, ws_kpi]:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if cell.row in [1, 2]:
                    continue
                max_len = max(max_len, len(val_str))
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Đặt độ rộng cố định hợp lý cho các cột chuẩn Database
    col_widths = {
        "A": 14,  # ID
        "B": 24,  # CREATED_AT
        "C": 12,  # MONTH
        "D": 24,  # SUBMITTER_NAME
        "E": 28,  # SUBMITTER_EMAIL
        "F": 20,  # FILE_NAME
        "G": 32,  # FILE_PATH
        "H": 14,  # MIME
        "I": 14,  # DELETE_TOKEN
        "J": 14,  # PROVIDER
        "K": 16,  # INVOICE_NUMBER
        "L": 14,  # INVOICE_DATE
        "M": 28,  # BILLING_EMAIL
        "N": 16,  # AMOUNT
        "O": 14,  # PLAN
        "P": 14,  # EMAIL_MATCH
        "Q": 14,  # WARNINGS
        "R": 14,  # AI_RAW
        "S": 14,  # STATUS
        "T": 16,  # ADMIN_NOTE
        "U": 16,  # REVIEWED_AT
        "V": 18,  # USERID
        "W": 45   # LOCAL_IMAGE_PATH
    }
    for col_l, w in col_widths.items():
        ws_detail.column_dimensions[col_l].width = w

    # 4. Lưu file báo cáo Excel vào thư mục Private
    tag = "TEST" if is_test_mode else "SYNCED"
    filename = f"Bao_cao_doi_soat_hoa_don_AI_{month_str}_{tag}.xlsx"
    out_file_path = excel_output_dir / filename

    try:
        wb.save(out_file_path)
    except PermissionError:
        # Nếu file đang bị mở trong Excel
        filename_alt = f"Bao_cao_doi_soat_hoa_don_AI_{month_str}_{tag}_{int(datetime.now().timestamp())}.xlsx"
        out_file_path = excel_output_dir / filename_alt
        wb.save(out_file_path)
    except Exception as e:
        # Fallback lưu vào project folder nếu đường dẫn private bị lỗi
        fallback_path = BASE_DIR / "output_debug" / filename
        fallback_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(fallback_path)
        out_file_path = fallback_path

    wb.close()
    return out_file_path

