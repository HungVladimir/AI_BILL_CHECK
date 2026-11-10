# ==============================================================================
# MODULE: AI BILL PROCESSOR - EXCEL DATA MATCHER
# File: ai_bill_processor/excel_matcher.py
# ==============================================================================
"""
Đọc File Excel 2 (Sheet 'Thanh toán' hoặc Active sheet):
  1. Trích xuất KnoxID (USERID) và Họ và Tên (Name).
  2. Bóc tách email từ cột 'Thanh toán' (chuẩn hóa @gmail.com hoặc giữ @yahoo.com.vn).
  3. Bóc tách đơn giá từng dịch vụ AI (ChatGPT Plus, Claude Pro, Gemini Pro) theo tháng yêu cầu.
  4. Ghép nối (Join) 1-1 với ảnh hóa đơn từ File 1 để tạo bản ghi cơ sở dữ liệu hoàn chỉnh.
"""
import re
from pathlib import Path
from typing import Dict, List, Any, Optional
import openpyxl
import pandas as pd

try:
    from .config import (
        PROVIDERS_CONFIG,
        extract_month_from_filename,
        convert_month_to_excel2_header
    )
    from .id_generator import generate_invoice_id
except ImportError:
    from config import (
        PROVIDERS_CONFIG,
        extract_month_from_filename,
        convert_month_to_excel2_header
    )
    from id_generator import generate_invoice_id


def normalize_name(name: Any) -> str:
    """Chuẩn hóa họ tên để ghép nối (bỏ khoảng trắng thừa, viết thường)."""
    if not name:
        return ""
    s = str(name).strip().lower()
    return re.sub(r'\s+', ' ', s)


def clean_email(raw_val: Any) -> str:
    """
    Quy tắc chuẩn hóa email:
      - Nếu đã có đuôi @yahoo.com.vn hoặc @gmail.com -> Giữ nguyên.
      - Nếu là username chưa có đuôi domain -> Tự động thêm @gmail.com.
      Ví dụ: 'nguyenthikimngan269' -> 'nguyenthikimngan269@gmail.com'
             'themeteor_xt@yahoo.com.vn' -> 'themeteor_xt@yahoo.com.vn'
    """
    if not raw_val or pd.isna(raw_val):
        return ""
    s = str(raw_val).strip()
    if not s or s.lower() in ["nan", "none", "-", "0"]:
        return ""

    s_lower = s.lower()
    if "@yahoo.com.vn" in s_lower or "@gmail.com" in s_lower:
        return s_lower
    if "@" in s_lower:
        return s_lower

    return f"{s_lower}@gmail.com"


class ExcelDataMatcher:
    """Bộ ghép nối đối soát dữ liệu giữa File 1 và File 2."""

    def __init__(self, excel2_path: str, month_str: Optional[str] = None):
        self.excel2_path = Path(excel2_path)
        if not self.excel2_path.exists():
            raise FileNotFoundError(f"Không tìm thấy file Excel 2: {self.excel2_path}")

        self.month_str = month_str  # '2026-10' hoặc None (tự trích xuất)
        self.month_col_header = convert_month_to_excel2_header(self.month_str) if self.month_str else None

    def load_users_data(self) -> Dict[str, Dict[str, Any]]:
        """
        Đọc File 2 và xây dựng từ điển người dùng:
        {
            "nguyễn thị kim ngân": {
                "knox_id": "kimngan.nguyen",
                "name": "Nguyễn Thị Kim Ngân",
                "email": "nguyenthikimngan269@gmail.com",
                "amounts": {
                    "chatgpt": 499000,
                    "claude": 0,
                    "gemini": 0
                }
            },
            ...
        }
        """
        wb = openpyxl.load_workbook(self.excel2_path, data_only=True)
        # Ưu tiên sheet có tên 'thanh toán' hoặc 'payment', nếu không lấy active
        target_sheet = None
        for sname in wb.sheetnames:
            if "thanh toán" in sname.lower() or "thanh toan" in sname.lower():
                target_sheet = wb[sname]
                break
        if not target_sheet:
            target_sheet = wb.active

        ws = target_sheet
        print(f"📄 [ExcelMatcher] Đang đọc File 2 tại sheet: '{ws.title}' (Tổng {ws.max_row} dòng)")

        # 1. Quét tìm dòng header tháng (Row 1) và dòng header provider (Row 2)
        month_header_row = 1
        sub_header_row = 2

        # Dò tìm vị trí các cột
        col_knoxid = None
        col_name = None
        col_email = None

        # Bản đồ các cột theo tháng: { 'chatgpt': col_idx, 'claude': col_idx, 'gemini': col_idx }
        target_month_col_map = {}

        # Dò tìm trong 5 dòng đầu
        for r in range(1, min(6, ws.max_row + 1)):
            for c in range(1, ws.max_column + 1):
                cell_val = str(ws.cell(r, c).value or "").strip()
                cell_lower = cell_val.lower()

                if "knoxid" in cell_lower:
                    col_knoxid = c
                elif cell_lower in ["name", "họ và tên", "ho va ten"]:
                    col_name = c
                elif "thanh toán" in cell_lower and not col_email and r <= 2:
                    # Cột email nằm dưới nhóm Gemini Pro / Đăng nhập / Thanh toán
                    col_email = c

        # Xác định tháng mục tiêu nếu chưa có
        if not self.month_col_header:
            # Tự tìm tháng xuất hiện gần nhất trong dòng 1 (ưu tiên cột có định dạng 'Mmm-YY' ví dụ 'Oct-26')
            for c in range(1, ws.max_column + 1):
                val = str(ws.cell(1, c).value or "").strip()
                if re.match(r'^[A-Za-z]{3}-\d{2}$', val):
                    self.month_col_header = val  # e.g. 'Oct-26'
            if not self.month_col_header:
                self.month_col_header = "Oct-26"

        print(f"🎯 [ExcelMatcher] Cột Tháng đối soát mục tiêu: '{self.month_col_header}'")

        # Tìm phạm vi các cột con của tháng này trong Header
        # File 2 có merged cell cho tháng: 'Oct-26' trải rộng trên các cột con: ChatGPT Plus, Claude Pro, Gemini Pro, Total
        month_start_col = None
        for c in range(1, ws.max_column + 1):
            val1 = str(ws.cell(1, c).value or "").strip()
            if self.month_col_header.lower() in val1.lower():
                month_start_col = c
                break

        if month_start_col:
            # Quét các cột con tiếp theo trong tháng
            for c in range(month_start_col, min(month_start_col + 6, ws.max_column + 1)):
                # Lấy tên cột con ở dòng 2 (hoặc dòng 1 nếu không merge)
                sub_val = str(ws.cell(2, c).value or ws.cell(1, c).value or "").strip().lower()
                for p_key, p_cfg in PROVIDERS_CONFIG.items():
                    if any(kw in sub_val for kw in p_cfg["keywords_file2"]):
                        target_month_col_map[p_key] = c

        print(f"📊 [ExcelMatcher] Vị trí cột đơn giá cho tháng '{self.month_col_header}': {target_month_col_map}")

        # 2. Đọc danh sách nhân sự và số tiền
        users_dict = {}
        data_start_row = max(month_header_row, sub_header_row) + 1

        for r in range(data_start_row, ws.max_row + 1):
            name_val = ws.cell(r, col_name or 3).value
            if not name_val:
                continue

            clean_name_key = normalize_name(name_val)
            knox_id = str(ws.cell(r, col_knoxid or 2).value or "").strip()
            email_val = clean_email(ws.cell(r, col_email or 5).value)

            # Lấy đơn giá từng dịch vụ
            amounts = {}
            for p_key, col_idx in target_month_col_map.items():
                raw_amt = ws.cell(r, col_idx).value
                try:
                    if raw_amt is not None and str(raw_amt).strip() not in ["", "-", "nan"]:
                        # Xử lý cả dạng chuỗi có dấu phẩy '499,000'
                        clean_amt_str = str(raw_amt).replace(",", "").replace(".", "").strip()
                        amounts[p_key] = float(clean_amt_str)
                    else:
                        amounts[p_key] = 0.0
                except Exception:
                    amounts[p_key] = 0.0

            users_dict[clean_name_key] = {
                "name": str(name_val).strip(),
                "knox_id": knox_id,
                "email": email_val,
                "amounts": amounts,
                "row_idx": r
            }

        wb.close()
        print(f"👥 [ExcelMatcher] Đã nạp thành công thông tin của {len(users_dict)} nhân sự từ File 2.")
        return users_dict

    def pair_and_build_records(self, extracted_invoices: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Ghép nối ảnh hóa đơn với dữ liệu File 2.
        Trả về danh sách các bản ghi hoàn chỉnh sẵn sàng nạp vào Database & MinIO:
        """
        users_dict = self.load_users_data()
        final_records = []
        month_db = self.month_str or "2026-10"

        for inv in extracted_invoices:
            sub_name = inv["submitter_name"]
            norm_key = normalize_name(sub_name)
            p_key = inv["provider_key"]
            p_name = inv["provider_name"]

            # Dò tìm thông tin nhân sự trong File 2
            user_info = users_dict.get(norm_key)

            # Nếu không tìm thấy tên chính xác, thử tìm kiếm gần đúng (tên chứa nhau)
            if not user_info:
                for k, v in users_dict.items():
                    if k in norm_key or norm_key in k:
                        user_info = v
                        break

            knox_id = user_info["knox_id"] if user_info else ""
            user_email = user_info["email"] if user_info else ""
            amount = user_info["amounts"].get(p_key, 0.0) if user_info else 0.0

            # Sinh ID ngẫu nhiên 10 ký tự tương đương LOCAL_AI.FN_MAKE_ID(10)
            inv_id = generate_invoice_id(10)

            # Đường dẫn chuẩn trên MinIO
            file_path = f"aibill/{month_db}/{inv_id}.png"

            record = {
                "ID": inv_id,
                "CREATED_AT": None,  # Sẽ dùng SYSDATE trong câu SQL
                "MONTH": month_db,
                "SUBMITTER_NAME": sub_name,
                "SUBMITTER_EMAIL": user_email,
                "FILE_NAME": inv["file_name"],
                "FILE_PATH": file_path,
                "MIME": "image/png",
                "DELETE_TOKEN": None,
                "PROVIDER": p_name,
                "INVOICE_NUMBER": None,
                "INVOICE_DATE": None,
                "BILLING_EMAIL": user_email,
                "AMOUNT": amount,
                "PLAN": inv["plan"],
                "EMAIL_MATCH": 1 if user_email else 0,
                "WARNINGS": None,
                "AI_RAW": None,
                "STATUS": "confirmed",
                "ADMIN_NOTE": None,
                "REVIEWED_AT": None,
                "USERID": knox_id,
                # Dữ liệu phục vụ upload MinIO
                "_raw_bytes": inv["raw_bytes"],
                "_width": inv["width"],
                "_height": inv["height"]
            }

            final_records.append(record)
            status_desc = f"KnoxID: {knox_id or 'CHƯA TÌM THẤY'} | Tiền: {amount:,.0f} VNĐ"
            print(f"🔗 [Matched] {sub_name} -> {p_name} ({status_desc}) -> ID: {inv_id}")

        return final_records
