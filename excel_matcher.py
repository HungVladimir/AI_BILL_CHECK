# ==============================================================================
# MODULE: AI BILL PROCESSOR - EXCEL DATA MATCHER
# File: ai_bill_processor/excel_matcher.py
# ==============================================================================
"""
Đọc File Excel 2 (Sheet 'Thanh toán' hoặc Active sheet):
  1. Tự động tìm Sheet chuẩn chứa danh sách KnoxID và đơn giá các tháng.
  2. Bóc tách email từ cột 'Thanh toán' (chuẩn hóa @gmail.com hoặc giữ @yahoo.com.vn).
  3. Bóc tách đơn giá từng dịch vụ AI (ChatGPT Plus, Claude Pro, Gemini Pro) theo tháng yêu cầu.
  4. Giới hạn chính xác phạm vi cột của tháng (Merged Cell / Total Boundary) tránh ghi đè sang tháng khác.
  5. Ghép nối (Join) 1-1 với ảnh hóa đơn từ File 1 để tạo bản ghi cơ sở dữ liệu hoàn chỉnh.
"""
import re
from datetime import datetime
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


def parse_currency_amount(raw_amt: Any) -> float:
    """
    Chuyển đổi an toàn giá trị ô tiền tệ sang float.
    Tránh lỗi nhân 10 khi số là float (e.g. 499000.0).
    Xử lý được cả số thực, số nguyên, chuỗi có dấu phẩy/chấm/đơn vị tiền tệ.
    """
    if raw_amt is None:
        return 0.0
    if isinstance(raw_amt, (int, float)):
        return float(raw_amt)

    s = str(raw_amt).strip()
    if not s or s.lower() in ["", "-", "nan", "none"]:
        return 0.0

    s_clean = re.sub(r'[^\d.,]', '', s)
    if not s_clean:
        return 0.0

    # Phân biệt dấu phân cách hàng ngàn và thập phân
    if ',' in s_clean and '.' in s_clean:
        if s_clean.rfind(',') > s_clean.rfind('.'):
            # Kiểu 1.234.567,89
            s_clean = s_clean.replace('.', '').replace(',', '.')
        else:
            # Kiểu 1,234,567.89
            s_clean = s_clean.replace(',', '')
    elif ',' in s_clean:
        parts = s_clean.split(',')
        if len(parts[-1]) == 3:
            s_clean = s_clean.replace(',', '')
        else:
            s_clean = s_clean.replace(',', '.')
    elif '.' in s_clean:
        parts = s_clean.split('.')
        if len(parts[-1]) == 3 and len(parts) > 1:
            s_clean = s_clean.replace('.', '')

    try:
        return float(s_clean)
    except Exception:
        return 0.0


def get_month_match_candidates(month_str: str) -> List[str]:
    """
    Sinh danh sách các biến thể có thể xuất hiện trong Header File 2:
    '2026-08' -> ['aug-26', 'aug 26', 'august-26', 'aug', '2026-08', '2026.08', '08/2026', '8/2026', 'tháng 8', 'tháng 08']
    """
    candidates = []
    if not month_str:
        return candidates

    clean_m = month_str.strip()
    candidates.append(clean_m.lower())
    candidates.append(clean_m.replace('-', '.').lower())

    try:
        dt = datetime.strptime(clean_m, "%Y-%m")
        candidates.append(dt.strftime("%b-%y").lower())  # 'aug-26'
        candidates.append(dt.strftime("%b %y").lower())  # 'aug 26'
        candidates.append(dt.strftime("%b_%y").lower())  # 'aug_26'
        candidates.append(dt.strftime("%B-%y").lower())  # 'august-26'
        candidates.append(dt.strftime("%B %y").lower())  # 'august 26'
        candidates.append(dt.strftime("%b").lower())     # 'aug'
        candidates.append(f"{dt.month}/{dt.year}".lower())
        candidates.append(f"{str(dt.month).zfill(2)}/{dt.year}".lower())
        candidates.append(f"tháng {dt.month}".lower())
        candidates.append(f"tháng {str(dt.month).zfill(2)}".lower())
        candidates.append(f"thang {dt.month}".lower())
        candidates.append(f"thang {str(dt.month).zfill(2)}".lower())
        candidates.append(f"t{dt.month}".lower())
        candidates.append(f"t{str(dt.month).zfill(2)}".lower())
    except Exception:
        pass

    return list(dict.fromkeys(candidates))


def select_best_sheet_matcher(wb, file_path: Path):
    """
    Tự động chọn Sheet chính xác chứa dữ liệu bảng thanh toán File 2:
    1. ƯU TIÊN TUYỆT ĐỐI: Sheet có tên 'Thanh toán' (hoặc 'thanh toan' / 'payment').
    2. Sheet chứa 'knoxid', 'name' hoặc từ khóa 'member', 'share'.
    3. Sheet active mặc định.
    """
    # 1. Kiểm tra sheet chuẩn 'Thanh toán' trước tiên
    for sname in wb.sheetnames:
        s_clean = normalize_name(sname)
        if "thanh toán" in s_clean or "thanh toan" in s_clean or "payment" in s_clean:
            return wb[sname]

    sheet_scores = []
    for sname in wb.sheetnames:
        score = 0
        s_lower = sname.lower()
        if any(k in s_lower for k in ["member", "share", "crew", "ai"]):
            score += 20

        # Kiểm tra nội dung 5 dòng đầu
        ws_temp = wb[sname]
        try:
            for r in range(1, min(7, ws_temp.max_row + 1)):
                row_vals = [str(ws_temp.cell(r, c).value or "").strip().lower() for c in range(1, min(25, ws_temp.max_column + 1))]
                if any("knoxid" in v for v in row_vals):
                    score += 40
                if any(k in v for v in row_vals for k in ["name", "họ và tên", "ho va ten"]):
                    score += 20
                if any(any(kw in v for kw in ["chatgpt", "claude", "gemini", "plus", "pro", "total"]) for v in row_vals):
                    score += 20
        except Exception:
            pass

        sheet_scores.append((score, sname))

    sheet_scores.sort(key=lambda x: x[0], reverse=True)
    if sheet_scores and sheet_scores[0][0] > 0:
        return wb[sheet_scores[0][1]]
    return wb.active


class ExcelDataMatcher:
    """Bộ ghép nối đối soát dữ liệu giữa File 1 và File 2."""

    def __init__(self, excel2_path: str, month_str: Optional[str] = None):
        self.excel2_path = Path(excel2_path)
        if not self.excel2_path.exists():
            raise FileNotFoundError(f"Không tìm thấy file Excel 2: {self.excel2_path}")

        self.month_str = month_str  # '2026-08', '2026-10' hoặc None
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
        ws = select_best_sheet_matcher(wb, self.excel2_path)

        print(f"📄 [ExcelMatcher] Đang đọc File 2: {self.excel2_path.name}")
        print(f"   - Danh sách sheet: {wb.sheetnames}")
        print(f"   - Sheet được chọn: '{ws.title}' (Tổng {ws.max_row} dòng, {ws.max_column} cột)")

        # 1. Dò tìm vị trí bắt đầu của các cột Tháng (Aug-26, Sep-26, Oct-26...)
        first_month_col = None
        for c in range(1, ws.max_column + 1):
            for r in range(1, 4):
                v_raw = ws.cell(r, c).value
                if isinstance(v_raw, datetime):
                    first_month_col = c
                    break
                v_str = str(v_raw or "").strip().lower()
                if re.match(r'^[A-Za-z]{3}-\d{2}$', v_str) or any(m in v_str for m in ["aug-26", "sep-26", "oct-26", "nov-26", "dec-26"]):
                    first_month_col = c
                    break
            if first_month_col:
                break

        # 2. Quét tìm vị trí các cột định danh: KnoxID, Name, và cột 'Thanh toán' (nằm trước cột tháng)
        col_knoxid = None
        col_name = None
        col_email = None
        account_email_cols = []
        scan_end_for_info = first_month_col if first_month_col else ws.max_column + 1

        for r in range(1, min(6, ws.max_row + 1)):
            for c in range(1, scan_end_for_info):
                cell_val = str(ws.cell(r, c).value or "").strip()
                cell_lower = cell_val.lower()

                if "knoxid" in cell_lower and not col_knoxid:
                    col_knoxid = c
                elif cell_lower in ["name", "họ và tên", "ho va ten", "họ tên"] and not col_name:
                    col_name = c
                elif ("thanh toán" in cell_lower or "thanh toan" in cell_lower) and r <= 3:
                    if c not in account_email_cols:
                        account_email_cols.append(c)

        if account_email_cols:
            # Ưu tiên cột thanh toán tài khoản (như Gemini Pro / ChatGPT Plus / Claude Pro -> Thanh toán)
            col_email = account_email_cols[-1]

        # Mặc định fallback nếu không thấy cột tên
        if not col_name:
            col_name = 3
        if not col_knoxid:
            col_knoxid = 2
        if not col_email:
            col_email = 5

        print(f"   - Cột định danh nhận diện: KnoxID=Cột {col_knoxid}, Họ tên=Cột {col_name}, Cột Email Thanh toán=Cột {col_email} (tất cả cột email: {account_email_cols or [col_email]})")

        # 3. Định vị cột Tháng mục tiêu (ví dụ '2026-08' -> 'Aug-26')
        month_candidates = get_month_match_candidates(self.month_str)
        month_start_col = None
        month_row = 1

        # Dò tìm cột tháng trong các dòng header (Dòng 1 đến dòng 3)
        for r in range(1, 4):
            for c in range(1, ws.max_column + 1):
                val_raw = ws.cell(r, c).value
                if isinstance(val_raw, datetime):
                    val_dt_str = val_raw.strftime("%Y-%m")
                    val_dt_mon = val_raw.strftime("%b-%y").lower()
                    if val_dt_str in month_candidates or val_dt_mon in month_candidates:
                        month_start_col = c
                        month_row = r
                        print(f"🎯 [ExcelMatcher] Tìm thấy cột Tháng (Date: {val_dt_mon}) tại Dòng {r}, Cột {c}")
                        break

                val = str(val_raw or "").strip().lower()
                if not val:
                    continue
                if any(cand == val or f" {cand} " in f" {val} " or val.startswith(cand) for cand in month_candidates):
                    month_start_col = c
                    month_row = r
                    print(f"🎯 [ExcelMatcher] Tìm thấy cột Tháng '{val}' tại Dòng {r}, Cột {c}")
                    break
            if month_start_col:
                break

        # Fallback nếu không khớp candidates: tìm cột theo mẫu 'Mmm-YY' (e.g. 'Oct-26', 'Aug-26')
        if not month_start_col:
            for c in range(1, ws.max_column + 1):
                val = str(ws.cell(1, c).value or ws.cell(2, c).value or "").strip()
                if re.match(r'^[A-Za-z]{3}-\d{2}$', val):
                    month_start_col = c
                    print(f"🎯 [ExcelMatcher] Fallback chọn cột Tháng '{val}' tại Cột {c}")
                    break

        if not month_start_col:
            print("⚠️ [ExcelMatcher Warning] Không tìm thấy cột Tháng tương ứng trong File 2! Mặc định dò tìm các cột có tên dịch vụ.")
            month_start_col = 6

        # 4. Xác định phạm vi các cột con của tháng này (tránh quét lấn sang tháng sau)
        # Kiểm tra merged cell tại (month_row, month_start_col)
        month_end_col = month_start_col + 5
        for mrange in ws.merged_cells.ranges:
            if mrange.min_row <= month_row <= mrange.max_row and mrange.min_col == month_start_col:
                month_end_col = mrange.max_col
                break

        # Đảm bảo không vượt quá max_column
        month_end_col = min(month_end_col, ws.max_column)

        # 5. Ánh xạ các cột nhà cung cấp (ChatGPT, Claude, Gemini) trong phạm vi tháng
        target_month_col_map = {}
        for c in range(month_start_col, month_end_col + 1):
            # Thu thập tiêu đề con ở các dòng (month_row, month_row + 1, month_row + 2)
            sub_texts = []
            for r_sub in range(month_row, month_row + 3):
                s_val = str(ws.cell(r_sub, c).value or "").strip().lower()
                if s_val and s_val not in sub_texts:
                    sub_texts.append(s_val)
            combined_sub = " ".join(sub_texts)

            # Nếu gặp cột Total thì dừng hoặc bỏ qua
            if "total" in combined_sub and not any(p in combined_sub for p in ["chatgpt", "claude", "gemini"]):
                continue

            for p_key, p_cfg in PROVIDERS_CONFIG.items():
                if p_key not in target_month_col_map:
                    if any(kw in combined_sub for kw in p_cfg["keywords_file2"]):
                        target_month_col_map[p_key] = c
                        break

        print(f"📊 [ExcelMatcher] Ánh xạ cột đơn giá tháng {self.month_str}: {target_month_col_map}")

        # 6. Đọc danh sách nhân sự và số tiền
        users_dict = {}
        data_start_row = month_row + 2

        for r in range(data_start_row, ws.max_row + 1):
            name_val = ws.cell(r, col_name).value
            if not name_val:
                continue

            clean_name_key = normalize_name(name_val)
            if any(k in clean_name_key for k in ["tổng", "total", "sum", "ghi chú", "note"]):
                continue

            knox_id = str(ws.cell(r, col_knoxid).value or "").strip()

            # Lấy email thanh toán: quét qua các cột account_email_cols nếu có
            raw_email_str = ""
            if account_email_cols:
                for c_em in account_email_cols:
                    v_em = ws.cell(r, c_em).value
                    if v_em and str(v_em).strip() not in ["", "-", "nan", "None", "0"]:
                        raw_email_str = str(v_em).strip()
                        break
            if not raw_email_str:
                raw_email_str = str(ws.cell(r, col_email).value or "").strip()

            email_val = clean_email(raw_email_str)

            # Lấy đơn giá từng dịch vụ
            amounts = {}
            for p_key, col_idx in target_month_col_map.items():
                raw_amt = ws.cell(r, col_idx).value
                amounts[p_key] = parse_currency_amount(raw_amt)

            users_dict[clean_name_key] = {
                "name": str(name_val).strip(),
                "knox_id": knox_id,
                "email": email_val,
                "amounts": amounts,
                "row_idx": r
            }

        wb.close()
        print(f"👥 [ExcelMatcher] Đã nạp thành công {len(users_dict)} nhân sự từ File 2.")
        return users_dict

    def pair_and_build_records(self, extracted_invoices: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Ghép nối ảnh hóa đơn với dữ liệu File 2.
        Trả về danh sách các bản ghi hoàn chỉnh sẵn sàng nạp vào Database & MinIO.
        """
        users_dict = self.load_users_data()
        final_records = []
        month_db = self.month_str or "2026-08"

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

            # Đường dẫn chuẩn trên MinIO: aibill/{YYYY-MM}/{ID}.png
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
                "STATUS": "confirmed" if (amount > 0 and knox_id) else "pending",
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
