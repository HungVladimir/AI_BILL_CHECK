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
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
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
    """Chuẩn hóa họ tên (NFC unicode, bỏ khoảng trắng thừa, viết thường)."""
    if not name:
        return ""
    nfc = unicodedata.normalize('NFC', str(name).strip().lower())
    return re.sub(r'\s+', ' ', nfc)


def remove_accents(text: Any) -> str:
    """Bỏ dấu tiếng Việt (NFD -> bỏ combining diacritics, đ/Đ -> d/D) phục vụ so khớp mờ."""
    if not text:
        return ""
    nfd = unicodedata.normalize('NFD', str(text).strip())
    no_acc = "".join(c for c in nfd if unicodedata.category(c) != 'Mn')
    no_acc = no_acc.replace('đ', 'd').replace('Đ', 'D')
    return re.sub(r'\s+', ' ', no_acc).strip().lower()


def normalize_knoxid(knoxid: Any) -> str:
    """Chuẩn hóa KnoxID (bỏ khoảng trắng thừa, viết thường)."""
    if not knoxid:
        return ""
    return str(knoxid).strip().lower()


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
        # Bỏ qua sheet tiêu chuẩn email khi chọn sheet thanh toán đơn giá
        if "tiêu chuẩn" in s_lower or "tieu chuan" in s_lower:
            continue
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


def load_email_standards(wb) -> Dict[str, Any]:
    """
    Đọc Sheet 'Tiêu chuẩn email' trong File 2 nếu có.
    Cấu trúc sheet:
      KnoxID | Name | AI_PROVIDER | AI_GEN | AI_TIER | EMAIL_PAYMENT
    Trả về dict phân loại hỗ trợ tra cứu đa chiều (KnoxID, Name có dấu/không dấu, Provider/Gen).
    """
    standards = {
        "by_name_gen": {},
        "by_name_provider": {},
        "by_name": {},
        "by_name_no_accent": {},
        "by_knoxid": {}
    }

    target_sheet = None
    for sname in wb.sheetnames:
        s_clean = normalize_name(sname)
        if any(k in s_clean for k in ["tiêu chuẩn email", "tieu chuan email", "tiêu chuẩn mail", "tieu chuan mail", "tiêu chuẩn", "tieu chuan"]):
            target_sheet = wb[sname]
            break

    if not target_sheet:
        return standards

    ws = target_sheet
    print(f"📧 [ExcelMatcher] Phát hiện Sheet '{ws.title}' (Chứa Tiêu chuẩn Email & KnoxID)")

    header_row = 1
    col_knoxid = None
    col_name = None
    col_provider = None
    col_gen = None
    col_tier = None
    col_email = None

    for r in range(1, min(6, ws.max_row + 1)):
        for c in range(1, min(20, ws.max_column + 1)):
            val = str(ws.cell(r, c).value or "").strip().lower()
            if not val:
                continue
            if "knoxid" in val or "knox_id" in val:
                col_knoxid = c
                header_row = r
            elif val in ["name", "họ và tên", "ho va ten", "họ tên"]:
                col_name = c
                header_row = r
            elif "provider" in val:
                col_provider = c
            elif "gen" in val:
                col_gen = c
            elif "tier" in val:
                col_tier = c
            elif any(k in val for k in ["email_payment", "email", "mail", "payment"]):
                col_email = c

    if not col_name:
        col_name = 2
    if not col_knoxid:
        col_knoxid = 1
    if not col_email:
        col_email = 6

    for r in range(header_row + 1, ws.max_row + 1):
        name_val = ws.cell(r, col_name).value
        if not name_val:
            continue
        name_str = str(name_val).strip()
        norm_name = normalize_name(name_str)
        no_acc_name = remove_accents(name_str)

        if not norm_name or any(k in norm_name for k in ["tổng", "total", "ghi chú"]):
            continue

        knox_val = str(ws.cell(r, col_knoxid).value or "").strip()
        prov_val = str(ws.cell(r, col_provider).value or "").strip() if col_provider else ""
        gen_val = str(ws.cell(r, col_gen).value or "").strip() if col_gen else ""
        tier_val = str(ws.cell(r, col_tier).value or "").strip() if col_tier else ""
        email_raw = ws.cell(r, col_email).value if col_email else ""
        email_val = clean_email(email_raw)

        item = {
            "knox_id": knox_val,
            "name": name_str,
            "ai_provider": prov_val,
            "ai_gen": gen_val,
            "ai_tier": tier_val,
            "email": email_val
        }

        norm_gen = normalize_name(gen_val)
        norm_prov = normalize_name(prov_val)
        norm_knox = normalize_knoxid(knox_val)

        if norm_gen:
            standards["by_name_gen"][(norm_name, norm_gen)] = item
            standards["by_name_gen"][(no_acc_name, norm_gen)] = item
            if "chatgpt" in norm_gen:
                standards["by_name_gen"][(norm_name, "chatgpt")] = item
                standards["by_name_gen"][(no_acc_name, "chatgpt")] = item
            elif "claude" in norm_gen:
                standards["by_name_gen"][(norm_name, "claude")] = item
                standards["by_name_gen"][(no_acc_name, "claude")] = item
            elif "gemini" in norm_gen:
                standards["by_name_gen"][(norm_name, "gemini")] = item
                standards["by_name_gen"][(no_acc_name, "gemini")] = item

        if norm_prov:
            standards["by_name_provider"][(norm_name, norm_prov)] = item
            standards["by_name_provider"][(no_acc_name, norm_prov)] = item
            if "openai" in norm_prov:
                standards["by_name_gen"][(norm_name, "chatgpt")] = item
                standards["by_name_gen"][(no_acc_name, "chatgpt")] = item
            elif "anthropic" in norm_prov:
                standards["by_name_gen"][(norm_name, "claude")] = item
                standards["by_name_gen"][(no_acc_name, "claude")] = item
            elif "google" in norm_prov:
                standards["by_name_gen"][(norm_name, "gemini")] = item
                standards["by_name_gen"][(no_acc_name, "gemini")] = item

        if norm_name not in standards["by_name"] or email_val:
            standards["by_name"][norm_name] = item
        if no_acc_name not in standards["by_name_no_accent"] or email_val:
            standards["by_name_no_accent"][no_acc_name] = item
        if norm_knox:
            standards["by_knoxid"][norm_knox] = item

    print(f"   👉 Đã nạp thông tin tiêu chuẩn của {len(standards['by_name'])} nhân sự từ Sheet Tiêu chuẩn email.")
    return standards


class ExcelDataMatcher:
    """Bộ ghép nối đối soát dữ liệu giữa File 1 và File 2."""

    def __init__(self, excel2_path: str, month_str: Optional[str] = None):
        self.excel2_path = Path(excel2_path)
        if not self.excel2_path.exists():
            raise FileNotFoundError(f"Không tìm thấy file Excel 2: {self.excel2_path}")

        self.month_str = month_str  # '2026-08', '2026-10' hoặc None
        self.month_col_header = convert_month_to_excel2_header(self.month_str) if self.month_str else None
        self.email_standards = {
            "by_name_gen": {},
            "by_name_provider": {},
            "by_name": {},
            "by_name_no_accent": {},
            "by_knoxid": {}
        }
        self.month_blocks = {}
        self.users_by_knoxid = {}
        self.users_by_name = {}
        self.users_by_name_no_accent = {}
        self.users_list = []

    def load_users_data(self) -> Dict[str, Dict[str, Any]]:
        """
        Đọc File 2 (Sheet 'Thanh toán' & 'Tiêu chuẩn email'):
        Tự động nhận diện mọi khối Tháng (Aug-26, Sep-26, Oct-26...) kết thúc bởi cột Total,
        ánh xạ chính xác các cột con ChatGPT Plus, Claude Pro, Gemini Pro, và bóc tách
        đơn giá theo từng nhân sự.
        """
        wb = openpyxl.load_workbook(self.excel2_path, data_only=True)
        # 0. Nạp thông tin từ sheet 'Tiêu chuẩn email' trước
        self.email_standards = load_email_standards(wb)

        ws = select_best_sheet_matcher(wb, self.excel2_path)

        print(f"📄 [ExcelMatcher] Đang đọc File 2: {self.excel2_path.name}")
        print(f"   - Danh sách sheet: {wb.sheetnames}")
        print(f"   - Sheet được chọn: '{ws.title}' (Tổng {ws.max_row} dòng, {ws.max_column} cột)")

        # 1. Tìm dòng Month Header (chứa Aug-26, Sep-26, Oct-26...) và dòng Sub-Header (chứa ChatGPT Plus, Claude Pro, Total...)
        month_header_row = None
        sub_header_row = None

        for r in range(1, min(7, ws.max_row + 1)):
            row_vals = [ws.cell(r, c).value for c in range(1, min(35, ws.max_column + 1))]
            has_month = False
            for v in row_vals:
                if isinstance(v, datetime):
                    has_month = True
                    break
                v_str = str(v or "").strip().lower()
                if re.match(r'^[A-Za-z]{3}[-_ ]\d{2}$', v_str) or any(m in v_str for m in ["aug-26", "sep-26", "oct-26", "nov-26", "dec-26"]):
                    has_month = True
                    break
            
            if has_month and not month_header_row:
                month_header_row = r
                sub_header_row = r + 1
                break

        if not month_header_row:
            month_header_row = 1
            sub_header_row = 2

        print(f"   - Header tháng tại Dòng {month_header_row}, Sub-header dịch vụ tại Dòng {sub_header_row}")

        # 2. Dò tìm cột tháng đầu tiên
        first_m_col = ws.max_column + 1
        for c in range(1, ws.max_column + 1):
            v = ws.cell(month_header_row, c).value
            if isinstance(v, datetime):
                first_m_col = c
                break
            v_str = str(v or "").strip().lower()
            if re.match(r'^[A-Za-z]{3}[-_ ]\d{2}$', v_str) or any(m in v_str for m in ["aug-26", "sep-26", "oct-26", "nov-26", "dec-26"]):
                first_m_col = c
                break

        # 3. Quét tìm vị trí các cột định danh: KnoxID, Name, và các cột Email tài khoản
        col_knoxid = None
        col_name = None
        col_email = None
        account_email_cols = []

        for r in range(1, sub_header_row + 1):
            for c in range(1, first_m_col):
                val = str(ws.cell(r, c).value or "").strip().lower()
                if not val:
                    continue
                if ("knoxid" in val or "knox_id" in val) and not col_knoxid:
                    col_knoxid = c
                elif val in ["name", "họ và tên", "ho va ten", "họ tên"] and not col_name:
                    col_name = c
                elif ("thanh toán" in val or "thanh toan" in val):
                    if c not in account_email_cols:
                        account_email_cols.append(c)

        if account_email_cols:
            col_email = account_email_cols[-1]

        if not col_knoxid:
            col_knoxid = 2
        if not col_name:
            col_name = 3
        if not col_email:
            col_email = 5

        print(f"   - Cột định danh: KnoxID=Cột {col_knoxid}, Họ tên=Cột {col_name}, Email=Cột {col_email}")

        # 4. Quét và nhận diện TOÀN BỘ các Khối Tháng (Month Blocks) trong File 2
        # Mỗi tháng bắt đầu ở ô Tháng (Row 1) và KẾT THÚC ở cột Total (Row 2 có 'total')
        self.month_blocks = {}
        c = first_m_col
        while c <= ws.max_column:
            val = ws.cell(month_header_row, c).value
            val_str = str(val or "").strip().lower()
            m_key = None

            if isinstance(val, datetime):
                m_key = val.strftime("%Y-%m")
            elif re.match(r'^[A-Za-z]{3}[-_ ]\d{2}$', val_str):
                try:
                    dt = datetime.strptime(val_str.replace('_', '-').replace(' ', '-'), "%b-%y")
                    m_key = dt.strftime("%Y-%m")
                except Exception:
                    m_key = val_str
            else:
                cand_list = ["aug", "sep", "oct", "nov", "dec", "jan", "feb", "mar", "apr", "may", "jun", "jul"]
                for eng_m in cand_list:
                    if eng_m in val_str:
                        m_key = extract_month_from_filename(val_str)
                        break

            if m_key:
                # Dò tìm cột kết thúc của khối tháng này (kết thúc tại cột Total)
                end_c = c
                for next_c in range(c, min(c + 8, ws.max_column + 1)):
                    sub_text = str(ws.cell(sub_header_row, next_c).value or "").strip().lower()
                    if "total" in sub_text or "tổng" in sub_text:
                        end_c = next_c
                        break

                # Ánh xạ các cột con (ChatGPT, Claude, Gemini, Total) trong khoảng [c, end_c]
                col_map = {}
                for sub_c in range(c, end_c + 1):
                    sub_val = str(ws.cell(sub_header_row, sub_c).value or "").strip().lower()
                    top_sub_val = str(ws.cell(month_header_row, sub_c).value or "").strip().lower()
                    comb = f"{top_sub_val} {sub_val}".strip()

                    if any(kw in comb for kw in ["chatgpt", "chatcpt", "gpt", "plus"]) and not any(k in comb for k in ["claude", "gemini", "total"]):
                        col_map["chatgpt"] = sub_c
                    elif any(kw in comb for kw in ["claude", "anthropic"]) and not any(k in comb for k in ["chatgpt", "gemini", "total"]):
                        col_map["claude"] = sub_c
                    elif any(kw in comb for kw in ["gemini", "google"]) and not any(k in comb for k in ["chatgpt", "claude", "total"]):
                        col_map["gemini"] = sub_c
                    elif any(kw in comb for kw in ["total", "tổng cộng", "tong cong", "tổng"]):
                        col_map["total"] = sub_c

                self.month_blocks[m_key] = {
                    "header": str(val),
                    "start_col": c,
                    "end_col": end_c,
                    "col_map": col_map
                }
                c = end_c + 1
            else:
                c += 1

        print(f"📊 [ExcelMatcher] Đã phát hiện {len(self.month_blocks)} khối tháng trong Sheet '{ws.title}':")
        for m, b in self.month_blocks.items():
            print(f"   - Tháng {m} ('{b['header']}'): Cột {b['start_col']} -> {b['end_col']} | Cột con: {b['col_map']}")

        # 5. Đọc dữ liệu từng dòng nhân sự và lưu vào bộ nhớ tra cứu đa chiều
        self.users_by_knoxid = {}
        self.users_by_name = {}
        self.users_by_name_no_accent = {}
        self.users_list = []
        data_start_row = sub_header_row + 1

        for r in range(data_start_row, ws.max_row + 1):
            name_val = ws.cell(r, col_name).value
            if not name_val:
                continue

            name_str = str(name_val).strip()
            norm_name = normalize_name(name_str)
            no_acc_name = remove_accents(name_str)

            if any(k in norm_name for k in ["tổng", "total", "sum", "ghi chú", "note"]):
                continue

            knox_val = str(ws.cell(r, col_knoxid).value or "").strip()
            norm_knox = normalize_knoxid(knox_val)

            # Lấy email thanh toán
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

            # Bóc tách đơn giá của TẤT CẢ các tháng đã nhận diện
            amounts_by_month = {}
            for m_key, b in self.month_blocks.items():
                col_map = b["col_map"]
                m_amts = {}
                for p_key, col_idx in col_map.items():
                    m_amts[p_key] = parse_currency_amount(ws.cell(r, col_idx).value)
                amounts_by_month[m_key] = m_amts

            user_item = {
                "name": name_str,
                "knox_id": knox_val,
                "email": email_val,
                "amounts_by_month": amounts_by_month,
                "row_idx": r
            }

            if norm_knox:
                self.users_by_knoxid[norm_knox] = user_item
            if norm_name:
                self.users_by_name[norm_name] = user_item
            if no_acc_name:
                self.users_by_name_no_accent[no_acc_name] = user_item
            self.users_list.append(user_item)

        wb.close()
        print(f"👥 [ExcelMatcher] Đã nạp thành công {len(self.users_list)} nhân sự từ File 2.")
        return self.users_by_name

    def pair_and_build_records(
        self,
        extracted_invoices: List[Dict[str, Any]],
        month_override: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """
        Ghép nối ảnh hóa đơn với dữ liệu File 2.
        Ưu tiên lấy KnoxID và Email từ Sheet 'Tiêu chuẩn email', lấy đơn giá từ Sheet 'Thanh toán'.
        Trả về danh sách 22 trường dữ liệu chuẩn của bảng LOCAL_AI.AI_INVOICE_CHECK.
        """
        if not self.month_blocks:
            self.load_users_data()

        final_records = []
        month_db = month_override or self.month_str or "2026-08"

        # Xác định khối tháng cần dùng trong File 2
        used_month = month_db
        if month_db not in self.month_blocks:
            cands = get_month_match_candidates(month_db)
            matched_m = None
            for cand in cands:
                for mb_key in self.month_blocks.keys():
                    if cand in mb_key or mb_key in cand:
                        matched_m = mb_key
                        break
                if matched_m:
                    break

            if matched_m:
                used_month = matched_m
            elif self.month_blocks:
                used_month = list(self.month_blocks.keys())[-1]
                print(f"⚠️ [ExcelMatcher Warning] Tháng {month_db} chưa có trong File 2, tự động dùng tháng gần nhất: {used_month}")

        header_desc = self.month_blocks.get(used_month, {}).get("header", used_month)
        print(f"🎯 [ExcelMatcher] Áp dụng bảng đơn giá tháng: {used_month} ('{header_desc}') cho đợt thanh toán {month_db}")

        for inv in extracted_invoices:
            sub_name = inv["submitter_name"]
            norm_name = normalize_name(sub_name)
            no_acc_name = remove_accents(sub_name)
            p_key = inv["provider_key"].lower()
            p_name = inv["provider_name"]

            # 1. Tìm thông tin trong Sheet 'Tiêu chuẩn email' trước
            std_info = None
            if self.email_standards:
                std_info = self.email_standards["by_name_gen"].get((norm_name, p_key))
                if not std_info:
                    std_info = self.email_standards["by_name_gen"].get((no_acc_name, p_key))
                if not std_info and p_key == "chatgpt":
                    std_info = self.email_standards["by_name_provider"].get((norm_name, "openai"))
                elif not std_info and p_key == "claude":
                    std_info = self.email_standards["by_name_provider"].get((norm_name, "anthropic"))
                elif not std_info and p_key == "gemini":
                    std_info = self.email_standards["by_name_provider"].get((norm_name, "google"))

                if not std_info:
                    std_info = self.email_standards["by_name"].get(norm_name)
                if not std_info:
                    std_info = self.email_standards["by_name_no_accent"].get(no_acc_name)

                if not std_info:
                    for k_name, v_info in self.email_standards["by_name"].items():
                        if k_name in norm_name or norm_name in k_name:
                            std_info = v_info
                            break

            knox_id = std_info["knox_id"] if std_info else ""
            user_email = std_info["email"] if std_info else ""

            # 2. Dò tìm thông tin nhân sự trong Sheet 'Thanh toán' (Ưu tiên theo KnoxID)
            user_info = None
            if knox_id:
                user_info = self.users_by_knoxid.get(normalize_knoxid(knox_id))
            if not user_info:
                user_info = self.users_by_name.get(norm_name)
            if not user_info:
                user_info = self.users_by_name_no_accent.get(no_acc_name)
            if not user_info:
                for u in self.users_list:
                    u_norm = normalize_name(u["name"])
                    u_no_acc = remove_accents(u["name"])
                    if norm_name in u_norm or u_norm in norm_name or no_acc_name in u_no_acc or u_no_acc in no_acc_name:
                        user_info = u
                        break

            if user_info:
                if not knox_id:
                    knox_id = user_info.get("knox_id", "")
                if not user_email:
                    user_email = user_info.get("email", "")

            # 3. Lấy đơn giá tháng used_month từ user_info
            amount = 0.0
            if user_info:
                m_amts = user_info["amounts_by_month"].get(used_month, {})
                amt = m_amts.get(p_key, 0.0)
                if amt > 0:
                    amount = amt
                else:
                    # Fallback an toàn:
                    # Nếu người này chỉ có 1 dịch vụ > 0 trong tháng -> lấy số đó
                    # Hoặc nếu cột Total > 0 -> lấy Total
                    other_positives = [v for k, v in m_amts.items() if k != "total" and v > 0]
                    if len(other_positives) == 1:
                        amount = other_positives[0]
                    elif m_amts.get("total", 0.0) > 0:
                        amount = m_amts.get("total", 0.0)
                    else:
                        amount = 0.0

            # 4. Sinh ID ngẫu nhiên 10 ký tự tương đương LOCAL_AI.FN_MAKE_ID(10)
            inv_id = generate_invoice_id(10)
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S.000")
            file_path = f"aibill/{month_db}/{inv_id}.png"

            # 22 Trường dữ liệu chuẩn theo đúng định dạng Database và yêu cầu của người dùng
            record = {
                "ID": inv_id,
                "CREATED_AT": now_str,
                "MONTH": month_db,
                "SUBMITTER_NAME": sub_name,
                "SUBMITTER_EMAIL": user_email,
                "FILE_NAME": inv.get("file_name", inv.get("original_filename", f"{p_key}.png")),
                "FILE_PATH": file_path,
                "MIME": "image/png",
                "DELETE_TOKEN": None,
                "PROVIDER": p_name,
                "INVOICE_NUMBER": None,
                "INVOICE_DATE": None,
                "BILLING_EMAIL": user_email,
                "AMOUNT": amount,
                "PLAN": None,
                "EMAIL_MATCH": 1,
                "WARNINGS": None,
                "AI_RAW": None,
                "STATUS": "confirmed",
                "ADMIN_NOTE": None,
                "REVIEWED_AT": None,
                "USERID": knox_id,
                # Dữ liệu nội bộ phục vụ upload MinIO / lưu file ảnh cục bộ
                "_raw_bytes": inv.get("raw_bytes") or inv.get("_raw_bytes", b""),
                "_width": inv.get("width", 0),
                "_height": inv.get("height", 0)
            }

            final_records.append(record)
            status_desc = f"KnoxID: {knox_id or 'CHƯA TÌM THẤY'} | Mail: {user_email or 'CHƯA TÌM THẤY'} | Tiền: {amount:,.0f} VNĐ"
            print(f"🔗 [Matched] {sub_name} -> {p_name} ({status_desc}) -> ID: {inv_id}")

        return final_records
