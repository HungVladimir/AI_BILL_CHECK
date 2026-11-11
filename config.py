# ==============================================================================
# MODULE: AI BILL PROCESSOR - CONFIGURATION
# File: ai_bill_processor/config.py
# ==============================================================================
import os
import re
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv, find_dotenv

# Nạp biến môi trường từ .env (ưu tiên .env của ai_bill_processor, sau đó tới backend/.env)
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")
load_dotenv(BASE_DIR.parent / "backend" / ".env")
load_dotenv(find_dotenv())

# --- 0. CẤU HÌNH ĐƯỜNG DẪN DỮ LIỆU PRIVATE (EXCEL) VÀ DỰ ÁN (IMAGES) ---
# Thư mục Private: Nơi lưu và đọc các file Excel (đầu vào & báo cáo đầu ra)
AI_BILL_PRIVATE_DIR = Path(
    os.getenv("AI_BILL_PRIVATE_DIR")
    or os.getenv("AI_BILL_DIR")
    or r"D:\Private_Documents\Doc\AI-Bill"
)

# Thư mục con trong Private dành riêng cho các file Excel:
DIR_EXCEL1 = AI_BILL_PRIVATE_DIR / "invoices_file1"        # File 1: Chứa ảnh hóa đơn các tháng
DIR_EXCEL2 = AI_BILL_PRIVATE_DIR / "members_file2"         # File 2: Chứa danh sách member share, đơn giá
DIR_EXCEL_OUTPUT = AI_BILL_PRIVATE_DIR / "output_reports"  # Nơi xuất file Excel báo cáo đối soát

# Thư mục lưu ảnh trích xuất: LUÔN LƯU TRONG FOLDER DỰ ÁN (PROJECT)
# (Lý do: Thư mục Private không hỗ trợ lưu file ảnh)
DIR_PROJECT_IMAGES = BASE_DIR / "extracted_images"

# Tên Sheet chuẩn theo quy ước nghiệp vụ:
SHEET_NAME_FILE1 = "Tổng hợp hóa đơn"       # File 1: Sheet chứa danh sách và ảnh hóa đơn
SHEET_NAME_FILE2 = "Thanh toán"             # File 2: Sheet chứa KnoxID, email và đơn giá các tháng
SHEET_NAME_EMAIL_STANDARD = "Tiêu chuẩn email"  # File 2: Sheet chứa KnoxID, Name, AI_PROVIDER, AI_GEN, AI_TIER, EMAIL_PAYMENT

# 22 Cột chuẩn Database LOCAL_AI.AI_INVOICE_CHECK:
DB_COLUMNS = [
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
    "USERID"
]

# --- 1. CẤU HÌNH NHÀ CUNG CẤP AI (CHATGPT, CLAUDE, GEMINI) ---
PROVIDERS_CONFIG = {
    "chatgpt": {
        "name": "ChatGPT",
        "keywords_file1": [
            "chatgpt", "chat gpt", "chat-gpt", "chatcpt", "cpt", "openai", "gpt", "gpt-4", "gpt4",
            "hóa đơn chatgpt", "hoa don chatgpt", "hóa đơn chat gpt", "hoa don chat gpt",
            "hóa đơn chatcpt", "hoa don chatcpt", "hóa đơn openai", "hoa don openai",
            "ảnh chatgpt", "minh chứng chatgpt", "bill chatgpt", "receipt chatgpt",
            "chatgpt plus", "chatgpt 4"
        ],
        "keywords_file2": [
            "chatgpt plus", "chatgpt", "chat gpt", "chat-gpt", "chatcpt", "gpt", "openai"
        ],
        "default_filename": "chatGpt.png",
        "plan": "ChatGPT Plus"
    },
    "claude": {
        "name": "Claude",
        "keywords_file1": [
            "claude", "claude pro", "claude 3", "claude ai", "anthropic",
            "hóa đơn claude", "hoa don claude", "hóa đơn claude pro", "hoa don claude pro",
            "hóa đơn anthropic", "ảnh claude", "minh chứng claude", "bill claude", "receipt claude"
        ],
        "keywords_file2": [
            "claude pro", "claude", "anthropic"
        ],
        "default_filename": "claude.png",
        "plan": "Claude Pro"
    },
    "gemini": {
        "name": "Gemini",
        "keywords_file1": [
            "gemini", "gemini pro", "gemini advanced", "google ai", "google",
            "hóa đơn gemini", "hoa don gemini", "hóa đơn google ai", "hoa don google ai",
            "ảnh gemini", "minh chứng gemini", "bill gemini", "receipt gemini"
        ],
        "keywords_file2": [
            "gemini pro", "gemini", "google ai pro", "google ai pro (5 tb)", "google ai"
        ],
        "default_filename": "gemini.png",
        "plan": "Gemini Pro"
    }
}

# --- 2. CẤU HÌNH MINIO OBJECT STORAGE ---
MINIO_ENDPOINT = os.getenv("ONEAI_MINIO_ENDPOINT") or os.getenv("MINIO_ENDPOINT")
MINIO_ACCESS_KEY = os.getenv("ONEAI_MINIO_ACCESS") or os.getenv("MINIO_ACCESS_KEY")
MINIO_SECRET_KEY = os.getenv("ONEAI_MINIO_SECRET") or os.getenv("MINIO_SECRET_KEY")
MINIO_BUCKET_NAME = os.getenv("ONEAI_MINIO_BUCKETNAME") or os.getenv("MINIO_BUCKET_NAME", "oneai")
MINIO_PREFIX = "aibill"  # Cấu trúc: aibill/{YYYY-MM}/{ID}.png

# --- 3. CẤU HÌNH ORACLE DATABASE (MBLDB) ---
DB_DNS = os.getenv("DB_DNS", "MBLDB")
DB_USER = os.getenv("DB_USER", "LOCAL_AI")
DB_PASSWORD = (
    os.getenv("DB_MBLDB_LOCAL_AI__PASSWORD")
    or os.getenv("DB_MBLDB_LOCAL_AI_PASSWORD")
    or os.getenv("DB_PASSWORD_YMS_AI")
    or "WKEHDGHK1!"
)
ORACLE_CLIENT_DIR = os.getenv("ORACLE_CLIENT_LIB_DIR", r"C:\oracle_client")
TABLE_NAME = "LOCAL_AI.AI_INVOICE_CHECK"


# --- 4. HÀM TIỆN ÍCH TRÍCH XUẤT THÁNG & ĐỒNG BỘ HEADER ---
def extract_month_from_filename(filename: str) -> str:
    """
    Tự động trích xuất tháng từ tên file dạng:
    '(AI Crew) thanh toán hóa đơn tháng 2026.08.xlsx' -> '2026-08'
    '(AI Crew) thanh toán hóa đơn tháng 2026.10.xlsx' -> '2026-10'
    '(AI Crew) thanh toán hóa đơn tháng 8.xlsx' -> '2026-08'
    '(AI Crew) thanh toán hóa đơn tháng 8.2026.xlsx' -> '2026-08'
    'Hoa don thang 8.2026.xlsx' -> '2026-08'
    'Aug-26.xlsx' -> '2026-08'
    'Oct-26.xlsx' -> '2026-10'
    """
    clean_name = Path(filename).name.lower()

    # 1. Mẫu YYYY.MM hoặc YYYY-MM hoặc YYYY_MM (e.g. 2026.08, 2026.8, 2026-10)
    m = re.search(r'(20\d{2})[._\-](\d{1,2})', clean_name)
    if m:
        year = m.group(1)
        month = m.group(2).zfill(2)
        if 1 <= int(month) <= 12:
            return f"{year}-{month}"

    # 2. Mẫu MM.YYYY hoặc MM-YYYY hoặc MM_YYYY (e.g. 08.2026, 8.2026, 10-2026)
    m = re.search(r'(?:^|[^\d])(\d{1,2})[._\-](20\d{2})', clean_name)
    if m:
        month = m.group(1).zfill(2)
        year = m.group(2)
        if 1 <= int(month) <= 12:
            return f"{year}-{month}"

    # 3. Mẫu tháng M.YYYY hoặc tháng M YYYY (e.g. tháng 8.2026, thang 8 2026, tháng 8)
    m = re.search(r'th[aá]ng\s*(\d{1,2})(?:[._\-\s]*(20\d{2}))?', clean_name)
    if m:
        month = m.group(1).zfill(2)
        year = m.group(2) or "2026"
        if 1 <= int(month) <= 12:
            return f"{year}-{month}"

    # 4. Mẫu tên tháng tiếng Anh (e.g. aug-26, sep-26, oct-26)
    eng_months = {
        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04', 'may': '05', 'jun': '06',
        'jul': '07', 'aug': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
    }
    for em, num in eng_months.items():
        if em in clean_name:
            m_yr = re.search(r'20\d{2}|\d{2}', clean_name)
            yr = f"20{m_yr.group(0)}" if (m_yr and len(m_yr.group(0)) == 2) else (m_yr.group(0) if m_yr else "2026")
            return f"{yr}-{num}"

    # Mặc định lấy tháng hiện tại nếu không nhận diện được
    return datetime.now().strftime("%Y-%m")


def convert_month_to_excel2_header(month_str: str) -> str:
    """
    Quy đổi định dạng tháng 'YYYY-MM' sang định dạng cột của File 2 (Mon-YY):
    '2026-10' -> 'Oct-26'
    '2026-09' -> 'Sep-26'
    '2026-08' -> 'Aug-26'
    '2026-11' -> 'Nov-26'
    '2026-12' -> 'Dec-26'
    """
    try:
        dt = datetime.strptime(month_str.strip(), "%Y-%m")
        return dt.strftime("%b-%y")  # 'Oct-26'
    except Exception:
        return month_str


def locate_ai_bill_files(
    private_dir: Path = AI_BILL_PRIVATE_DIR,
    target_month: str = None
):
    """
    Tự động dò tìm File 1 và File 2 trong thư mục Private:
    Hỗ trợ 2 mô hình lưu trữ:
      Mô hình A (Thư mục chung): Cả 2 file nằm ngay trong private_dir
      Mô hình B (Phân loại): File 1 nằm trong invoices_file1/, File 2 nằm trong members_file2/
    Trả về: (file1_path, file2_path, dir_excel_output, dir_project_images)
    """
    private_dir = Path(private_dir)
    dir_f1 = private_dir / "invoices_file1"
    dir_f2 = private_dir / "members_file2"
    dir_excel_out = private_dir / "output_reports"
    dir_img_out = DIR_PROJECT_IMAGES

    # Tập hợp các thư mục có thể chứa File 1
    scan_dirs_f1 = [dir_f1] if dir_f1.exists() else []
    scan_dirs_f1.append(private_dir)

    # Tập hợp các thư mục có thể chứa File 2
    scan_dirs_f2 = [dir_f2] if dir_f2.exists() else []
    scan_dirs_f2.append(private_dir)

    f1_candidates = []
    for sdir in scan_dirs_f1:
        if sdir.exists():
            for f in sdir.glob("*.xlsx"):
                if f.name.startswith("~$") or "bao_cao" in f.name.lower() or "report" in f.name.lower():
                    continue
                name_l = f.name.lower()
                if any(k in name_l for k in ["thanh toán hóa đơn", "hóa đơn", "crew", "bill"]):
                    f1_candidates.append(f)

    f2_candidates = []
    for sdir in scan_dirs_f2:
        if sdir.exists():
            for f in sdir.glob("*.xlsx"):
                if f.name.startswith("~$") or "bao_cao" in f.name.lower() or "report" in f.name.lower():
                    continue
                name_l = f.name.lower()
                if any(k in name_l for k in ["member", "share", "chi tiết", "payment", "knox", "boq"]):
                    f2_candidates.append(f)

    # Ưu tiên tìm File 1 theo target_month nếu có truyền
    selected_f1 = None
    if target_month:
        month_clean = target_month.replace("-", "").replace(".", "")
        for f in f1_candidates:
            if month_clean in f.name or target_month in f.name:
                selected_f1 = f
                break
    if not selected_f1 and f1_candidates:
        selected_f1 = f1_candidates[0]

    selected_f2 = f2_candidates[0] if f2_candidates else None

    return selected_f1, selected_f2, dir_excel_out, dir_img_out


