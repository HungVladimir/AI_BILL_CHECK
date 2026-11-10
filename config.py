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

# --- 0. CẤU HÌNH ĐƯỜNG DẪN DỮ LIỆU PRIVATE (LOCAL & TEST) ---
# Thư mục gốc chứa dữ liệu thực tế trên máy cá nhân của bạn
AI_BILL_PRIVATE_DIR = Path(
    os.getenv("AI_BILL_PRIVATE_DIR")
    or os.getenv("AI_BILL_DIR")
    or r"D:\Private_Documents\Doc\AI-Bill"
)

# Cấu trúc thư mục con phân loại rõ ràng (tùy chọn hoặc tự động nhận diện):
DIR_EXCEL1 = AI_BILL_PRIVATE_DIR / "invoices_file1"       # File 1: Chứa ảnh hóa đơn các tháng
DIR_EXCEL2 = AI_BILL_PRIVATE_DIR / "members_file2"        # File 2: Chứa danh sách member share, đơn giá
DIR_OUTPUT = AI_BILL_PRIVATE_DIR / "output_reports"       # Thư mục xuất file Excel báo cáo & ảnh bóc tách

# --- 1. CẤU HÌNH NHÀ CUNG CẤP AI (CHATGPT, CLAUDE, GEMINI) ---
PROVIDERS_CONFIG = {
    "chatgpt": {
        "name": "ChatGPT",
        "keywords_file1": [
            "chatgpt", "chat gpt", "chat-gpt", "openai", "gpt", "gpt-4", "gpt4",
            "hóa đơn chatgpt", "hoa don chatgpt", "hóa đơn chat gpt", "hoa don chat gpt",
            "hóa đơn openai", "hoa don openai", "ảnh chatgpt", "minh chứng chatgpt",
            "bill chatgpt", "receipt chatgpt", "chatgpt plus", "chatgpt 4"
        ],
        "keywords_file2": [
            "chatgpt plus", "chatgpt", "chat gpt", "chat-gpt", "gpt", "openai"
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
    '(AI Crew) thanh toán hóa đơn tháng 2026.10.xlsx' -> '2026-10'
    '(AI Crew) thanh toán hóa đơn tháng 2026.09.xlsx' -> '2026-09'
    '(AI Crew) thanh toán hóa đơn tháng 2026-08.xlsx' -> '2026-08'
    """
    clean_name = Path(filename).name
    # Tìm mẫu 4 số năm . hoặc - 2 số tháng
    match = re.search(r'(\d{4})[._](\d{1,2})', clean_name)
    if match:
        year = match.group(1)
        month = match.group(2).zfill(2)
        return f"{year}-{month}"
    
    # Dự phòng: tìm năm và tháng cách nhau bởi chữ
    match_fallback = re.search(r'tháng\s*(\d{1,2})[._\-\s]+(\d{4})', clean_name, re.IGNORECASE)
    if match_fallback:
        month = match_fallback.group(1).zfill(2)
        year = match_fallback.group(2)
        return f"{year}-{month}"

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
    Trả về: (file1_path, file2_path, output_reports_dir)
    """
    private_dir = Path(private_dir)
    dir_f1 = private_dir / "invoices_file1"
    dir_f2 = private_dir / "members_file2"
    dir_out = private_dir / "output_reports"

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

    return selected_f1, selected_f2, dir_out


