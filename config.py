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

# --- 1. CẤU HÌNH NHÀ CUNG CẤP AI (CHATGPT, CLAUDE, GEMINI) ---
PROVIDERS_CONFIG = {
    "chatgpt": {
        "name": "ChatGPT",
        "keywords_file1": ["chatgpt", "hóa đơn chatgpt", "hoa don chatgpt", "gpt"],
        "keywords_file2": ["chatgpt plus", "chatgpt", "gpt"],
        "default_filename": "chatGpt.png",
        "plan": "ChatGPT Plus"
    },
    "claude": {
        "name": "Claude",
        "keywords_file1": ["claude", "hóa đơn claude", "hoa don claude"],
        "keywords_file2": ["claude pro", "claude"],
        "default_filename": "claude.png",
        "plan": "Claude Pro"
    },
    "gemini": {
        "name": "Gemini",
        "keywords_file1": ["gemini", "hóa đơn gemini", "hoa don gemini", "google ai"],
        "keywords_file2": ["gemini pro", "gemini", "google ai pro (5 tb)"],
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

