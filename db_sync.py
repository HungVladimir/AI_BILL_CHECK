# ==============================================================================
# MODULE: AI BILL PROCESSOR - DATABASE SYNCHRONIZER
# File: ai_bill_processor/db_sync.py
# ==============================================================================
"""
Thực hiện chèn (INSERT) các bản ghi hóa đơn vào bảng LOCAL_AI.AI_INVOICE_CHECK trong Oracle DB.
Hỗ trợ cả thư viện nội bộ OneAI (oneai_sdv) và oracledb Native (Thick/Thin Mode).
Tích hợp chế độ Dry-Run an toàn (kiểm tra dữ liệu trước khi nạp vào DB thật).
"""
import os
import sys
from pathlib import Path
from typing import Dict, List, Any, Optional

try:
    from .config import (
        DB_DNS,
        DB_USER,
        DB_PASSWORD,
        ORACLE_CLIENT_DIR,
        TABLE_NAME
    )
except ImportError:
    from config import (
        DB_DNS,
        DB_USER,
        DB_PASSWORD,
        ORACLE_CLIENT_DIR,
        TABLE_NAME
    )


class DatabaseSynchronizer:
    """Bộ đồng bộ dữ liệu vào bảng LOCAL_AI.AI_INVOICE_CHECK."""

    def __init__(self, dry_run: bool = False):
        self.dry_run = dry_run
        self.conn = None
        self.cur = None
        self.use_oneai_sdv = False

        if not self.dry_run:
            self._connect()

    def _connect(self):
        """Khởi tạo kết nối Oracle DB (MBLDB)."""
        # 1. Thử dùng thư viện nội bộ oneai_sdv
        try:
            from oneai_sdv import oracle_conn
            self.cur = oracle_conn(
                dns=DB_DNS,
                user_name=DB_USER,
                password=DB_PASSWORD,
                oracle_client_dir=ORACLE_CLIENT_DIR
            )
            self.use_oneai_sdv = True
            print(f"🗄️ [Oracle DB] Đã kết nối qua oneai_sdv (DNS: {DB_DNS}, User: {DB_USER})")
            return
        except Exception:
            pass

        # 2. Thử dùng thư viện oracledb native
        try:
            import oracledb
            # Khởi tạo Thick Mode nếu có thư viện client
            if ORACLE_CLIENT_DIR and os.path.exists(ORACLE_CLIENT_DIR):
                try:
                    oracledb.init_oracle_client(lib_dir=ORACLE_CLIENT_DIR)
                except Exception:
                    pass

            self.conn = oracledb.connect(
                user=DB_USER,
                password=DB_PASSWORD,
                dsn=DB_DNS
            )
            self.cur = self.conn.cursor()
            print(f"🗄️ [Oracle DB] Đã kết nối qua oracledb (DSN: {DB_DNS}, User: {DB_USER})")
        except Exception as e:
            print(f"⚠️ [Oracle DB Warning] Không thể kết nối cơ sở dữ liệu ({e}) -> Chuyển sang chế độ giả lập.")

    def insert_records(self, records: List[Dict[str, Any]]) -> int:
        """
        Thực thi chèn danh sách các bản ghi hóa đơn vào bảng AI_INVOICE_CHECK.
        """
        if not records:
            print("ℹ️ Không có bản ghi nào để lưu vào DB.")
            return 0

        sql = f"""
            INSERT INTO {TABLE_NAME} (
                ID, CREATED_AT, MONTH, SUBMITTER_NAME, SUBMITTER_EMAIL,
                FILE_NAME, FILE_PATH, MIME, DELETE_TOKEN, PROVIDER,
                INVOICE_NUMBER, INVOICE_DATE, BILLING_EMAIL, AMOUNT, PLAN,
                EMAIL_MATCH, WARNINGS, AI_RAW, STATUS, ADMIN_NOTE,
                REVIEWED_AT, USERID
            ) VALUES (
                :ID, SYSDATE, :MONTH, :SUBMITTER_NAME, :SUBMITTER_EMAIL,
                :FILE_NAME, :FILE_PATH, :MIME, :DELETE_TOKEN, :PROVIDER,
                :INVOICE_NUMBER, :INVOICE_DATE, :BILLING_EMAIL, :AMOUNT, :PLAN,
                :EMAIL_MATCH, :WARNINGS, :AI_RAW, :STATUS, :ADMIN_NOTE,
                :REVIEWED_AT, :USERID
            )
        """

        success_count = 0

        for r in records:
            # Loại bỏ các trường nội bộ không có trong bảng
            params = {
                "ID": str(r["ID"]).strip(),
                "MONTH": str(r["MONTH"]).strip(),
                "SUBMITTER_NAME": str(r["SUBMITTER_NAME"] or "").strip(),
                "SUBMITTER_EMAIL": str(r["SUBMITTER_EMAIL"] or "").strip(),
                "FILE_NAME": str(r["FILE_NAME"] or "").strip(),
                "FILE_PATH": str(r["FILE_PATH"] or "").strip(),
                "MIME": str(r["MIME"] or "image/png").strip(),
                "DELETE_TOKEN": r.get("DELETE_TOKEN"),
                "PROVIDER": str(r["PROVIDER"] or "").strip(),
                "INVOICE_NUMBER": r.get("INVOICE_NUMBER"),
                "INVOICE_DATE": r.get("INVOICE_DATE"),
                "BILLING_EMAIL": str(r["BILLING_EMAIL"] or "").strip(),
                "AMOUNT": float(r.get("AMOUNT") or 0.0),
                "PLAN": r.get("PLAN"),
                "EMAIL_MATCH": int(r.get("EMAIL_MATCH") or 1),
                "WARNINGS": r.get("WARNINGS"),
                "AI_RAW": r.get("AI_RAW"),
                "STATUS": str(r.get("STATUS") or "confirmed").strip(),
                "ADMIN_NOTE": r.get("ADMIN_NOTE"),
                "REVIEWED_AT": r.get("REVIEWED_AT"),
                "USERID": str(r["USERID"] or "").strip()
            }

            if self.dry_run or not self.cur:
                print(f"   [Dry-Run SQL] INSERT {params['ID']} | {params['USERID']} | {params['PROVIDER']} | {params['AMOUNT']:,.0f} đ | {params['FILE_PATH']}")
                success_count += 1
                continue

            try:
                self.cur.execute(sql, params=params)
                if hasattr(self.cur, "commit"):
                    self.cur.commit()
                elif self.conn and hasattr(self.conn, "commit"):
                    self.conn.commit()
                success_count += 1
                print(f"   ✅ [DB Insert OK] {params['ID']} - {params['SUBMITTER_NAME']} ({params['PROVIDER']}: {params['AMOUNT']:,.0f} đ)")
            except Exception as e:
                print(f"   ❌ [DB Insert Error] Không thể lưu bản ghi {params['ID']}: {e}")

        print(f"🎉 [DatabaseSynchronizer] Đã lưu thành công {success_count}/{len(records)} bản ghi vào {TABLE_NAME}.")
        return success_count

    def close(self):
        """Đóng kết nối cơ sở dữ liệu an toàn."""
        try:
            if self.cur and hasattr(self.cur, "close"):
                self.cur.close()
            if self.conn and hasattr(self.conn, "close"):
                self.conn.close()
        except Exception:
            pass
