# ==============================================================================
# MODULE: AI BILL PROCESSOR - MINIO OBJECT STORAGE UPLOADER
# File: ai_bill_processor/minio_client.py
# ==============================================================================
"""
Quản lý tải ảnh hóa đơn lên MinIO S3:
  - Đường dẫn đích: aibill/{YYYY-MM}/{ID}.png
  - Định dạng Content-Type: image/png
  - Tự động fallback lưu file cục bộ (Dry-Run / Debug) nếu MinIO offline hoặc chưa bật.
"""
import io
import os
from pathlib import Path
from typing import Dict, Any, Optional
from PIL import Image

try:
    from .config import (
        MINIO_ENDPOINT,
        MINIO_ACCESS_KEY,
        MINIO_SECRET_KEY,
        MINIO_BUCKET_NAME,
        BASE_DIR
    )
except ImportError:
    from config import (
        MINIO_ENDPOINT,
        MINIO_ACCESS_KEY,
        MINIO_SECRET_KEY,
        MINIO_BUCKET_NAME,
        BASE_DIR
    )


class MinioUploader:
    """Bộ điều phối upload ảnh hóa đơn lên MinIO."""

    def __init__(self, dry_run: bool = False):
        self.dry_run = dry_run
        self.bucket_name = MINIO_BUCKET_NAME or "oneai"
        self.client = None
        self._init_client()

    def _init_client(self):
        """Khởi tạo kết nối tới máy chủ MinIO."""
        if self.dry_run:
            print("📦 [MinIO] Chế độ Dry-Run: Sẽ không đẩy ảnh lên MinIO thật.")
            return

        # 1. Thử dùng minio_object của OneAI Platform nếu có
        try:
            from minio_object import minio_object as _oneai_obj
            if _oneai_obj and hasattr(_oneai_obj, "client"):
                self.client = _oneai_obj.client
                self.bucket_name = getattr(_oneai_obj, "bucket_name", self.bucket_name)
                print(f"📦 [MinIO] Đã tích hợp OneAI minio_object (Bucket: {self.bucket_name})")
                return
        except Exception:
            pass

        # 2. Khởi tạo trực tiếp từ biến môi trường
        try:
            from minio import Minio
            if MINIO_ENDPOINT and MINIO_ACCESS_KEY and MINIO_SECRET_KEY:
                endpoint_clean = MINIO_ENDPOINT.replace("http://", "").replace("https://", "").strip()
                self.client = Minio(
                    endpoint_clean,
                    access_key=MINIO_ACCESS_KEY,
                    secret_key=MINIO_SECRET_KEY,
                    secure=False
                )
                print(f"📦 [MinIO] Đã kết nối MinIO Server ({endpoint_clean}, Bucket: {self.bucket_name})")
            else:
                print("ℹ️ [MinIO] Chưa cấu hình đủ MINIO_ENDPOINT/ACCESS/SECRET -> Bật chế độ lưu đĩa cục bộ.")
        except Exception as e:
            print(f"⚠️ [MinIO Warning] Không thể kết nối MinIO ({e}) -> Sẽ lưu ảnh cục bộ.")

    def upload_image(self, object_key: str, raw_bytes: bytes, output_debug_dir: Optional[Path] = None) -> bool:
        """
        Upload ảnh lên MinIO:
          - object_key: 'aibill/2026-10/62a87c311d.png'
          - raw_bytes: byte stream ảnh gốc
        """
        # Đảm bảo dữ liệu là chuẩn PNG
        try:
            im = Image.open(io.BytesIO(raw_bytes))
            png_io = io.BytesIO()
            im.save(png_io, format="PNG")
            upload_data = png_io.getvalue()
        except Exception:
            upload_data = raw_bytes

        # Nếu có thư mục debug, lưu 1 bản sao cục bộ để kiểm tra
        if output_debug_dir:
            try:
                local_file = output_debug_dir / object_key
                local_file.parent.mkdir(parents=True, exist_ok=True)
                with open(local_file, "wb") as f:
                    f.write(upload_data)
            except Exception as e:
                print(f"⚠️ [Debug Save Warning] {e}")

        # Nếu là chế độ Dry-Run hoặc MinIO chưa kết nối, dừng ở đây
        if self.dry_run or not self.client:
            print(f"   [Dry-Run/Local] Đã chuẩn bị ảnh: {object_key} ({len(upload_data):,} bytes)")
            return True

        # Thực hiện upload thật lên MinIO
        try:
            # Đảm bảo bucket tồn tại
            if not self.client.bucket_exists(self.bucket_name):
                self.client.make_bucket(self.bucket_name)

            data_stream = io.BytesIO(upload_data)
            self.client.put_object(
                bucket_name=self.bucket_name,
                object_name=object_key,
                data=data_stream,
                length=len(upload_data),
                content_type="image/png"
            )
            print(f"   ☁️ [MinIO Upload OK] {object_key}")
            return True
        except Exception as e:
            print(f"❌ [MinIO Upload Error] Thất bại khi upload {object_key}: {e}")
            return False
