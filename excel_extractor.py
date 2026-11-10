# ==============================================================================
# MODULE: AI BILL PROCESSOR - RAW IMAGE EXTRACTOR
# File: ai_bill_processor/excel_extractor.py
# ==============================================================================
"""
Trích xuất ảnh hóa đơn RAW nguyên bản từ File Excel 1.
Bảo toàn 100% độ sắc nét, kích thước pixel và bitrate gốc từ gói nén OpenXML.
Tự động ánh xạ tọa độ ô (Row, Col) sang Nhân sự (Họ và Tên) và Dịch vụ (ChatGPT, Claude, Gemini).
"""
import io
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Any, Optional
import openpyxl
from PIL import Image

try:
    from .config import PROVIDERS_CONFIG
except ImportError:
    from config import PROVIDERS_CONFIG


def normalize_text(text: Any) -> str:
    """Chuẩn hóa chuỗi văn bản (bỏ khoảng trắng thừa, chữ thường)."""
    if text is None:
        return ""
    s = str(text).strip()
    return re.sub(r'\s+', ' ', s)


class ExcelImageExtractor:
    """Bộ giải mã trích xuất ảnh hóa đơn từ File Excel 1."""

    def __init__(self, excel_path: str):
        self.excel_path = Path(excel_path)
        if not self.excel_path.exists():
            raise FileNotFoundError(f"Không tìm thấy file Excel: {self.excel_path}")

    def extract_invoices(self) -> List[Dict[str, Any]]:
        """
        Trích xuất toàn bộ hóa đơn từ File 1.
        Trả về danh sách các đối tượng hóa đơn:
        [
            {
                "row_idx": 2,
                "submitter_name": "Nguyễn Thị Kim Ngân",
                "gen_id": "24420206",
                "team": "DTM Team(SDV)",
                "provider": "ChatGPT",
                "file_name": "chatGpt.png",
                "raw_bytes": b'...',
                "image_format": "PNG",
                "width": 1170,
                "height": 2532
            },
            ...
        ]
        """
        # 1. Đọc dữ liệu bảng tính để xác định cột và danh sách nhân sự
        wb = openpyxl.load_workbook(self.excel_path, data_only=True)
        ws = wb.active

        # Dò tìm chỉ số cột theo tiêu đề Header (dòng 1 hoặc dòng đầu tiên có chữ)
        header_row = 1
        col_map = {}
        for r in range(1, min(10, ws.max_row + 1)):
            row_vals = [normalize_text(ws.cell(r, c).value).lower() for c in range(1, ws.max_column + 1)]
            if any("họ và tên" in v or "tên" in v for v in row_vals):
                header_row = r
                for c_idx, val in enumerate(row_vals, start=1):
                    col_map[c_idx] = val
                break

        col_name_idx = None
        col_genid_idx = None
        col_team_idx = None
        provider_cols = {}  # {c_idx: 'chatgpt' / 'claude' / 'gemini'}

        for c_idx, val in col_map.items():
            if any(k in val for k in ["họ và tên", "ho va ten", "họ tên"]):
                col_name_idx = c_idx
            elif any(k in val for k in ["genid", "mã nv", "manv"]):
                col_genid_idx = c_idx
            elif "team" in val:
                col_team_idx = c_idx

            # Dò cột nhà cung cấp
            for p_key, p_cfg in PROVIDERS_CONFIG.items():
                if any(kw in val for kw in p_cfg["keywords_file1"]):
                    provider_cols[c_idx] = p_key

        if not col_name_idx:
            # Fallback nếu không thấy header rõ: gán cột 2 là Họ và Tên theo chuẩn ảnh chụp
            col_name_idx = 2

        print(f"📋 [ExcelExtractor] Đã nhận diện Header tại dòng {header_row}:")
        print(f"   - Cột Họ và Tên: Cột {col_name_idx}")
        print(f"   - Các cột Hóa đơn AI nhận diện: {provider_cols}")

        # 2. Xây dựng bản đồ tọa độ ảnh: (row, col) -> raw_bytes
        cell_images = self._extract_anchored_images_openxml()

        results = []

        # 3. Duyệt qua từng dòng nhân sự và gom hóa đơn
        for r in range(header_row + 1, ws.max_row + 1):
            name_val = ws.cell(r, col_name_idx).value
            submitter_name = normalize_text(name_val)
            if not submitter_name:
                continue

            gen_id = normalize_text(ws.cell(r, col_genid_idx).value) if col_genid_idx else ""
            team = normalize_text(ws.cell(r, col_team_idx).value) if col_team_idx else ""

            # Kiểm tra xem dòng này có ảnh ở các cột nhà cung cấp không
            for c_idx, p_key in provider_cols.items():
                p_cfg = PROVIDERS_CONFIG[p_key]
                # Tọa độ trong OpenXML DrawingML luôn là 0-indexed:
                # Row 1-indexed r -> OpenXML row r - 1
                # Col 1-indexed c_idx -> OpenXML col c_idx - 1
                img_data = cell_images.get((r - 1, c_idx - 1))

                if img_data:
                    raw_bytes = img_data["bytes"]
                    # Kiểm tra kích thước và định dạng ảnh
                    try:
                        im = Image.open(io.BytesIO(raw_bytes))
                        img_format = im.format or "PNG"
                        w, h = im.size
                    except Exception:
                        img_format = "PNG"
                        w, h = 0, 0

                    results.append({
                        "row_idx": r,
                        "submitter_name": submitter_name,
                        "gen_id": gen_id,
                        "team": team,
                        "provider_key": p_key,
                        "provider_name": p_cfg["name"],
                        "plan": p_cfg["plan"],
                        "file_name": p_cfg["default_filename"],
                        "raw_bytes": raw_bytes,
                        "image_format": img_format,
                        "width": w,
                        "height": h
                    })
                    print(f"   📸 Tìm thấy hóa đơn {p_cfg['name']} của: {submitter_name} (Kích thước: {w}x{h} px)")

        wb.close()
        print(f"✅ [ExcelExtractor] Đã trích xuất thành công {len(results)} ảnh hóa đơn từ {self.excel_path.name}")
        return results

    def _extract_anchored_images_openxml(self) -> Dict[tuple, Dict[str, Any]]:
        """
        Phân tích trực tiếp gói nén OpenXML để gắn kết chính xác 100% từng ảnh
        vào đúng tọa độ ô (row_idx, col_idx).
        Bảo toàn byte gốc không suy hao.
        """
        cell_images = {}
        try:
            with zipfile.ZipFile(self.excel_path, 'r') as zf:
                # Tìm file drawing (thường là xl/drawings/drawing1.xml)
                drawing_files = [f for f in zf.namelist() if f.startswith('xl/drawings/drawing') and f.endswith('.xml')]
                if not drawing_files:
                    return cell_images

                for drawing_path in drawing_files:
                    rels_path = drawing_path.replace('xl/drawings/', 'xl/drawings/_rels/') + '.rels'
                    if rels_path not in zf.namelist():
                        continue

                    # Đọc file rels để map rId -> target file trong media
                    rels_xml = zf.read(rels_path)
                    rels_tree = ET.fromstring(rels_xml)
                    rid_to_target = {}
                    for rel in rels_tree:
                        rid = rel.attrib.get('Id')
                        target = rel.attrib.get('Target')
                        if rid and target:
                            # Chuẩn hóa đường dẫn media: xử lý cả '../media/image1.png' và '/xl/media/image1.png'
                            clean_target = target.lstrip('/').replace('../', '')
                            if not clean_target.startswith('xl/'):
                                clean_target = 'xl/' + clean_target
                            rid_to_target[rid] = clean_target

                    # Đọc drawing.xml để lấy tọa độ from (col, row)
                    draw_xml = zf.read(drawing_path)
                    draw_tree = ET.fromstring(draw_xml)

                    # Namespaces của OpenXML Drawing
                    ns = {
                        'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                    }

                    # Quét cả twoCellAnchor và oneCellAnchor
                    for anchor in draw_tree.findall('.//xdr:twoCellAnchor', ns) + draw_tree.findall('.//xdr:oneCellAnchor', ns):
                        from_node = anchor.find('xdr:from', ns)
                        if from_node is None:
                            continue

                        col_node = from_node.find('xdr:col', ns)
                        row_node = from_node.find('xdr:row', ns)
                        if col_node is None or row_node is None:
                            continue

                        c_idx = int(col_node.text)
                        r_idx = int(row_node.text)

                        # Tìm thẻ blip chứa liên kết r:embed
                        blip = anchor.find('.//a:blip', ns)
                        if blip is None:
                            continue

                        embed_id = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if not embed_id or embed_id not in rid_to_target:
                            continue

                        media_path = rid_to_target[embed_id]
                        if media_path in zf.namelist():
                            raw_data = zf.read(media_path)
                            cell_images[(r_idx, c_idx)] = {
                                "bytes": raw_data,
                                "media_path": media_path
                            }
        except Exception as e:
            print(f"⚠️ [ExcelExtractor Warning] Lỗi đọc OpenXML Drawing: {e}")

        return cell_images
