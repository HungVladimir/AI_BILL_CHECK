# ==============================================================================
# MODULE: AI BILL PROCESSOR - RAW IMAGE EXTRACTOR
# File: ai_bill_processor/excel_extractor.py
# ==============================================================================
"""
Trích xuất ảnh hóa đơn RAW nguyên bản từ File Excel 1.
Bảo toàn 100% độ sắc nét, kích thước pixel và bitrate gốc từ gói nén OpenXML.
Tự động ánh xạ tọa độ ô (Row, Col) sang Nhân sự (Họ và Tên) và Dịch vụ (ChatGPT, Claude, Gemini).
Tương thích với Header đơn tầng, Header đa tầng (Merged Header), và cơ chế Fallback tự động dò cột ảnh.
"""
import io
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
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


def get_cell_effective_value(ws, row: int, col: int) -> str:
    """
    Lấy giá trị của ô tại (row, col).
    Nếu ô nằm trong vùng Merged Cells và bị rỗng, kế thừa giá trị từ ô góc trên-trái.
    """
    val = ws.cell(row, col).value
    if val is not None and str(val).strip() != "":
        return normalize_text(val)

    # Kiểm tra các dải merge
    for mrange in ws.merged_cells.ranges:
        if mrange.min_row <= row <= mrange.max_row and mrange.min_col <= col <= mrange.max_col:
            top_val = ws.cell(mrange.min_row, mrange.min_col).value
            if top_val is not None:
                return normalize_text(top_val)
    return ""


def select_best_sheet(wb, file_path: Path):
    """
    Tự động chọn Sheet chính xác chứa dữ liệu hóa đơn trong File 1:
    1. ƯU TIÊN TUYỆT ĐỐI: Sheet có tên 'Tổng hợp hóa đơn' (hoặc 'tong hop hoa don').
    2. Sheet có tên chứa từ khóa 'tổng hợp' hoặc 'tong hop'.
    3. Sheet có tên chứa từ khóa: 'thanh toán', 'hóa đơn', 'crew', 'bill', hoặc tháng/năm.
    4. Sheet có nhiều dòng và có cột 'Họ và tên'.
    5. Sheet active mặc định.
    """
    # 1. Kiểm tra sheet chuẩn 'Tổng hợp hóa đơn' trước tiên
    for sname in wb.sheetnames:
        s_clean = normalize_text(sname).lower()
        if "tổng hợp hóa đơn" in s_clean or "tong hop hoa don" in s_clean:
            return wb[sname]

    for sname in wb.sheetnames:
        s_clean = normalize_text(sname).lower()
        if "tổng hợp" in s_clean or "tong hop" in s_clean:
            return wb[sname]

    sheet_scores = []
    clean_name = file_path.name.lower()

    for sname in wb.sheetnames:
        score = 0
        s_lower = sname.lower()

        # Điểm thưởng theo tên sheet
        if any(k in s_lower for k in ["thanh toán", "thanh toan", "hóa đơn", "hoa don", "crew", "aibill", "bill"]):
            score += 20
        # Nếu tên file có tháng (e.g. 2026.08, 08, 10) mà tên sheet cũng có
        match_num = re.findall(r'\d+', clean_name)
        for num in match_num:
            if len(num) >= 2 and num in s_lower:
                score += 10

        # Kiểm tra nội dung dòng đầu của sheet
        ws_temp = wb[sname]
        try:
            for r in range(1, min(6, ws_temp.max_row + 1)):
                row_vals = [normalize_text(ws_temp.cell(r, c).value).lower() for c in range(1, min(20, ws_temp.max_column + 1))]
                if any(k in v for v in row_vals for k in ["họ và tên", "ho va ten", "họ tên", "full name"]):
                    score += 30
                    break
        except Exception:
            pass

        sheet_scores.append((score, sname))

    # Sắp xếp theo điểm giảm dần
    sheet_scores.sort(key=lambda x: x[0], reverse=True)
    if sheet_scores and sheet_scores[0][0] > 0:
        best_sheet_name = sheet_scores[0][1]
        return wb[best_sheet_name]

    return wb.active


class ExcelImageExtractor:
    """Bộ giải mã trích xuất ảnh hóa đơn từ File Excel 1."""

    def __init__(self, excel_path: str):
        self.excel_path = Path(excel_path)
        if not self.excel_path.exists():
            raise FileNotFoundError(f"Không tìm thấy file Excel: {self.excel_path}")

    def extract_invoices(self) -> List[Dict[str, Any]]:
        """
        Trích xuất toàn bộ hóa đơn từ File 1.
        Tự động xử lý đa tầng Header, Merged Cells và liên kết ảnh OpenXML Drawing.
        """
        # 1. Đọc dữ liệu bảng tính
        wb = openpyxl.load_workbook(self.excel_path, data_only=True)
        ws = select_best_sheet(wb, self.excel_path)

        print(f"📄 [ExcelExtractor] Đang đọc file: {self.excel_path.name}")
        print(f"   - Danh sách sheet trong file: {wb.sheetnames}")
        print(f"   - Sheet được chọn: '{ws.title}' (Tổng {ws.max_row} dòng, {ws.max_column} cột)")

        # 2. Xây dựng bản đồ tọa độ ảnh từ OpenXML: (row_0indexed, col_0indexed) -> raw_bytes
        cell_images = self._extract_anchored_images_openxml()
        cols_with_images_0indexed = sorted(list(set(c for (r, c) in cell_images.keys())))
        cols_with_images_1indexed = [c + 1 for c in cols_with_images_0indexed]

        print(f"📸 [OpenXML Drawing] Tìm thấy {len(cell_images)} ảnh gốc trong OpenXML.")
        if cols_with_images_1indexed:
            print(f"   - Các cột có ảnh gắn neo (1-indexed): {cols_with_images_1indexed}")
        else:
            print("   ⚠️ Lưu ý: Không phát hiện thẻ DrawingML nào có ảnh trong OpenXML!")

        # 3. Quét đa tầng Header (Hỗ trợ Merged Header từ dòng 1 đến 5)
        start_header_row = 1
        end_header_row = 1

        # Dò tìm dòng bắt đầu Header (chứa 'STT', 'Họ và tên', 'Mã NV'...)
        for r in range(1, min(7, ws.max_row + 1)):
            row_vals = [normalize_text(get_cell_effective_value(ws, r, c)).lower() for c in range(1, min(25, ws.max_column + 1))]
            if any(k in v for v in row_vals for k in ["họ và tên", "ho va ten", "họ tên", "name", "stt", "họ và tên nhân sự"]):
                start_header_row = r
                break

        # Dò dòng kết thúc Header (quét tiếp 3 dòng xem có sub-header hoặc tên provider không)
        end_header_row = start_header_row
        for r in range(start_header_row + 1, min(start_header_row + 4, ws.max_row + 1)):
            row_vals = [normalize_text(ws.cell(r, c).value).lower() for c in range(1, min(25, ws.max_column + 1))]
            # Nếu dòng này chứa từ khóa AI provider hoặc các từ khóa phân loại hóa đơn
            has_ai_sub = any(
                any(kw in v for p in PROVIDERS_CONFIG.values() for kw in p["keywords_file1"])
                or any(k in v for k in ["chatgpt", "claude", "gemini", "gpt", "hóa đơn", "minh chứng", "bill", "ảnh"])
                for v in row_vals
            )
            if has_ai_sub:
                end_header_row = r

        # Gộp tiêu đề đa tầng cho từng cột từ start_header_row -> end_header_row
        col_combined_headers = {}
        for c in range(1, ws.max_column + 1):
            parts = []
            for r in range(start_header_row, end_header_row + 1):
                val = get_cell_effective_value(ws, r, c).lower()
                if val and val not in parts:
                    parts.append(val)
            col_combined_headers[c] = " ".join(parts).strip()

        # In thông tin các cột nhận diện được để dễ kiểm tra
        print(f"📋 [ExcelExtractor] Vùng Header nhận diện từ Dòng {start_header_row} đến Dòng {end_header_row}:")
        for c in range(1, min(15, ws.max_column + 1)):
            c_header = col_combined_headers.get(c, "")
            if c_header:
                print(f"   - Cột {c}: '{c_header}'")

        # 4. Xác định cột Nhân sự và cột Hóa đơn AI
        col_name_idx = None
        col_genid_idx = None
        col_team_idx = None
        provider_cols = {}  # {c_idx_1based: 'chatgpt' / 'claude' / 'gemini'}

        for c, text in col_combined_headers.items():
            if not col_name_idx and any(k in text for k in ["họ và tên", "ho va ten", "họ tên", "full name"]):
                col_name_idx = c
            elif not col_genid_idx and any(k in text for k in ["genid", "gen_id", "mã nv", "manv", "id"]):
                col_genid_idx = c
            elif not col_team_idx and any(k in text for k in ["team", "bộ phận", "phòng ban", "dept"]):
                col_team_idx = c

            # Dò tìm theo từ khóa của từng nhà cung cấp
            for p_key, p_cfg in PROVIDERS_CONFIG.items():
                if any(kw in text for kw in p_cfg["keywords_file1"]):
                    provider_cols[c] = p_key
                    break

        if not col_name_idx:
            # Fallback mặc định: cột 2 thường là Họ và tên
            col_name_idx = 2
            print(f"   ℹ️ Không thấy cột Họ và Tên rõ ràng, mặc định gán Cột {col_name_idx}")

        # 5. SMART FALLBACK NẾU HEADER KHÔNG CÓ TỪ KHÓA AI:
        # Nếu provider_cols bị rỗng NHƯNG OpenXML lại phát hiện các cột thực sự có ảnh
        if not provider_cols and cols_with_images_1indexed:
            print("⚠️ [Smart Fallback] Tiêu đề Header không chứa từ khóa AI chuẩn, nhưng OpenXML phát hiện ảnh!")
            # Duyệt qua các cột có ảnh để cố gắng ánh xạ
            available_providers = ["chatgpt", "claude", "gemini"]
            p_assign_idx = 0
            for c_img in cols_with_images_1indexed:
                header_text = col_combined_headers.get(c_img, "")
                matched_p = None
                for p_key, p_cfg in PROVIDERS_CONFIG.items():
                    if any(kw in header_text for kw in p_cfg["keywords_file1"]):
                        matched_p = p_key
                        break

                if not matched_p:
                    # Gán theo thứ tự các Provider phổ biến
                    if p_assign_idx < len(available_providers):
                        matched_p = available_providers[p_assign_idx]
                        p_assign_idx += 1

                if matched_p:
                    provider_cols[c_img] = matched_p
                    print(f"   👉 Tự động gán Cột {c_img} ('{header_text}') -> Nhà cung cấp: {matched_p.upper()}")

        print(f"🎯 [ExcelExtractor] Kết quả phân tích cột:")
        print(f"   - Cột Họ và Tên: Cột {col_name_idx}")
        print(f"   - Các cột Hóa đơn AI: {provider_cols}")

        if not provider_cols:
            print("❌ [ExcelExtractor Warning] Không tìm thấy bất kỳ cột hóa đơn AI nào trong sheet này!")
            wb.close()
            return []

        # 6. Duyệt qua từng dòng nhân sự và trích xuất hóa đơn
        results = []
        data_start_row = end_header_row + 1

        for r in range(data_start_row, ws.max_row + 1):
            name_val = get_cell_effective_value(ws, r, col_name_idx)
            submitter_name = normalize_text(name_val)
            if not submitter_name:
                continue

            # Bỏ qua nếu dòng này là dòng tổng cộng / footer
            if any(k in submitter_name.lower() for k in ["tổng", "total", "ghi chú", "note", "kết luận"]):
                continue

            gen_id = normalize_text(get_cell_effective_value(ws, r, col_genid_idx)) if col_genid_idx else ""
            team = normalize_text(get_cell_effective_value(ws, r, col_team_idx)) if col_team_idx else ""

            # Kiểm tra xem dòng này có ảnh ở các cột nhà cung cấp không
            for c_idx, p_key in provider_cols.items():
                p_cfg = PROVIDERS_CONFIG[p_key]
                # Tọa độ OpenXML là 0-indexed: Row r -> r - 1, Col c_idx -> c_idx - 1
                img_data = cell_images.get((r - 1, c_idx - 1))

                if img_data:
                    raw_bytes = img_data["bytes"]
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
                    print(f"   📸 [Trích xuất thành công] Dòng {r}: {submitter_name} | {p_cfg['name']} ({w}x{h} px)")

        wb.close()
        print(f"✅ [ExcelExtractor] Đã trích xuất thành công {len(results)} ảnh hóa đơn từ {self.excel_path.name}")
        return results

    def _extract_anchored_images_openxml(self) -> Dict[Tuple[int, int], Dict[str, Any]]:
        """
        Phân tích trực tiếp gói nén OpenXML để gắn kết chính xác 100% từng ảnh
        vào đúng tọa độ ô (row_0indexed, col_0indexed).
        Bảo toàn byte gốc không suy hao.
        """
        cell_images = {}
        try:
            with zipfile.ZipFile(self.excel_path, 'r') as zf:
                # Tìm tất cả các file drawing trong xl/drawings/
                drawing_files = [f for f in zf.namelist() if 'drawings/drawing' in f and f.endswith('.xml')]
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
                            clean_target = target.lstrip('/').replace('../', '')
                            if not clean_target.startswith('xl/'):
                                clean_target = 'xl/' + clean_target
                            rid_to_target[rid] = clean_target

                    # Đọc drawing.xml để lấy tọa độ from (col, row)
                    draw_xml = zf.read(drawing_path)
                    draw_tree = ET.fromstring(draw_xml)

                    ns = {
                        'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                    }

                    anchors = draw_tree.findall('.//xdr:twoCellAnchor', ns) + draw_tree.findall('.//xdr:oneCellAnchor', ns)
                    for anchor in anchors:
                        from_node = anchor.find('xdr:from', ns)
                        if from_node is None:
                            continue

                        col_node = from_node.find('xdr:col', ns)
                        row_node = from_node.find('xdr:row', ns)
                        if col_node is None or row_node is None or not col_node.text or not row_node.text:
                            continue

                        c_idx = int(col_node.text)
                        r_idx = int(row_node.text)

                        # Tìm thẻ blip
                        blip = anchor.find('.//a:blip', ns)
                        if blip is None:
                            continue

                        # Lấy thuộc tính embed id bất kể namespace
                        embed_id = None
                        for attr_k, attr_v in blip.attrib.items():
                            if attr_k.endswith('embed'):
                                embed_id = attr_v
                                break

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
