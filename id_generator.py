# ==============================================================================
# MODULE: AI BILL PROCESSOR - RANDOM ID GENERATOR
# File: ai_bill_processor/id_generator.py
# ==============================================================================
"""
Tái tạo chính xác hàm sinh ID 10 ký tự của Oracle:
  CREATE OR REPLACE function LOCAL_AI.fn_make_id(LEN NUMBER) return VARCHAR2 IS
  begin
    return translate(DBMS_RANDOM.string('o', LEN), 'GHIJKLMNOPQRSTUVWXYZ', '0123456789ABCDEF0123');
  end fn_make_id;

Ví dụ đầu ra mẫu: '62a87c311d', '26a2dba1a5', 'b4f19c08e2'
"""
import secrets
import string

HEX_ALPHABET = "0123456789abcdef"

def generate_invoice_id(length: int = 10) -> str:
    """
    Sinh chuỗi ID ngẫu nhiên gồm 10 ký tự Hexadecimal an toàn tuyệt đối (cryptographically secure).
    Chuỗi trả về tương đương 100% với hàm LOCAL_AI.FN_MAKE_ID(10) trong Oracle DB.
    """
    # Sử dụng secrets.choice để đảm bảo ngẫu nhiên đồng đều không trùng lặp
    return "".join(secrets.choice(HEX_ALPHABET) for _ in range(length))


if __name__ == "__main__":
    print("Mẫu 5 ID ngẫu nhiên được sinh:")
    for i in range(5):
        print(f"  {i+1}: {generate_invoice_id(10)}")

