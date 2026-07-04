"""
Cấu hình vị trí cột cho các file đầu vào của quy trình kiểm tra thuế.

QUY ƯỚC THAM CHIẾU CỘT:
  - Số nguyên (int)  -> vị trí cột theo THỨ TỰ, đếm từ 1 (cột 1 = cột đầu tiên).
                        Ví dụ: 42 = cột thứ 42 trong file TMS.
  - Chuỗi (str)      -> khớp theo TÊN tiêu đề cột (không phân biệt hoa/thường,
                        khớp gần đúng "chứa").

Người dùng có thể chỉnh lại toàn bộ vị trí này ngay trên giao diện web
nếu file thực tế có bố cục khác.
"""

# ---------------------------------------------------------------------------
# 1) FILE DỮ LIỆU CHÍNH (bảng kê / tờ khai hóa đơn mua vào cần kiểm tra)
# ---------------------------------------------------------------------------
MAIN = {
    "header_row": 1,  # dòng chứa tiêu đề (đếm từ 1)
    "cols": {
        "mst": "MST",                    # Mã số thuế người bán
        "invoice_no": "Số hóa đơn",      # Số hóa đơn
        "invoice_date": "Ngày hóa đơn",  # Ngày lập hóa đơn
        "pretax": "Giá trị chưa thuế",   # Giá trị hàng hóa/dịch vụ chưa thuế
        "vat": "Thuế GTGT",              # Thuế GTGT theo tờ khai
        "goods": "Tên hàng hóa",         # Tên hàng hóa / mặt hàng
    },
}

# ---------------------------------------------------------------------------
# 2) FILE EXPORT TỪ TMS (tra cứu trạng thái người nộp thuế)
# ---------------------------------------------------------------------------
TMS = {
    "header_row": 1,
    "mst_col": 1,          # cột chứa MST trong file TMS
    "status_col": 42,      # cột 42 - Trạng thái người nộp thuế
    "close_date_col": 51,  # cột 51 - Ngày đóng trạng thái tổ chức
}

# ---------------------------------------------------------------------------
# 3) FILE DANH SÁCH DOANH NGHIỆP CÓ DẤU HIỆU RỦI RO
# ---------------------------------------------------------------------------
RISK = {
    "header_row": 1,
    "mst_col": 1,   # cột chứa MST
    "doc_col": 5,   # cột 5 - Văn bản (số/ký hiệu văn bản cảnh báo rủi ro)
}

# ---------------------------------------------------------------------------
# 4) FILE HÓA ĐƠN ĐIỆN TỬ (tra trạng thái hóa đơn, tổng tiền thuế)
# ---------------------------------------------------------------------------
EINVOICE = {
    "header_row": 1,
    "cols": {
        "mst": "MST người bán",
        "invoice_no": "Số hóa đơn",
        "invoice_date": "Ngày hóa đơn",
        "status": "Trạng thái hóa đơn",
        "total_tax": "Tổng tiền thuế",
    },
}

# ---------------------------------------------------------------------------
# 5) TỪ KHÓA MẶT HÀNG "KHÔNG PHỤC VỤ SẢN XUẤT KINH DOANH"
# ---------------------------------------------------------------------------
GOODS_KEYWORDS = [
    "golf",
    "quà",
    "tặng",
    "biếu",
    "rượu",
    "wine",
    "bia",
    "beer",
    "thuốc lá",
    "cigar",
    "sân golf",
]

# ---------------------------------------------------------------------------
# 6) TỪ KHÓA NHẬN DIỆN TRẠNG THÁI HÓA ĐƠN (tra ở file HĐĐT)
# ---------------------------------------------------------------------------
INVOICE_STATUS = {
    # nhãn hiển thị -> danh sách từ khóa nhận diện trong cột trạng thái
    "Bị thay thế": ["thay thế", "bị thay thế"],
    "Bị điều chỉnh": ["điều chỉnh", "bị điều chỉnh"],
    "Bị xóa bỏ / hủy": ["xóa bỏ", "hủy", "huỷ", "xoá bỏ"],
}

# Nhãn dùng khi không tìm thấy hóa đơn trong file HĐĐT
STATUS_NOT_FOUND = "Không tìm thấy"
