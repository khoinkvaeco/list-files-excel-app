"""
Cấu hình vị trí cột cho các file đầu vào của quy trình kiểm tra thuế.

QUY ƯỚC THAM CHIẾU CỘT:
  - Số nguyên (int)  -> vị trí cột theo THỨ TỰ, đếm từ 1 (cột 1 = cột đầu tiên).
  - Chuỗi (str)      -> khớp theo TÊN tiêu đề cột (không phân biệt hoa/thường,
                        khớp gần đúng "chứa").

Cấu hình mặc định dưới đây đã được căn theo bộ file mẫu thực tế:
  - Bảng kê hóa đơn mua vào (mẫu 01-1/HT).
  - Danh sách DN có dấu hiệu rủi ro (3 sheet).
  - File hóa đơn điện tử (export từ hoadondientu.gdt.gov.vn).
Người dùng vẫn có thể chỉnh lại trên giao diện web nếu file khác bố cục.
"""

# ---------------------------------------------------------------------------
# 1) FILE DỮ LIỆU CHÍNH (Bảng kê hóa đơn mua vào)
#    Tiêu đề nằm ở dòng 9-10, dữ liệu bắt đầu từ dòng 11 -> header_row = 10.
#    Tham chiếu cột theo VỊ TRÍ (số thứ tự) vì bảng kê có tiêu đề gộp nhiều dòng.
# ---------------------------------------------------------------------------
MAIN = {
    "header_row": 10,
    "cols": {
        "mst": 9,            # [9]  Mã số thuế người bán
        "invoice_no": 5,     # [5]  Số hóa đơn (bản có số 0 đầu, vd 0000611)
        "invoice_date": 7,   # [7]  Ngày, tháng, năm lập hóa đơn
        "pretax": 14,        # [14] Giá trị HHDV mua vào chưa có thuế GTGT
        "vat": 17,           # [17] Tiền thuế GTGT
        "goods": 10,         # [10] Tên hàng hóa, dịch vụ
    },
}

# ---------------------------------------------------------------------------
# 2) FILE EXPORT TỪ TMS (tra cứu trạng thái người nộp thuế) — TÙY CHỌN
#    Chưa có file mẫu; giữ vị trí theo mô tả (cột 42, 51). Điều chỉnh khi có file.
# ---------------------------------------------------------------------------
TMS = {
    "header_row": 1,
    "mst_col": 1,          # cột chứa MST trong file TMS
    "status_col": 42,      # cột 42 - Trạng thái người nộp thuế
    "close_date_col": 51,  # cột 51 - Ngày đóng trạng thái tổ chức
}

# ---------------------------------------------------------------------------
# 3) FILE DANH SÁCH DOANH NGHIỆP CÓ DẤU HIỆU RỦI RO
#    Gồm nhiều sheet, mỗi sheet một bố cục -> khai báo từng sheet.
#    doc_col:  int  -> cột chứa số/ký hiệu văn bản (theo từng dòng)
#              str  -> nhãn văn bản cố định gán cho mọi dòng của sheet đó
# ---------------------------------------------------------------------------
RISK = {
    "sheets": [
        {
            "name": "Cơ quan điều tra - UPDATE",
            "header_row": 1,
            "mst_col": 2,
            "doc_col": 6,  # cột "Văn bản"
        },
        {
            "name": "Công văn 4532",
            "header_row": 3,
            "mst_col": 2,
            "doc_col": "Công văn 4532",
        },
        {
            "name": "524 DN",
            "header_row": 4,
            "mst_col": 1,
            "doc_col": "Công văn 1798/TCT-TTKT ngày 16/5/2023",
        },
    ],
}

# ---------------------------------------------------------------------------
# 4) FILE HÓA ĐƠN ĐIỆN TỬ (tra trạng thái hóa đơn, tổng tiền thuế)
#    Khớp theo TÊN cột (export chuẩn của hoadondientu.gdt.gov.vn).
#    Sẽ tinh chỉnh khi có file export hợp lệ (file mẫu hiện tại rỗng).
# ---------------------------------------------------------------------------
EINVOICE = {
    "header_row": 1,
    "cols": {
        "mst": "Mã số thuế người bán",
        "invoice_no": "Số hóa đơn",
        "invoice_date": "Ngày lập",
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
    "Bị thay thế": ["thay thế", "bị thay thế"],
    "Bị điều chỉnh": ["điều chỉnh", "bị điều chỉnh"],
    "Bị xóa bỏ / hủy": ["xóa bỏ", "hủy", "huỷ", "xoá bỏ"],
}

# Nhãn dùng khi không tìm thấy hóa đơn trong file HĐĐT
STATUS_NOT_FOUND = "Không tìm thấy"
