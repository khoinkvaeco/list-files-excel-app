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
# 1) FILE DỮ LIỆU CHÍNH (Bảng kê hóa đơn mua vào, mẫu 01-1/HT)
#    Dữ liệu thường nằm ở sheet "BKMV"; tiêu đề ở dòng 10-11, dữ liệu từ dòng 12.
#    Tham chiếu cột theo VỊ TRÍ vì bảng kê có tiêu đề gộp nhiều dòng và mỗi file
#    có thể thêm cột phụ khác nhau -> chọn sheet & chỉnh cột trên giao diện nếu cần.
# ---------------------------------------------------------------------------
MAIN = {
    "sheet": None,           # None -> tự đoán (ưu tiên sheet tên chứa "BK")
    "header_row": 11,
    "cols": {
        "mst": 12,           # [12] Mã số thuế người bán (bản giữ số 0 đầu)
        "invoice_no": 6,     # [6]  Số hóa đơn
        "invoice_date": 7,   # [7]  Ngày, tháng, năm lập hóa đơn
        "pretax": 17,        # [17] Giá trị HHDV mua vào chưa có thuế GTGT
        "vat": 19,           # [19] Tiền thuế GTGT
        "goods": 13,         # [13] Tên hàng hóa, dịch vụ
    },
}

# ---------------------------------------------------------------------------
# 2) FILE EXPORT TỪ TMS (Danh bạ người nộp thuế) — TÙY CHỌN
#    Định dạng SpreadsheetML (.xls là XML), sheet "NNT", tiêu đề dòng 2.
# ---------------------------------------------------------------------------
TMS = {
    "header_row": 2,
    "mst_col": 4,          # [4]  Mã số thuế
    "status_col": 45,      # [45] Trạng thái ĐKT tổ chức
    "close_date_col": 54,  # [54] Ngày đóng trạng thái tổ chức
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
    "header_row": 4,
    "cols": {
        "mst": 7,           # [7]  MST người bán/người xuất hàng
        "invoice_no": 4,    # [4]  Số hóa đơn
        "invoice_date": 5,  # [5]  Ngày lập
        "status": 19,       # [19] Trạng thái hóa đơn
        "total_tax": 13,    # [13] Tổng tiền thuế
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
