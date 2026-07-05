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

# ---------------------------------------------------------------------------
# 7) CHÍNH SÁCH GIẢM THUẾ GTGT 10% -> 8%
# ---------------------------------------------------------------------------
# Các giai đoạn được giảm còn 8% (đã gộp khoảng liền kề). Ngoài các khoảng này
# thuế suất phổ thông là 10%.
VAT_REDUCED_PERIODS = [
    ("2022-02-01", "2022-12-31"),  # NĐ 15/2022
    ("2023-07-01", "2023-12-31"),  # NĐ 44/2023
    ("2024-01-01", "2024-12-31"),  # NĐ 94/2023 + NĐ 72/2024
    ("2025-07-01", "2026-12-31"),  # NĐ 174/2024 (174/2025)
]

# Nhóm hàng hóa/dịch vụ KHÔNG được giảm thuế (vẫn 10%) theo Phụ lục I, II, III
# của NĐ 15/2022 (được các NĐ sau dùng lại). Đây là bộ TỪ KHÓA nhận diện gần
# đúng trên mô tả hàng hóa (giữ nguyên dấu, khớp theo ranh giới từ) — mang tính
# TRỢ GIÚP, cần người kiểm tra xác nhận theo mã HS/ngành khi cần.
VAT_EXCLUDE_KEYWORDS = {
    "PL I - Viễn thông": ["viễn thông", "cước viễn thông"],
    "PL I - Tài chính, ngân hàng": ["ngân hàng", "tín dụng", "cho vay", "lãi vay", "dịch vụ tài chính"],
    "PL I - Chứng khoán": ["chứng khoán", "môi giới chứng khoán", "trái phiếu", "cổ phiếu"],
    "PL I - Bảo hiểm": ["bảo hiểm"],
    "PL I - Bất động sản": ["bất động sản", "quyền sử dụng đất", "chuyển nhượng đất"],
    "PL I - Kim loại": ["kim loại", "sắt", "thép", "inox", "nhôm", "gang", "tôn"],
    "PL I - Khai khoáng": ["khoáng sản", "quặng", "khai khoáng"],
    "PL I - Than cốc, dầu tinh chế": ["than cốc", "dầu mỏ", "dầu diesel", "dầu do", "dầu fo"],
    "PL I - Hóa chất": ["hóa chất", "hoá chất"],
    "PL II - Thuốc lá": ["thuốc lá", "xì gà", "cigar"],
    "PL II - Rượu, bia": ["rượu", "bia", "wine", "beer"],
    "PL II - Ô tô, xe máy": ["ô tô", "ôtô", "xe ô tô", "mô tô", "xe máy phân khối lớn"],
    "PL II - Tàu bay, du thuyền": ["tàu bay", "máy bay", "du thuyền"],
    "PL II - Xăng": ["xăng"],
    "PL II - Vàng mã, bài lá": ["vàng mã", "hàng mã", "bài lá"],
    "PL II - Dịch vụ TTĐB": ["vũ trường", "massage", "karaoke", "casino", "đặt cược", "xổ số", "golf", "sân golf", "gôn"],
    "PL III - Công nghệ thông tin": [
        "công nghệ thông tin", "máy vi tính", "máy tính", "laptop", "máy in",
        "linh kiện điện tử", "bán dẫn", "mạch điện tử", "thẻ thông minh", "thiết bị ngoại vi",
    ],
}
