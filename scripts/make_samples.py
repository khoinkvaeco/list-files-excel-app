"""
Tạo bộ file Excel MẪU để thử nghiệm công cụ (KHÔNG phải dữ liệu thật).

Bộ mẫu này mô phỏng đúng bố cục các file thực tế:
  - Bảng kê hóa đơn mua vào: tiêu đề ở dòng 9-10, dữ liệu từ dòng 11.
  - Danh sách DN rủi ro: 3 sheet khác cấu trúc.
  - File hóa đơn điện tử: sheet "DanhSach", tiêu đề dòng 1.

Chạy:  python scripts/make_samples.py   -> kết quả trong thư mục samples/
"""

import os

import openpyxl

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(HERE, "samples")
os.makedirs(OUT, exist_ok=True)


def _set(ws, row, col, value):
    ws.cell(row=row, column=col, value=value)


def make_main():
    """Bảng kê hóa đơn mua vào (mẫu). Cột theo vị trí thật."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BK MV- 2020-2025 (đã sửa)"

    # Dòng 1-8: khối tiêu đề biểu mẫu (để trống phần lớn)
    _set(ws, 1, 2, "BẢNG KÊ HOÁ ĐƠN, CHỨNG TỪ HÀNG HOÁ, DỊCH VỤ MUA VÀO")
    _set(ws, 4, 4, "CÔNG TY MẪU DEMO")
    _set(ws, 5, 4, "0304436870")

    # Dòng 9: nhãn cho các cột nằm ở hàng trên của tiêu đề gộp
    _set(ws, 9, 8, "Tên người bán")
    _set(ws, 9, 9, "Mã số thuế người bán")
    _set(ws, 9, 10, "Tên hàng hóa, dịch vụ")
    _set(ws, 9, 14, "Giá trị HHDV mua vào chưa có thuế GTGT")
    _set(ws, 9, 17, "Tiền thuế GTGT")

    # Dòng 10: nhãn hàng dưới của tiêu đề gộp
    _set(ws, 10, 2, "STT")
    _set(ws, 10, 3, "Mẫu số")
    _set(ws, 10, 4, "Ký hiệu")
    _set(ws, 10, 5, "Số")
    _set(ws, 10, 7, "Ngày, tháng, năm")

    # Dòng 11+: dữ liệu. (STT, MẫuSố, KýHiệu, Số HĐ, số bỏ 0, Ngày, Tên, MST, Hàng, ... , pretax(14), ..., vat(17))
    rows = [
        # stt, mau, kyhieu, soHD,      ngay,        ten,               mst,             hang,                       pretax,     vat
        (1, "1", "AA/20E", "0000123", "2020-10-23", "Ngân hàng ABC", "0301224067", "Phí ngân hàng", 525000, 0),
        (2, "1", "AA/20E", "0000123", "2020-10-23", "Ngân hàng ABC", "0301224067", "Phí ngân hàng", 525000, 0),  # trùng
        (3, "1", "ME/20E", "0000045", "2021-02-20", "Công ty Rượu XYZ", "0312816241", "Chai rượu vang biếu khách", 5000000, 500000),
        (4, "1", "ME/20E", "0000046", "2021-03-01", "Công ty Golf", "0313025323", "Bộ gậy chơi golf", 2000000, 200000),
        (5, "1", "HK/18P", "0000789", "2022-04-10", "Công ty Vật tư", "0400111222", "Thép hộp tăng cường (không trùng golf)", 8000000, 800000),
        (6, "1", "HK/18P", "0000790", "2022-06-15", "Công ty Quà", "0400111222", "Quà tặng khách hàng", 3000000, 300000),
        (7, "1", "CA/19P", "0000001", "2023-05-05", "Công ty VPP", "9999999999", "Văn phòng phẩm", 1000000, 100000),
    ]
    r = 11
    for stt, mau, kyhieu, sohd, ngay, ten, mst, hang, pretax, vat in rows:
        _set(ws, r, 2, stt)
        _set(ws, r, 3, mau)
        _set(ws, r, 4, kyhieu)
        _set(ws, r, 5, sohd)
        _set(ws, r, 6, sohd.lstrip("0"))
        _set(ws, r, 7, ngay)
        _set(ws, r, 8, ten)
        _set(ws, r, 9, mst)
        _set(ws, r, 10, hang)
        _set(ws, r, 14, pretax)
        _set(ws, r, 17, vat)
        r += 1

    wb.save(os.path.join(OUT, "01_du_lieu_chinh.xlsx"))


def make_risk():
    """Danh sách DN rủi ro — 3 sheet khác cấu trúc."""
    wb = openpyxl.Workbook()

    # Sheet 1
    ws1 = wb.active
    ws1.title = "Cơ quan điều tra - UPDATE"
    hdr = ["STT", "MÃ SỐ THUẾ", "TÊN CÔNG TY", "MÃ SỐ THUẾ", "ĐƠN VỊ QUẢN LÝ THUẾ",
           "Văn bản", "Ngày", "Cơ quan ban hành", "Vụ án liên quan", "Mức rủi ro"]
    for j, h in enumerate(hdr, start=1):
        _set(ws1, 1, j, h)
    data1 = [
        (1, "0312816241", "CÔNG TY RƯỢU XYZ", "1396/YC-ANĐT-P6"),
        (2, "0313025323", "CÔNG TY GOLF", "1396/YC-ANĐT-P6"),
    ]
    for i, (stt, mst, ten, vb) in enumerate(data1, start=2):
        _set(ws1, i, 1, stt); _set(ws1, i, 2, mst); _set(ws1, i, 3, ten); _set(ws1, i, 6, vb)

    # Sheet 2: tiêu đề ở dòng 3
    ws2 = wb.create_sheet("Công văn 4532")
    _set(ws2, 3, 1, "Tên người nộp thuế"); _set(ws2, 3, 2, "Mã số thuế")
    _set(ws2, 3, 3, "Cơ quan thuế"); _set(ws2, 3, 4, "Mức rủi ro")
    _set(ws2, 4, 1, "CÔNG TY VPP"); _set(ws2, 4, 2, "9999999999"); _set(ws2, 4, 4, 3)

    # Sheet 3: tiêu đề "Mã số thuế" ở dòng 4
    ws3 = wb.create_sheet("524 DN")
    _set(ws3, 2, 2, "524 DN RỦI RO THEO Công văn số 1798/TCT-TTKT ngày 16/5/2023")
    _set(ws3, 4, 1, "Mã số thuế")
    _set(ws3, 5, 1, "0400111222")

    wb.save(os.path.join(OUT, "03_ds_rui_ro.xlsx"))


def make_einvoice():
    """File hóa đơn điện tử hợp lệ (mẫu) — sheet DanhSach, tiêu đề dòng 1."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DanhSach"
    hdr = ["Mã số thuế người bán", "Số hóa đơn", "Ngày lập", "Trạng thái hóa đơn", "Tổng tiền thuế"]
    for j, h in enumerate(hdr, start=1):
        _set(ws, 1, j, h)
    rows = [
        ("0301224067", "123", "2020-10-23", "Đã cấp mã hóa đơn", 0),
        ("0312816241", "45", "2021-02-20", "Hóa đơn bị thay thế", 500000),
        ("0313025323", "46", "2021-03-01", "Đã cấp mã hóa đơn", 250000),  # lệch thuế 50k
        ("0400111222", "789", "2022-04-10", "Hóa đơn đã bị xóa bỏ", 800000),
        # HĐ 790 và 1 không có -> "không tìm thấy"
    ]
    for i, row in enumerate(rows, start=2):
        for j, v in enumerate(row, start=1):
            _set(ws, i, j, v)
    wb.save(os.path.join(OUT, "04_hoa_don_dien_tu.xlsx"))


if __name__ == "__main__":
    make_main()
    make_risk()
    make_einvoice()
    print("Đã tạo file mẫu trong:", OUT)
