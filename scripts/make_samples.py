"""
Tạo bộ file Excel MẪU để thử nghiệm công cụ (không phải dữ liệu thật).

Chạy:  python scripts/make_samples.py
Kết quả nằm trong thư mục samples/.
"""

import os

import pandas as pd

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(HERE, "samples")
os.makedirs(OUT, exist_ok=True)


def make_main():
    rows = [
        # MST,        Số HĐ,   Ngày HĐ,     chưa thuế, GTGT,    hàng hóa
        ["0101234567", "0000123", "2026-01-15", 10_000_000, 1_000_000, "Dịch vụ tư vấn"],
        ["0101234567", "123", "2026-01-15", 10_000_000, 1_000_000, "Dịch vụ tư vấn (trùng)"],
        ["0312345678", "0045", "2026-02-20", 5_000_000, 500_000, "Chai rượu vang biếu khách"],
        ["0312345678", "0046", "2026-03-01", 2_000_000, 200_000, "Bộ gậy chơi golf"],
        ["0400111222", "0789", "2026-04-10", 8_000_000, 800_000, "Vật tư sản xuất"],
        ["0400111222", "0790", "2026-06-15", 3_000_000, 300_000, "Quà tặng khách hàng"],
        ["9999999999", "0001", "2026-05-05", 1_000_000, 100_000, "Văn phòng phẩm"],
    ]
    df = pd.DataFrame(
        rows,
        columns=["MST", "Số hóa đơn", "Ngày hóa đơn", "Giá trị chưa thuế", "Thuế GTGT", "Tên hàng hóa"],
    )
    df.to_excel(os.path.join(OUT, "01_du_lieu_chinh.xlsx"), index=False)


def make_tms():
    # file TMS phải có ít nhất 51 cột; nhồi cột trống để đúng vị trí 42, 51
    ncols = 55
    header = [f"Cột {i+1}" for i in range(ncols)]
    header[0] = "MST"
    header[41] = "Trạng thái người nộp thuế"  # cột 42 (index 41)
    header[50] = "Ngày đóng trạng thái tổ chức"  # cột 51 (index 50)

    def row(mst, status, close):
        r = [""] * ncols
        r[0] = mst
        r[41] = status
        r[50] = close
        return r

    rows = [
        row("0101234567", "Đang hoạt động", ""),
        row("0312345678", "Ngừng hoạt động nhưng chưa đóng MST", "2026-02-25"),
        row("0400111222", "Người nộp thuế ngừng hoạt động", "2026-05-01"),
    ]
    df = pd.DataFrame(rows, columns=header)
    df.to_excel(os.path.join(OUT, "02_tms.xlsx"), index=False)


def make_risk():
    ncols = 6
    header = [f"Cột {i+1}" for i in range(ncols)]
    header[0] = "MST"
    header[4] = "Văn bản"  # cột 5 (index 4)

    def row(mst, doc):
        r = [""] * ncols
        r[0] = mst
        r[4] = doc
        return r

    rows = [
        row("0312345678", "CV số 123/CT-TTKT ngày 01/03/2026"),
        row("9999999999", "TB số 456/TCT ngày 15/04/2026"),
    ]
    df = pd.DataFrame(rows, columns=header)
    df.to_excel(os.path.join(OUT, "03_ds_rui_ro.xlsx"), index=False)


def make_einvoice():
    rows = [
        ["0101234567", "123", "2026-01-15", "Đã cấp mã hóa đơn", 1_000_000],
        ["0312345678", "45", "2026-02-20", "Hóa đơn bị thay thế", 500_000],
        ["0312345678", "46", "2026-03-01", "Đã cấp mã hóa đơn", 250_000],  # lệch thuế
        ["0400111222", "789", "2026-04-10", "Hóa đơn đã bị xóa bỏ", 800_000],
        # HĐ 0790 và 0001 không có trong file HĐĐT -> "không tìm thấy"
    ]
    df = pd.DataFrame(
        rows,
        columns=["MST người bán", "Số hóa đơn", "Ngày hóa đơn", "Trạng thái hóa đơn", "Tổng tiền thuế"],
    )
    df.to_excel(os.path.join(OUT, "04_hoa_don_dien_tu.xlsx"), index=False)


if __name__ == "__main__":
    make_main()
    make_tms()
    make_risk()
    make_einvoice()
    print("Đã tạo file mẫu trong:", OUT)
