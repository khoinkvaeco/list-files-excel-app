"""
Chuyển XML TỜ KHAI THUẾ (HTKK, vd 05/QTT-TNCN) sang Excel nhiều sheet.

Cách làm (theo bộ công cụ chuẩn):
  - Mỗi nhóm thẻ lặp lại (>=2 lần, có cấu trúc con) -> 1 sheet bảng kê, tên sheet
    theo thẻ cha (vd PLuc_05_1_BK_QTT), tiêu đề tiếng Việt + dòng mã cột gốc.
  - Các giá trị lá không lặp -> sheet "ThongTinChung" (Nhóm | Chỉ tiêu | Giá trị).
  - Nếu có các chỉ tiêu ct16..ct41 -> dựng sheet "ToKhaiChinh" đúng mẫu 05/QTT-TNCN.
  - Giữ dạng chữ cho mã có số 0 đầu / số quá dài (MST, CCCD); còn lại thành số.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET

import pandas as pd

# ---------------------------------------------------------------------------
# Mẫu 05/QTT-TNCN
# ---------------------------------------------------------------------------
FORM05_TITLE = "Mẫu số: 05/QTT-TNCN"
FORM05_SECTIONS = [
    ("I. NGHĨA VỤ KHẤU TRỪ THUẾ CỦA TỔ CHỨC, CÁ NHÂN TRẢ THU NHẬP", [
        ("1", "Tổng số người lao động", "ct16", "Người"),
        ("", "Trong đó: Cá nhân cư trú có hợp đồng lao động", "ct17", "Người"),
        ("2", "Tổng số cá nhân đã khấu trừ thuế [18]=[19]+[20]", "ct18", "Người"),
        ("2.1", "Cá nhân cư trú", "ct19", "Người"),
        ("2.2", "Cá nhân không cư trú", "ct20", "Người"),
        ("3", "Tổng số cá nhân thuộc diện được miễn, giảm thuế theo Hiệp định tránh đánh thuế hai lần", "ct21", "Người"),
        ("4", "Tổng số cá nhân giảm trừ gia cảnh", "ct22", "Người"),
        ("5", "Tổng thu nhập chịu thuế trả cho cá nhân [23]=[24]+[25]", "ct23", "VNĐ"),
        ("5.1", "Cá nhân cư trú", "ct24", "VNĐ"),
        ("5.2", "Cá nhân không cư trú", "ct25", "VNĐ"),
        ("5.3", "Trong đó: Tổng thu nhập chịu thuế từ tiền phí mua bảo hiểm nhân thọ, bảo hiểm không bắt buộc khác của DN bảo hiểm không thành lập tại Việt Nam cho người lao động", "ct26", "VNĐ"),
        ("6", "Trong đó tổng thu nhập chịu thuế được miễn theo quy định của Hợp đồng dầu khí", "ct27", "VNĐ"),
        ("7", "Tổng thu nhập chịu thuế trả cho cá nhân thuộc diện phải khấu trừ thuế [28]=[29]+[30]", "ct28", "VNĐ"),
        ("7.1", "Cá nhân cư trú", "ct29", "VNĐ"),
        ("7.2", "Cá nhân không cư trú", "ct30", "VNĐ"),
        ("8", "Tổng số thuế thu nhập cá nhân đã khấu trừ [31]=[32]+[33]", "ct31", "VNĐ"),
        ("8.1", "Cá nhân cư trú", "ct32", "VNĐ"),
        ("8.2", "Cá nhân không cư trú", "ct33", "VNĐ"),
        ("8.3", "Trong đó: Tổng số thuế TNCN đã khấu trừ trên tiền phí mua bảo hiểm nhân thọ, bảo hiểm không bắt buộc khác của DN bảo hiểm không thành lập tại Việt Nam cho người lao động", "ct34", "VNĐ"),
    ]),
    ("II. NGHĨA VỤ QUYẾT TOÁN THAY CHO CÁ NHÂN", [
        ("1", "Tổng số cá nhân uỷ quyền cho tổ chức, cá nhân trả thu nhập quyết toán thay", "ct35", "Người"),
        ("2", "Tổng số thuế thu nhập cá nhân đã khấu trừ", "ct36", "VNĐ"),
        ("", "Trong đó: Số thuế TNCN đã khấu trừ tại tổ chức trước khi điều chuyển (trường hợp có đánh dấu vào chỉ tiêu [04])", "ct37", "VNĐ"),
        ("3", "Tổng số thuế thu nhập cá nhân phải nộp", "ct38", "VNĐ"),
        ("4", "Tổng số thuế TNCN được miễn do cá nhân có số thuế còn phải nộp sau ủy quyền quyết toán từ 50.000 đồng trở xuống", "ct39", "VNĐ"),
        ("5", "Tổng số thuế thu nhập cá nhân còn phải nộp [40]=([38]–[36]–[39])>0", "ct40", "VNĐ"),
        ("6", "Tổng số thuế thu nhập cá nhân đã nộp thừa [41]=([38]–[36]–[39])<0", "ct41", "VNĐ"),
    ]),
]
FORM_CT = {r[2] for _, rows in FORM05_SECTIONS for r in rows}

FIELD_LABELS = {
    "maDVu": "Mã dịch vụ", "tenDVu": "Tên dịch vụ", "pbanDVu": "Phiên bản",
    "maTKhai": "Mã tờ khai", "tenTKhai": "Tên tờ khai", "moTaBMau": "Mô tả biểu mẫu",
    "pbanTKhaiXML": "Phiên bản tờ khai", "loaiTKhai": "Loại tờ khai", "soLan": "Số lần",
    "kieuKy": "Kiểu kỳ", "kyKKhai": "Kỳ kê khai",
    "kyKKhaiTuThang": "Kỳ từ tháng", "kyKKhaiDenThang": "Kỳ đến tháng",
    "maCQTNoiNop": "Mã CQT nơi nộp", "tenCQTNoiNop": "Tên CQT nơi nộp",
    "ngayLapTKhai": "Ngày lập tờ khai", "ngayKy": "Ngày ký",
    "mst": "Mã số thuế", "tenNNT": "Tên người nộp thuế", "dchiNNT": "Địa chỉ",
    "maHuyenNNT": "Mã quận/huyện", "tenHuyenNNT": "Quận/Huyện",
    "maTinhNNT": "Mã tỉnh", "tenTinhNNT": "Tỉnh/Thành phố",
    "dthoaiNNT": "Điện thoại", "faxNNT": "Fax", "emailNNT": "Email",
    "toChucCoQTTTheoUyQuyen": "Tổ chức QTT theo uỷ quyền [04]",
    "ma_THQuyetToan": "Mã trường hợp quyết toán", "ten_THQuyetToan": "Trường hợp quyết toán",
}

COL_LABELS_051 = {
    "coDieuChinhSoLieu": "Có điều chỉnh số liệu",
    "ct07": "Họ và tên", "ct08": "Mã số thuế",
    "ct09a_ma": "Mã loại giấy tờ", "ct09a_ten": "Loại giấy tờ",
    "ct09": "Số CMND/CCCD/Hộ chiếu",
    "ct10": "Cá nhân uỷ quyền QT thay", "ct11": "CN nước ngoài uỷ quyền QT <12 tháng",
    "ct12": "Thu nhập chịu thuế",
    "ct13": "TNCT trước điều chuyển", "ct14": "TNCT miễn theo Hiệp định",
    "ct15": "TNCT miễn theo HĐ dầu khí", "ct16": "Số lượng NPT giảm trừ",
    "ct17": "Tổng giảm trừ gia cảnh", "ct18": "Từ thiện, nhân đạo, khuyến học",
    "ct19": "Bảo hiểm được trừ", "ct20": "Quỹ hưu trí tự nguyện được trừ",
    "ct21": "Thu nhập tính thuế", "ct22": "Số thuế TNCN đã khấu trừ",
    "ct23": "Số thuế đã khấu trừ trước điều chuyển", "ct24": "Tổng số thuế phải nộp",
    "ct25": "Số thuế đã nộp thừa", "ct26": "Số thuế còn phải nộp",
    "ct27": "Được miễn thuế còn phải nộp ≤50.000đ",
}
COL_LABELS_052 = {
    "coDieuChinhSoLieu": "Có điều chỉnh số liệu",
    "ct07": "Họ và tên", "ct08": "Mã số thuế",
    "ct09a_ma": "Mã loại giấy tờ", "ct09a_ten": "Loại giấy tờ",
    "ct09": "Số CMND/CCCD/Hộ chiếu",
    "ct10": "Cá nhân không cư trú", "ct11": "Thu nhập chịu thuế",
    "ct12": "TNCT được miễn theo Hiệp định", "ct13": "TNCT được miễn theo HĐ dầu khí",
    "ct14": "Thu nhập tính thuế", "ct15": "Số thuế TNCN đã khấu trừ",
    "ct16": "Được miễn thuế còn phải nộp ≤50.000đ",
}
COL_LABELS_053 = {
    "ct07": "Họ và tên NNT", "ct08": "Mã số thuế NNT",
    "ct09": "Họ và tên người phụ thuộc", "ct10": "Ngày sinh NPT",
    "ct11": "Mã số thuế NPT", "ct12_ma": "Mã quan hệ", "ct12_ten": "Quan hệ với NNT",
    "ct13": "Số CMND/CCCD/Hộ chiếu NPT",
    "ct14_ma": "Mã loại giấy tờ", "ct14_ten": "Loại giấy tờ NPT",
    "ct15": "Thời điểm bắt đầu tính giảm trừ", "ct16": "Thời điểm kết thúc tính giảm trừ",
}


def _col_label_map(title: str):
    if "05_1" in title:
        return COL_LABELS_051
    if "05_2" in title:
        return COL_LABELS_052
    if "05_3" in title:
        return COL_LABELS_053
    return None


# ---------------------------------------------------------------------------
# Duyệt cây XML
# ---------------------------------------------------------------------------
def _local(el) -> str:
    return el.tag.rsplit("}", 1)[-1] if "}" in el.tag else el.tag


def _repeated_child_name(el):
    """Thẻ con lặp >=2 lần và có cấu trúc con -> tên 'dòng' của bảng kê."""
    counts = {}
    for k in el:
        if len(k) > 0:
            counts[_local(k)] = counts.get(_local(k), 0) + 1
    best, mx = None, 0
    for n, c in counts.items():
        if c > mx:
            best, mx = n, c
    return best if mx >= 2 else None


def _collect_tables(el, tables):
    name = _repeated_child_name(el)
    if name:
        tables.append({"title": _local(el), "rows": [c for c in el if _local(c) == name]})
        return
    for c in el:
        _collect_tables(c, tables)


def _leaf_values(el, prefix, out):
    kids = list(el)
    if not kids:
        v = (el.text or "").strip()
        if v:
            out.append((prefix, v))
        return
    rep = _repeated_child_name(el)
    for c in kids:
        if rep and _local(c) == rep:
            continue
        p = f"{prefix}/{_local(c)}" if prefix else _local(c)
        _leaf_values(c, p, out)


def maybe_number(v: str):
    """Giữ text cho mã có số 0 đầu hoặc >15 chữ số (MST/CCCD); còn lại thành số."""
    if re.fullmatch(r"-?\d+", v):
        digits = v.lstrip("-")
        if (digits.startswith("0") and digits != "0") or len(digits) > 15:
            return v
        try:
            return int(v)
        except ValueError:
            return v
    if re.fullmatch(r"-?\d+\.\d+", v):
        try:
            return float(v)
        except ValueError:
            return v
    return v


# ---------------------------------------------------------------------------
# API chính
# ---------------------------------------------------------------------------
def build_sheets(xml_text: str) -> dict[str, pd.DataFrame]:
    """Chuyển XML tờ khai -> dict {tên sheet: DataFrame}."""
    root = ET.fromstring(xml_text)

    tables: list = []
    _collect_tables(root, tables)
    fields: list = []
    _leaf_values(root, "", fields)

    sheets: dict[str, pd.DataFrame] = {}

    # gom ct16..ct41 bất kể nằm nhánh nào
    ct_val = {}
    for path, value in fields:
        leaf = path.rsplit("/", 1)[-1]
        if leaf in FORM_CT and leaf not in ct_val:
            ct_val[leaf] = value
    has_form = bool(ct_val)

    if has_form:
        rows = [[FORM05_TITLE, "", "", "", ""]]
        for heading, sec_rows in FORM05_SECTIONS:
            rows.append([heading, "", "", "", ""])
            for stt, ten, ct, dv in sec_rows:
                raw = ct_val.get(ct)
                rows.append([stt, ten, f"[{ct.replace('ct', '')}]", dv,
                             "" if raw is None else maybe_number(raw)])
        sheets["ToKhaiChinh"] = pd.DataFrame(
            rows, columns=["STT", "Chỉ tiêu", "Mã chỉ tiêu", "Đơn vị tính", "Số người/Số tiền"])

    # thông tin chung (loại các ct đã vào tờ khai chính)
    info_rows = []
    for path, value in fields:
        parts = path.split("/")
        leaf = parts[-1]
        if has_form and leaf in FORM_CT:
            continue
        group = parts[-2] if len(parts) >= 2 else ""
        info_rows.append([group, FIELD_LABELS.get(leaf, leaf), maybe_number(value)])
    if info_rows:
        sheets["ThongTinChung"] = pd.DataFrame(info_rows, columns=["Nhóm", "Chỉ tiêu", "Giá trị"])

    # các bảng kê
    for t in tables:
        cols: list[str] = []
        for r in t["rows"]:
            for c in r:
                n = _local(c)
                if n not in cols:
                    cols.append(n)
        labels = _col_label_map(t["title"])
        header = ["STT"] + [labels.get(c, c) if labels else c for c in cols]
        body = []
        if labels:
            body.append([""] + cols)  # dòng mã cột gốc
        for i, r in enumerate(t["rows"], start=1):
            m = {_local(c): (c.text or "").strip() for c in r}
            body.append([i] + [maybe_number(m.get(col, "")) for col in cols])
        sheets[t["title"][:31] or "Sheet"] = pd.DataFrame(body, columns=header)

    return sheets
