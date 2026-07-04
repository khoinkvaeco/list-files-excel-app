"""
Tiện ích dùng chung: đọc/ghi Excel, chuẩn hóa MST, chuẩn hóa số hóa đơn,
phân giải tham chiếu cột (theo vị trí hoặc theo tên).
"""

from __future__ import annotations

import io
import re
from typing import Union

import pandas as pd

ColumnRef = Union[int, str]


# ---------------------------------------------------------------------------
# ĐỌC / GHI FILE
# ---------------------------------------------------------------------------
def read_excel(file, header_row: int = 1, sheet_name=0) -> pd.DataFrame:
    """Đọc file Excel/CSV thành DataFrame.

    ``header_row`` đếm từ 1 (dòng tiêu đề). ``file`` có thể là đường dẫn
    hoặc đối tượng file (ví dụ file upload của Streamlit).
    """
    name = getattr(file, "name", str(file)).lower()
    if name.endswith(".csv"):
        return pd.read_csv(file, header=header_row - 1, dtype=str, keep_default_na=False)
    return pd.read_excel(
        file,
        header=header_row - 1,
        sheet_name=sheet_name,
        dtype=object,
        engine="openpyxl",
    )


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Ghi nhiều DataFrame thành 1 file Excel nhiều sheet, trả về bytes."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            # tên sheet Excel tối đa 31 ký tự
            safe_name = str(sheet_name)[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# PHÂN GIẢI CỘT
# ---------------------------------------------------------------------------
def resolve_column(df: pd.DataFrame, ref: ColumnRef) -> str:
    """Trả về TÊN cột thực tế trong ``df`` từ tham chiếu ``ref``.

    - int  : vị trí cột đếm từ 1.
    - str  : khớp theo tên (ưu tiên khớp chính xác, sau đó khớp "chứa",
             không phân biệt hoa/thường và khoảng trắng thừa).
    """
    if isinstance(ref, int):
        idx = ref - 1
        if idx < 0 or idx >= len(df.columns):
            raise IndexError(
                f"File chỉ có {len(df.columns)} cột, không có cột thứ {ref}."
            )
        return df.columns[idx]

    target = _norm_text(ref)
    cols = list(df.columns)
    # khớp chính xác
    for c in cols:
        if _norm_text(c) == target:
            return c
    # khớp chứa
    for c in cols:
        if target in _norm_text(c):
            return c
    raise KeyError(
        f"Không tìm thấy cột tên gần đúng '{ref}'. "
        f"Các cột hiện có: {cols}"
    )


def _norm_text(x) -> str:
    return re.sub(r"\s+", " ", str(x)).strip().lower()


# ---------------------------------------------------------------------------
# CHUẨN HÓA MST
# ---------------------------------------------------------------------------
def clean_mst(value) -> str:
    """Chuẩn hóa 1 mã số thuế.

    - Bỏ khoảng trắng, ký tự lạ, giữ dạng chuỗi (bảo toàn số 0 đầu).
    - MST 10 số: nếu bị mất số 0 đầu (còn 9 số) sẽ được bù về 10 số.
    - MST đơn vị phụ thuộc dạng ``xxxxxxxxxx-xxx`` được giữ nguyên định dạng.
    """
    if value is None:
        return ""
    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return ""

    # tách phần mã đơn vị phụ thuộc (sau dấu "-")
    parts = re.split(r"[-\s]+", s)
    main = re.sub(r"\D", "", parts[0])
    if not main:
        return ""
    # bù số 0 đầu cho MST 10 số nếu bị Excel cắt mất
    if len(main) == 9:
        main = main.zfill(10)

    if len(parts) > 1:
        branch = re.sub(r"\D", "", parts[1])
        if branch:
            return f"{main}-{branch.zfill(3)}"
    return main


def clean_mst_series(s: pd.Series) -> pd.Series:
    return s.map(clean_mst)


# ---------------------------------------------------------------------------
# CHUẨN HÓA SỐ HÓA ĐƠN
# ---------------------------------------------------------------------------
def strip_leading_zeros(value) -> str:
    """Xóa các số 0 ở đầu số hóa đơn. '0000123' -> '123', '0' -> '0'."""
    if value is None:
        return ""
    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return ""
    # nếu là số thực dạng '123.0' thì bỏ phần thập phân .0
    m = re.fullmatch(r"(\d+)\.0+", s)
    if m:
        s = m.group(1)
    stripped = s.lstrip("0")
    return stripped if stripped != "" else "0"


def strip_leading_zeros_series(s: pd.Series) -> pd.Series:
    return s.map(strip_leading_zeros)


# ---------------------------------------------------------------------------
# CHUẨN HÓA NGÀY & SỐ TIỀN
# ---------------------------------------------------------------------------
def parse_date(value):
    """Chuyển giá trị về pandas Timestamp; lỗi trả về NaT.

    Ưu tiên định dạng ngày kiểu Việt Nam (dd/mm/yyyy) qua ``dayfirst=True``,
    đồng thời bỏ qua cảnh báo khi gặp chuỗi ISO đã rõ định dạng.
    """
    import warnings

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        return pd.to_datetime(value, errors="coerce", dayfirst=True)


def parse_amount(value) -> float:
    """Chuyển giá trị tiền (có thể chứa dấu phẩy ngăn cách) về float."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        try:
            return float(value)
        except (TypeError, ValueError):
            return 0.0
    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return 0.0
    # bỏ ký tự ngăn cách hàng nghìn, giữ dấu chấm thập phân
    s = re.sub(r"[^\d,.\-]", "", s)
    # nếu có cả '.' và ',' -> giả định ',' là ngăn cách nghìn
    if "," in s and "." in s:
        s = s.replace(",", "")
    elif "," in s:
        # chỉ có ',' -> có thể là thập phân kiểu VN hoặc ngăn cách nghìn
        # nếu ',' đứng cách cuối 3 chữ số coi như ngăn cách nghìn
        if re.search(r",\d{3}(\D|$)", s):
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_amount_series(s: pd.Series) -> pd.Series:
    return s.map(parse_amount)
