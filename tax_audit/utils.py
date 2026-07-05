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
    hoặc đối tượng file (ví dụ file upload của Streamlit). ``sheet_name`` có
    thể là chỉ số hoặc tên sheet.
    """
    name = getattr(file, "name", str(file)).lower()

    if name.endswith(".csv"):
        _rewind(file)
        return pd.read_csv(file, header=header_row - 1, dtype=str, keep_default_na=False)

    # SpreadsheetML (Excel 2003 XML) — nhiều file .xls từ TMS/eTax ở định dạng này
    if _is_spreadsheetml(file):
        return _read_spreadsheetml(file, header_row, sheet_name)

    engines = _engine_candidates(name)
    last_err = None
    for engine in engines:
        try:
            _rewind(file)
            return pd.read_excel(
                file,
                header=header_row - 1,
                sheet_name=sheet_name,
                dtype=object,
                engine=engine,
            )
        except Exception as e:  # noqa: BLE001
            last_err = e

    # Fallback cuối: file "xls" thực chất là bảng HTML
    try:
        return _read_html_table(file, header_row, sheet_name)
    except Exception:  # noqa: BLE001
        if last_err:
            raise last_err
        raise


def list_sheets(file) -> list[str]:
    """Trả về danh sách tên sheet của file Excel (chịu nhiều định dạng)."""
    name = getattr(file, "name", str(file)).lower()
    if name.endswith(".csv"):
        return ["Sheet1"]
    if _is_spreadsheetml(file):
        return _spreadsheetml_sheets(file)
    for engine in _engine_candidates(name):
        try:
            _rewind(file)
            return pd.ExcelFile(file, engine=engine).sheet_names
        except Exception:  # noqa: BLE001
            continue
    # HTML-"xls": không có khái niệm sheet, đặt tên mặc định
    return ["Sheet1"]


# ---------------------------------------------------------------------------
# SpreadsheetML (Excel 2003 XML)
# ---------------------------------------------------------------------------
_SSML_NS = "{urn:schemas-microsoft-com:office:spreadsheet}"


def _read_all_bytes(file) -> bytes:
    _rewind(file)
    if hasattr(file, "read"):
        data = file.read()
        _rewind(file)
        return data if isinstance(data, bytes) else data.encode("utf-8", "ignore")
    with open(file, "rb") as fh:
        return fh.read()


def _is_spreadsheetml(file) -> bool:
    """Nhận biết file định dạng SpreadsheetML qua vài KB đầu."""
    try:
        _rewind(file)
        if hasattr(file, "read"):
            head = file.read(4096)
            _rewind(file)
            if isinstance(head, str):
                head = head.encode("utf-8", "ignore")
        else:
            with open(file, "rb") as fh:
                head = fh.read(4096)
    except Exception:  # noqa: BLE001
        return False
    return b"urn:schemas-microsoft-com:office:spreadsheet" in head


def _spreadsheetml_worksheets(root):
    return list(root.iter(_SSML_NS + "Worksheet"))


def _spreadsheetml_sheets(file) -> list[str]:
    import xml.etree.ElementTree as ET

    root = ET.fromstring(_read_all_bytes(file))
    names = [w.get(_SSML_NS + "Name") or f"Sheet{i+1}"
             for i, w in enumerate(_spreadsheetml_worksheets(root))]
    return names or ["Sheet1"]


def _read_spreadsheetml(file, header_row: int, sheet_name) -> pd.DataFrame:
    """Đọc 1 sheet của file SpreadsheetML thành DataFrame (giữ đúng vị trí cột)."""
    import xml.etree.ElementTree as ET

    root = ET.fromstring(_read_all_bytes(file))
    worksheets = _spreadsheetml_worksheets(root)
    if not worksheets:
        raise ValueError("Không tìm thấy worksheet trong file SpreadsheetML.")

    chosen = None
    if isinstance(sheet_name, str):
        for w in worksheets:
            if (w.get(_SSML_NS + "Name") or "") == sheet_name:
                chosen = w
                break
    elif isinstance(sheet_name, int) and 0 <= sheet_name < len(worksheets):
        chosen = worksheets[sheet_name]
    chosen = chosen or worksheets[0]

    table = chosen.find(_SSML_NS + "Table")
    if table is None:
        return pd.DataFrame()

    rows = []
    max_col = 0
    for r in table.findall(_SSML_NS + "Row"):
        cells = {}
        c = 0
        for cell in r.findall(_SSML_NS + "Cell"):
            idx = cell.get(_SSML_NS + "Index")
            c = int(idx) if idx else c + 1
            data_el = cell.find(_SSML_NS + "Data")
            cells[c] = data_el.text if (data_el is not None and data_el.text is not None) else ""
            merge = cell.get(_SSML_NS + "MergeAcross")
            if merge:
                c += int(merge)
        rows.append(cells)
        if cells:
            max_col = max(max_col, max(cells))

    matrix = [[row.get(c, None) for c in range(1, max_col + 1)] for row in rows]
    if not matrix:
        return pd.DataFrame()

    hidx = header_row - 1
    if hidx >= len(matrix):
        raise IndexError(f"Dòng tiêu đề {header_row} vượt quá số dòng của sheet.")
    header = [("" if h is None else str(h)) for h in matrix[hidx]]
    data = matrix[hidx + 1:]
    return pd.DataFrame(data, columns=_dedupe_columns(header)).astype(object)


def _dedupe_columns(names: list[str]) -> list[str]:
    """Bảo đảm tên cột không trùng/không rỗng (giống cách pandas xử lý)."""
    seen: dict[str, int] = {}
    out = []
    for i, n in enumerate(names):
        base = n if n.strip() else f"Unnamed: {i}"
        if base in seen:
            seen[base] += 1
            out.append(f"{base}.{seen[base]}")
        else:
            seen[base] = 0
            out.append(base)
    return out


def _engine_candidates(name: str) -> list[str]:
    """Chọn thứ tự engine đọc theo phần mở rộng file."""
    if name.endswith((".xlsx", ".xlsm")):
        return ["openpyxl"]
    if name.endswith(".xls"):
        return ["xlrd"]
    return ["openpyxl", "xlrd"]


def _read_html_table(file, header_row: int, sheet_name) -> pd.DataFrame:
    """Đọc file 'xls'/'html' thực chất là bảng HTML.

    Lấy bảng có nhiều cột nhất. ``sheet_name`` bị bỏ qua vì HTML không có sheet.
    """
    _rewind(file)
    data = file.read() if hasattr(file, "read") else open(file, "rb").read()
    if isinstance(data, bytes):
        for enc in ("utf-8", "utf-16", "latin-1"):
            try:
                text = data.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            text = data.decode("utf-8", errors="ignore")
    else:
        text = data

    tables = pd.read_html(io.StringIO(text), header=header_row - 1)
    if not tables:
        raise ValueError("Không tìm thấy bảng dữ liệu trong file HTML.")
    best = max(tables, key=lambda t: t.shape[1])
    return best.astype(object)


def _rewind(file) -> None:
    """Đưa con trỏ đọc về đầu (cần cho file upload đọc nhiều lần)."""
    if hasattr(file, "seek"):
        try:
            file.seek(0)
        except (OSError, ValueError):
            pass


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

    Nhận diện đồng thời hai kiểu ngày:
      - ISO / có năm ở đầu (yyyy-mm-dd, yyyy/mm/dd, kèm giờ)  -> dayfirst=False
      - Kiểu Việt Nam (dd/mm/yyyy, dd-mm-yyyy)                -> dayfirst=True
    Tránh lỗi hiểu nhầm '2022-02-10' thành ngày 02 tháng 10.
    """
    import warnings

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        if isinstance(value, pd.Series):
            s = value.astype("string")
            iso_mask = s.str.match(r"^\s*\d{4}[-/]\d{1,2}[-/]\d{1,2}").fillna(False)
            day_first = pd.to_datetime(value, errors="coerce", dayfirst=True)
            iso = pd.to_datetime(value, errors="coerce", dayfirst=False)
            return day_first.where(~iso_mask, iso)
        # giá trị đơn
        text = "" if value is None else str(value)
        dayfirst = not bool(re.match(r"^\s*\d{4}[-/]\d{1,2}[-/]\d{1,2}", text))
        return pd.to_datetime(value, errors="coerce", dayfirst=dayfirst)


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
