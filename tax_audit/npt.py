"""
Kiểm tra NGƯỜI PHỤ THUỘC (NPT) — đối chiếu Phụ lục 05-3/BK-QTT-TNCN với dữ liệu
NPT từ TMS.

Quy trình:
  1. Với mỗi dòng NPT trong Phụ lục 3, tra dữ liệu TMS theo MST NPT (nếu không
     có thì theo MST NNT + tên NPT chuẩn hóa) -> điền "TMS Từ tháng", "TMS Đến
     tháng" vào các cột mới.
  2. So khớp thời gian: kỳ giảm trừ kê khai (Từ tháng/Đến tháng của PL3) phải
     nằm TRONG kỳ đã đăng ký trên TMS (TMS đến tháng trống = còn hiệu lực)
     -> cột "Hợp lệ" TRUE/FALSE.
  3. Tổng hợp số NPT TRUE theo từng MST NNT -> điền vào cột cuối Phụ lục 1.
"""

from __future__ import annotations

import re
import unicodedata

import pandas as pd

from . import utils


# ---------------------------------------------------------------------------
# Tiện ích
# ---------------------------------------------------------------------------
def _norm_name(s) -> str:
    """Chuẩn hóa tên: bỏ dấu, thường hóa, gộp khoảng trắng."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    t = unicodedata.normalize("NFD", str(s))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.replace("đ", "d").replace("Đ", "D")
    return re.sub(r"\s+", " ", t).strip().lower()


def month_index(value):
    """Đổi giá trị 'tháng' bất kỳ -> chỉ số tháng (year*12+month) hoặc None.

    Nhận: datetime/Timestamp; 'mm/yyyy' hoặc 'm/yyyy'; 'yyyy-mm'; ngày đầy đủ.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if not s or s.lower() in ("nan", "nat", "none"):
        return None
    m = re.fullmatch(r"(\d{1,2})\s*/\s*(\d{4})", s)
    if m:
        mm, yy = int(m.group(1)), int(m.group(2))
        return yy * 12 + mm if 1 <= mm <= 12 else None
    m = re.fullmatch(r"(\d{4})\s*[-/]\s*(\d{1,2})", s)
    if m:
        yy, mm = int(m.group(1)), int(m.group(2))
        return yy * 12 + mm if 1 <= mm <= 12 else None
    dd = utils.parse_date(value)
    if dd is not None and not pd.isna(dd):
        return int(dd.year) * 12 + int(dd.month)
    return None


def fmt_month(idx) -> str:
    if idx is None:
        return ""
    y, m = divmod(int(idx) - 1, 12)
    return f"{m + 1:02d}/{y}"


def _digits(v) -> str:
    return re.sub(r"\D", "", "" if v is None else str(v))


# ---------------------------------------------------------------------------
# Bước 0: đọc dữ liệu TMS NPT thành bảng tra cứu
# ---------------------------------------------------------------------------
def build_tms_npt(tms_df: pd.DataFrame, cfg: dict):
    """Từ file TMS NPT tạo tra cứu.

    ``cfg``: {mst_nnt, ten_npt, mst_npt, tu_thang, den_thang} (tham chiếu cột).
    Trả về (by_mst, by_name):
      by_mst : {MST NPT -> [rec]}
      by_name: {(MST NNT, tên chuẩn hóa) -> [rec]}
      rec = {"from": idx|None, "to": idx|None, "mst_nnt": str}
    """
    c = {k: utils.resolve_column(tms_df, v) for k, v in cfg.items()}
    by_mst: dict = {}
    by_name: dict = {}
    for _, row in tms_df.iterrows():
        mst_npt = _digits(row[c["mst_npt"]])
        mst_nnt = _digits(row[c["mst_nnt"]])
        name = _norm_name(row[c["ten_npt"]])
        if not mst_npt and not name:
            continue
        rec = {
            "from": month_index(row[c["tu_thang"]]),
            "to": month_index(row[c["den_thang"]]),
            "mst_nnt": mst_nnt,
        }
        if rec["from"] is None and rec["to"] is None:
            continue  # dòng tiêu đề phụ / trống
        if mst_npt:
            by_mst.setdefault(mst_npt, []).append(rec)
        if name and mst_nnt:
            by_name.setdefault((mst_nnt, name), []).append(rec)
    return by_mst, by_name


# ---------------------------------------------------------------------------
# Bước 1 + 2: đối chiếu Phụ lục 3
# ---------------------------------------------------------------------------
def reconcile_pl3(pl3_df: pd.DataFrame, cfg: dict, by_mst, by_name):
    """Đối chiếu từng dòng PL3 với TMS. ``cfg``: {mst_nnt, ten_npt, mst_npt,
    tu_thang, den_thang}.

    Trả về (DataFrame PL3 gốc + 4 cột mới, dict {MST NNT -> số NPT TRUE}).
    """
    c = {k: utils.resolve_column(pl3_df, v) for k, v in cfg.items()}

    tms_from, tms_to, matched_by, valid = [], [], [], []
    true_counts: dict = {}

    for _, row in pl3_df.iterrows():
        mst_npt = _digits(row[c["mst_npt"]])
        mst_nnt = _digits(row[c["mst_nnt"]])
        name = _norm_name(row[c["ten_npt"]])
        d_from = month_index(row[c["tu_thang"]])
        d_to = month_index(row[c["den_thang"]])

        # bỏ qua dòng không phải dữ liệu (dòng mã kỹ thuật ct07/ct08..., dòng trống)
        raw_mst_nnt = "" if row[c["mst_nnt"]] is None else str(row[c["mst_nnt"]]).strip()
        is_ct_row = bool(re.fullmatch(r"ct\d+(_\w+)?", raw_mst_nnt, flags=re.I))
        if is_ct_row or (not mst_nnt and not mst_npt and not name):
            tms_from.append(""); tms_to.append(""); matched_by.append(""); valid.append("")
            continue

        recs, src = [], ""
        if mst_npt and mst_npt in by_mst:
            recs, src = by_mst[mst_npt], "MST NPT"
        elif name and (mst_nnt, name) in by_name:
            recs, src = by_name[(mst_nnt, name)], "Tên NPT"

        if not recs:
            tms_from.append(""); tms_to.append("")
            matched_by.append("Không tìm thấy"); valid.append("FALSE")
        else:
            # hiển thị kỳ TMS (bản ghi đầu; nhiều kỳ thì nối)
            tms_from.append("; ".join(fmt_month(r["from"]) for r in recs))
            tms_to.append("; ".join(fmt_month(r["to"]) for r in recs))
            matched_by.append(src)
            ok = False
            if d_from is not None:
                d_end = d_to if d_to is not None else d_from
                for r in recs:
                    lo = r["from"] if r["from"] is not None else d_from
                    hi = r["to"]  # None = còn hiệu lực
                    if d_from >= lo and (hi is None or d_end <= hi):
                        ok = True
                        break
            valid.append("TRUE" if ok else "FALSE")

        if valid[-1] == "TRUE" and mst_nnt:
            true_counts[mst_nnt] = true_counts.get(mst_nnt, 0) + 1

    out = pl3_df.copy()
    out["TMS Từ tháng"] = tms_from
    out["TMS Đến tháng"] = tms_to
    out["Nguồn khớp"] = matched_by
    out["Hợp lệ"] = valid
    return out, true_counts


# ---------------------------------------------------------------------------
# Bước 3: điền tổng NPT TRUE vào cột cuối Phụ lục 1
# ---------------------------------------------------------------------------
def apply_pl1(pl1_df: pd.DataFrame, mst_ref, true_counts: dict) -> pd.DataFrame:
    """Thêm cột cuối 'Số NPT hợp lệ (TRUE)' vào Phụ lục 1 theo MST NNT."""
    col = utils.resolve_column(pl1_df, mst_ref)
    out = pl1_df.copy()
    out["Số NPT hợp lệ (TRUE)"] = [
        true_counts.get(_digits(v), "") if _digits(v) else ""
        for v in out[col]
    ]
    return out
