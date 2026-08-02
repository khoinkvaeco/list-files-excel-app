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
    # mm/yyyy, mm.yyyy, mm-yyyy (TMS xuất dạng '01.2023')
    m = re.fullmatch(r"(\d{1,2})\s*[./-]\s*(\d{4})", s)
    if m:
        mm, yy = int(m.group(1)), int(m.group(2))
        return yy * 12 + mm if 1 <= mm <= 12 else None
    m = re.fullmatch(r"(\d{4})\s*[./-]\s*(\d{1,2})", s)
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


def parse_periods(from_val, to_val):
    """Tách các kỳ đăng ký đóng gói bằng ';' (hoặc xuống dòng) thành danh sách
    (từ_idx, đến_idx).

    Ví dụ TMS thô: Từ tháng = '01/2016;01/2023', Đến tháng = '; 10/2023'
      -> kỳ 1: 01/2016 -> (12/2022, ngay trước kỳ sau)
         kỳ 2: 01/2023 -> 10/2023
    'Đến tháng' trống ở kỳ cuối = còn hiệu lực (None)."""
    def _split(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return [""]
        return [s.strip() for s in re.split(r"[;\n]", str(v))]

    fs, ts = _split(from_val), _split(to_val)
    n = max(len(fs), len(ts))
    froms = [month_index(fs[i]) if i < len(fs) else None for i in range(n)]
    tos = [month_index(ts[i]) if i < len(ts) else None for i in range(n)]
    periods = []
    for i in range(n):
        f, t = froms[i], tos[i]
        if f is None and t is None:
            continue
        if t is None and i + 1 < n and froms[i + 1] is not None:
            t = froms[i + 1] - 1  # kết thúc ngay trước kỳ đăng ký kế tiếp
        periods.append((f, t))
    return periods


def _months_span(d_from, d_to) -> int:
    """Số tháng đủ điều kiện (bao gồm cả tháng đầu và cuối)."""
    if d_from is None:
        return 0
    end = d_to if d_to is not None else d_from
    return max(0, end - d_from + 1)


def _period_date(raw, month_idx, is_start):
    """Ngày mốc của kỳ giảm trừ: ưu tiên ct15/ct16 (ngày); nếu trống thì suy ra
    từ ct21/ct22 (tháng) — đầu tháng nếu is_start, cuối tháng nếu ngược lại."""
    import calendar
    d = utils.parse_date(raw)
    if d is not None and not pd.isna(d):
        return pd.Timestamp(d)
    if month_idx is None:
        return None
    y, m = divmod(int(month_idx) - 1, 12)
    m += 1
    day = 1 if is_start else calendar.monthrange(y, m)[1]
    return pd.Timestamp(year=y, month=m, day=day)


def build_hokd(hk_df: pd.DataFrame, cfg: dict) -> dict:
    """File 'MST NPT là hộ cá nhân kinh doanh' -> {MST NPT -> giá trị cột BB}.

    ``cfg``: {mst, bb} (tham chiếu cột). Có mặt khóa = MST thuộc hộ CNKD; giá trị
    là 'Ngày đóng trạng thái tổ chức' (có thể trống)."""
    c = {k: utils.resolve_column(hk_df, v) for k, v in cfg.items()}
    out: dict = {}
    for _, row in hk_df.iterrows():
        mst = _mst_key(row[c["mst"]])
        if not mst:
            continue
        out[mst] = row[c["bb"]]
    return out


def _digits(v) -> str:
    return re.sub(r"\D", "", "" if v is None else str(v))


def _mst_key(v) -> str:
    """Khóa so khớp MST — bỏ số 0 đầu để khớp được cả khi file lưu MST dạng số
    (mất số 0 đầu) lẫn dạng chữ (giữ số 0), vd '0301164435' == 301164435."""
    d = _digits(v)
    return d.lstrip("0") if d else d


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
        mst_npt = _mst_key(row[c["mst_npt"]])
        mst_nnt = _mst_key(row[c["mst_nnt"]])
        name = _norm_name(row[c["ten_npt"]])
        if not mst_npt and not name:
            continue
        # TMS thô có thể gói nhiều kỳ đăng ký bằng ';' -> tách thành nhiều rec
        periods = parse_periods(row[c["tu_thang"]], row[c["den_thang"]])
        if not periods:
            continue  # dòng tiêu đề phụ / trống
        for p_from, p_to in periods:
            rec = {"from": p_from, "to": p_to, "mst_nnt": mst_nnt}
            if mst_npt:
                by_mst.setdefault(mst_npt, []).append(rec)
            if name and mst_nnt:
                by_name.setdefault((mst_nnt, name), []).append(rec)
    # cùng 1 MST NPT: 'Đến tháng' trống = liên tục đến 'Từ tháng' của dòng kế
    for recs in by_mst.values():
        _fill_open_periods(recs)
    for recs in by_name.values():
        _fill_open_periods(recs)
    return by_mst, by_name


def _fill_open_periods(recs: list) -> None:
    """Sắp xếp theo 'từ tháng'; kỳ có 'đến tháng' trống được nối liên tục đến
    ngay trước 'từ tháng' của kỳ kế tiếp (kỳ cuối để trống = còn hiệu lực)."""
    recs.sort(key=lambda r: (r["from"] is None, r["from"] or 0))
    for i, r in enumerate(recs):
        if r["to"] is None:
            for j in range(i + 1, len(recs)):
                nf = recs[j]["from"]
                if nf is not None and (r["from"] is None or nf > r["from"]):
                    r["to"] = nf - 1
                    break


# ---------------------------------------------------------------------------
# Bước 1 + 2: đối chiếu Phụ lục 3
# ---------------------------------------------------------------------------
def reconcile_pl3(pl3_df: pd.DataFrame, cfg: dict, by_mst, by_name, hokd_map=None):
    """Đối chiếu từng dòng PL3 với TMS. ``cfg``: {mst_nnt, ten_npt, mst_npt,
    tu_thang, den_thang}.

    Trả về (DataFrame PL3 + cột mới, true_counts, total_counts, months_by_nnt).
    """
    c = {k: utils.resolve_column(pl3_df, v) for k, v in cfg.items()}
    hokd_map = hokd_map or {}
    c15 = c.get("tu_ngay")   # ct15 = Thời điểm bắt đầu tính giảm trừ (nếu có)
    c16 = c.get("den_ngay")  # ct16 = Thời điểm kết thúc tính giảm trừ (nếu có)

    tms_from, tms_to, matched_by, valid, reasons = [], [], [], [], []
    months_col: list = []  # số tháng đủ ĐK của từng dòng
    row_nnt: list = []  # MST NNT của từng dòng (để dựng cột tỉ lệ TRUE/Tổng)
    true_counts: dict = {}
    total_counts: dict = {}
    months_by_nnt: dict = {}  # MST NNT -> [số tháng của các NPT hợp lệ]

    for _, row in pl3_df.iterrows():
        mst_npt = _mst_key(row[c["mst_npt"]])
        mst_nnt = _mst_key(row[c["mst_nnt"]])
        name = _norm_name(row[c["ten_npt"]])
        # kỳ giảm trừ kê khai: ưu tiên ct21/ct22, nếu trống thì lấy ct15/ct16
        d_from = month_index(row[c["tu_thang"]])
        if d_from is None and c15 is not None:
            d_from = month_index(row[c15])
        d_to = month_index(row[c["den_thang"]])
        if d_to is None and c16 is not None:
            d_to = month_index(row[c16])

        # bỏ qua dòng không phải dữ liệu (dòng mã kỹ thuật ct07/ct08..., dòng trống)
        raw_mst_nnt = "" if row[c["mst_nnt"]] is None else str(row[c["mst_nnt"]]).strip()
        is_ct_row = bool(re.fullmatch(r"(ct\d+(_\w+)?|\[\d+\])", raw_mst_nnt, flags=re.I))
        if is_ct_row or (not mst_nnt and not mst_npt and not name):
            tms_from.append(""); tms_to.append(""); matched_by.append("")
            valid.append(""); reasons.append(""); row_nnt.append("")
            months_col.append("")
            continue
        row_nnt.append(mst_nnt)

        recs, src = [], ""
        if mst_npt and mst_npt in by_mst:
            recs, src = by_mst[mst_npt], "MST NPT"
        elif name and (mst_nnt, name) in by_name:
            recs, src = by_name[(mst_nnt, name)], "Tên NPT"

        if not recs:
            tms_from.append(""); tms_to.append("")
            matched_by.append("Không tìm thấy")
            ok = False
            reason = "Không có trên TMS"
        else:
            # hiển thị kỳ TMS (bản ghi đầu; nhiều kỳ thì nối)
            tms_from.append("; ".join(fmt_month(r["from"]) for r in recs))
            tms_to.append("; ".join(fmt_month(r["to"]) for r in recs))
            matched_by.append(src)
            ok = False
            why = []
            if d_from is None:
                why.append("Thiếu Từ tháng kê khai")
            else:
                d_end = d_to if d_to is not None else d_from
                for r in recs:
                    lo = r["from"] if r["from"] is not None else d_from
                    hi = r["to"]  # None = còn hiệu lực
                    if d_from >= lo and (hi is None or d_end <= hi):
                        ok = True
                        break
                if not ok:
                    for r in recs:
                        lo, hi = r["from"], r["to"]
                        if lo is not None and d_from < lo:
                            why.append(f"Kê khai từ {fmt_month(d_from)} trước kỳ TMS ({fmt_month(lo)})")
                        if hi is not None and d_end > hi:
                            why.append(f"Kê khai đến {fmt_month(d_end)} sau kỳ TMS ({fmt_month(hi)})")
                    if not why:
                        why.append("Kỳ kê khai ngoài kỳ đăng ký TMS")
            reason = "" if ok else "; ".join(dict.fromkeys(why))

        # Ghi đè theo file "MST NPT là hộ cá nhân kinh doanh" (nếu MST NPT có trong file)
        if mst_npt and mst_npt in hokd_map:
            bb = utils.parse_date(hokd_map[mst_npt])
            end_d = _period_date(row[c16] if c16 else None, d_to if d_to is not None else d_from, is_start=False)
            if bb is None or pd.isna(bb):
                ok, reason = False, "NPT là hộ CNKD"
            elif end_d is not None and end_d < bb:
                ok, reason = False, "NPT là hộ CNKD"
            else:
                ok = True
                reason = "đủ thời gian npt hộ CNKD từ ngày " + bb.strftime("%d/%m/%Y")

        valid.append("TRUE" if ok else "FALSE")
        reasons.append(reason)

        # số tháng đủ ĐK: chỉ tính cho dòng hợp lệ (theo kỳ kê khai ct21/ct22)
        n_months = _months_span(d_from, d_to) if valid[-1] == "TRUE" else 0
        months_col.append(f"{n_months:02d} tháng" if n_months else "")

        if mst_nnt:
            total_counts[mst_nnt] = total_counts.get(mst_nnt, 0) + 1
            if valid[-1] == "TRUE":
                true_counts[mst_nnt] = true_counts.get(mst_nnt, 0) + 1
                months_by_nnt.setdefault(mst_nnt, []).append(n_months)

    # cột tỉ lệ NPT hợp lệ theo từng NNT (TRUE/Tổng kê khai) — chuyển từ PL1 sang PL3
    ratio_col = []
    for m in row_nnt:
        if m and m in total_counts:
            t = true_counts.get(m, 0)
            n = total_counts.get(m, 0)
            ratio_col.append(f"{t}/{n}" if n else str(t))
        else:
            ratio_col.append("")

    out = pl3_df.copy()
    # thứ tự cột mới: Nguồn khớp -> TMS Từ/Đến tháng -> Hợp lệ
    #                 -> Số tháng đủ ĐK NPT -> Số NPT (TRUE/Tổng kê khai) -> Lý do sai
    out["Nguồn khớp"] = matched_by
    out["TMS Từ tháng"] = tms_from
    out["TMS Đến tháng"] = tms_to
    out["Hợp lệ"] = valid
    out["Số tháng đủ ĐK NPT"] = months_col
    out["Số NPT (TRUE/Tổng kê khai)"] = ratio_col
    out["Lý do sai"] = reasons
    return out, true_counts, total_counts, months_by_nnt


# ---------------------------------------------------------------------------
# Bước 3: điền tổng NPT TRUE vào cột cuối Phụ lục 1
# ---------------------------------------------------------------------------
def _months_note(month_list) -> str:
    """Ghi chú số tháng đủ ĐK cho các NPT chưa đủ 12 tháng (đủ 12 -> để trống)."""
    from collections import Counter
    partial = [m for m in (month_list or []) if 0 < m < 12]
    if not partial:
        return ""
    return "; ".join(
        f"có {cnt} mst npt đủ {m:02d} tháng" for m, cnt in sorted(Counter(partial).items())
    )


def apply_pl1(pl1_df: pd.DataFrame, mst_ref, true_counts: dict,
              total_counts: dict | None = None,
              months_by_nnt: dict | None = None) -> pd.DataFrame:
    """Thêm cột cuối vào Phụ lục 1 theo MST NNT:
    - 'Số NPT hợp lệ (TRUE)'
    - 'Số tháng đủ ĐK NPT (nếu <12)' — ghi chú các NPT chưa đủ 12 tháng.

    (Cột tỉ lệ 'Số NPT (TRUE/Tổng kê khai)' đã chuyển sang hiển thị ở PL3.)"""
    col = utils.resolve_column(pl1_df, mst_ref)
    out = pl1_df.copy()
    total_counts = total_counts or {}
    months_by_nnt = months_by_nnt or {}
    trues, notes = [], []
    for v in out[col]:
        d = _mst_key(v)
        if not d or (d not in true_counts and d not in total_counts):
            trues.append(""); notes.append("")
            continue
        trues.append(true_counts.get(d, 0))
        notes.append(_months_note(months_by_nnt.get(d)))
    out["Số NPT hợp lệ (TRUE)"] = trues
    out["Số tháng đủ ĐK NPT (nếu <12)"] = notes
    return out


# ---------------------------------------------------------------------------
# Phụ lục 05-2: tính lại thuế TNCN và chênh lệch
# ---------------------------------------------------------------------------
def _is_nonresident(v) -> bool:
    """'Cá nhân không cư trú' = true/x/1 -> không cư trú (thuế suất 20%)."""
    s = _norm_name(v)
    return s in ("true", "x", "1", "co", "có", "khong cu tru")


def process_pl2(pl2_df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """Thêm cột 'Tính thuế TNCN (PL2)' và 'Chênh lệch' vào Phụ lục 05-2.

    ``cfg``: {income, nonresident, withheld} (tham chiếu cột).
    - Cư trú (Cá nhân không cư trú = false) -> thuế = Thu nhập chịu thuế × 10%.
    - Không cư trú (true) -> thuế = Thu nhập chịu thuế × 20%.
    - Chênh lệch = Tính thuế TNCN − Số thuế TNCN đã khấu trừ."""
    c = {k: utils.resolve_column(pl2_df, v) for k, v in cfg.items()}
    out = pl2_df.copy()
    calc, diff = [], []
    for _, row in pl2_df.iterrows():
        raw_inc = row[c["income"]]
        raw_str = "" if raw_inc is None else str(raw_inc).strip()
        # bỏ qua dòng mã kỹ thuật (ct11.../[12]...) và dòng trống
        if re.fullmatch(r"(ct\d+(_\w+)?|\[\d+\])", raw_str, flags=re.I) or raw_str == "":
            calc.append(""); diff.append("")
            continue
        inc = utils.parse_amount(raw_inc)
        rate = 0.20 if _is_nonresident(row[c["nonresident"]]) else 0.10
        tax = round(inc * rate)
        withheld = round(utils.parse_amount(row[c["withheld"]]))
        calc.append(tax)
        diff.append(tax - withheld)
    out["Tính thuế TNCN (PL2)"] = calc
    out["Chênh lệch"] = diff
    return out


# ---------------------------------------------------------------------------
# Ủy quyền quyết toán đúng/sai — đối chiếu PL1 với file "nhiều nguồn TN"
# ---------------------------------------------------------------------------
def _find_col(df: pd.DataFrame, *cands):
    """Tìm cột theo tên tiêu đề — độc lập với chế độ cột toàn cục (file 'nhiều
    nguồn TN' các sheet có bố cục cột khác nhau nhưng tên tiêu đề nhất quán).

    Ưu tiên khớp CHÍNH XÁC (sau khi bỏ dấu) trước, rồi mới đến khớp CHỨA — tránh
    'MST cá nhân' khớp nhầm 'CQT quản lý MST cá nhân', hay 'Thuế TNCN' khớp nhầm
    'Thu nhập chịu thuế TNCN'."""
    norm = {col: _norm_name(col) for col in df.columns}
    cns = [_norm_name(c) for c in cands if _norm_name(c)]
    for cn in cns:  # khớp chính xác
        for col in df.columns:
            if norm[col] == cn:
                return col
    for cn in cns:  # khớp chứa
        for col in df.columns:
            if cn in norm[col]:
                return col
    return None


# thứ tự ưu tiên khi 1 MST xuất hiện ở nhiều năm/sheet
_MS_PRIORITY = [
    "nhiều nguồn tn",
    "05-2 dưới 2 triệu, không khấu trừ thuế",
    "Đã khấu trừ TNCN 10%",
]


def _ms_verdict(rows_051, rows_052, c_inc, c_tax) -> str:
    """Xét kết quả 'nhiều nguồn TN' cho 1 MST trong 1 sheet (năm)."""
    if len(rows_051) >= 2:
        return "nhiều nguồn tn"
    if len(rows_051) == 1 and len(rows_052) >= 1:
        verdicts = []
        for r in rows_052:
            inc = utils.parse_amount(r[c_inc]) if c_inc is not None else 0.0
            tax = utils.parse_amount(r[c_tax]) if c_tax is not None else 0.0
            if inc < 2_000_000:
                verdicts.append("05-2 dưới 2 triệu, không khấu trừ thuế")
            elif tax + 1 >= 0.1 * inc:  # Thuế TNCN = 10% × TNCT (có dung sai làm tròn)
                verdicts.append("Đã khấu trừ TNCN 10%")
            else:  # Thuế TNCN < 10% × TNCT
                verdicts.append("nhiều nguồn tn")
        for p in _MS_PRIORITY:
            if p in verdicts:
                return p
        return verdicts[0] if verdicts else ""
    return ""  # chỉ 1 nguồn (1 bảng kê 05-1) — không gắn cờ


def classify_multi_source(sheets: dict) -> dict:
    """Phân loại 'ủy quyền quyết toán đúng/sai' từ file 'nhiều nguồn TN'.

    ``sheets``: {tên_sheet -> DataFrame} (mỗi sheet là 1 năm). Nhóm theo
    'MST cá nhân' trong từng sheet, áp quy tắc, rồi gộp qua các năm theo ưu tiên.
    Trả về {MST cá nhân -> kết quả}.
    """
    per: dict = {}
    for _name, df in sheets.items():
        if df is None or df.empty:
            continue
        c_mst = _find_col(df, "MST cá nhân", "MST ca nhan", "Mã số thuế")
        c_src = _find_col(df, "Nguồn dữ liệu")
        c_inc = _find_col(df, "Thu nhập chịu thuế")
        c_tax = _find_col(df, "Thuế TNCN", "số thuế TNCN đã khấu trừ")
        if c_mst is None or c_src is None:
            continue
        groups: dict = {}
        for _, row in df.iterrows():
            mst = _mst_key(row[c_mst])
            if not mst:
                continue
            src = "" if row[c_src] is None else str(row[c_src])
            groups.setdefault(mst, []).append((src, row))
        for mst, items in groups.items():
            r051 = [r for s, r in items if "05-1" in s]
            r052 = [r for s, r in items if "05-2" in s]
            v = _ms_verdict(r051, r052, c_inc, c_tax)
            if v:
                per.setdefault(mst, []).append(v)
    out: dict = {}
    for mst, vs in per.items():
        chosen = next((p for p in _MS_PRIORITY if p in vs), vs[0] if vs else "")
        out[mst] = chosen
    return out


def apply_multi_source(pl1_df: pd.DataFrame, mst_ref, result_map: dict) -> pd.DataFrame:
    """VLOOKUP kết quả 'nhiều nguồn TN' theo MST cá nhân vào cột mới của PL1."""
    col = utils.resolve_column(pl1_df, mst_ref)
    out = pl1_df.copy()
    out["Ủy quyền QT (nhiều nguồn TN)"] = [result_map.get(_mst_key(v), "") for v in out[col]]
    return out
