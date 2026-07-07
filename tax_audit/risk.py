"""
Tác vụ 6: Dò MST với danh sách doanh nghiệp có dấu hiệu rủi ro.

Danh sách rủi ro thường gồm NHIỀU sheet với bố cục khác nhau; mỗi sheet được
khai báo riêng trong config (tên sheet, dòng tiêu đề, cột MST, cột/nhãn văn bản).
"""

from __future__ import annotations

import pandas as pd

from . import utils


def build_risk_lookup(file, cfg: dict) -> pd.DataFrame:
    """Đọc toàn bộ các sheet rủi ro trong ``file`` và gộp thành bảng tra cứu.

    ``cfg`` là dict RISK trong config, có khóa "sheets" là list các spec:
        {name, header_row, mst_col, doc_col}
    - doc_col là int  -> lấy văn bản theo từng dòng ở cột đó.
    - doc_col là str  -> gán nhãn văn bản cố định cho mọi dòng của sheet.

    Trả về DataFrame: MST (chuẩn hóa), Văn bản rủi ro.
    """
    available = set(utils.list_sheets(file))
    frames = []

    for spec in cfg["sheets"]:
        name = spec["name"]
        if name not in available:
            # bỏ qua sheet không tồn tại thay vì lỗi (file có thể thay đổi)
            continue
        df = utils.read_excel(file, header_row=spec["header_row"], sheet_name=name)
        if df.empty:
            continue

        try:
            mst_col = utils.resolve_column(df, spec["mst_col"])
        except (IndexError, KeyError):
            continue
        mst = utils.clean_mst_series(df[mst_col])

        doc_ref = spec.get("doc_col")
        # thử coi doc_col là 1 CỘT (số thứ tự / chữ cái / tên tiêu đề); nếu không
        # khớp được cột nào thì coi là NHÃN văn bản cố định cho cả sheet.
        doc = None
        if doc_ref not in (None, ""):
            try:
                doc_col = utils.resolve_column(df, doc_ref)
                doc = df[doc_col].astype(str).str.strip()
            except (IndexError, KeyError):
                doc = None
        if doc is None:
            doc = pd.Series([str(doc_ref) if doc_ref else str(name)] * len(df))

        part = pd.DataFrame(
            {"MST (chuẩn hóa)": mst.values, "Văn bản rủi ro": doc.values}
        )
        part = part[part["MST (chuẩn hóa)"] != ""]
        frames.append(part)

    if not frames:
        return pd.DataFrame(columns=["MST (chuẩn hóa)", "Văn bản rủi ro"])

    combined = pd.concat(frames, ignore_index=True)
    # gộp nhiều văn bản của cùng 1 MST
    out = (
        combined.groupby("MST (chuẩn hóa)")["Văn bản rủi ro"]
        .apply(lambda s: "; ".join(sorted(set(v for v in s if v and v.lower() != "nan"))))
        .reset_index()
    )
    return out


def enrich_with_risk(df: pd.DataFrame, risk_lookup: pd.DataFrame) -> pd.DataFrame:
    """Ghép thông tin rủi ro vào dữ liệu chính.

    Thêm cột 'Văn bản rủi ro' và 'DN rủi ro' (bool).
    """
    df = df.copy()
    if "MST (chuẩn hóa)" not in df.columns:
        raise ValueError("Cần chạy xử lý MST trước (thiếu cột 'MST (chuẩn hóa)').")

    merged = df.merge(risk_lookup, on="MST (chuẩn hóa)", how="left")
    merged["DN rủi ro"] = merged["Văn bản rủi ro"].notna() & (
        merged["Văn bản rủi ro"].astype(str).str.strip() != ""
    )
    return merged


def risky_invoices(enriched: pd.DataFrame) -> pd.DataFrame:
    """Lọc ra các hóa đơn có MST thuộc DS rủi ro."""
    if "DN rủi ro" not in enriched.columns:
        return enriched.iloc[0:0]
    return enriched[enriched["DN rủi ro"] == True].copy()  # noqa: E712
