"""
Tác vụ 6: Dò MST với danh sách doanh nghiệp có dấu hiệu rủi ro,
lấy cột 5 (văn bản cảnh báo).
"""

from __future__ import annotations

import pandas as pd

from . import utils


def build_risk_lookup(risk_df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """Từ file DS rủi ro, trích: MST -> Văn bản (cột 5)."""
    mst_col = utils.resolve_column(risk_df, cfg["mst_col"])
    doc_col = utils.resolve_column(risk_df, cfg["doc_col"])

    out = pd.DataFrame(
        {
            "MST (chuẩn hóa)": utils.clean_mst_series(risk_df[mst_col]),
            "Văn bản rủi ro": risk_df[doc_col].astype(str).str.strip(),
        }
    )
    out = out[out["MST (chuẩn hóa)"] != ""]
    # gộp nhiều văn bản của cùng 1 MST
    out = (
        out.groupby("MST (chuẩn hóa)")["Văn bản rủi ro"]
        .apply(lambda s: "; ".join(sorted(set(v for v in s if v and v != "nan"))))
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
