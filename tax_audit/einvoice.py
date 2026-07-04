"""
Tác vụ 10, 11, 12: Tra file hóa đơn điện tử (HĐĐT).

- Lấy trạng thái hóa đơn, tổng tiền thuế.
- Phân loại: bị thay thế / bị xóa bỏ / không tìm thấy.
- Tính chênh lệch tiền thuế = Tổng tiền thuế (HĐĐT) - Thuế GTGT (tờ khai).
"""

from __future__ import annotations

import unicodedata

import pandas as pd

from . import utils


def _norm(text) -> str:
    s = str(text)
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower().strip()


def build_einvoice_lookup(einv_df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """Từ file HĐĐT tạo bảng tra cứu theo khóa (MST + Số HĐ + Ngày HĐ)."""
    cols = cfg["cols"]
    mst_col = utils.resolve_column(einv_df, cols["mst"])
    no_col = utils.resolve_column(einv_df, cols["invoice_no"])
    date_col = utils.resolve_column(einv_df, cols["invoice_date"])
    status_col = utils.resolve_column(einv_df, cols["status"])
    tax_col = utils.resolve_column(einv_df, cols["total_tax"])

    out = pd.DataFrame(
        {
            "k_mst": utils.clean_mst_series(einv_df[mst_col]),
            "k_no": utils.strip_leading_zeros_series(einv_df[no_col]),
            "k_date": utils.parse_date(einv_df[date_col]).dt.strftime("%Y-%m-%d"),
            "Trạng thái hóa đơn": einv_df[status_col].astype(str).str.strip(),
            "Tổng tiền thuế (HĐĐT)": utils.parse_amount_series(einv_df[tax_col]),
        }
    )
    out = out[out["k_no"] != ""]
    out = out.drop_duplicates(subset=["k_mst", "k_no", "k_date"], keep="last")
    return out.reset_index(drop=True)


def classify_status(status: str, status_cfg: dict, not_found_label: str) -> str:
    """Quy trạng thái thô về nhóm chuẩn (thay thế / điều chỉnh / xóa bỏ / OK)."""
    if status is None or str(status).strip() == "" or pd.isna(status):
        return not_found_label
    norm = _norm(status)
    for label, keywords in status_cfg.items():
        for kw in keywords:
            if _norm(kw) in norm:
                return label
    return "Đang hoạt động / hợp lệ"


def enrich_with_einvoice(
    df: pd.DataFrame,
    einv_lookup: pd.DataFrame,
    main_cols: dict,
    status_cfg: dict,
    not_found_label: str,
) -> pd.DataFrame:
    """Ghép dữ liệu HĐĐT vào dữ liệu chính và tính toán.

    Thêm các cột:
      - Trạng thái hóa đơn, Nhóm trạng thái
      - Tổng tiền thuế (HĐĐT)
      - Thuế GTGT (tờ khai)
      - Chênh lệch tiền thuế
      - Cảnh báo HĐ (bị thay thế / xóa bỏ / không tìm thấy)
    """
    df = df.copy()

    no_col = utils.resolve_column(df, main_cols["invoice_no"])
    date_col = utils.resolve_column(df, main_cols["invoice_date"])
    vat_col = utils.resolve_column(df, main_cols["vat"])

    if "MST (chuẩn hóa)" in df.columns:
        mst_key = df["MST (chuẩn hóa)"]
    else:
        mst_key = utils.clean_mst_series(df[utils.resolve_column(df, main_cols["mst"])])

    df["_k_mst"] = mst_key.values
    df["_k_no"] = utils.strip_leading_zeros_series(df[no_col]).values
    df["_k_date"] = utils.parse_date(df[date_col]).dt.strftime("%Y-%m-%d").values

    lookup = einv_lookup.rename(columns={"k_mst": "_k_mst", "k_no": "_k_no", "k_date": "_k_date"})
    merged = df.merge(lookup, on=["_k_mst", "_k_no", "_k_date"], how="left")

    merged["Nhóm trạng thái"] = merged["Trạng thái hóa đơn"].map(
        lambda s: classify_status(s, status_cfg, not_found_label)
    )

    merged["Thuế GTGT (tờ khai)"] = utils.parse_amount_series(merged[vat_col])
    total_tax = merged["Tổng tiền thuế (HĐĐT)"]
    merged["Chênh lệch tiền thuế"] = total_tax.fillna(0) - merged["Thuế GTGT (tờ khai)"]
    # nếu không tìm thấy HĐ thì để chênh lệch trống thay vì âm gây hiểu nhầm
    merged.loc[total_tax.isna(), "Chênh lệch tiền thuế"] = pd.NA

    problem_labels = set(status_cfg.keys()) | {not_found_label}
    merged["Cảnh báo trạng thái HĐ"] = merged["Nhóm trạng thái"].isin(problem_labels)
    merged["Có chênh lệch thuế"] = merged["Chênh lệch tiền thuế"].fillna(0).abs() > 0.5

    merged = merged.drop(columns=["_k_mst", "_k_no", "_k_date"])
    return merged


def status_problems(enriched: pd.DataFrame) -> pd.DataFrame:
    """Lọc các HĐ bị thay thế / xóa bỏ / không tìm thấy."""
    if "Cảnh báo trạng thái HĐ" not in enriched.columns:
        return enriched.iloc[0:0]
    return enriched[enriched["Cảnh báo trạng thái HĐ"] == True].copy()  # noqa: E712


def tax_differences(enriched: pd.DataFrame) -> pd.DataFrame:
    """Lọc các HĐ có chênh lệch tiền thuế."""
    if "Có chênh lệch thuế" not in enriched.columns:
        return enriched.iloc[0:0]
    return enriched[enriched["Có chênh lệch thuế"] == True].copy()  # noqa: E712
