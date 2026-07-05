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
    """Ghép trạng thái + tổng tiền thuế HĐĐT vào dữ liệu chính (theo từng dòng).

    Thêm các cột (mức DÒNG bảng kê):
      - Trạng thái hóa đơn, Nhóm trạng thái
      - Tổng tiền thuế (HĐĐT)  (giá trị cả hóa đơn, lặp lại trên mỗi dòng)
      - Thuế GTGT (tờ khai)     (thuế của riêng dòng đó)
      - Cảnh báo trạng thái HĐ

    Lưu ý: chênh lệch tiền thuế KHÔNG tính ở mức dòng vì bảng kê ghi 1 dòng/mặt
    hàng còn HĐĐT ghi 1 dòng/hóa đơn — dùng ``reconcile_tax`` để đối chiếu ở mức
    hóa đơn (cộng thuế các dòng theo từng hóa đơn).
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

    problem_labels = set(status_cfg.keys()) | {not_found_label}
    merged["Cảnh báo trạng thái HĐ"] = merged["Nhóm trạng thái"].isin(problem_labels)

    merged = merged.drop(columns=["_k_mst", "_k_no", "_k_date"])
    return merged


def reconcile_tax(enriched: pd.DataFrame, main_cols: dict) -> pd.DataFrame:
    """Đối chiếu tiền thuế ở MỨC HÓA ĐƠN.

    Gộp các dòng bảng kê theo (MST, số HĐ, ngày), cộng thuế GTGT tờ khai rồi so
    với tổng tiền thuế trên HĐĐT (giá trị cả hóa đơn).

    Trả về bảng mỗi dòng 1 hóa đơn với cột 'Chênh lệch tiền thuế'.
    """
    if "Tổng tiền thuế (HĐĐT)" not in enriched.columns:
        return enriched.iloc[0:0]

    date_col = utils.resolve_column(enriched, main_cols["invoice_date"])
    tmp = pd.DataFrame(
        {
            "MST (chuẩn hóa)": enriched.get("MST (chuẩn hóa)"),
            "Số HĐ (bỏ 0 đầu)": enriched.get("Số HĐ (bỏ 0 đầu)"),
            "Ngày HĐ": utils.parse_date(enriched[date_col]).dt.strftime("%Y-%m-%d"),
            "Thuế GTGT (tờ khai)": enriched["Thuế GTGT (tờ khai)"],
            "Tổng tiền thuế (HĐĐT)": enriched["Tổng tiền thuế (HĐĐT)"],
            "Trạng thái hóa đơn": enriched.get("Trạng thái hóa đơn"),
            "Nhóm trạng thái": enriched.get("Nhóm trạng thái"),
        }
    )
    grouped = tmp.groupby(["MST (chuẩn hóa)", "Số HĐ (bỏ 0 đầu)", "Ngày HĐ"], dropna=False)
    out = grouped.agg(
        **{
            "Số dòng bảng kê": ("Thuế GTGT (tờ khai)", "size"),
            "Tổng thuế GTGT (tờ khai)": ("Thuế GTGT (tờ khai)", "sum"),
            "Tổng tiền thuế (HĐĐT)": ("Tổng tiền thuế (HĐĐT)", "first"),
            "Trạng thái hóa đơn": ("Trạng thái hóa đơn", "first"),
            "Nhóm trạng thái": ("Nhóm trạng thái", "first"),
        }
    ).reset_index()

    total = out["Tổng tiền thuế (HĐĐT)"]
    out["Chênh lệch tiền thuế"] = total - out["Tổng thuế GTGT (tờ khai)"]
    out.loc[total.isna(), "Chênh lệch tiền thuế"] = pd.NA  # không tìm thấy HĐ
    out["Có chênh lệch thuế"] = out["Chênh lệch tiền thuế"].abs() > 0.5
    return out


def status_problems(enriched: pd.DataFrame, main_cols: dict) -> pd.DataFrame:
    """Danh sách hóa đơn (mức hóa đơn) bị thay thế / xóa bỏ / không tìm thấy."""
    if "Cảnh báo trạng thái HĐ" not in enriched.columns:
        return enriched.iloc[0:0]
    date_col = utils.resolve_column(enriched, main_cols["invoice_date"])
    prob = enriched[enriched["Cảnh báo trạng thái HĐ"] == True].copy()  # noqa: E712
    prob["Ngày HĐ"] = utils.parse_date(prob[date_col]).dt.strftime("%Y-%m-%d")
    cols = ["MST (chuẩn hóa)", "Số HĐ (bỏ 0 đầu)", "Ngày HĐ",
            "Trạng thái hóa đơn", "Nhóm trạng thái"]
    cols = [c for c in cols if c in prob.columns]
    return prob[cols].drop_duplicates().reset_index(drop=True)


def tax_differences(reconciled: pd.DataFrame) -> pd.DataFrame:
    """Lọc các hóa đơn có chênh lệch tiền thuế (từ kết quả reconcile_tax)."""
    if "Có chênh lệch thuế" not in reconciled.columns:
        return reconciled.iloc[0:0]
    return reconciled[reconciled["Có chênh lệch thuế"] == True].copy()  # noqa: E712
