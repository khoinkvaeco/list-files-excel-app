"""
Tác vụ 8 & 9: Chuẩn hóa số hóa đơn (xóa số 0 đầu) và tìm hóa đơn kê khai trùng.

Khóa trùng: (Số HĐ đã bỏ số 0 đầu, Ngày HĐ, MST, Giá trị chưa thuế).
"""

from __future__ import annotations

import pandas as pd

from . import utils


def normalize_invoice_no(df: pd.DataFrame, invoice_no_col) -> pd.DataFrame:
    """Thêm cột 'Số HĐ (bỏ 0 đầu)' từ cột số hóa đơn gốc."""
    df = df.copy()
    col = utils.resolve_column(df, invoice_no_col)
    df["Số HĐ (bỏ 0 đầu)"] = utils.strip_leading_zeros_series(df[col])
    return df


def find_duplicates(
    df: pd.DataFrame,
    invoice_no_col,
    invoice_date_col,
    mst_col,
    pretax_col,
) -> pd.DataFrame:
    """Đánh dấu các hóa đơn kê khai trùng.

    Thêm cột:
      - Nhóm trùng (ID nhóm, để trống nếu không trùng)
      - Số bản ghi trùng
      - HĐ kê khai trùng (bool)
    """
    df = df.copy()

    no_col = utils.resolve_column(df, invoice_no_col)
    date_col = utils.resolve_column(df, invoice_date_col)
    mst_c = utils.resolve_column(df, mst_col)
    pretax_c = utils.resolve_column(df, pretax_col)

    key = pd.DataFrame(
        {
            "k_no": utils.strip_leading_zeros_series(df[no_col]),
            "k_date": utils.parse_date(df[date_col]).dt.strftime("%Y-%m-%d"),
            "k_mst": utils.clean_mst_series(df[mst_c]),
            "k_pretax": utils.parse_amount_series(df[pretax_c]).round(0),
        }
    )
    # chỉ xét các dòng có đủ thông tin định danh
    valid = (key["k_no"] != "") & (key["k_mst"] != "")

    group_id = pd.Series([""] * len(df), index=df.index, dtype=object)
    dup_count = pd.Series([0] * len(df), index=df.index, dtype=int)

    grouped = key[valid].groupby(["k_no", "k_date", "k_mst", "k_pretax"])
    gid = 0
    for _, idx in grouped.groups.items():
        if len(idx) > 1:
            gid += 1
            group_id.loc[idx] = f"TRÙNG-{gid:03d}"
            dup_count.loc[idx] = len(idx)

    df["Nhóm trùng"] = group_id
    df["Số bản ghi trùng"] = dup_count
    df["HĐ kê khai trùng"] = group_id != ""
    return df


def duplicate_invoices(flagged: pd.DataFrame) -> pd.DataFrame:
    """Lọc các dòng thuộc nhóm trùng, sắp xếp theo nhóm."""
    if "HĐ kê khai trùng" not in flagged.columns:
        return flagged.iloc[0:0]
    out = flagged[flagged["HĐ kê khai trùng"] == True].copy()  # noqa: E712
    return out.sort_values("Nhóm trùng").reset_index(drop=True)
