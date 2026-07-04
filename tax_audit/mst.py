"""
Tác vụ 1 & 2: Xử lý MST và tạo sheet danh sách MST để tra cứu TMS.
"""

from __future__ import annotations

import pandas as pd

from . import utils


def process_mst(df: pd.DataFrame, mst_col) -> pd.DataFrame:
    """Chuẩn hóa cột MST ngay trên dữ liệu chính.

    Thêm cột ``MST (chuẩn hóa)`` bên cạnh cột MST gốc.
    """
    df = df.copy()
    col = utils.resolve_column(df, mst_col)
    df["MST (chuẩn hóa)"] = utils.clean_mst_series(df[col])
    return df


def build_mst_sheet(df: pd.DataFrame, mst_col) -> pd.DataFrame:
    """Tạo sheet danh sách MST duy nhất (đã chuẩn hóa) để copy sang tra TMS.

    Trả về DataFrame gồm: MST, Số lần xuất hiện.
    """
    col = utils.resolve_column(df, mst_col)
    mst = utils.clean_mst_series(df[col])
    mst = mst[mst != ""]
    counts = mst.value_counts()
    out = (
        counts.rename_axis("MST")
        .reset_index(name="Số dòng hóa đơn")
        .sort_values("MST")
        .reset_index(drop=True)
    )
    return out
