"""
Tác vụ 3, 4, 5: Tra cứu trạng thái người nộp thuế từ file TMS.

- Cột 42: Trạng thái người nộp thuế.
- Cột 51: Ngày đóng trạng thái tổ chức (ngày ngừng/không hoạt động).
- Cảnh báo hóa đơn xuất SAU ngày đóng trạng thái.
"""

from __future__ import annotations

import pandas as pd

from . import utils


def build_tms_lookup(tms_df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """Từ file TMS, trích ra bảng tra cứu: MST -> trạng thái, ngày đóng.

    ``cfg`` là dict TMS trong config (mst_col, status_col, close_date_col).
    """
    mst_col = utils.resolve_column(tms_df, cfg["mst_col"])
    status_col = utils.resolve_column(tms_df, cfg["status_col"])
    close_col = utils.resolve_column(tms_df, cfg["close_date_col"])

    out = pd.DataFrame(
        {
            "MST (chuẩn hóa)": utils.clean_mst_series(tms_df[mst_col]),
            "Trạng thái NNT": tms_df[status_col].astype(str).str.strip(),
            "Ngày đóng trạng thái": utils.parse_date(tms_df[close_col]),
        }
    )
    out = out[out["MST (chuẩn hóa)"] != ""]
    # 1 MST có thể xuất hiện nhiều dòng -> giữ dòng có ngày đóng mới nhất
    out = (
        out.sort_values("Ngày đóng trạng thái")
        .drop_duplicates(subset=["MST (chuẩn hóa)"], keep="last")
        .reset_index(drop=True)
    )
    return out


def enrich_with_taxpayer_status(
    df: pd.DataFrame,
    tms_lookup: pd.DataFrame,
    invoice_date_col,
) -> pd.DataFrame:
    """Ghép trạng thái NNT + ngày đóng vào dữ liệu chính và tạo cảnh báo.

    Yêu cầu ``df`` đã có cột 'MST (chuẩn hóa)'. Thêm các cột:
      - Trạng thái NNT
      - Ngày đóng trạng thái
      - Cảnh báo HĐ sau ngày đóng (bool)
      - Số ngày xuất sau khi đóng
    """
    df = df.copy()
    if "MST (chuẩn hóa)" not in df.columns:
        raise ValueError("Cần chạy xử lý MST trước (thiếu cột 'MST (chuẩn hóa)').")

    merged = df.merge(tms_lookup, on="MST (chuẩn hóa)", how="left")

    inv_col = utils.resolve_column(df, invoice_date_col)
    inv_date = utils.parse_date(merged[inv_col])
    close_date = merged["Ngày đóng trạng thái"]

    after = inv_date.notna() & close_date.notna() & (inv_date > close_date)
    merged["Cảnh báo HĐ sau ngày đóng"] = after
    # Số ngày xuất sau khi đóng = Ngày đóng trạng thái - Ngày hóa đơn
    merged["Số ngày xuất sau khi đóng"] = (close_date - inv_date).dt.days.where(after)

    return merged


def invoices_after_close(enriched: pd.DataFrame) -> pd.DataFrame:
    """Lọc ra các dòng hóa đơn xuất sau ngày đóng trạng thái NNT."""
    if "Cảnh báo HĐ sau ngày đóng" not in enriched.columns:
        return enriched.iloc[0:0]
    return enriched[enriched["Cảnh báo HĐ sau ngày đóng"] == True].copy()  # noqa: E712
