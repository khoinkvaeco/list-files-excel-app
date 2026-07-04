"""
Điều phối toàn bộ quy trình kiểm tra thuế.

Nhận các DataFrame đầu vào + cấu hình, chạy lần lượt 12 tác vụ và trả về
dict các kết quả (mỗi kết quả là 1 DataFrame -> 1 sheet trong file xuất).
"""

from __future__ import annotations

import pandas as pd

import config as C
from tax_audit import mst, taxpayer, risk, goods, invoice, einvoice


def run_pipeline(
    main_df: pd.DataFrame,
    tms_df: pd.DataFrame | None = None,
    risk_df: pd.DataFrame | None = None,
    einv_df: pd.DataFrame | None = None,
    cfg=None,
) -> dict:
    """Chạy quy trình. File TMS/rủi ro/HĐĐT là tùy chọn — thiếu file nào
    thì bỏ qua các tác vụ tương ứng.

    Trả về dict:
      {
        "data": DataFrame dữ liệu chính đã bổ sung mọi cột phân tích,
        "sheets": {tên sheet -> DataFrame} để xuất Excel,
        "summary": {tên chỉ tiêu -> số lượng},
      }
    """
    cfg = cfg or C
    main_cols = cfg.MAIN["cols"]
    summary: dict[str, int] = {}
    sheets: dict[str, pd.DataFrame] = {}

    # --- Tác vụ 1 & 8: xử lý MST + chuẩn hóa số hóa đơn -----------------
    data = mst.process_mst(main_df, main_cols["mst"])
    data = invoice.normalize_invoice_no(data, main_cols["invoice_no"])

    # --- Tác vụ 2: sheet danh sách MST để tra TMS ----------------------
    mst_sheet = mst.build_mst_sheet(main_df, main_cols["mst"])
    sheets["MST tra cứu"] = mst_sheet
    summary["Số MST duy nhất"] = len(mst_sheet)

    # --- Tác vụ 3, 4, 5: trạng thái NNT từ TMS -------------------------
    if tms_df is not None:
        tms_lookup = taxpayer.build_tms_lookup(tms_df, cfg.TMS)
        data = taxpayer.enrich_with_taxpayer_status(
            data, tms_lookup, main_cols["invoice_date"]
        )
        sheets["Trạng thái NNT (TMS)"] = tms_lookup
        after = taxpayer.invoices_after_close(data)
        sheets["HĐ xuất sau ngày đóng"] = after
        summary["HĐ xuất sau ngày đóng trạng thái"] = len(after)

    # --- Tác vụ 6: dò danh sách rủi ro ---------------------------------
    if risk_df is not None:
        risk_lookup = risk.build_risk_lookup(risk_df, cfg.RISK)
        data = risk.enrich_with_risk(data, risk_lookup)
        risky = risk.risky_invoices(data)
        sheets["HĐ của DN rủi ro"] = risky
        summary["HĐ của DN có dấu hiệu rủi ro"] = len(risky)

    # --- Tác vụ 7: mặt hàng không phục vụ SXKD -------------------------
    data = goods.flag_non_business_goods(
        data, main_cols["goods"], cfg.GOODS_KEYWORDS
    )
    suspicious = goods.suspicious_goods(data)
    sheets["Mặt hàng nghi ngờ"] = suspicious
    summary["Dòng mặt hàng nghi ngờ"] = len(suspicious)

    # --- Tác vụ 9: hóa đơn kê khai trùng -------------------------------
    data = invoice.find_duplicates(
        data,
        main_cols["invoice_no"],
        main_cols["invoice_date"],
        main_cols["mst"],
        main_cols["pretax"],
    )
    dups = invoice.duplicate_invoices(data)
    sheets["HĐ kê khai trùng"] = dups
    summary["Dòng HĐ kê khai trùng"] = len(dups)

    # --- Tác vụ 10, 11, 12: tra HĐĐT -----------------------------------
    if einv_df is not None:
        einv_lookup = einvoice.build_einvoice_lookup(einv_df, cfg.EINVOICE)
        data = einvoice.enrich_with_einvoice(
            data,
            einv_lookup,
            main_cols,
            cfg.INVOICE_STATUS,
            cfg.STATUS_NOT_FOUND,
        )
        problems = einvoice.status_problems(data)
        diffs = einvoice.tax_differences(data)
        sheets["HĐ sai trạng thái"] = problems
        sheets["Chênh lệch tiền thuế"] = diffs
        summary["HĐ bị thay thế/xóa bỏ/không tìm thấy"] = len(problems)
        summary["HĐ có chênh lệch tiền thuế"] = len(diffs)

    # sheet dữ liệu tổng hợp đặt lên đầu
    ordered = {"Dữ liệu tổng hợp": data}
    ordered.update(sheets)

    return {"data": data, "sheets": ordered, "summary": summary}
