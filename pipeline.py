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
    risk_file=None,
    einv_df: pd.DataFrame | None = None,
    cfg=None,
) -> dict:
    """Chạy quy trình. File TMS/rủi ro/HĐĐT là tùy chọn — thiếu file nào
    thì bỏ qua các tác vụ tương ứng.

    ``risk_file`` là ĐỐI TƯỢNG FILE (không phải DataFrame) vì danh sách rủi ro
    gồm nhiều sheet cần đọc riêng.

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

    # Loại bỏ các dòng KHÔNG phải hóa đơn (dòng "Tổng cộng", dòng trống ở cuối
    # bảng kê): giữ lại dòng có MST hoặc có số hóa đơn. Tránh dòng tổng cộng làm
    # nhân đôi tiền thuế và gây sai các chỉ tiêu.
    has_mst = data["MST (chuẩn hóa)"].astype(str).str.strip() != ""
    has_no = data["Số HĐ (bỏ 0 đầu)"].astype(str).str.strip() != ""
    data = data[has_mst | has_no].reset_index(drop=True)

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
    if risk_file is not None:
        risk_lookup = risk.build_risk_lookup(risk_file, cfg.RISK)
        data = risk.enrich_with_risk(data, risk_lookup)
        risky = risk.risky_invoices(data)
        sheets["DS MST rủi ro"] = risk_lookup
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
    # bỏ qua nếu file HĐĐT rỗng (không có dòng dữ liệu) để tránh gắn nhầm
    # nhãn "không tìm thấy" cho toàn bộ hóa đơn.
    if einv_df is not None and not einv_df.empty:
        einv_lookup = einvoice.build_einvoice_lookup(einv_df, cfg.EINVOICE)
        data = einvoice.enrich_with_einvoice(
            data,
            einv_lookup,
            main_cols,
            cfg.INVOICE_STATUS,
            cfg.STATUS_NOT_FOUND,
        )
        problems = einvoice.status_problems(data, main_cols)
        reconciled = einvoice.reconcile_tax(data, main_cols)
        diffs = einvoice.tax_differences(reconciled)
        sheets["HĐ sai trạng thái"] = problems
        sheets["Đối chiếu thuế theo HĐ"] = reconciled
        sheets["Chênh lệch tiền thuế"] = diffs
        summary["HĐ bị thay thế/xóa bỏ/không tìm thấy"] = len(problems)
        summary["HĐ có chênh lệch tiền thuế"] = len(diffs)

    # --- Sheet tổng hợp kết quả rà soát (giống "KQ rà BKMV") ------------
    from tax_audit import utils as _u

    pretax_col = _u.resolve_column(data, main_cols["pretax"])
    vat_col = _u.resolve_column(data, main_cols["vat"])
    tong_chua_thue = float(_u.parse_amount_series(data[pretax_col]).sum())
    tong_thue = float(_u.parse_amount_series(data[vat_col]).sum())

    kq_rows = [
        ("Số dòng bảng kê", len(data)),
        ("Số hóa đơn (số MST × số HĐ)", summary.get("Số MST duy nhất", "")),
        ("Tổng giá trị chưa thuế (bảng kê)", round(tong_chua_thue)),
        ("Tổng tiền thuế GTGT (bảng kê)", round(tong_thue)),
    ]
    # đưa toàn bộ chỉ tiêu cảnh báo vào bảng tổng hợp
    for k, v in summary.items():
        if k != "Số MST duy nhất":
            kq_rows.append((k, v))
    kq_df = pd.DataFrame(kq_rows, columns=["Chỉ tiêu", "Giá trị"])

    # sheet KQ rà soát + dữ liệu tổng hợp đặt lên đầu
    ordered = {"KQ rà soát": kq_df, "Dữ liệu tổng hợp": data}
    ordered.update(sheets)

    return {"data": data, "sheets": ordered, "summary": summary}
