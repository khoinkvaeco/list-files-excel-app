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

    reconciled = None

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

    # --- Sheet "Bảng kê đã xử lý": A-T gốc + 9 cột kết quả U-AC ----------
    from tax_audit import utils as _u

    bxl = _build_processed_sheet(data, main_cols, reconciled)

    # --- Sheet tổng hợp kết quả rà soát (giống "KQ rà BKMV") ------------

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

    # thứ tự sheet: KQ rà soát -> Bảng kê đã xử lý -> dữ liệu tổng hợp -> ...
    ordered = {
        "KQ rà soát": kq_df,
        "Bảng kê đã xử lý": bxl,
        "Dữ liệu tổng hợp": data,
    }
    ordered.update(sheets)

    return {"data": data, "sheets": ordered, "summary": summary}


# Các giai đoạn được giảm thuế GTGT 10% -> 8% (đã gộp khoảng liền kề).
# Ngoài các khoảng này, thuế suất phổ thông là 10%.
VAT_REDUCED_PERIODS = [
    ("2022-02-01", "2022-12-31"),  # NĐ 15/2022
    ("2023-07-01", "2023-12-31"),  # NĐ 44/2023
    ("2024-01-01", "2024-12-31"),  # NĐ 94/2023 + NĐ 72/2024
    ("2025-07-01", "2026-12-31"),  # NĐ 174/2024 (174/2025)
]


def _infer_rate(pretax_val: float, vat_val: float):
    """Suy ra thuế suất từ (tiền thuế / giá trị chưa thuế). None nếu không rõ."""
    if not pretax_val:
        return None
    r = vat_val / pretax_val
    for target in (0.10, 0.08, 0.05, 0.0):
        if abs(r - target) < 0.005:
            return target
    return round(r, 4)


def _in_reduced_period(ts) -> bool:
    import pandas as _pd

    if ts is None or _pd.isna(ts):
        return False
    for start, end in VAT_REDUCED_PERIODS:
        if _pd.Timestamp(start) <= ts <= _pd.Timestamp(end):
            return True
    return False


def _vat_policy_check(ts, pretax_val: float, vat_val: float) -> str:
    """Đối chiếu thuế suất thực tế với chính sách giảm 8% theo ngày hóa đơn."""
    rate = _infer_rate(pretax_val, vat_val)
    if rate is None:
        return ""
    reduced = _in_reduced_period(ts)
    if abs(rate - 0.10) < 0.005:
        return "10% - rà lại (đang kỳ giảm 8%)" if reduced else "10% - đúng"
    if abs(rate - 0.08) < 0.005:
        return "8% - đúng kỳ giảm" if reduced else "8% - NGOÀI kỳ giảm (?)"
    if abs(rate - 0.05) < 0.005:
        return "5%"
    if abs(rate - 0.0) < 0.005:
        return "0% / không thuế"
    return f"{rate * 100:.1f}% - khác"


def _build_processed_sheet(data, main_cols, reconciled):
    """Tạo 'Bảng kê đã xử lý' = cột A-T gốc + 9 cột kết quả U-AC tự điền."""
    from tax_audit import utils as _u

    n = len(data)

    def col(name, default=""):
        return data[name] if name in data.columns else pd.Series([default] * n, index=data.index)

    # cột A-T (20 cột đầu của bảng kê gốc)
    orig = data.iloc[:, : min(20, data.shape[1])].copy()

    pretax_col = _u.resolve_column(data, main_cols["pretax"])
    vat_col = _u.resolve_column(data, main_cols["vat"])
    pre = _u.parse_amount_series(data[pretax_col])
    vatv = _u.parse_amount_series(data[vat_col])

    # U (21): Loại hàng hóa nghi ngờ
    loai_hang = col("Từ khóa khớp")
    # V (22): Trạng thái NNT
    trang_thai = col("Trạng thái NNT")
    # W (23): ngày liên quan (ngày đóng trạng thái)
    ngay_lq = col("Ngày đóng trạng thái")
    if "Ngày đóng trạng thái" in data.columns:
        ngay_lq = _u.parse_date(data["Ngày đóng trạng thái"]).dt.strftime("%d/%m/%Y").fillna("")
    # X (24): Hóa đơn rủi ro
    rui_ro = col("Văn bản rủi ro")
    # Y (25): Kê khai trùng
    ke_khai_trung = col("Nhóm trùng")
    # Z (26): tra HĐĐT (trạng thái)
    tra_hddt = col("Nhóm trạng thái")
    # AA (27): tiền thuế trên HĐĐT
    tien_thue_hddt = col("Tổng tiền thuế (HĐĐT)")
    aa = _u.parse_amount_series(tien_thue_hddt) if "Tổng tiền thuế (HĐĐT)" in data.columns else None

    # AB (28): Chênh lệch với BK = AA - S (tiền thuế HĐĐT - tiền thuế GTGT bảng kê)
    if aa is not None and "Tổng tiền thuế (HĐĐT)" in data.columns:
        raw = data["Tổng tiền thuế (HĐĐT)"]
        chenh_lech = (aa.values - vatv.values)
        # để trống khi không tra được HĐ trên HĐĐT
        chenh_lech = pd.Series(chenh_lech, index=data.index).where(raw.notna().values, pd.NA)
    else:
        chenh_lech = pd.Series([pd.NA] * n, index=data.index)

    # AC (29): check 10%, 8% theo chính sách giảm thuế + ngày hóa đơn
    date_col = _u.resolve_column(data, main_cols["invoice_date"])
    inv_dates = _u.parse_date(data[date_col])
    check_rate = [
        _vat_policy_check(d, p, v)
        for d, p, v in zip(inv_dates.tolist(), pre.tolist(), vatv.tolist())
    ]

    result = pd.DataFrame({
        "Loại hàng hóa": loai_hang.values if hasattr(loai_hang, "values") else loai_hang,
        "Trạng thái": trang_thai.values if hasattr(trang_thai, "values") else trang_thai,
        "Ngày liên quan": ngay_lq.values if hasattr(ngay_lq, "values") else ngay_lq,
        "Hóa đơn rủi ro": rui_ro.values if hasattr(rui_ro, "values") else rui_ro,
        "Kê khai trùng": ke_khai_trung.values if hasattr(ke_khai_trung, "values") else ke_khai_trung,
        "Tra HĐĐT": tra_hddt.values if hasattr(tra_hddt, "values") else tra_hddt,
        "Tiền thuế trên HĐĐT": tien_thue_hddt.values if hasattr(tien_thue_hddt, "values") else tien_thue_hddt,
        "Chênh lệch với BK": chenh_lech,
        "Check 10%, 8%": check_rate,
    }, index=data.index)

    return pd.concat([orig, result], axis=1)
