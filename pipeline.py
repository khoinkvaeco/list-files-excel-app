"""
Điều phối toàn bộ quy trình kiểm tra thuế.

Nhận các DataFrame đầu vào + cấu hình, chạy lần lượt 12 tác vụ và trả về
dict các kết quả (mỗi kết quả là 1 DataFrame -> 1 sheet trong file xuất).
"""

from __future__ import annotations

import re

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
    # đếm theo từng nhóm (Không phục vụ SXKD / Quà tặng)
    if isinstance(cfg.GOODS_KEYWORDS, dict) and "Nhóm nghi ngờ" in data.columns:
        for label in cfg.GOODS_KEYWORDS:
            n = data["Nhóm nghi ngờ"].astype(str).str.contains(
                re.escape(label), na=False
            ).sum()
            if n:
                summary[f"  • {label}"] = int(n)

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

    # --- Kiểm tra thuế suất (dùng cho cột AC và sheet KQ) --------------
    from tax_audit import utils as _u, vat_policy

    date_col = _u.resolve_column(data, main_cols["invoice_date"])
    goods_col = _u.resolve_column(data, main_cols["goods"])
    pretax_col = _u.resolve_column(data, main_cols["pretax"])
    vat_col = _u.resolve_column(data, main_cols["vat"])
    _pre = _u.parse_amount_series(data[pretax_col])
    _vat = _u.parse_amount_series(data[vat_col])
    _dates = _u.parse_date(data[date_col])
    _goods = data[goods_col].astype(str)
    _periods = getattr(cfg, "VAT_REDUCED_PERIODS", [])
    _excl = getattr(cfg, "VAT_EXCLUDE_KEYWORDS", {})
    data["Check thuế suất"] = [
        vat_policy.check(d, p, v, g, _periods, _excl)
        for d, p, v, g in zip(_dates.tolist(), _pre.tolist(), _vat.tolist(), _goods.tolist())
    ]

    # --- Sheet "Bảng kê đã xử lý": A-T gốc + 9 cột kết quả U-AC ----------
    bxl = _build_processed_sheet(data, main_cols, reconciled, cfg)

    # --- Sheet "KQ rà soát" theo mẫu (từng loại vi phạm + giải thích) ---
    kq_df = _build_kq_sheet(data, main_cols, cfg)

    # thứ tự sheet: KQ rà soát -> Bảng kê đã xử lý -> dữ liệu tổng hợp -> ...
    ordered = {
        "KQ rà soát": kq_df,
        "Bảng kê đã xử lý": bxl,
        "Dữ liệu tổng hợp": data,
    }
    ordered.update(sheets)

    return {"data": data, "sheets": ordered, "summary": summary}


def _build_kq_sheet(data, main_cols, cfg):
    """Bảng 'KQ rà soát': mỗi dòng 1 loại dấu hiệu + tổng chưa thuế/thuế +
    số hóa đơn (đã bỏ trùng) + giải thích cách tìm ra."""
    from tax_audit import utils as _u

    pretax_col = _u.resolve_column(data, main_cols["pretax"])
    vat_col = _u.resolve_column(data, main_cols["vat"])
    date_col = _u.resolve_column(data, main_cols["invoice_date"])

    pre = _u.parse_amount_series(data[pretax_col])
    vat = _u.parse_amount_series(data[vat_col])
    mst = (
        data["MST (chuẩn hóa)"]
        if "MST (chuẩn hóa)" in data.columns
        else _u.clean_mst_series(data[_u.resolve_column(data, main_cols["mst"])])
    ).astype(str)
    no = (
        data["Số HĐ (bỏ 0 đầu)"]
        if "Số HĐ (bỏ 0 đầu)" in data.columns
        else _u.strip_leading_zeros_series(data[_u.resolve_column(data, main_cols["invoice_no"])])
    ).astype(str)
    datestr = _u.parse_date(data[date_col]).dt.strftime("%Y-%m-%d").fillna("")
    invkey = mst.str.strip() + "|" + no.str.strip() + "|" + datestr
    has_id = (mst.str.strip() != "") | (no.str.strip() != "")

    def agg(mask):
        m = mask.fillna(False) if hasattr(mask, "fillna") else mask
        p = float(pre[m].sum())
        v = float(vat[m].sum())
        n = int(invkey[m & has_id].nunique())
        return round(p), round(v), n

    rows = []
    all_mask = pd.Series(True, index=data.index)
    p, v, n = agg(all_mask)
    rows.append(["Bảng kê GTGT (tổng)", p, v, n, "Tổng toàn bộ bảng kê"])

    # Theo trạng thái người nộp thuế (TMS)
    if "Trạng thái NNT" in data.columns:
        stt = data["Trạng thái NNT"].map(lambda x: "" if pd.isna(x) else str(x).strip())
        for val in sorted(set(stt)):
            low = val.lower()
            if low in ("", "nan", "none") or "đang hoạt động" in low:
                continue
            p, v, n = agg(stt == val)
            rows.append([val, p, v, n, "Từ trạng thái NNT (tra TMS theo MST)"])

    # DN rủi ro
    if "DN rủi ro" in data.columns:
        p, v, n = agg(data["DN rủi ro"] == True)  # noqa: E712
        if n or p or v:
            rows.append(["Hóa đơn của DN rủi ro", p, v, n,
                         "Từ dò MST với danh sách DN có dấu hiệu rủi ro"])

    # Kê khai trùng
    if "HĐ kê khai trùng" in data.columns:
        p, v, n = agg(data["HĐ kê khai trùng"] == True)  # noqa: E712
        if n:
            rows.append(["Hóa đơn kê khai trùng", p, v, n,
                         "Từ trùng số HĐ + ngày + MST + giá trị chưa thuế"])

    # Mặt hàng nghi ngờ theo nhóm
    if "Nhóm nghi ngờ" in data.columns and isinstance(cfg.GOODS_KEYWORDS, dict):
        grp = data["Nhóm nghi ngờ"].astype(str)
        for label in cfg.GOODS_KEYWORDS:
            mask = grp.str.contains(re.escape(label), na=False)
            if mask.any():
                p, v, n = agg(mask)
                rows.append([label, p, v, n, f'Từ loại hàng hóa "{label}"'])

    # Trạng thái hóa đơn điện tử
    if "Nhóm trạng thái" in data.columns:
        grp = data["Nhóm trạng thái"].map(lambda x: "" if pd.isna(x) else str(x).strip())
        problem = set(cfg.INVOICE_STATUS.keys()) | {cfg.STATUS_NOT_FOUND}
        for val in sorted(set(grp) & problem):
            p, v, n = agg(grp == val)
            rows.append([f"Hóa đơn {val}", p, v, n, "Từ trạng thái hóa đơn điện tử"])

    # Rà soát thuế suất 10%/8%
    if "Check thuế suất" in data.columns:
        chk = data["Check thuế suất"].astype(str)
        mask = chk.str.contains("SAI", na=False) | chk.str.contains("rà lại", na=False)
        if mask.any():
            p, v, n = agg(mask)
            rows.append(["Rà soát thuế suất 10%/8%", p, v, n,
                         "Từ kiểm tra thuế suất theo ngày HĐ + nhóm hàng (chính sách giảm 8%)"])

    return pd.DataFrame(
        rows,
        columns=["Chỉ tiêu", "Giá trị chưa thuế", "Thuế GTGT",
                 "Số hóa đơn (đã bỏ trùng)", "Giải thích"],
    )


def _build_processed_sheet(data, main_cols, reconciled, cfg):
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

    # U (21): Loại hàng hóa nghi ngờ (tên nhóm: Không phục vụ SXKD / Quà tặng)
    loai_hang = col("Nhóm nghi ngờ")
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

    # AC (29): check 10%, 8% — đã tính sẵn ở cột "Check thuế suất"
    check_rate = col("Check thuế suất").tolist()

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
