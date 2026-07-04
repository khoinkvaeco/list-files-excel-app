"""
Ứng dụng web (Streamlit) hỗ trợ kiểm tra thuế trên file Excel.

Chạy:  streamlit run app.py
"""

from __future__ import annotations

import copy

import pandas as pd
import streamlit as st

import config as C
from tax_audit import utils
from pipeline import run_pipeline

st.set_page_config(page_title="Công cụ kiểm tra thuế", page_icon="📊", layout="wide")

st.title("📊 Công cụ hỗ trợ kiểm tra thuế")
st.caption(
    "Tải lên bảng kê hóa đơn cần kiểm tra và các file tra cứu, "
    "công cụ sẽ tự động thực hiện các bước rà soát và xuất ra file Excel kết quả."
)


# ---------------------------------------------------------------------------
# CẤU HÌNH CỘT (có thể chỉnh trực tiếp)
# ---------------------------------------------------------------------------
def build_runtime_config():
    """Tạo bản sao config, cho phép người dùng chỉnh vị trí cột trên sidebar."""
    cfg = copy.deepcopy(C)

    with st.sidebar:
        st.header("⚙️ Cấu hình cột")
        st.caption(
            "Nhập TÊN cột hoặc SỐ THỨ TỰ cột (đếm từ 1). "
            "Để mặc định nếu file đúng bố cục chuẩn."
        )

        with st.expander("File dữ liệu chính", expanded=False):
            for k, label in [
                ("mst", "MST người bán"),
                ("invoice_no", "Số hóa đơn"),
                ("invoice_date", "Ngày hóa đơn"),
                ("pretax", "Giá trị chưa thuế"),
                ("vat", "Thuế GTGT"),
                ("goods", "Tên hàng hóa"),
            ]:
                cfg.MAIN["cols"][k] = _col_input(f"main_{k}", label, C.MAIN["cols"][k])
            cfg.MAIN["header_row"] = st.number_input(
                "Dòng tiêu đề", 1, 50, C.MAIN["header_row"], key="main_hdr"
            )

        with st.expander("File TMS", expanded=False):
            cfg.TMS["mst_col"] = _col_input("tms_mst", "Cột MST", C.TMS["mst_col"])
            cfg.TMS["status_col"] = _col_input(
                "tms_status", "Cột trạng thái NNT (mặc định 42)", C.TMS["status_col"]
            )
            cfg.TMS["close_date_col"] = _col_input(
                "tms_close", "Cột ngày đóng trạng thái (mặc định 51)", C.TMS["close_date_col"]
            )
            cfg.TMS["header_row"] = st.number_input(
                "Dòng tiêu đề (TMS)", 1, 50, C.TMS["header_row"], key="tms_hdr"
            )

        with st.expander("File DS rủi ro", expanded=False):
            cfg.RISK["mst_col"] = _col_input("risk_mst", "Cột MST", C.RISK["mst_col"])
            cfg.RISK["doc_col"] = _col_input(
                "risk_doc", "Cột văn bản (mặc định 5)", C.RISK["doc_col"]
            )
            cfg.RISK["header_row"] = st.number_input(
                "Dòng tiêu đề (rủi ro)", 1, 50, C.RISK["header_row"], key="risk_hdr"
            )

        with st.expander("File hóa đơn điện tử", expanded=False):
            for k, label in [
                ("mst", "MST người bán"),
                ("invoice_no", "Số hóa đơn"),
                ("invoice_date", "Ngày hóa đơn"),
                ("status", "Trạng thái hóa đơn"),
                ("total_tax", "Tổng tiền thuế"),
            ]:
                cfg.EINVOICE["cols"][k] = _col_input(
                    f"einv_{k}", label, C.EINVOICE["cols"][k]
                )
            cfg.EINVOICE["header_row"] = st.number_input(
                "Dòng tiêu đề (HĐĐT)", 1, 50, C.EINVOICE["header_row"], key="einv_hdr"
            )

        with st.expander("Từ khóa mặt hàng nghi ngờ", expanded=False):
            kw = st.text_area(
                "Mỗi từ khóa 1 dòng",
                value="\n".join(C.GOODS_KEYWORDS),
                key="kw",
                height=180,
            )
            cfg.GOODS_KEYWORDS = [x.strip() for x in kw.splitlines() if x.strip()]

    return cfg


def _col_input(key: str, label: str, default):
    """Ô nhập tham chiếu cột: nếu là số -> vị trí, nếu là chữ -> tên."""
    raw = st.text_input(label, value=str(default), key=key)
    raw = raw.strip()
    if raw.isdigit():
        return int(raw)
    return raw


# ---------------------------------------------------------------------------
# TẢI FILE
# ---------------------------------------------------------------------------
cfg = build_runtime_config()

st.subheader("1️⃣ Tải file lên")
col1, col2 = st.columns(2)
with col1:
    main_file = st.file_uploader(
        "📄 File dữ liệu chính (bảng kê hóa đơn cần kiểm tra) — bắt buộc",
        type=["xlsx", "xls", "csv"],
    )
    tms_file = st.file_uploader(
        "🏢 File TMS (trạng thái người nộp thuế) — tùy chọn",
        type=["xlsx", "xls", "csv"],
    )
with col2:
    risk_file = st.file_uploader(
        "⚠️ File danh sách DN có dấu hiệu rủi ro — tùy chọn",
        type=["xlsx", "xls", "csv"],
    )
    einv_file = st.file_uploader(
        "🧾 File hóa đơn điện tử — tùy chọn",
        type=["xlsx", "xls", "csv"],
    )


def _load(file, header_row):
    if file is None:
        return None
    return utils.read_excel(file, header_row=header_row)


st.subheader("2️⃣ Chạy kiểm tra")
if st.button("▶️ Bắt đầu kiểm tra", type="primary", disabled=main_file is None):
    if main_file is None:
        st.warning("Vui lòng tải lên file dữ liệu chính.")
        st.stop()

    try:
        main_df = _load(main_file, cfg.MAIN["header_row"])
        tms_df = _load(tms_file, cfg.TMS["header_row"])
        risk_df = _load(risk_file, cfg.RISK["header_row"])
        einv_df = _load(einv_file, cfg.EINVOICE["header_row"])
    except Exception as e:  # noqa: BLE001
        st.error(f"Lỗi khi đọc file: {e}")
        st.stop()

    try:
        result = run_pipeline(main_df, tms_df, risk_df, einv_df, cfg=cfg)
    except Exception as e:  # noqa: BLE001
        st.error(f"Lỗi trong quá trình xử lý: {e}")
        st.exception(e)
        st.stop()

    st.session_state["result"] = result

# ---------------------------------------------------------------------------
# HIỂN THỊ KẾT QUẢ
# ---------------------------------------------------------------------------
if "result" in st.session_state:
    result = st.session_state["result"]
    summary = result["summary"]
    sheets = result["sheets"]

    st.subheader("3️⃣ Kết quả")

    # bảng chỉ tiêu tổng hợp
    if summary:
        cols = st.columns(min(len(summary), 4))
        for i, (k, v) in enumerate(summary.items()):
            cols[i % len(cols)].metric(k, v)

    # tải file kết quả
    excel_bytes = utils.to_excel_bytes(sheets)
    st.download_button(
        "⬇️ Tải file Excel kết quả (nhiều sheet)",
        data=excel_bytes,
        file_name="ket_qua_kiem_tra_thue.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    # xem trước từng sheet
    st.markdown("#### Xem trước các sheet kết quả")
    tab_names = list(sheets.keys())
    tabs = st.tabs(tab_names)
    for tab, name in zip(tabs, tab_names):
        with tab:
            df = sheets[name]
            st.caption(f"{len(df):,} dòng")
            st.dataframe(df.head(500), use_container_width=True)
else:
    st.info("Tải file dữ liệu chính và nhấn **Bắt đầu kiểm tra** để xem kết quả.")
