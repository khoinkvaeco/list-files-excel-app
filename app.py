"""
Ứng dụng web (Streamlit) hỗ trợ kiểm tra thuế trên file Excel.

Chạy:  streamlit run app.py
"""

from __future__ import annotations

import copy
import types

import pandas as pd
import streamlit as st

import config as C
from tax_audit import utils, xml_invoice
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
    """Tạo bản sao config, cho phép người dùng chỉnh vị trí cột trên sidebar.

    Không deepcopy cả module ``config`` (module không thể copy/pickle) — chỉ sao
    chép các giá trị dữ liệu cần thiết vào một namespace có thể chỉnh sửa.
    """
    cfg = types.SimpleNamespace(
        MAIN=copy.deepcopy(C.MAIN),
        TMS=copy.deepcopy(C.TMS),
        RISK=copy.deepcopy(C.RISK),
        EINVOICE=copy.deepcopy(C.EINVOICE),
        GOODS_KEYWORDS=list(C.GOODS_KEYWORDS),
        INVOICE_STATUS=copy.deepcopy(C.INVOICE_STATUS),
        STATUS_NOT_FOUND=C.STATUS_NOT_FOUND,
    )

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

        with st.expander("File DS rủi ro (nhiều sheet)", expanded=False):
            st.caption(
                "Danh sách rủi ro gồm nhiều sheet, mỗi sheet một bố cục. "
                "Chỉnh tên sheet / dòng tiêu đề / cột MST / cột (hoặc nhãn) văn bản:"
            )
            for i, spec in enumerate(cfg.RISK["sheets"]):
                st.markdown(f"**Sheet {i + 1}**")
                spec["name"] = st.text_input(
                    "Tên sheet", value=spec["name"], key=f"risk_name_{i}"
                )
                spec["header_row"] = st.number_input(
                    "Dòng tiêu đề", 1, 50, spec["header_row"], key=f"risk_hdr_{i}"
                )
                spec["mst_col"] = _col_input(
                    f"risk_mst_{i}", "Cột MST", spec["mst_col"]
                )
                spec["doc_col"] = _col_input(
                    f"risk_doc_{i}", "Cột văn bản (số) hoặc nhãn văn bản (chữ)",
                    spec["doc_col"],
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


def _load(file, header_row, sheet_name=0):
    if file is None:
        return None
    return utils.read_excel(file, header_row=header_row, sheet_name=sheet_name)


def _default_sheet_index(sheets: list[str]) -> int:
    """Đoán sheet chứa dữ liệu bảng kê (ưu tiên tên chứa 'BKMV' hoặc bắt đầu 'BK')."""
    upper = [str(s).upper() for s in sheets]
    for i, s in enumerate(upper):
        if "BKMV" in s or "BK M" in s or s.startswith("BK"):
            return i
    return 0


# ---------------------------------------------------------------------------
# TAB 1: KIỂM TRA BẢNG KÊ
# ---------------------------------------------------------------------------
def render_audit(cfg):
    st.subheader("1️⃣ Tải file lên")
    col1, col2 = st.columns(2)
    with col1:
        main_file = st.file_uploader(
            "📄 File dữ liệu chính (bảng kê hóa đơn cần kiểm tra) — bắt buộc",
            type=["xlsx", "xls", "csv"],
        )
        main_sheet = 0
        if main_file is not None:
            try:
                sheets = utils.list_sheets(main_file)
                if len(sheets) > 1:
                    main_sheet = st.selectbox(
                        "→ Chọn sheet chứa bảng kê",
                        sheets,
                        index=_default_sheet_index(sheets),
                    )
                else:
                    main_sheet = sheets[0]
            except Exception:  # noqa: BLE001
                main_sheet = 0
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
            "🧾 File hóa đơn điện tử — tùy chọn (có thể tạo từ tab XML → Excel)",
            type=["xlsx", "xls", "csv"],
        )

    st.subheader("2️⃣ Chạy kiểm tra")
    if st.button("▶️ Bắt đầu kiểm tra", type="primary", disabled=main_file is None):
        if main_file is None:
            st.warning("Vui lòng tải lên file dữ liệu chính.")
            st.stop()

        try:
            main_df = _load(main_file, cfg.MAIN["header_row"], sheet_name=main_sheet)
            tms_df = _load(tms_file, cfg.TMS["header_row"])
            einv_df = _load(einv_file, cfg.EINVOICE["header_row"])
        except Exception as e:  # noqa: BLE001
            st.error(f"Lỗi khi đọc file: {e}")
            st.stop()

        if einv_file is not None and (einv_df is None or einv_df.empty):
            st.warning(
                "⚠️ File hóa đơn điện tử không có dữ liệu (rỗng) — bỏ qua các bước tra "
                "trạng thái HĐ và chênh lệch thuế. Vui lòng export lại file HĐĐT."
            )

        try:
            # danh sách rủi ro gồm nhiều sheet -> truyền thẳng file
            result = run_pipeline(main_df, tms_df, risk_file, einv_df, cfg=cfg)
        except Exception as e:  # noqa: BLE001
            st.error(f"Lỗi trong quá trình xử lý: {e}")
            st.exception(e)
            st.stop()

        st.session_state["result"] = result

    if "result" in st.session_state:
        result = st.session_state["result"]
        summary = result["summary"]
        sheets = result["sheets"]

        st.subheader("3️⃣ Kết quả")
        if summary:
            cols = st.columns(min(len(summary), 4))
            for i, (k, v) in enumerate(summary.items()):
                cols[i % len(cols)].metric(k, v)

        excel_bytes = utils.to_excel_bytes(sheets)
        st.download_button(
            "⬇️ Tải file Excel kết quả (nhiều sheet)",
            data=excel_bytes,
            file_name="ket_qua_kiem_tra_thue.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

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


# ---------------------------------------------------------------------------
# TAB 2: CHUYỂN XML HÓA ĐƠN -> EXCEL
# ---------------------------------------------------------------------------
def render_xml_converter():
    st.subheader("Chuyển hóa đơn điện tử XML → Excel")
    st.caption(
        "Tải lên một hoặc nhiều file XML hóa đơn điện tử (chuẩn Tổng cục Thuế). "
        "Công cụ trích xuất thành bảng Excel; có thể dùng luôn làm *File hóa đơn "
        "điện tử* ở tab Kiểm tra bảng kê."
    )

    xml_files = st.file_uploader(
        "📎 File XML hóa đơn (chọn nhiều file cùng lúc)",
        type=["xml"],
        accept_multiple_files=True,
    )

    if st.button(
        "🔄 Chuyển sang Excel", type="primary", disabled=not xml_files
    ):
        try:
            df_inv, df_items = xml_invoice.parse_files(xml_files)
        except Exception as e:  # noqa: BLE001
            st.error(f"Lỗi khi đọc XML: {e}")
            st.exception(e)
            st.stop()

        if df_inv.empty:
            st.warning("Không trích xuất được hóa đơn nào từ các file XML đã tải.")
            st.stop()

        st.session_state["xml_result"] = (df_inv, df_items)

    if "xml_result" in st.session_state:
        df_inv, df_items = st.session_state["xml_result"]

        c1, c2 = st.columns(2)
        c1.metric("Số hóa đơn đọc được", len(df_inv))
        c2.metric("Số dòng hàng hóa", len(df_items))

        sheets = {"Hóa đơn": df_inv}
        if not df_items.empty:
            sheets["Chi tiết hàng hóa"] = df_items
        excel_bytes = utils.to_excel_bytes(sheets)
        st.download_button(
            "⬇️ Tải file Excel hóa đơn",
            data=excel_bytes,
            file_name="hoa_don_dien_tu_tu_xml.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        st.markdown("#### Xem trước")
        st.markdown("**Bảng hóa đơn**")
        st.dataframe(df_inv.head(500), use_container_width=True)
        if not df_items.empty:
            st.markdown("**Chi tiết hàng hóa**")
            st.dataframe(df_items.head(500), use_container_width=True)


# ---------------------------------------------------------------------------
# ĐIỀU HƯỚNG
# ---------------------------------------------------------------------------
cfg = build_runtime_config()
tab_audit, tab_xml = st.tabs(["🔍 Kiểm tra bảng kê", "🔄 XML hóa đơn → Excel"])
with tab_audit:
    render_audit(cfg)
with tab_xml:
    render_xml_converter()
