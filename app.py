"""
Ứng dụng web (Streamlit) hỗ trợ kiểm tra thuế trên file Excel.

Chạy:  streamlit run app.py
"""

from __future__ import annotations

import copy
import json
import os
import types

import pandas as pd
import streamlit as st

import config as C
from tax_audit import utils, xml_invoice, merge as merge_mod
from pipeline import run_pipeline

# File lưu cấu hình cột do người dùng chỉnh (nằm cạnh app.py, theo từng máy)
USER_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_config.json")

# Các key widget cấu hình cột — dùng khi khôi phục mặc định
_CFG_WIDGET_PREFIXES = ("main_", "tms_", "einv_", "risk_", "kw")


def load_user_config() -> dict:
    """Đọc cấu hình đã lưu (nếu có)."""
    try:
        with open(USER_CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}


def save_user_config(cfg) -> None:
    """Lưu các phần cấu hình cột có thể chỉnh ra file JSON."""
    data = {
        "MAIN": {"header_row": cfg.MAIN["header_row"], "cols": cfg.MAIN["cols"]},
        "TMS": {
            "header_row": cfg.TMS["header_row"],
            "mst_col": cfg.TMS["mst_col"],
            "status_col": cfg.TMS["status_col"],
            "close_date_col": cfg.TMS["close_date_col"],
        },
        "EINVOICE": {"header_row": cfg.EINVOICE["header_row"], "cols": cfg.EINVOICE["cols"]},
        "RISK": {"sheets": cfg.RISK["sheets"]},
        "GOODS_KEYWORDS": cfg.GOODS_KEYWORDS,
        "COL_MODE": getattr(cfg, "_col_mode", "letter"),
    }
    with open(USER_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def saved_col_mode() -> str:
    return load_user_config().get("COL_MODE", "letter")


def _risk_name_sheets():
    """Sheet rủi ro mặc định khi chỉ định cột theo TÊN (giữ nhãn cố định ở sheet 2,3)."""
    out = []
    for i, sp in enumerate(C.RISK["sheets"]):
        ns = dict(sp)
        ns["mst_col"] = C.RISK_NAME_COLS["mst_col"]
        ns["doc_col"] = C.RISK_NAME_COLS["doc_col"] if i == 0 else sp["doc_col"]
        out.append(ns)
    return out


def _base_config(mode: str) -> dict:
    """Giá trị mặc định cho các ô cấu hình theo CHẾ ĐỘ (letter/name), phủ bởi
    cấu hình đã lưu (chỉ khi cùng chế độ)."""
    if mode == "name":
        base = {
            "MAIN": {"header_row": C.MAIN["header_row"], "cols": dict(C.MAIN_NAME_COLS)},
            "TMS": {"header_row": C.TMS["header_row"], **C.TMS_NAME_COLS},
            "EINVOICE": {"header_row": C.EINVOICE["header_row"], "cols": dict(C.EINVOICE_NAME_COLS)},
            "RISK": {"sheets": _risk_name_sheets()},
            "GOODS_KEYWORDS": copy.deepcopy(C.GOODS_KEYWORDS),
        }
    else:
        base = {
            "MAIN": copy.deepcopy(C.MAIN),
            "TMS": copy.deepcopy(C.TMS),
            "EINVOICE": copy.deepcopy(C.EINVOICE),
            "RISK": copy.deepcopy(C.RISK),
            "GOODS_KEYWORDS": copy.deepcopy(C.GOODS_KEYWORDS),
        }

    saved = load_user_config()
    if saved.get("COL_MODE", "letter") != mode:
        return base  # cấu hình đã lưu thuộc chế độ khác -> dùng mặc định của chế độ này

    if saved.get("MAIN"):
        base["MAIN"]["header_row"] = saved["MAIN"].get("header_row", base["MAIN"]["header_row"])
        base["MAIN"]["cols"].update(saved["MAIN"].get("cols", {}))
    if saved.get("TMS"):
        base["TMS"].update({k: v for k, v in saved["TMS"].items() if k in base["TMS"]})
    if saved.get("EINVOICE"):
        base["EINVOICE"]["header_row"] = saved["EINVOICE"].get("header_row", base["EINVOICE"]["header_row"])
        base["EINVOICE"]["cols"].update(saved["EINVOICE"].get("cols", {}))
    if saved.get("RISK", {}).get("sheets"):
        base["RISK"]["sheets"] = saved["RISK"]["sheets"]
    if saved.get("GOODS_KEYWORDS"):
        base["GOODS_KEYWORDS"] = saved["GOODS_KEYWORDS"]
    return base

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
        GOODS_KEYWORDS=copy.deepcopy(C.GOODS_KEYWORDS),
        INVOICE_STATUS=copy.deepcopy(C.INVOICE_STATUS),
        STATUS_NOT_FOUND=C.STATUS_NOT_FOUND,
        VAT_REDUCED_PERIODS=list(C.VAT_REDUCED_PERIODS),
        VAT_EXCLUDE_KEYWORDS=copy.deepcopy(C.VAT_EXCLUDE_KEYWORDS),
    )

    with st.sidebar:
        st.header("⚙️ Cấu hình cột")
        if os.path.exists(USER_CONFIG_PATH):
            st.caption("✅ Đang dùng cấu hình đã lưu của bạn.")

        _mode_labels = ["Chữ cái cột Excel (A, B, C…)", "Theo tên tiêu đề"]
        _default_mode_idx = 1 if saved_col_mode() == "name" else 0
        mode_label = st.radio(
            "Kiểu chỉ định cột", _mode_labels, index=_default_mode_idx, key="col_mode_radio"
        )
        col_mode = "name" if mode_label == _mode_labels[1] else "letter"
        cfg._col_mode = col_mode

        # đổi kiểu chỉ định -> xóa giá trị ô cũ để nạp lại mặc định của kiểu mới
        if st.session_state.get("_prev_col_mode") not in (None, col_mode):
            for key in list(st.session_state.keys()):
                if str(key).startswith(_CFG_WIDGET_PREFIXES):
                    del st.session_state[key]
        st.session_state["_prev_col_mode"] = col_mode

        if col_mode == "name":
            st.caption(
                "Nhập **tên tiêu đề cột** (khớp gần đúng, ví dụ 'Ngày hóa đơn', "
                "'Mã số thuế người bán'). Nhớ đặt đúng **Dòng tiêu đề** của file."
            )
        else:
            st.caption(
                "Nhập **chữ cái cột Excel** (A, B, C, … như trên thanh cột Excel). "
                "Cũng chấp nhận số thứ tự cột. Để mặc định nếu file đúng bố cục chuẩn."
            )

        base = _base_config(col_mode)

        with st.expander("File dữ liệu chính", expanded=False):
            for k, label in [
                ("mst", "MST người bán"),
                ("seller", "Tên người bán"),
                ("invoice_no", "Số hóa đơn"),
                ("invoice_date", "Ngày hóa đơn"),
                ("pretax", "Giá trị chưa thuế"),
                ("vat", "Thuế GTGT"),
                ("goods", "Tên hàng hóa"),
                ("period", "Cột lấy năm kiểm tra (tùy chọn, vd cột kỳ 'T1.2022')"),
            ]:
                cfg.MAIN["cols"][k] = _col_input(
                    f"main_{k}", label, base["MAIN"]["cols"].get(k, "")
                )
            cfg.MAIN["header_row"] = st.number_input(
                "Dòng tiêu đề", 1, 50, base["MAIN"]["header_row"], key="main_hdr"
            )

        with st.expander("File TMS", expanded=False):
            cfg.TMS["mst_col"] = _col_input("tms_mst", "Cột MST", base["TMS"]["mst_col"])
            cfg.TMS["status_col"] = _col_input(
                "tms_status", "Cột trạng thái NNT (mặc định AS)", base["TMS"]["status_col"]
            )
            cfg.TMS["close_date_col"] = _col_input(
                "tms_close", "Cột ngày đóng trạng thái (mặc định BB)", base["TMS"]["close_date_col"]
            )
            cfg.TMS["header_row"] = st.number_input(
                "Dòng tiêu đề (TMS)", 1, 50, base["TMS"]["header_row"], key="tms_hdr"
            )

        with st.expander("File DS rủi ro (nhiều sheet)", expanded=False):
            st.caption(
                "Danh sách rủi ro gồm nhiều sheet, mỗi sheet một bố cục. "
                "Chỉnh tên sheet / dòng tiêu đề / cột MST / cột (hoặc nhãn) văn bản:"
            )
            cfg.RISK["sheets"] = []
            for i, spec in enumerate(base["RISK"]["sheets"]):
                st.markdown(f"**Sheet {i + 1}**")
                new_spec = {
                    "name": st.text_input("Tên sheet", value=spec["name"], key=f"risk_name_{i}"),
                    "header_row": st.number_input(
                        "Dòng tiêu đề", 1, 50, spec["header_row"], key=f"risk_hdr_{i}"
                    ),
                    "mst_col": _col_input(f"risk_mst_{i}", "Cột MST", spec["mst_col"]),
                    "doc_col": _col_input(
                        f"risk_doc_{i}", "Cột văn bản (chữ cái) hoặc nhãn văn bản",
                        spec["doc_col"],
                    ),
                }
                cfg.RISK["sheets"].append(new_spec)

        with st.expander("File hóa đơn điện tử", expanded=False):
            for k, label in [
                ("mst", "MST người bán"),
                ("invoice_no", "Số hóa đơn"),
                ("invoice_date", "Ngày hóa đơn"),
                ("status", "Trạng thái hóa đơn"),
                ("total_tax", "Tổng tiền thuế"),
            ]:
                cfg.EINVOICE["cols"][k] = _col_input(
                    f"einv_{k}", label, base["EINVOICE"]["cols"][k]
                )
            cfg.EINVOICE["header_row"] = st.number_input(
                "Dòng tiêu đề (HĐĐT)", 1, 50, base["EINVOICE"]["header_row"], key="einv_hdr"
            )

        with st.expander("Từ khóa mặt hàng nghi ngờ (2 nhóm)", expanded=False):
            base_kw = base["GOODS_KEYWORDS"]
            if not isinstance(base_kw, dict):  # tương thích cấu hình cũ (list)
                base_kw = {"Không phục vụ SXKD": list(base_kw), "Quà tặng": []}
            cfg.GOODS_KEYWORDS = {}
            for gi, (grp_label, grp_kw) in enumerate(base_kw.items()):
                st.markdown(f"**{grp_label}**")
                txt = st.text_area(
                    "Mỗi từ khóa 1 dòng",
                    value="\n".join(grp_kw),
                    key=f"kw_{gi}",
                    height=140,
                    label_visibility="collapsed",
                )
                cfg.GOODS_KEYWORDS[grp_label] = [
                    x.strip() for x in txt.splitlines() if x.strip()
                ]

        # --- Lưu / khôi phục cấu hình ---
        st.divider()
        c1, c2 = st.columns(2)
        if c1.button("💾 Lưu cấu hình", width="stretch"):
            save_user_config(cfg)
            st.success("Đã lưu. Lần sau mở app sẽ tự dùng lại cấu hình này.")
        if c2.button("↩️ Về mặc định", width="stretch"):
            if os.path.exists(USER_CONFIG_PATH):
                os.remove(USER_CONFIG_PATH)
            for key in list(st.session_state.keys()):
                if str(key).startswith(_CFG_WIDGET_PREFIXES):
                    del st.session_state[key]
            st.rerun()

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


def show_df(df, n: int = 500):
    """Hiển thị bảng an toàn với Arrow: ép tên cột & các cột hỗn hợp về chuỗi.

    Tránh lỗi 'Could not convert ... to int64' khi cột có kiểu dữ liệu lẫn lộn
    (ví dụ số hóa đơn '01GTKT0/005' lẫn số) và cảnh báo tên cột hỗn hợp kiểu.
    """
    d = df.head(n).copy()
    d.columns = [str(c) for c in d.columns]
    for c in d.columns:
        if d[c].dtype == object:
            d[c] = d[c].map(lambda x: "" if pd.isna(x) else str(x))
    st.dataframe(d, width="stretch")


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
                show_df(df)
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
            st.warning("Không trích xuất được dữ liệu nào từ các file XML đã tải.")
            st.stop()

        # nếu chỉ có cột lỗi -> báo lỗi cụ thể
        if set(df_inv.columns) <= {"File nguồn", "Lỗi"} and "Lỗi" in df_inv.columns:
            st.error("Không đọc được file XML:")
            show_df(df_inv)
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
        show_df(df_inv)
        if not df_items.empty:
            st.markdown("**Chi tiết hàng hóa**")
            show_df(df_items)


# ---------------------------------------------------------------------------
# TAB 3: GỘP NHIỀU FILE / SHEET EXCEL
# ---------------------------------------------------------------------------
def render_merge():
    st.subheader("Gộp nhiều file / sheet Excel thành 1 bảng tổng hợp")
    st.caption(
        "Chọn nhiều file Excel (mỗi file lấy sheet đầu), hoặc 1 file để gộp tất cả "
        "các sheet của nó, thành một bảng duy nhất."
    )

    mode = st.radio(
        "Kiểu gộp",
        ["Gộp nhiều FILE (mỗi file 1 bảng)", "Gộp các SHEET trong 1 file"],
        horizontal=True,
    )
    merge_by_files = mode.startswith("Gộp nhiều FILE")

    files = st.file_uploader(
        "📎 File Excel" + (" (chọn nhiều file)" if merge_by_files else " (1 file nhiều sheet)"),
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=merge_by_files,
    )

    c1, c2 = st.columns(2)
    with c1:
        opt_blank = st.checkbox("Bỏ dòng trống", value=True)
        opt_source = st.checkbox("Thêm cột 'Nguồn' (tên file/sheet)", value=True)
    with c2:
        opt_dedupe = st.checkbox("Bỏ dòng trùng lặp hoàn toàn", value=False)
        header_rows = st.number_input(
            "Số dòng tiêu đề dùng chung (0 = không có)", 0, 50, 1
        )

    out_name = st.text_input("Tên file kết quả", value="TongHop.xlsx")

    has_files = bool(files) if merge_by_files else files is not None
    if st.button("⬇️ Gộp & tải file tổng hợp", type="primary", disabled=not has_files):
        try:
            if merge_by_files:
                units = merge_mod.units_from_files(files)
            else:
                units = merge_mod.units_from_sheets(files)
            out, stats = merge_mod.merge_units(
                units,
                blank=opt_blank,
                header_rows=int(header_rows),
                source=opt_source,
                dedupe=opt_dedupe,
            )
        except Exception as e:  # noqa: BLE001
            st.error(f"Lỗi khi gộp: {e}")
            st.exception(e)
            st.stop()

        if not out:
            st.warning("Không có dữ liệu để gộp.")
            st.stop()

        st.session_state["merge_out"] = (out, stats, out_name)

    if "merge_out" in st.session_state:
        out, stats, out_name = st.session_state["merge_out"]
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Số nguồn", stats["nguồn"])
        m2.metric("Dòng giữ lại", stats["giữ"])
        m3.metric("Bỏ trống", stats["bỏ_trống"])
        m4.metric("Bỏ trùng", stats["bỏ_trùng"])

        name = out_name.strip() or "TongHop.xlsx"
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
        st.download_button(
            "⬇️ Tải file tổng hợp",
            data=utils.matrix_to_excel_bytes(out, "TongHop"),
            file_name=name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.markdown("#### Xem trước (tối đa 500 dòng)")
        show_df(pd.DataFrame(out).head(500))


# ---------------------------------------------------------------------------
# ĐIỀU HƯỚNG
# ---------------------------------------------------------------------------
cfg = build_runtime_config()
# áp dụng cách hiểu tham chiếu cột (chữ cái Excel hay tên tiêu đề) cho toàn bộ xử lý
utils.set_column_mode(getattr(cfg, "_col_mode", "letter"))
tab_audit, tab_xml, tab_merge = st.tabs(
    ["🔍 Kiểm tra bảng kê", "🔄 XML hóa đơn → Excel", "📁 Gộp file Excel"]
)
with tab_audit:
    render_audit(cfg)
with tab_xml:
    render_xml_converter()
with tab_merge:
    render_merge()
