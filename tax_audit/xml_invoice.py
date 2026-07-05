"""
Đọc hóa đơn điện tử dạng XML (chuẩn Tổng cục Thuế) và chuyển thành bảng Excel.

Hỗ trợ:
  - File XML một hóa đơn hoặc nhiều hóa đơn (thẻ <HDon> / <DLHDon>).
  - File "thông điệp" bọc ngoài (<TDiep> ... chứa hóa đơn bên trong).
  - Bỏ qua namespace (so khớp theo tên thẻ cục bộ) để chịu được nhiều biến thể.

Trả về 2 bảng:
  - Bảng hóa đơn (mỗi dòng 1 hóa đơn) — tên cột khớp cấu hình EINVOICE để có thể
    dùng luôn làm "file hóa đơn điện tử" trong quy trình kiểm tra.
  - Bảng chi tiết hàng hóa (mỗi dòng 1 mặt hàng).
"""

from __future__ import annotations

import xml.etree.ElementTree as ET

import pandas as pd


# ---------------------------------------------------------------------------
# Tiện ích truy vấn XML không phụ thuộc namespace
# ---------------------------------------------------------------------------
def _local(tag: str) -> str:
    """Lấy tên thẻ cục bộ, bỏ phần namespace {..}."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _first(node, *names: str):
    """Trả về phần tử con/cháu ĐẦU TIÊN có tên cục bộ nằm trong ``names``."""
    if node is None:
        return None
    target = set(names)
    for el in node.iter():
        if el is node:
            continue
        if _local(el.tag) in target:
            return el
    return None


def _text(node, *names: str) -> str:
    """Trả về text của phần tử con/cháu đầu tiên khớp tên; rỗng nếu không có."""
    el = _first(node, *names)
    if el is None:
        return ""
    return (el.text or "").strip()


def _iter_invoices(root):
    """Sinh ra từng phần tử hóa đơn trong cây XML.

    Ưu tiên các thẻ <DLHDon> (dữ liệu hóa đơn); nếu không có thì <HDon>;
    nếu vẫn không có nhưng tài liệu chứa <SHDon> thì coi cả gốc là 1 hóa đơn.
    """
    for tag in ("DLHDon", "HDon", "HoaDon", "Invoice", "Inv"):
        nodes = [el for el in root.iter() if _local(el.tag) == tag]
        if nodes:
            return nodes
    if _first(root, "SHDon") is not None or _local(root.tag) in ("DLHDon", "HDon"):
        return [root]
    return []


# ---------------------------------------------------------------------------
# Trích xuất 1 hóa đơn
# ---------------------------------------------------------------------------
def _parse_invoice(inv, source: str) -> tuple[dict, list[dict]]:
    nban = _first(inv, "NBan")   # người bán
    nmua = _first(inv, "NMua")   # người mua

    header = {
        "Ký hiệu mẫu số": _text(inv, "KHMSHDon"),
        "Ký hiệu hóa đơn": _text(inv, "KHHDon"),
        "Số hóa đơn": _text(inv, "SHDon"),
        "Ngày lập": _text(inv, "NLap"),
        "MST người bán": _text(nban, "MST"),
        "Tên người bán": _text(nban, "Ten"),
        "MST người mua": _text(nmua, "MST"),
        "Tên người mua": _text(nmua, "Ten"),
        "Tổng tiền chưa thuế": _text(inv, "TgTCThue", "TgTienChuaThue"),
        "Tổng tiền thuế": _text(inv, "TgTThue", "TgTienThue"),
        "Tổng tiền thanh toán": _text(inv, "TgTTTBSo", "TgTienThanhToan"),
        "Trạng thái hóa đơn": _text(inv, "TTHDon", "TThai", "TrangThai"),
        "File nguồn": source,
    }

    # gom tên hàng hóa để phục vụ dò mặt hàng nghi ngờ
    items = []
    ds = _first(inv, "DSHHDVu")
    hhdvu_nodes = (
        [el for el in ds.iter() if _local(el.tag) == "HHDVu"] if ds is not None else []
    )
    goods_names = []
    for it in hhdvu_nodes:
        ten = _text(it, "THHDVu", "TenHHDVu", "Ten")
        goods_names.append(ten)
        items.append(
            {
                "Số hóa đơn": header["Số hóa đơn"],
                "MST người bán": header["MST người bán"],
                "STT": _text(it, "STT"),
                "Tên hàng hóa": ten,
                "Đơn vị tính": _text(it, "DVTinh", "DonViTinh"),
                "Số lượng": _text(it, "SLuong", "SoLuong"),
                "Đơn giá": _text(it, "DGia", "DonGia"),
                "Thành tiền": _text(it, "ThTien", "ThanhTien"),
                "Thuế suất": _text(it, "TSuat", "ThueSuat"),
                "File nguồn": source,
            }
        )
    header["Tên hàng hóa"] = " | ".join(n for n in goods_names if n)
    return header, items


# ---------------------------------------------------------------------------
# API chính
# ---------------------------------------------------------------------------
def parse_xml_bytes(data: bytes, source: str = "") -> tuple[list[dict], list[dict]]:
    """Phân tích nội dung XML (bytes) -> (danh sách hóa đơn, danh sách hàng hóa)."""
    text = data.decode("utf-8", errors="ignore") if isinstance(data, bytes) else data
    root = ET.fromstring(text)
    invoices, items = [], []
    for inv in _iter_invoices(root):
        h, its = _parse_invoice(inv, source)
        if any(v for k, v in h.items() if k not in ("File nguồn",)):
            invoices.append(h)
            items.extend(its)
    return invoices, items


# ---------------------------------------------------------------------------
# CHẾ ĐỘ TỔNG QUÁT: tự dò phần tử lặp lại và trải phẳng thành cột
# (dùng khi XML không theo cấu trúc hóa đơn đơn lẻ chuẩn, ví dụ bảng kê/bảng
#  tổng hợp doanh nghiệp gửi)
# ---------------------------------------------------------------------------
def _leaf_children(el) -> int:
    return sum(1 for c in el if len(c) == 0)


def _detect_record_elements(root) -> list:
    """Tìm nhóm phần tử 'bản ghi' lặp lại nhiều nhất và có nhiều trường lá.

    Trả về danh sách các phần tử được coi là 1 dòng dữ liệu.
    """
    by_tag: dict[str, list] = {}
    for el in root.iter():
        if len(el) > 0:  # có phần tử con
            by_tag.setdefault(_local(el.tag), []).append(el)

    best_tag, best_score, best_els = None, -1.0, []
    for tag, els in by_tag.items():
        if len(els) < 2:
            continue
        leaf_counts = [_leaf_children(e) for e in els]
        avg_leaves = sum(leaf_counts) / len(els)
        if avg_leaves < 2:
            continue
        score = len(els) * avg_leaves
        if score > best_score:
            best_tag, best_score, best_els = tag, score, els
    return best_els


def _flatten_record(el) -> dict:
    """Trải phẳng mọi thẻ lá trong 1 bản ghi thành {tên thẻ: giá trị}."""
    row: dict[str, str] = {}
    for c in el.iter():
        if c is el or len(c) > 0:
            continue
        key = _local(c.tag)
        val = (c.text or "").strip()
        if key in row:
            i = 2
            while f"{key}_{i}" in row:
                i += 1
            key = f"{key}_{i}"
        row[key] = val
    return row


def parse_generic_bytes(data: bytes, source: str = "") -> list[dict]:
    """Trải phẳng XML bất kỳ: tìm phần tử lặp lại, mỗi phần tử -> 1 dòng."""
    text = data.decode("utf-8", errors="ignore") if isinstance(data, bytes) else data
    root = ET.fromstring(text)
    records = _detect_record_elements(root)
    rows = []
    for r in records:
        row = _flatten_record(r)
        if row:
            row["File nguồn"] = source
            rows.append(row)
    return rows


def _read_bytes(f) -> bytes:
    if hasattr(f, "read"):
        if hasattr(f, "seek"):
            try:
                f.seek(0)
            except (OSError, ValueError):
                pass
        data = f.read()
        return data if isinstance(data, bytes) else str(data).encode("utf-8", "ignore")
    with open(f, "rb") as fh:
        return fh.read()


def parse_files(files) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Đọc nhiều file XML (đường dẫn hoặc file upload) -> 2 DataFrame.

    Ưu tiên cấu trúc hóa đơn chuẩn (TCT). Nếu không nhận ra, tự động chuyển sang
    CHẾ ĐỘ TỔNG QUÁT (dò phần tử lặp lại) để vẫn xuất được dữ liệu ra Excel.

    Trả về (df_hoa_don, df_hang_hoa). Ở chế độ tổng quát, df_hang_hoa rỗng.
    """
    all_inv, all_items, generic_rows, errors = [], [], [], []
    for f in files:
        source = getattr(f, "name", str(f))
        try:
            data = _read_bytes(f)
        except Exception as e:  # noqa: BLE001
            errors.append({"File nguồn": source, "Lỗi": f"Không đọc được file: {e}"})
            continue
        try:
            inv, items = parse_xml_bytes(data, source)
        except ET.ParseError as e:
            errors.append({"File nguồn": source, "Lỗi": f"XML không hợp lệ: {e}"})
            continue
        if inv:
            all_inv.extend(inv)
            all_items.extend(items)
        else:
            # không nhận ra cấu trúc hóa đơn chuẩn -> chế độ tổng quát
            try:
                generic_rows.extend(parse_generic_bytes(data, source))
            except Exception as e:  # noqa: BLE001
                errors.append({"File nguồn": source, "Lỗi": f"Không trải phẳng được: {e}"})

    inv_cols = [
        "Ký hiệu mẫu số", "Ký hiệu hóa đơn", "Số hóa đơn", "Ngày lập",
        "MST người bán", "Tên người bán", "MST người mua", "Tên người mua",
        "Tổng tiền chưa thuế", "Tổng tiền thuế", "Tổng tiền thanh toán",
        "Trạng thái hóa đơn", "Tên hàng hóa", "File nguồn",
    ]

    if all_inv:
        df_inv = pd.DataFrame(all_inv)
        for c in inv_cols:
            if c not in df_inv.columns:
                df_inv[c] = ""
        keep = inv_cols + [c for c in df_inv.columns if c not in inv_cols]
        df_inv = df_inv[keep]
        return df_inv, pd.DataFrame(all_items)

    if generic_rows:
        # chế độ tổng quát: trả bảng đã trải phẳng
        df_inv = pd.DataFrame(generic_rows)
        return df_inv, pd.DataFrame()

    # không có gì -> trả bảng lỗi (nếu có) để người dùng biết nguyên nhân
    return pd.DataFrame(errors), pd.DataFrame()
