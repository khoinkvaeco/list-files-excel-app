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
    for tag in ("DLHDon", "HDon"):
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


def parse_files(files) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Đọc nhiều file XML (đường dẫn hoặc file upload) -> 2 DataFrame.

    Trả về (df_hoa_don, df_hang_hoa).
    """
    all_inv, all_items = [], []
    for f in files:
        source = getattr(f, "name", str(f))
        if hasattr(f, "read"):
            if hasattr(f, "seek"):
                try:
                    f.seek(0)
                except (OSError, ValueError):
                    pass
            data = f.read()
        else:
            with open(f, "rb") as fh:
                data = fh.read()
        try:
            inv, items = parse_xml_bytes(data, source)
        except ET.ParseError as e:
            inv, items = [], []
            all_inv.append({"Số hóa đơn": "", "File nguồn": source, "Lỗi": f"XML không hợp lệ: {e}"})
        all_inv.extend(inv)
        all_items.extend(items)

    inv_cols = [
        "Ký hiệu mẫu số", "Ký hiệu hóa đơn", "Số hóa đơn", "Ngày lập",
        "MST người bán", "Tên người bán", "MST người mua", "Tên người mua",
        "Tổng tiền chưa thuế", "Tổng tiền thuế", "Tổng tiền thanh toán",
        "Trạng thái hóa đơn", "Tên hàng hóa", "File nguồn",
    ]
    df_inv = pd.DataFrame(all_inv)
    if not df_inv.empty:
        # đảm bảo đủ cột và đúng thứ tự
        for c in inv_cols:
            if c not in df_inv.columns:
                df_inv[c] = ""
        keep = inv_cols + [c for c in df_inv.columns if c not in inv_cols]
        df_inv = df_inv[keep]

    df_items = pd.DataFrame(all_items)
    return df_inv, df_items
