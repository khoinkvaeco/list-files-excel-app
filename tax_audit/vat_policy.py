"""
Chính sách giảm thuế GTGT 10% -> 8%: kiểm tra thuế suất trên hóa đơn có phù hợp
với (a) giai đoạn được giảm theo ngày hóa đơn và (b) mặt hàng có thuộc diện được
giảm hay không (Phụ lục I, II, III của NĐ 15/2022 và các NĐ sau).

Lưu ý: phân loại mặt hàng dựa trên TỪ KHÓA trong mô tả -> mang tính TRỢ GIÚP,
cần người kiểm tra xác nhận theo mã HS/ngành khi cần.
"""

from __future__ import annotations

import re

import pandas as pd


def _norm(text) -> str:
    """Hạ chữ thường, giữ nguyên dấu tiếng Việt."""
    return str(text).lower()


def in_reduced_period(ts, periods) -> bool:
    """Ngày hóa đơn có nằm trong giai đoạn được giảm 8% không."""
    if ts is None or pd.isna(ts):
        return False
    for start, end in periods:
        if pd.Timestamp(start) <= ts <= pd.Timestamp(end):
            return True
    return False


def classify_excluded(goods_text, exclude_keywords: dict) -> str:
    """Trả về nhóm Phụ lục nếu mặt hàng thuộc diện KHÔNG được giảm; '' nếu không.

    So khớp theo ranh giới từ, giữ nguyên dấu (tránh khớp nhầm).
    """
    text = _norm(goods_text)
    if not text.strip():
        return ""
    for category, keywords in exclude_keywords.items():
        for kw in keywords:
            nk = _norm(kw)
            if re.search(rf"(?<!\w){re.escape(nk)}(?!\w)", text):
                return category
    return ""


def infer_rate(pretax_val: float, vat_val: float):
    """Suy ra thuế suất từ (tiền thuế / giá trị chưa thuế). None nếu không rõ."""
    if not pretax_val:
        return None
    r = vat_val / pretax_val
    for target in (0.10, 0.08, 0.05, 0.0):
        if abs(r - target) < 0.005:
            return target
    return round(r, 4)


def check(ts, pretax_val, vat_val, goods_text, periods, exclude_keywords) -> str:
    """Verdict kiểm tra thuế suất, kết hợp ngày + nhóm mặt hàng loại trừ."""
    rate = infer_rate(pretax_val, vat_val)
    if rate is None:
        return ""

    excluded = classify_excluded(goods_text, exclude_keywords)
    reduced_period = in_reduced_period(ts, periods)

    is10 = abs(rate - 0.10) < 0.005
    is8 = abs(rate - 0.08) < 0.005

    if excluded:
        # mặt hàng nghi thuộc diện KHÔNG được giảm -> phải 10%
        # (khớp theo từ khóa nên dùng 'rà lại', cần người xác nhận)
        if is8:
            return f"8% - rà lại: nhóm không được giảm? ({excluded})"
        if is10:
            return f"10% - nhóm không được giảm ({excluded})"
        return _other_label(rate)

    # mặt hàng thuộc diện có thể được giảm
    if is10:
        return "10% - rà lại (đang kỳ giảm, có thể phải 8%)" if reduced_period else "10% - đúng"
    if is8:
        return "8% - đúng kỳ giảm" if reduced_period else "8% - SAI: ngoài kỳ giảm, phải 10%"
    return _other_label(rate)


def _other_label(rate: float) -> str:
    if abs(rate - 0.05) < 0.005:
        return "5%"
    if abs(rate - 0.0) < 0.005:
        return "0% / không chịu thuế"
    return f"{rate * 100:.1f}% - khác"
