"""
Tác vụ 7: Tìm mặt hàng "không phục vụ sản xuất kinh doanh".

Dò từ khóa (golf, quà, tặng, biếu, rượu, wine, ...) trong cột tên hàng hóa.
"""

from __future__ import annotations

import re

import pandas as pd

from . import utils


def _norm(text) -> str:
    """Chỉ hạ chữ thường, GIỮ NGUYÊN dấu tiếng Việt.

    Không bỏ dấu vì bỏ dấu khiến 'tặng' (quà tặng) trùng với 'tầng' (tầng lầu),
    'tăng' (tăng cường)... gây rất nhiều cảnh báo sai.
    """
    return str(text).lower()


def flag_non_business_goods(
    df: pd.DataFrame, goods_col, keyword_groups
) -> pd.DataFrame:
    """Đánh dấu các dòng có mặt hàng chứa từ khóa nghi ngờ, theo NHÓM.

    ``keyword_groups`` là dict {tên nhóm: [từ khóa]} (ví dụ 'Không phục vụ SXKD',
    'Quà tặng'); cũng chấp nhận list (khi đó gộp thành 1 nhóm 'Nghi ngờ').

    Thêm cột:
      - Mặt hàng nghi ngờ (bool)
      - Nhóm nghi ngờ (tên nhóm khớp, vd 'Quà tặng')
      - Từ khóa khớp (các từ khóa tìm thấy)
    So khớp theo ranh giới từ, phân biệt dấu để tránh cảnh báo sai.
    """
    df = df.copy()
    col = utils.resolve_column(df, goods_col)

    if isinstance(keyword_groups, (list, tuple)):
        keyword_groups = {"Nghi ngờ": list(keyword_groups)}

    groups = [
        (label, [(kw, _norm(kw)) for kw in kws if str(kw).strip()])
        for label, kws in keyword_groups.items()
    ]

    def match(cell):
        text = _norm(cell)
        if not text:
            return ("", "")
        found_kw, found_grp = [], []
        for label, kws in groups:
            hits = [
                original
                for original, nk in kws
                if re.search(rf"(?<!\w){re.escape(nk)}(?!\w)", text)
            ]
            if hits:
                found_kw.extend(hits)
                found_grp.append(label)
        return (", ".join(found_grp), ", ".join(found_kw))

    res = df[col].map(match)
    df["Nhóm nghi ngờ"] = res.map(lambda t: t[0])
    df["Từ khóa khớp"] = res.map(lambda t: t[1])
    df["Mặt hàng nghi ngờ"] = df["Nhóm nghi ngờ"] != ""
    return df


def suspicious_goods(flagged: pd.DataFrame) -> pd.DataFrame:
    """Lọc các dòng có mặt hàng nghi ngờ."""
    if "Mặt hàng nghi ngờ" not in flagged.columns:
        return flagged.iloc[0:0]
    return flagged[flagged["Mặt hàng nghi ngờ"] == True].copy()  # noqa: E712
