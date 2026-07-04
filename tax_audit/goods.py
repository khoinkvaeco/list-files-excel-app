"""
Tác vụ 7: Tìm mặt hàng "không phục vụ sản xuất kinh doanh".

Dò từ khóa (golf, quà, tặng, biếu, rượu, wine, ...) trong cột tên hàng hóa.
"""

from __future__ import annotations

import re
import unicodedata

import pandas as pd

from . import utils


def _strip_accents(text: str) -> str:
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _norm(text) -> str:
    return _strip_accents(str(text)).lower()


def flag_non_business_goods(
    df: pd.DataFrame, goods_col, keywords: list[str]
) -> pd.DataFrame:
    """Đánh dấu các dòng có mặt hàng chứa từ khóa nghi ngờ.

    Thêm cột:
      - Mặt hàng nghi ngờ (bool)
      - Từ khóa khớp (chuỗi các từ khóa tìm thấy)
    So khớp bỏ dấu tiếng Việt để bắt cả 'ruou' lẫn 'rượu'.
    """
    df = df.copy()
    col = utils.resolve_column(df, goods_col)

    norm_keywords = [(kw, _norm(kw)) for kw in keywords if str(kw).strip()]

    def match(cell):
        text = _norm(cell)
        if not text:
            return ""
        found = []
        for original, nk in norm_keywords:
            # khớp theo ranh giới từ để tránh khớp nhầm (vd 'bia' trong 'biar')
            if re.search(rf"(?<![a-z0-9]){re.escape(nk)}(?![a-z0-9])", text):
                found.append(original)
        return ", ".join(found)

    matched = df[col].map(match)
    df["Từ khóa khớp"] = matched
    df["Mặt hàng nghi ngờ"] = matched != ""
    return df


def suspicious_goods(flagged: pd.DataFrame) -> pd.DataFrame:
    """Lọc các dòng có mặt hàng nghi ngờ."""
    if "Mặt hàng nghi ngờ" not in flagged.columns:
        return flagged.iloc[0:0]
    return flagged[flagged["Mặt hàng nghi ngờ"] == True].copy()  # noqa: E712
