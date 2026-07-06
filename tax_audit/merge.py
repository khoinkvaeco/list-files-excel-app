"""
Gộp nhiều file/sheet Excel thành 1 bảng tổng hợp (tích hợp từ bộ công cụ HTML).

Tùy chọn:
  - blank       : bỏ dòng trống.
  - header_rows : số dòng tiêu đề (lấy từ nguồn ĐẦU TIÊN, dùng chung; các nguồn
                  sau cũng bỏ đi số dòng tiêu đề này).
  - source      : thêm cột "Nguồn" (tên file/sheet) ở đầu mỗi dòng.
  - dedupe      : bỏ dòng trùng lặp hoàn toàn.
"""

from __future__ import annotations

import json

from . import utils


def _is_blank_row(row) -> bool:
    return not row or all(
        c is None or str(c).strip() == "" for c in row
    )


def units_from_files(files) -> list[tuple[str, list]]:
    """Mỗi file -> 1 nguồn, lấy sheet đầu tiên. Trả về [(nhãn, ma trận)]."""
    units = []
    for f in files:
        label = getattr(f, "name", str(f))
        units.append((label, utils.read_matrix(f, sheet_name=0)))
    return units


def units_from_sheets(file) -> list[tuple[str, list]]:
    """Mỗi sheet trong 1 file -> 1 nguồn. Trả về [(tên sheet, ma trận)]."""
    units = []
    for sn in utils.list_sheets(file):
        units.append((sn, utils.read_matrix(file, sheet_name=sn)))
    return units


def merge_units(
    units: list[tuple[str, list]],
    blank: bool = True,
    header_rows: int = 0,
    source: bool = False,
    dedupe: bool = False,
) -> tuple[list[list], dict]:
    """Gộp các nguồn thành 1 ma trận. Trả về (ma trận kết quả, thống kê)."""
    header = None
    header_saved = False
    all_rows = []
    seen = set() if dedupe else None
    stats = {"nguồn": len(units), "giữ": 0, "bỏ_trống": 0, "bỏ_trùng": 0, "chi tiết": []}

    for label, rows in units:
        rows = rows or []
        if header_rows > 0:
            if not header_saved:
                header = rows[:header_rows]
                data_rows = rows[header_rows:]
                header_saved = True
            else:
                data_rows = rows[header_rows:]
        else:
            data_rows = list(rows)

        kept = dropped = dup = 0
        for r in data_rows:
            if blank and _is_blank_row(r):
                dropped += 1
                continue
            if dedupe:
                key = json.dumps([("" if c is None else str(c)) for c in r], ensure_ascii=False)
                if key in seen:
                    dup += 1
                    continue
                seen.add(key)
            all_rows.append(([label] + list(r)) if source else list(r))
            kept += 1
        stats["giữ"] += kept
        stats["bỏ_trống"] += dropped
        stats["bỏ_trùng"] += dup
        stats["chi tiết"].append({"nguồn": label, "giữ": kept, "bỏ_trống": dropped, "bỏ_trùng": dup})

    out = []
    if header:
        out = [(["Nguồn"] + list(h)) if source else list(h) for h in header]
    out.extend(all_rows)
    return out, stats
