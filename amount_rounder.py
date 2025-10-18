from __future__ import annotations
import math
import os
import json
from datetime import datetime
from typing import Optional, Set, Dict, List, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet

# ---------------- Utils -----------------
DAYS_HEADER_VARIANTS: Set[str] = {
    "no of days", "no. of days", "number of days", "days", "no days", "no of day"
}

def normalize_header(s: Optional[str]) -> str:
    if s is None:
        return ""
    return " ".join(str(s).replace("\n", " ").strip().lower().split())

def round_half_up_amount(x: float, base: int) -> float:
    sign = -1 if x < 0 else 1
    a = abs(x)
    return sign * (math.floor(a / base + 0.5) * base)

def round_half_up_int(x: float) -> int:
    sign = -1 if x < 0 else 1
    a = abs(x)
    return int(sign * math.floor(a + 0.5))

def is_percentage_cell(cell) -> bool:
    fmt = (cell.number_format or "").lower()
    return "%" in fmt

def header_map(ws: Worksheet, header_row: int) -> Dict[int, str]:
    hdrs = {}
    max_col = ws.max_column or 0
    for c in range(1, max_col + 1):
        hdrs[c] = normalize_header(ws.cell(row=header_row, column=c).value)
    return hdrs

# ------------- Core Rounding -------------

def _should_fallback_to_thousand(val: float, lakh_edge_threshold: float) -> bool:
    # If rounding to lakh would zero it out, use thousand instead
    return abs(val) < lakh_edge_threshold


def _round_amount(val: float, mode: str, lakh_edge_threshold: float) -> float:
    if mode == "thousand":
        return round(val / 1000, 2)
    elif mode == "lakh":
        if _should_fallback_to_thousand(val, lakh_edge_threshold):
            return round(val / 1000, 2)
        return round(val / 100000, 2)
    elif mode == "auto":
        # Auto logic: values >= 1 lakh shown in lakhs, else in thousands
        if abs(val) >= 100000:
            return round(val / 100000, 2)
        elif abs(val) >= lakh_edge_threshold:
            return round(val / 1000, 2)
        else:
            return val
    else:
        return val



def process_excel(
    input_bytes: bytes,
    mode: str = "thousand",                  # thousand | lakh | auto
    header_row: int = 1,
    lakh_edge_threshold: float = 50000,       # < threshold → thousand fallback in lakh mode
    sheets_include: Optional[List[str]] = None,
    sheets_exclude: Optional[List[str]] = None,
) -> Tuple[bytes, dict]:
    """
    Process an Excel file given as bytes and return (output_bytes, summary_dict).
    Automatically rounds 'No of days' on Depreciation sheets.
    """
    import io

    # Load workbook twice: evaluated values only
    bio = io.BytesIO(input_bytes)
    wb = load_workbook(bio, data_only=True)

    # Filter sheets if requested
    def _sheet_allowed(name: str) -> bool:
        allowed = True
        if sheets_include:
            allowed = any(pat.lower() in name.lower() for pat in sheets_include)
        if sheets_exclude and any(pat.lower() in name.lower() for pat in sheets_exclude):
            allowed = False
        return allowed

    summary = {
        "mode": mode,
        "header_row": header_row,
        "lakh_edge_threshold": lakh_edge_threshold,
        "sheets": {},
        "totals": {"cells_seen": 0, "cells_rounded": 0},
    }

    for ws in wb.worksheets:
        if not _sheet_allowed(ws.title):
            continue

        title = (ws.title or "").strip()
        title_l = title.lower()
        is_depr = "depreciation" in title_l
        hdrs = header_map(ws, header_row) if ws.max_row >= header_row else {}

        days_cols: Set[int] = set()
        if is_depr and hdrs:
            for c, h in hdrs.items():
                if (h in DAYS_HEADER_VARIANTS) or ("day" in h and "holiday" not in h):
                    days_cols.add(c)

        ws_stats = {"cells_seen": 0, "cells_rounded": 0}

        for row in ws.iter_rows():
            for cell in row:
                # Skip non-master merged cells
                if isinstance(cell, MergedCell):
                    continue

                v = cell.value
                ws_stats["cells_seen"] += 1

                # Depreciation day rounding
                if is_depr and days_cols and cell.column in days_cols and isinstance(v, (int, float)) and not isinstance(v, bool):
                    new_v = round_half_up_int(float(v))
                    # if new_v != v:
                    #     cell.value = new_v
                    #     ws_stats["days_rounded"] += 1
                    continue

                # Amount rounding for significant numerics
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    if is_percentage_cell(cell):
                        continue
                    if abs(v) >= 100:  # treat as amount-like
                        new_v = _round_amount(float(v), mode, lakh_edge_threshold)
                        if new_v != v:
                            cell.value = new_v
                            ws_stats["cells_rounded"] += 1

        summary["sheets"][title] = ws_stats
        summary["totals"]["cells_seen"] += ws_stats["cells_seen"]
        summary["totals"]["cells_rounded"] += ws_stats["cells_rounded"]
        # summary["totals"]["days_rounded"] += ws_stats["days_rounded"]

    # Save to bytes
    out_bio = io.BytesIO()
    wb.save(out_bio)
    out_bio.seek(0)

    return out_bio.read(), summary


# -------- Helpers for saving outputs/logs on server/UI ----------

def save_outputs(
    output_bytes: bytes,
    summary: dict,
    original_filename: str,
    out_dir: str = "backup",
    tag: str = "rounded",
) -> Tuple[str, str]:
    os.makedirs(out_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.splitext(os.path.basename(original_filename))[0]
    xlsx_path = os.path.join(out_dir, f"{base}__{tag}__{stamp}.xlsx")
    json_path = os.path.join(out_dir, f"{base}__{tag}__{stamp}.json")

    with open(xlsx_path, "wb") as f:
        f.write(output_bytes)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    return xlsx_path, json_path