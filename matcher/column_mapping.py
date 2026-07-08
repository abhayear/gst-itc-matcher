"""Flexible column detection for Purchase Register and GSTR-2A/2B files."""

from __future__ import annotations

import re
from typing import Iterable

import pandas as pd

REQUIRED_FIELDS = (
    "gstin",
    "supplier_name",
    "invoice_no",
    "invoice_date",
    "taxable_value",
    "igst",
    "cgst",
    "sgst",
)

FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "gstin": (
        "supplier gstin",
        "gstin of supplier",
        "gstin/uin",
        "party gstin",
        "vendor gstin",
        "gstin",
        "supplier's gstin",
    ),
    "supplier_name": (
        "supplier name",
        "trade/legal name",
        "legal name",
        "party name",
        "vendor name",
        "name of supplier",
        "supplier",
    ),
    "invoice_no": (
        "invoice no",
        "invoice no.",
        "invoice number",
        "bill no",
        "bill no.",
        "document number",
        "inv no",
        "inv no.",
    ),
    "invoice_date": (
        "invoice date",
        "bill date",
        "document date",
        "inv date",
        "date",
    ),
    "taxable_value": (
        "taxable value",
        "taxable amount",
        "invoice value",
        "total taxable value",
        "taxable val",
    ),
    "igst": (
        "igst",
        "integrated tax",
        "integrated tax(₹)",
        "integrated tax (₹)",
    ),
    "cgst": (
        "cgst",
        "central tax",
        "central tax(₹)",
        "central tax (₹)",
    ),
    "sgst": (
        "sgst",
        "state/ut tax",
        "state tax",
        "state/ut tax(₹)",
        "state/ut tax (₹)",
    ),
}


def _normalize_header(header: str) -> str:
    text = str(header).strip().lower()
    text = text.replace("₹", "").replace("rs.", "").replace("rs", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _score_header(header: str, aliases: Iterable[str]) -> int:
    normalized = _normalize_header(header)
    best = 0
    for alias in aliases:
        alias_norm = _normalize_header(alias)
        if normalized == alias_norm:
            return 100
        if alias_norm in normalized or normalized in alias_norm:
            best = max(best, 60)
    return best


def detect_header_row(df: pd.DataFrame, max_rows: int = 15) -> int:
    best_row = 0
    best_score = -1
    limit = min(max_rows, len(df))
    for row_idx in range(limit):
        row = df.iloc[row_idx]
        score = 0
        for cell in row:
            cell_text = _normalize_header(cell)
            for aliases in FIELD_ALIASES.values():
                if _score_header(cell_text, aliases) >= 60:
                    score += 1
        if score > best_score:
            best_score = score
            best_row = row_idx
    return best_row


def read_excel_with_headers(path_or_buffer, sheet_name: str | int = 0) -> pd.DataFrame:
    raw = pd.read_excel(path_or_buffer, sheet_name=sheet_name, header=None, dtype=object)
    header_row = detect_header_row(raw)
    headers = [_normalize_header(col) for col in raw.iloc[header_row].tolist()]
    data = raw.iloc[header_row + 1 :].copy()
    data.columns = headers
    data = data.dropna(how="all").reset_index(drop=True)
    return data


def map_columns(df: pd.DataFrame) -> dict[str, str]:
    mapping: dict[str, str] = {}
    used_columns: set[str] = set()

    for field in REQUIRED_FIELDS:
        best_col = None
        best_score = 0
        for col in df.columns:
            if col in used_columns:
                continue
            score = _score_header(col, FIELD_ALIASES[field])
            if score > best_score:
                best_score = score
                best_col = col
        if best_col and best_score >= 60:
            mapping[field] = best_col
            used_columns.add(best_col)

    missing = [field for field in REQUIRED_FIELDS if field not in mapping]
    if missing:
        readable = ", ".join(missing)
        raise ValueError(
            f"Could not detect required columns: {readable}. "
            "Please ensure your file includes GSTIN, invoice number, date, and tax columns."
        )
    return mapping


def extract_standard_frame(df: pd.DataFrame, mapping: dict[str, str]) -> pd.DataFrame:
    extracted = pd.DataFrame(
        {field: df[mapping[field]] for field in REQUIRED_FIELDS}
    )
    return extracted
