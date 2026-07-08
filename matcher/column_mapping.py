"""Flexible column detection for Purchase Register and GSTR-2A/2B files."""

from __future__ import annotations

import re
from io import BytesIO
from typing import Any, Iterable

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
        "gstin of supplier / isd",
    ),
    "supplier_name": (
        "supplier name",
        "trade/legal name",
        "trade/legal name of supplier",
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
        "note number",
        "debit note number",
    ),
    "invoice_date": (
        "invoice date",
        "bill date",
        "document date",
        "inv date",
        "note date",
        "date",
    ),
    "taxable_value": (
        "taxable value",
        "taxable amount",
        "invoice value",
        "total taxable value",
        "taxable val",
        "value",
    ),
    "igst": (
        "igst",
        "integrated tax",
        "integrated tax amount",
        "igst amount",
    ),
    "cgst": (
        "cgst",
        "central tax",
        "central tax amount",
        "cgst amount",
    ),
    "sgst": (
        "sgst",
        "state/ut tax",
        "state tax",
        "state/ ut tax",
        "sgst amount",
    ),
}

# GST portal GSTR-2B / GSTR-2A invoice sheets
GSTR_DATA_SHEET_KEYWORDS = (
    "b2b",
    "b2ba",
    "cdnr",
    "cdnra",
    "impg",
    "impga",
    "impgsez",
    "isd",
    "isda",
)

KEYWORD_FALLBACKS: dict[str, tuple[str, ...]] = {
    "igst": ("integrated", "igst"),
    "cgst": ("central", "cgst"),
    "sgst": ("state", "sgst", "ut tax"),
    "taxable_value": ("taxable", "invoice value"),
}

# Columns that indicate a non-GSTR file (payment register, PR export, etc.)
NON_GSTR_COLUMN_HINTS = (
    "payment type",
    "balance due",
    "received / paid",
    "transaction type",
    "payment status",
    "received paid amount",
    "total amount",
    "description",
)


def _normalize_header(header: str) -> str:
    if header is None or (isinstance(header, float) and pd.isna(header)):
        return ""
    text = str(header).strip().lower()
    text = re.sub(r"\([^)]*\)", "", text)
    text = text.replace("₹", "").replace("rs.", "").replace("rs", "").replace("/", " / ")
    text = re.sub(r"[^\w\s/]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _score_header(header: str, aliases: Iterable[str]) -> int:
    normalized = _normalize_header(header)
    if not normalized or normalized.startswith("unnamed"):
        return 0
    best = 0
    for alias in aliases:
        alias_norm = _normalize_header(alias)
        if normalized == alias_norm:
            return 100
        if alias_norm in normalized or normalized in alias_norm:
            best = max(best, 60)
    return best


def _keyword_score(header: str, keywords: Iterable[str]) -> int:
    normalized = _normalize_header(header)
    if not normalized:
        return 0
    hits = sum(1 for keyword in keywords if keyword in normalized)
    return hits * 40


def detect_header_row(df: pd.DataFrame, max_rows: int = 20) -> int:
    best_row = 0
    best_score = -1
    limit = min(max_rows, len(df))
    for row_idx in range(limit):
        row = df.iloc[row_idx]
        score = 0
        for cell in row:
            cell_text = _normalize_header(cell)
            if not cell_text:
                continue
            for aliases in FIELD_ALIASES.values():
                if _score_header(cell_text, aliases) >= 60:
                    score += 1
        if score > best_score:
            best_score = score
            best_row = row_idx
    return best_row


def _read_sheet_raw(path_or_buffer: Any, sheet_name: str | int) -> pd.DataFrame:
    raw = pd.read_excel(path_or_buffer, sheet_name=sheet_name, header=None, dtype=object)
    header_row = detect_header_row(raw)
    headers = [_normalize_header(col) for col in raw.iloc[header_row].tolist()]
    data = raw.iloc[header_row + 1 :].copy()
    data.columns = headers
    data = data.dropna(how="all").reset_index(drop=True)
    # Drop rows that repeat header labels
    if not data.empty:
        first_cell = _normalize_header(data.iloc[0].get(headers[0], ""))
        if first_cell in FIELD_ALIASES["gstin"]:
            data = data.iloc[1:].reset_index(drop=True)
    return data


def _looks_like_non_gstr_export(df: pd.DataFrame) -> bool:
    columns_text = " ".join(_normalize_header(col) for col in df.columns)
    hits = sum(1 for hint in NON_GSTR_COLUMN_HINTS if hint in columns_text)
    has_tax_columns = any(
        _score_header(col, FIELD_ALIASES[field]) >= 40 or _keyword_score(col, KEYWORD_FALLBACKS.get(field, ()))
        for col in df.columns
        for field in ("igst", "cgst", "sgst", "taxable_value")
    )
    return hits >= 2 and not has_tax_columns


def _format_column_error(df: pd.DataFrame, missing: list[str]) -> str:
    found = ", ".join(col for col in df.columns if col)[:500]
    if _looks_like_non_gstr_export(df):
        return (
            f"Could not detect required columns: {', '.join(missing)}. "
            "This file looks like a Payment Register or Purchase Register export, "
            "not a GSTR-2A/2B download from the GST portal. "
            f"Found columns: {found}. "
            "Download GSTR-2B Excel from GST Portal: Services → Returns → GSTR-2B → Download Excel."
        )
    return (
        f"Could not detect required columns: {', '.join(missing)}. "
        f"Found columns: {found}. "
        "Please ensure your file includes GSTIN, invoice number, date, and tax columns "
        "(Taxable Value, IGST, CGST, SGST)."
    )


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
            if field in KEYWORD_FALLBACKS:
                score = max(score, _keyword_score(col, KEYWORD_FALLBACKS[field]))
            if score > best_score:
                best_score = score
                best_col = col
        if best_col and best_score >= 40:
            mapping[field] = best_col
            used_columns.add(best_col)

    missing = [field for field in REQUIRED_FIELDS if field not in mapping]
    if missing:
        raise ValueError(_format_column_error(df, missing))
    return mapping


def extract_standard_frame(df: pd.DataFrame, mapping: dict[str, str]) -> pd.DataFrame:
    return pd.DataFrame({field: df[mapping[field]] for field in REQUIRED_FIELDS})


def read_excel_with_headers(path_or_buffer: Any, sheet_name: str | int = 0) -> pd.DataFrame:
    return _read_sheet_raw(path_or_buffer, sheet_name)


def _is_gstr_data_sheet(sheet_name: str) -> bool:
    lowered = sheet_name.lower().replace(" ", "").replace("-", "")
    return any(keyword in lowered for keyword in GSTR_DATA_SHEET_KEYWORDS)


def read_gstr_excel(path_or_buffer: Any, filename: str = "") -> pd.DataFrame:
    """Read GSTR-2A/2B Excel, including multi-sheet GST portal downloads."""
    if hasattr(path_or_buffer, "read"):
        data = path_or_buffer.read()
        source: Any = BytesIO(data)
    else:
        source = path_or_buffer

    workbook = pd.ExcelFile(source)
    sheet_names = workbook.sheet_names
    skip_sheets = {"summary", "readme", "read me", "help", "index", "cover"}
    data_sheets = [name for name in sheet_names if _is_gstr_data_sheet(name)]
    candidate_sheets = data_sheets or [
        name for name in sheet_names if not any(s in name.lower() for s in skip_sheets)
    ]

    frames: list[pd.DataFrame] = []
    errors: list[str] = []
    non_gstr_detected = False

    for sheet in candidate_sheets:
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            sheet_df = _read_sheet_raw(source, sheet)
            if _looks_like_non_gstr_export(sheet_df):
                non_gstr_detected = True
                continue
            mapping = map_columns(sheet_df)
            frames.append(extract_standard_frame(sheet_df, mapping))
        except ValueError as exc:
            errors.append(f"sheet '{sheet}': {exc}")
            continue

    if frames:
        return pd.concat(frames, ignore_index=True)

    prefix = f"File '{filename}': " if filename else ""
    if non_gstr_detected and not errors:
        raise ValueError(
            f"{prefix}This file looks like a Payment Register or Purchase Register export, "
            "not GSTR-2A/2B from the GST portal. "
            "Remove it from GSTR upload and use only GST portal GSTR-2B Excel downloads."
        )
    detail = errors[0] if errors else "No readable invoice sheet found."
    raise ValueError(
        f"{prefix}Could not read GSTR-2A/2B Excel. {detail} "
        "Download from GST Portal: Services → Returns → GSTR-2B → Download Excel (B2B)."
    )
