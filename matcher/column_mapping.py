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
        "gstin of supplier",
        "supplier gstin",
        "gstin/uin",
        "party gstin",
        "vendor gstin",
        "gstin",
        "supplier's gstin",
        "gstin of supplier / isd",
    ),
    "supplier_name": (
        "trade/legal name of the supplier",
        "trade/legal name of supplier",
        "trade/legal name",
        "supplier name",
        "legal name",
        "party name",
        "vendor name",
        "name of supplier",
        "supplier",
    ),
    "invoice_no": (
        "invoice number",
        "invoice no",
        "invoice no.",
        "note number",
        "bill no",
        "bill no.",
        "document number",
        "inv no",
        "inv no.",
        "debit note number",
        "credit note number",
    ),
    "invoice_date": (
        "invoice date",
        "note date",
        "bill date",
        "document date",
        "inv date",
        "date",
    ),
    "taxable_value": (
        "taxable value",
        "taxable amount",
        "total taxable value",
        "taxable val",
    ),
    "igst": (
        "integrated tax",
        "igst",
        "integrated tax amount",
        "igst amount",
    ),
    "cgst": (
        "central tax",
        "cgst",
        "central tax amount",
        "cgst amount",
    ),
    "sgst": (
        "state/ut tax",
        "state/ ut tax",
        "state tax",
        "sgst",
        "sgst amount",
    ),
}

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

GSTR_SKIP_SHEET_KEYWORDS = (
    "read me",
    "readme",
    "summary",
    "help",
    "index",
    "cover",
    "instructions",
    "eco",
    "ecoa",
    "tds",
    "tdsa",
    "tcs",
)

KEYWORD_FALLBACKS: dict[str, tuple[str, ...]] = {
    "igst": ("integrated", "igst"),
    "cgst": ("central", "cgst"),
    "sgst": ("state", "sgst", "ut tax"),
    "taxable_value": ("taxable value", "taxable"),
}

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

SECTION_ROW_HINTS = (
    "supplier wise",
    "supplier-wise",
    "table ",
    "gstr-2b",
    "gstr 2b",
    "gstr-2a",
    "tax period",
    "financial year",
    "grand total",
    "note:",
    "remarks",
    "all tables",
    "download excel",
    "goods and services tax",
    "government of india",
)

GSTIN_PATTERN = re.compile(r"^[0-9]{2}[A-Z0-9]{13}$")


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


def _row_cells(row: pd.Series) -> list[str]:
    return [_normalize_header(cell) for cell in row.tolist()]


def _score_header_row_cells(cells: list[str]) -> tuple[int, int]:
    matched_fields: set[str] = set()
    for cell in cells:
        if not cell:
            continue
        for field, aliases in FIELD_ALIASES.items():
            score = _score_header(cell, aliases)
            if score >= 60:
                matched_fields.add(field)
            elif field in KEYWORD_FALLBACKS and _keyword_score(cell, KEYWORD_FALLBACKS[field]) >= 40:
                matched_fields.add(field)
    critical = int(
        "gstin" in matched_fields
        and "invoice_no" in matched_fields
        and bool({"igst", "cgst", "sgst", "taxable_value"} & matched_fields)
    )
    return len(matched_fields), critical


def _build_headers(*rows: list[str]) -> list[str]:
    if not rows:
        return []
    if len(rows) == 1:
        return rows[0]

    headers: list[str] = []
    group = ""
    max_len = max(len(row) for row in rows)
    for idx in range(max_len):
        parts = [row[idx] if idx < len(row) else "" for row in rows]
        parts = [part for part in parts if part]
        if not parts:
            headers.append(_normalize_header(group))
            continue
        header = parts[-1]
        if len(parts) > 1 and parts[0] not in header:
            header = parts[-1]
        group = parts[0] if len(parts) == 1 and parts[0] not in {
            "invoice details",
            "credit note/debit note details",
            "credit note debit note details",
        } else group
        headers.append(_normalize_header(header))
    return headers


def _is_section_or_title_row(row: pd.Series) -> bool:
    cells = [str(v).strip() for v in row.tolist() if pd.notna(v) and str(v).strip()]
    if not cells:
        return True
    if len(cells) <= 3:
        text = _normalize_header(" ".join(cells))
        if any(hint in text for hint in SECTION_ROW_HINTS):
            return True
        if text.startswith("total") or text.startswith("grand total"):
            return True
    return False


def _is_valid_gstin(value: Any) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    text = re.sub(r"[^0-9A-Z]", "", str(value).strip().upper())
    return bool(GSTIN_PATTERN.match(text))


def _dedupe_headers(headers: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    result: list[str] = []
    for header in headers:
        base = header or "column"
        count = seen.get(base, 0)
        seen[base] = count + 1
        result.append(base if count == 0 else f"{base}_{count + 1}")
    return result


def _filter_invoice_rows(df: pd.DataFrame, mapping: dict[str, str]) -> pd.DataFrame:
    gstin_col = mapping["gstin"]
    mask = df[gstin_col].apply(_is_valid_gstin)
    filtered = df[mask].copy().reset_index(drop=True)
    return filtered


def _dataframe_from_header(raw: pd.DataFrame, header_rows: list[int]) -> pd.DataFrame:
    header_parts = [_row_cells(raw.iloc[row_idx]) for row_idx in header_rows]
    headers = _build_headers(*header_parts)
    headers = _dedupe_headers(headers)
    data_start = max(header_rows) + 1
    data = raw.iloc[data_start:].copy()
    if data.empty:
        return pd.DataFrame()

    width = min(len(headers), data.shape[1])
    data = data.iloc[:, :width]
    data.columns = headers[:width]
    data = data.dropna(how="all").reset_index(drop=True)

    keep_rows = [idx for idx, row in data.iterrows() if not _is_section_or_title_row(row)]
    if keep_rows:
        data = data.iloc[keep_rows].reset_index(drop=True)

    if not data.empty:
        first_cell = _normalize_header(data.iloc[0].iloc[0])
        if _score_header(first_cell, FIELD_ALIASES["gstin"]) >= 60:
            data = data.iloc[1:].reset_index(drop=True)

    return data


def _header_row_combinations(scan_limit: int) -> list[list[int]]:
    combos: list[list[int]] = []
    for start in range(scan_limit):
        combos.append([start])
        if start + 1 < scan_limit:
            combos.append([start, start + 1])
        if start + 2 < scan_limit:
            combos.append([start, start + 1, start + 2])
    return combos


def _find_best_gstr_table(raw: pd.DataFrame) -> pd.DataFrame | None:
    best_df: pd.DataFrame | None = None
    best_key = (-1, -1, -1)

    scan_limit = min(len(raw), 30)
    for header_rows in _header_row_combinations(scan_limit):
        candidate = _dataframe_from_header(raw, header_rows)
        if candidate.empty:
            continue
        try:
            mapping = map_columns(candidate)
            filtered = _filter_invoice_rows(candidate, mapping)
            if filtered.empty:
                continue
            std = extract_standard_frame(filtered, mapping)
        except ValueError:
            continue
        field_count, critical = _score_header_row_cells(list(candidate.columns))
        key = (critical, field_count, len(std))
        if key > best_key:
            best_key = key
            best_df = std

    return best_df


def detect_header_row(df: pd.DataFrame, max_rows: int = 30) -> int:
    best_row = 0
    best_key = (-1, -1)
    limit = min(max_rows, len(df))
    for row_idx in range(limit):
        score = _score_header_row_cells(_row_cells(df.iloc[row_idx]))
        if score > best_key:
            best_key = score
            best_row = row_idx
    return best_row


def _read_sheet_raw(path_or_buffer: Any, sheet_name: str | int) -> pd.DataFrame:
    raw = pd.read_excel(path_or_buffer, sheet_name=sheet_name, header=None, dtype=object)
    best = _find_best_gstr_table(raw)
    if best is not None and not best.empty:
        return best

    header_row = detect_header_row(raw)
    candidate = _dataframe_from_header(raw, [header_row])
    if candidate.empty:
        return candidate
    mapping = map_columns(candidate)
    filtered = _filter_invoice_rows(candidate, mapping)
    return extract_standard_frame(filtered, mapping)


def _looks_like_non_gstr_export(df: pd.DataFrame) -> bool:
    columns_text = " ".join(_normalize_header(col) for col in df.columns)
    hits = sum(1 for hint in NON_GSTR_COLUMN_HINTS if hint in columns_text)
    has_tax_columns = any(
        _score_header(col, FIELD_ALIASES[field]) >= 40
        or _keyword_score(col, KEYWORD_FALLBACKS.get(field, ()))
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
            if field == "taxable_value" and "invoice value" in _normalize_header(col):
                score = max(score - 30, 0)
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
    if any(skip in lowered for skip in GSTR_SKIP_SHEET_KEYWORDS):
        return False
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
    data_sheets = [name for name in sheet_names if _is_gstr_data_sheet(name)]

    frames: list[pd.DataFrame] = []
    errors: list[str] = []
    non_gstr_detected = False

    for sheet in data_sheets:
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            sheet_df = _read_sheet_raw(source, sheet)
            if sheet_df.empty:
                continue
            if _looks_like_non_gstr_export(sheet_df):
                non_gstr_detected = True
                continue
            frames.append(sheet_df)
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
    detail = errors[0] if errors else "No readable invoice rows found in B2B/B2BA/CDNR sheets."
    raise ValueError(
        f"{prefix}Could not read GSTR-2A/2B Excel. {detail} "
        "Use the GST portal download with B2B, B2BA, and CDNR sheets."
    )
