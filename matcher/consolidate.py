"""Consolidate multiple register files (Purchase Register or GSTR-2A/2B)."""

from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd

from .column_mapping import extract_standard_frame, map_columns, read_excel_with_headers

FIELD_TO_OUTPUT = {
    "gstin": "Supplier GSTIN",
    "supplier_name": "Supplier Name",
    "invoice_no": "Invoice No.",
    "invoice_date": "Invoice Date",
    "taxable_value": "Taxable Value",
    "igst": "IGST",
    "cgst": "CGST",
    "sgst": "SGST",
}

PR_LABEL_COLUMN = "register_type"
GSTR_LABEL_COLUMN = "gstr_source"
PR_LABEL_OUTPUT = "Register Type"
GSTR_LABEL_OUTPUT = "Period"


def label_from_filename(filename: str) -> str:
    """Derive a period label from an uploaded file name."""
    stem = Path(filename).stem.replace("_", " ").replace("-", " ").strip()
    patterns = (
        r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s-]*(\d{4})\b",
        r"\b(\d{1,2})[\s-/](\d{4})\b",
        r"\b(\d{4})[\s-/](\d{1,2})\b",
        r"\b((?:19|20)\d{2})\b",
    )
    lowered = stem.lower()
    for pattern in patterns:
        match = re.search(pattern, lowered, re.I)
        if match:
            parts = [part for part in match.groups() if part]
            if len(parts) == 2 and parts[0].isdigit() and len(parts[0]) <= 2:
                return f"{parts[0].zfill(2)}-{parts[1]}"
            if len(parts) == 2 and parts[1].isdigit() and len(parts[1]) <= 2:
                return f"{parts[1].zfill(2)}-{parts[0]}"
            return "-".join(parts).title()
    cleaned = re.sub(r"\bgstr[\s-]*2[ab]?\b", "", stem, flags=re.I).strip()
    return cleaned or stem


def gstr_summary_caption(consolidated: pd.DataFrame) -> str:
    counts = consolidated[GSTR_LABEL_COLUMN].value_counts()
    parts = [f"{count} from {period}" for period, count in counts.items()]
    return " + ".join(parts) + " invoices"


def _read_standard_register(source: Any, source_type: str, label_column: str) -> pd.DataFrame:
    raw = read_excel_with_headers(source)
    mapping = map_columns(raw)
    std = extract_standard_frame(raw, mapping)
    std[label_column] = source_type
    return std


def _consolidate_registers(
    sources: list[tuple[Any, str]],
    label_column: str,
    empty_message: str,
) -> pd.DataFrame:
    if not sources:
        raise ValueError(empty_message)

    frames: list[pd.DataFrame] = []
    for source, source_type in sources:
        if hasattr(source, "seek"):
            source.seek(0)
        frames.append(_read_standard_register(source, source_type, label_column))

    return pd.concat(frames, ignore_index=True)


def consolidate_purchase_registers(sources: list[tuple[Any, str]]) -> pd.DataFrame:
    return _consolidate_registers(
        sources,
        PR_LABEL_COLUMN,
        "At least one Purchase Register file is required.",
    )


def consolidate_gstr_registers(sources: list[tuple[Any, str]]) -> pd.DataFrame:
    return _consolidate_registers(
        sources,
        GSTR_LABEL_COLUMN,
        "At least one GSTR-2A/2B file is required.",
    )


def consolidated_to_display(
    consolidated: pd.DataFrame,
    label_column: str,
    label_output: str,
) -> pd.DataFrame:
    display = pd.DataFrame({output: consolidated[field] for field, output in FIELD_TO_OUTPUT.items()})
    display[label_output] = consolidated[label_column]
    return display[[*FIELD_TO_OUTPUT.values(), label_output]]


def consolidated_pr_to_display(consolidated: pd.DataFrame) -> pd.DataFrame:
    return consolidated_to_display(consolidated, PR_LABEL_COLUMN, PR_LABEL_OUTPUT)


def consolidated_gstr_to_display(consolidated: pd.DataFrame) -> pd.DataFrame:
    return consolidated_to_display(consolidated, GSTR_LABEL_COLUMN, GSTR_LABEL_OUTPUT)


def _export_consolidated(
    consolidated: pd.DataFrame,
    label_column: str,
    label_output: str,
    sheet_name: str,
    header_color: str,
) -> bytes:
    display = consolidated_to_display(consolidated, label_column, label_output)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        display.to_excel(writer, sheet_name=sheet_name, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        header_fmt = workbook.add_format({"bold": True, "bg_color": header_color, "border": 1})
        for col_num, value in enumerate(display.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
        worksheet.autofilter(0, 0, len(display), len(display.columns) - 1)
        worksheet.freeze_panes(1, 0)
    buffer.seek(0)
    return buffer.getvalue()


def export_consolidated_purchase_register(consolidated: pd.DataFrame) -> bytes:
    return _export_consolidated(
        consolidated,
        PR_LABEL_COLUMN,
        PR_LABEL_OUTPUT,
        "Consolidated PR",
        "#E2EFDA",
    )


def export_consolidated_gstr(consolidated: pd.DataFrame) -> bytes:
    return _export_consolidated(
        consolidated,
        GSTR_LABEL_COLUMN,
        GSTR_LABEL_OUTPUT,
        "Consolidated GSTR",
        "#FCE4D6",
    )
