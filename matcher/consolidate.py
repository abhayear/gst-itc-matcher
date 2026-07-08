"""Consolidate multiple Purchase Register files (e.g. Sales + Service)."""

from __future__ import annotations

from io import BytesIO
from typing import Any

import pandas as pd

from .column_mapping import extract_standard_frame, map_columns, read_excel_with_headers

CONSOLIDATED_COLUMNS = [
    "Supplier GSTIN",
    "Supplier Name",
    "Invoice No.",
    "Invoice Date",
    "Taxable Value",
    "IGST",
    "CGST",
    "SGST",
    "Register Type",
]

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


def _read_standard_purchase_register(source: Any, register_type: str) -> pd.DataFrame:
    raw = read_excel_with_headers(source)
    mapping = map_columns(raw)
    std = extract_standard_frame(raw, mapping)
    std["register_type"] = register_type
    return std


def consolidate_purchase_registers(
    sources: list[tuple[Any, str]],
) -> pd.DataFrame:
    if not sources:
        raise ValueError("At least one Purchase Register file is required.")

    frames: list[pd.DataFrame] = []
    for source, register_type in sources:
        if hasattr(source, "seek"):
            source.seek(0)
        frames.append(_read_standard_purchase_register(source, register_type))

    consolidated = pd.concat(frames, ignore_index=True)
    return consolidated


def consolidated_to_display(consolidated: pd.DataFrame) -> pd.DataFrame:
    display = pd.DataFrame(
        {
            FIELD_TO_OUTPUT[field]: consolidated[field]
            for field in FIELD_TO_OUTPUT
        }
    )
    display["Register Type"] = consolidated["register_type"]
    return display[CONSOLIDATED_COLUMNS]


def export_consolidated_purchase_register(consolidated: pd.DataFrame) -> bytes:
    display = consolidated_to_display(consolidated)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        display.to_excel(writer, sheet_name="Consolidated PR", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Consolidated PR"]
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#E2EFDA", "border": 1})
        for col_num, value in enumerate(display.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
        worksheet.autofilter(0, 0, len(display), len(display.columns) - 1)
        worksheet.freeze_panes(1, 0)
    buffer.seek(0)
    return buffer.getvalue()
