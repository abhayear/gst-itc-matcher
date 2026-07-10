"""Core GST ITC invoice matching engine."""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from typing import Any

import pandas as pd

from .column_mapping import read_gstr_excel, read_purchase_register_excel
from .consolidate import consolidate_gstr_registers, consolidate_purchase_registers
from .itc_dashboard import (
    ITCDashboardSummary,
    action_plan_dataframe,
    build_itc_claim_plan,
    build_itc_dashboard,
    claim_plan_rows,
    dashboard_summary_rows,
    enrich_match_result,
    generate_action_plan,
    generate_ai_recommendations,
    generate_optimization_insights,
)
from .normalize import (
    amounts_equal,
    normalize_amount,
    normalize_date,
    normalize_gstin,
    normalize_invoice_no,
)

MATCH_FULLY = "Fully Matched"
MATCH_TAX = "Tax Mismatch"
MATCH_GSTIN = "GSTIN Mismatch"
MATCH_INV_NO = "Inv No Mismatch"
MATCH_INV_DATE = "Inv Date Mismatch"
MATCH_NOT = "Not Matched"
MATCH_DUPLICATE = "Duplicate Invoice"
MATCH_MISSING_IN_PR = "Missing in Purchase Register"

OUTPUT_COLUMNS = [
    "Supplier GSTIN",
    "Supplier Name",
    "Invoice No.",
    "Invoice Date",
    "Taxable Value",
    "IGST",
    "CGST",
    "SGST",
    "Purchase Register",
    "GSTR-2A/2B",
    "Match Status",
    "Available ITC",
    "Eligible ITC",
    "Non-Eligible ITC",
    "ITC Category",
    "ITC Taken",
    "AI Recommendation",
    "Remarks",
]


@dataclass
class MatchSummary:
    total_rows: int
    fully_matched: int
    tax_mismatch: int
    gstin_mismatch: int
    inv_no_mismatch: int
    inv_date_mismatch: int
    not_matched: int
    duplicate: int
    missing_in_pr: int
    total_itc_igst: float
    total_itc_cgst: float
    total_itc_sgst: float
    total_itc: float


def _prepare_records(df: pd.DataFrame, source_label: str) -> pd.DataFrame:
    records = df.copy()
    records["_gstin"] = records["gstin"].map(normalize_gstin)
    records["_invoice_no"] = records["invoice_no"].map(normalize_invoice_no)
    records["_invoice_date"] = records["invoice_date"].map(normalize_date)
    records["_taxable"] = records["taxable_value"].map(normalize_amount)
    records["_igst"] = records["igst"].map(normalize_amount)
    records["_cgst"] = records["cgst"].map(normalize_amount)
    records["_sgst"] = records["sgst"].map(normalize_amount)
    records["_source"] = source_label
    records["_match_key"] = records.apply(
        lambda row: f"{row['_gstin']}|{row['_invoice_no']}", axis=1
    )
    records["_full_key"] = records.apply(
        lambda row: f"{row['_gstin']}|{row['_invoice_no']}|{row['_invoice_date']}",
        axis=1
    )
    return records


def _taxes_match(pr_row: pd.Series, gstr_row: pd.Series, tolerance: float) -> bool:
    return (
        amounts_equal(pr_row["_taxable"], gstr_row["_taxable"], tolerance)
        and amounts_equal(pr_row["_igst"], gstr_row["_igst"], tolerance)
        and amounts_equal(pr_row["_cgst"], gstr_row["_cgst"], tolerance)
        and amounts_equal(pr_row["_sgst"], gstr_row["_sgst"], tolerance)
    )


def _format_amount(value: float) -> float:
    return round(value, 2)


def _build_output_row(
    pr_row: pd.Series | None,
    gstr_row: pd.Series | None,
    status: str,
    itc_taken: float,
    remarks: str,
) -> dict[str, Any]:
    base = pr_row if pr_row is not None else gstr_row
    assert base is not None

    pr_flag = "Yes" if pr_row is not None else "No"
    gstr_flag = "Yes" if gstr_row is not None else "No"

    return {
        "Supplier GSTIN": base["gstin"],
        "Supplier Name": base["supplier_name"],
        "Invoice No.": base["invoice_no"],
        "Invoice Date": base["invoice_date"],
        "Taxable Value": _format_amount(normalize_amount(base["taxable_value"])),
        "IGST": _format_amount(normalize_amount(base["igst"])),
        "CGST": _format_amount(normalize_amount(base["cgst"])),
        "SGST": _format_amount(normalize_amount(base["sgst"])),
        "Purchase Register": pr_flag,
        "GSTR-2A/2B": gstr_flag,
        "Match Status": status,
        "ITC Taken": _format_amount(itc_taken),
        "Remarks": remarks,
    }


def _compute_itc(
    pr_row: pd.Series, status: str, gstr_row: pd.Series | None
) -> tuple[float, float, float, float, str]:
    igst = pr_row["_igst"]
    cgst = pr_row["_cgst"]
    sgst = pr_row["_sgst"]
    total_tax = igst + cgst + sgst

    if status == MATCH_FULLY:
        return igst, cgst, sgst, total_tax, "Eligible ITC - full match with GSTR-2A/2B."

    if status == MATCH_TAX and gstr_row is not None:
        igst = min(pr_row["_igst"], gstr_row["_igst"])
        cgst = min(pr_row["_cgst"], gstr_row["_cgst"])
        sgst = min(pr_row["_sgst"], gstr_row["_sgst"])
        return igst, cgst, sgst, igst + cgst + sgst, "Partial ITC - tax mismatch. Verify before claiming."

    if status == MATCH_INV_DATE:
        return igst, cgst, sgst, total_tax, "ITC held - invoice date mismatch. Verify with supplier."

    if status in (MATCH_INV_NO, MATCH_GSTIN):
        return 0.0, 0.0, 0.0, 0.0, "ITC not taken - key field mismatch. Verify invoice details."

    if status == MATCH_DUPLICATE:
        return 0.0, 0.0, 0.0, 0.0, "Duplicate entry in Purchase Register. Remove duplicate claim."

    if status == MATCH_NOT:
        return 0.0, 0.0, 0.0, 0.0, "Not found in GSTR-2A/2B. Follow up with supplier before claiming."

    return 0.0, 0.0, 0.0, 0.0, ""


def _attach_tax_breakup(
    row: dict[str, Any],
    pr_row: pd.Series | None,
    gstr_row: pd.Series | None,
) -> None:
    if pr_row is not None:
        row["_pr_igst"] = _format_amount(pr_row["_igst"])
        row["_pr_cgst"] = _format_amount(pr_row["_cgst"])
        row["_pr_sgst"] = _format_amount(pr_row["_sgst"])
    else:
        row["_pr_igst"] = 0.0
        row["_pr_cgst"] = 0.0
        row["_pr_sgst"] = 0.0

    if gstr_row is not None:
        row["_gstr_igst"] = _format_amount(gstr_row["_igst"])
        row["_gstr_cgst"] = _format_amount(gstr_row["_cgst"])
        row["_gstr_sgst"] = _format_amount(gstr_row["_sgst"])
    else:
        row["_gstr_igst"] = 0.0
        row["_gstr_cgst"] = 0.0
        row["_gstr_sgst"] = 0.0


def _find_best_gstr_match(
    pr_row: pd.Series,
    gstr_df: pd.DataFrame,
    used_gstr_indices: set[int],
) -> tuple[pd.Series | None, str]:
    candidates = gstr_df[~gstr_df.index.isin(used_gstr_indices)]

    full_matches = candidates[candidates["_full_key"] == pr_row["_full_key"]]
    if len(full_matches) == 1:
        return full_matches.iloc[0], MATCH_FULLY
    if len(full_matches) > 1:
        return full_matches.iloc[0], MATCH_DUPLICATE

    key_matches = candidates[candidates["_match_key"] == pr_row["_match_key"]]
    if not key_matches.empty:
        row = key_matches.iloc[0]
        if row["_invoice_date"] != pr_row["_invoice_date"]:
            return row, MATCH_INV_DATE
        return row, MATCH_TAX

    inv_date_matches = candidates[
        (candidates["_gstin"] == pr_row["_gstin"])
        & (candidates["_invoice_date"] == pr_row["_invoice_date"])
    ]
    if not inv_date_matches.empty:
        return inv_date_matches.iloc[0], MATCH_INV_NO

    inv_no_only = candidates[candidates["_invoice_no"] == pr_row["_invoice_no"]]
    if not inv_no_only.empty:
        return inv_no_only.iloc[0], MATCH_GSTIN

    return None, MATCH_NOT


def match_invoices(
    purchase_register: pd.DataFrame,
    gstr_data: pd.DataFrame,
    tax_tolerance: float = 1.0,
) -> tuple[pd.DataFrame, MatchSummary, ITCDashboardSummary]:
    pr = _prepare_records(purchase_register, "PR")
    gstr = _prepare_records(gstr_data, "GSTR")

    duplicate_keys = pr["_full_key"].value_counts()
    duplicate_keys = set(duplicate_keys[duplicate_keys > 1].index)
    seen_keys: set[str] = set()

    used_gstr_indices: set[int] = set()
    output_rows: list[dict[str, Any]] = []

    for _, pr_row in pr.iterrows():
        if pr_row["_full_key"] in duplicate_keys:
            if pr_row["_full_key"] in seen_keys:
                _, _, _, itc, remarks = _compute_itc(pr_row, MATCH_DUPLICATE, None)
                row = _build_output_row(pr_row, None, MATCH_DUPLICATE, itc, remarks)
                row["_itc_igst"] = 0.0
                row["_itc_cgst"] = 0.0
                row["_itc_sgst"] = 0.0
                _attach_tax_breakup(row, pr_row, None)
                output_rows.append(row)
                continue
            seen_keys.add(pr_row["_full_key"])

        gstr_row, preliminary_status = _find_best_gstr_match(pr_row, gstr, used_gstr_indices)

        if gstr_row is not None:
            used_gstr_indices.add(gstr_row.name)
            if preliminary_status == MATCH_FULLY and not _taxes_match(pr_row, gstr_row, tax_tolerance):
                status = MATCH_TAX
            else:
                status = preliminary_status
        else:
            status = MATCH_NOT
            gstr_row = None

        itc_igst, itc_cgst, itc_sgst, itc, remarks = _compute_itc(pr_row, status, gstr_row)
        row = _build_output_row(pr_row, gstr_row, status, itc, remarks)
        row["_itc_igst"] = _format_amount(itc_igst)
        row["_itc_cgst"] = _format_amount(itc_cgst)
        row["_itc_sgst"] = _format_amount(itc_sgst)
        _attach_tax_breakup(row, pr_row, gstr_row)
        output_rows.append(row)

    for idx, gstr_row in gstr.iterrows():
        if idx in used_gstr_indices:
            continue
        row = _build_output_row(
            None,
            gstr_row,
            MATCH_MISSING_IN_PR,
            0.0,
            "Present in GSTR-2A/2B but missing from Purchase Register.",
        )
        row["_itc_igst"] = 0.0
        row["_itc_cgst"] = 0.0
        row["_itc_sgst"] = 0.0
        _attach_tax_breakup(row, None, gstr_row)
        output_rows.append(row)

    full_result = enrich_match_result(pd.DataFrame(output_rows))
    dashboard = build_itc_dashboard(full_result, gstr)
    summary = _build_summary(full_result)
    result = full_result[OUTPUT_COLUMNS].copy()
    return result, summary, dashboard


def _build_summary(full_result: pd.DataFrame) -> MatchSummary:
    status_counts = full_result["Match Status"].value_counts().to_dict()
    return MatchSummary(
        total_rows=len(full_result),
        fully_matched=status_counts.get(MATCH_FULLY, 0),
        tax_mismatch=status_counts.get(MATCH_TAX, 0),
        gstin_mismatch=status_counts.get(MATCH_GSTIN, 0),
        inv_no_mismatch=status_counts.get(MATCH_INV_NO, 0),
        inv_date_mismatch=status_counts.get(MATCH_INV_DATE, 0),
        not_matched=status_counts.get(MATCH_NOT, 0),
        duplicate=status_counts.get(MATCH_DUPLICATE, 0),
        missing_in_pr=status_counts.get(MATCH_MISSING_IN_PR, 0),
        total_itc_igst=_format_amount(full_result["_itc_igst"].sum()),
        total_itc_cgst=_format_amount(full_result["_itc_cgst"].sum()),
        total_itc_sgst=_format_amount(full_result["_itc_sgst"].sum()),
        total_itc=_format_amount(full_result["ITC Taken"].sum()),
    )


def load_and_match(
    purchase_file: Any,
    gstr_file: Any,
    tax_tolerance: float = 1.0,
) -> tuple[pd.DataFrame, MatchSummary, ITCDashboardSummary]:
    pr_std = read_purchase_register_excel(purchase_file)
    gstr_std = read_gstr_excel(gstr_file)
    return match_invoices(pr_std, gstr_std, tax_tolerance=tax_tolerance)


def load_and_match_consolidated(
    purchase_sources: list[tuple[Any, str]],
    gstr_file: Any,
    tax_tolerance: float = 1.0,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, MatchSummary, ITCDashboardSummary]:
    pr_std = consolidate_purchase_registers(purchase_sources)
    gstr_std = read_gstr_excel(gstr_file)
    result, summary, dashboard = match_invoices(pr_std, gstr_std, tax_tolerance=tax_tolerance)
    return pr_std, gstr_std, result, summary, dashboard


def load_and_match_with_consolidation(
    purchase_sources: list[tuple[Any, str]],
    gstr_sources: list[tuple[Any, str]],
    tax_tolerance: float = 1.0,
) -> tuple[pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame, MatchSummary, ITCDashboardSummary]:
    pr_std = consolidate_purchase_registers(purchase_sources)
    gstr_std = consolidate_gstr_registers(gstr_sources)
    result, summary, dashboard = match_invoices(pr_std, gstr_std, tax_tolerance=tax_tolerance)
    return pr_std, gstr_std, result, summary, dashboard

def _summary_rows(
    summary: MatchSummary,
    dashboard: ITCDashboardSummary | None = None,
    result: pd.DataFrame | None = None,
) -> list[tuple[str, Any]]:
    rows = [
        ("Total Rows", summary.total_rows),
        ("Fully Matched", summary.fully_matched),
        ("Tax Mismatch", summary.tax_mismatch),
        ("GSTIN Mismatch", summary.gstin_mismatch),
        ("Inv No Mismatch", summary.inv_no_mismatch),
        ("Inv Date Mismatch", summary.inv_date_mismatch),
        ("Not Matched", summary.not_matched),
        ("Duplicate Invoice", summary.duplicate),
        ("Missing in Purchase Register", summary.missing_in_pr),
        ("", ""),
        ("Eligible ITC - IGST", summary.total_itc_igst),
        ("Eligible ITC - CGST", summary.total_itc_cgst),
        ("Eligible ITC - SGST", summary.total_itc_sgst),
        ("Total ITC Taken", summary.total_itc),
    ]
    if dashboard is not None:
        rows.extend([("", ""), *dashboard_summary_rows(dashboard)])
        if result is not None:
            plan = build_itc_claim_plan(dashboard, result)
            rows.extend([("", ""), *claim_plan_rows(plan)])
    return rows


def export_to_excel(
    result: pd.DataFrame,
    summary: MatchSummary,
    dashboard: ITCDashboardSummary | None = None,
) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        result.to_excel(writer, sheet_name="ITC Matching", index=False)
        summary_df = pd.DataFrame(_summary_rows(summary, dashboard, result), columns=["Metric", "Value"])
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        if dashboard is not None:
            plan = build_itc_claim_plan(dashboard, result)
            tips = generate_ai_recommendations(dashboard, result)
            tips_df = pd.DataFrame({"AI Recommendation": tips})
            tips_df.to_excel(writer, sheet_name="AI Recommendations", index=False)
            opt_insights = generate_optimization_insights(dashboard, plan, result)
            opt_df = pd.DataFrame({"Optimization Insight": opt_insights})
            opt_df.to_excel(writer, sheet_name="ITC Optimization", index=False)
            claim_df = pd.DataFrame(claim_plan_rows(plan), columns=["Metric", "Value"])
            claim_df.to_excel(writer, sheet_name="ITC Claim Plan", index=False)
            actions = generate_action_plan(dashboard, result)
            action_plan_dataframe(actions).to_excel(writer, sheet_name="Action Plan", index=False)

        workbook = writer.book
        worksheet = writer.sheets["ITC Matching"]
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        for col_num, value in enumerate(result.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
        worksheet.autofilter(0, 0, len(result), len(result.columns) - 1)
        worksheet.freeze_panes(1, 0)

    buffer.seek(0)
    return buffer.getvalue()


def export_to_csv(
    result: pd.DataFrame,
    summary: MatchSummary,
    dashboard: ITCDashboardSummary | None = None,
) -> bytes:
    buffer = BytesIO()
    buffer.write("\ufeff".encode("utf-8"))
    result.to_csv(buffer, index=False, lineterminator="\n")
    buffer.write(b"\n")
    summary_df = pd.DataFrame(_summary_rows(summary, dashboard, result), columns=["Metric", "Value"])
    summary_df.to_csv(buffer, index=False, lineterminator="\n")
    buffer.seek(0)
    return buffer.getvalue()


def _pdf_format_cell(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float):
        return f"{value:,.2f}"
    return str(value).replace("\n", " ")[:120]


def export_to_pdf(
    result: pd.DataFrame,
    summary: MatchSummary,
    dashboard: ITCDashboardSummary | None = None,
) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=24,
        rightMargin=24,
        topMargin=24,
        bottomMargin=24,
    )
    styles = getSampleStyleSheet()
    story: list[Any] = [
        Paragraph("GST ITC Matching Report", styles["Title"]),
        Spacer(1, 0.2 * inch),
    ]

    summary_table_data = [["Metric", "Value"]] + [
        [metric, _pdf_format_cell(value)] for metric, value in _summary_rows(summary, dashboard, result)
    ]
    summary_table = Table(summary_table_data, colWidths=[3.2 * inch, 2.0 * inch])
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E1F2")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    story.extend([summary_table, Spacer(1, 0.25 * inch), Paragraph("Matching Results", styles["Heading2"])])

    headers = list(result.columns)
    table_data = [headers] + [
        [_pdf_format_cell(value) for value in row] for row in result.itertuples(index=False, name=None)
    ]
    page_width = landscape(A4)[0] - 48
    col_width = page_width / max(len(headers), 1)
    results_table = Table(table_data, colWidths=[col_width] * len(headers), repeatRows=1)
    results_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E1F2")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 6),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    story.append(results_table)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()
