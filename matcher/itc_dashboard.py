"""ITC dashboard metrics, categorization, and AI-style recommendations."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import pandas as pd

from .normalize import normalize_amount

MATCH_FULLY = "Fully Matched"
MATCH_TAX = "Tax Mismatch"
MATCH_GSTIN = "GSTIN Mismatch"
MATCH_INV_NO = "Inv No Mismatch"
MATCH_INV_DATE = "Inv Date Mismatch"
MATCH_NOT = "Not Matched"
MATCH_DUPLICATE = "Duplicate Invoice"
MATCH_MISSING_IN_PR = "Missing in Purchase Register"

ITC_ELIGIBLE = "Eligible"
ITC_PENDING = "Pending Review"
ITC_NON_ELIGIBLE = "Non-Eligible"

DASHBOARD_COLUMNS = [
    "Available ITC",
    "Eligible ITC",
    "Non-Eligible ITC",
    "ITC Category",
    "AI Recommendation",
]


@dataclass
class ITCDashboardSummary:
    available_igst: float
    available_cgst: float
    available_sgst: float
    available_total: float
    eligible_igst: float
    eligible_cgst: float
    eligible_sgst: float
    eligible_total: float
    non_eligible_igst: float
    non_eligible_cgst: float
    non_eligible_sgst: float
    non_eligible_total: float
    pending_igst: float
    pending_cgst: float
    pending_sgst: float
    pending_total: float
    eligible_invoices: int
    pending_invoices: int
    non_eligible_invoices: int
    claim_rate_pct: float


def _round(value: float) -> float:
    return round(value, 2)


def _row_tax(prefix: str, row: pd.Series) -> float:
    return _round(
        normalize_amount(row.get(f"{prefix}igst", 0))
        + normalize_amount(row.get(f"{prefix}cgst", 0))
        + normalize_amount(row.get(f"{prefix}sgst", 0))
    )


def _categorize_itc(status: str) -> str:
    if status == MATCH_FULLY:
        return ITC_ELIGIBLE
    if status in (MATCH_TAX, MATCH_INV_DATE):
        return ITC_PENDING
    return ITC_NON_ELIGIBLE


def _row_recommendation(
    status: str,
    supplier_name: str,
    invoice_no: str,
    eligible: float,
    available: float,
    non_eligible: float,
) -> str:
    inv = invoice_no or "invoice"
    party = supplier_name or "supplier"

    if status == MATCH_FULLY:
        return (
            f"Claim full ITC of ₹{eligible:,.2f} in GSTR-3B Table 4. "
            f"{inv} from {party} is fully matched with GSTR-2B."
        )
    if status == MATCH_TAX:
        return (
            f"Verify tax breakup before claiming. Claim only ₹{eligible:,.2f} "
            f"(lower of books vs GSTR-2B). ₹{non_eligible:,.2f} is blocked until "
            f"{party} amends the return or you correct the purchase entry."
        )
    if status == MATCH_INV_DATE:
        return (
            f"Hold ITC of ₹{eligible:,.2f} until invoice date is reconciled with {party}. "
            "Do not claim until GSTIN, invoice number, and date all align with GSTR-2B."
        )
    if status == MATCH_NOT:
        return (
            f"Do not claim ITC on {inv}. Not found in GSTR-2B — ask {party} to file "
            "or amend GSTR-1 so the invoice appears in your next 2B download."
        )
    if status == MATCH_MISSING_IN_PR:
        return (
            f"₹{available:,.2f} ITC available in GSTR-2B but missing from your Purchase Register. "
            f"Book the purchase for {inv} from {party}, then re-run reconciliation before claiming."
        )
    if status == MATCH_GSTIN:
        return (
            f"GSTIN mismatch on {inv}. Confirm supplier GSTIN in Vyapar and GSTR-2B "
            "before claiming any ITC."
        )
    if status == MATCH_INV_NO:
        return (
            f"Invoice number mismatch for {party}. Match exact bill number "
            f"({inv}) with GSTR-2B before claiming ITC."
        )
    if status == MATCH_DUPLICATE:
        return (
            f"Duplicate entry for {inv} in Purchase Register. Remove the duplicate "
            "and claim ITC only once."
        )
    return "Review this invoice manually before claiming ITC."


def enrich_match_result(full_result: pd.DataFrame) -> pd.DataFrame:
    """Add available / eligible / non-eligible ITC columns and recommendations."""
    enriched = full_result.copy()
    rows: list[dict[str, Any]] = []

    for _, row in enriched.iterrows():
        status = row["Match Status"]
        gstr_tax = _row_tax("_gstr_", row)
        pr_tax = _row_tax("_pr_", row)
        eligible = _round(normalize_amount(row.get("ITC Taken", 0)))
        category = _categorize_itc(status)

        if status == MATCH_MISSING_IN_PR:
            available = gstr_tax
            non_eligible = gstr_tax
        elif status == MATCH_NOT:
            available = 0.0
            non_eligible = pr_tax
        elif gstr_tax > 0:
            available = gstr_tax
            non_eligible = _round(max(gstr_tax - eligible, 0.0))
        else:
            available = pr_tax
            non_eligible = _round(max(pr_tax - eligible, 0.0))

        if category == ITC_NON_ELIGIBLE and status not in (MATCH_NOT, MATCH_MISSING_IN_PR):
            non_eligible = max(non_eligible, available)

        recommendation = _row_recommendation(
            status,
            str(row.get("Supplier Name", "")),
            str(row.get("Invoice No.", "")),
            eligible,
            available,
            non_eligible,
        )
        rows.append(
            {
                "Available ITC": available,
                "Eligible ITC": eligible,
                "Non-Eligible ITC": non_eligible,
                "ITC Category": category,
                "AI Recommendation": recommendation,
            }
        )

    dashboard_df = pd.DataFrame(rows, index=enriched.index)
    for column in DASHBOARD_COLUMNS:
        enriched[column] = dashboard_df[column]
    return enriched


def _gstr_tax_totals(gstr_df: pd.DataFrame) -> tuple[float, float, float, float]:
    igst = _round(gstr_df["igst"].map(normalize_amount).sum())
    cgst = _round(gstr_df["cgst"].map(normalize_amount).sum())
    sgst = _round(gstr_df["sgst"].map(normalize_amount).sum())
    return igst, cgst, sgst, _round(igst + cgst + sgst)


def build_itc_dashboard(enriched: pd.DataFrame, gstr_df: pd.DataFrame) -> ITCDashboardSummary:
    gstr_igst, gstr_cgst, gstr_sgst, gstr_total = _gstr_tax_totals(gstr_df)

    eligible_mask = enriched["ITC Category"] == ITC_ELIGIBLE
    pending_mask = enriched["ITC Category"] == ITC_PENDING
    non_eligible_mask = enriched["ITC Category"] == ITC_NON_ELIGIBLE

    eligible_total = _round(enriched.loc[eligible_mask, "Eligible ITC"].sum())
    pending_total = _round(enriched.loc[pending_mask, "Eligible ITC"].sum())
    row_non_eligible = _round(enriched["Non-Eligible ITC"].sum())

    eligible_igst = _round(enriched.loc[eligible_mask, "_itc_igst"].sum())
    eligible_cgst = _round(enriched.loc[eligible_mask, "_itc_cgst"].sum())
    eligible_sgst = _round(enriched.loc[eligible_mask, "_itc_sgst"].sum())

    pending_igst = _round(enriched.loc[pending_mask, "_itc_igst"].sum())
    pending_cgst = _round(enriched.loc[pending_mask, "_itc_cgst"].sum())
    pending_sgst = _round(enriched.loc[pending_mask, "_itc_sgst"].sum())

    non_eligible_total = _round(max(gstr_total - eligible_total - pending_total, row_non_eligible))
    non_eligible_igst = _round(max(gstr_igst - eligible_igst - pending_igst, 0.0))
    non_eligible_cgst = _round(max(gstr_cgst - eligible_cgst - pending_cgst, 0.0))
    non_eligible_sgst = _round(max(gstr_sgst - eligible_sgst - pending_sgst, 0.0))

    claim_rate = _round((eligible_total / gstr_total * 100) if gstr_total else 0.0)

    return ITCDashboardSummary(
        available_igst=gstr_igst,
        available_cgst=gstr_cgst,
        available_sgst=gstr_sgst,
        available_total=gstr_total,
        eligible_igst=eligible_igst,
        eligible_cgst=eligible_cgst,
        eligible_sgst=eligible_sgst,
        eligible_total=eligible_total,
        non_eligible_igst=non_eligible_igst,
        non_eligible_cgst=non_eligible_cgst,
        non_eligible_sgst=non_eligible_sgst,
        non_eligible_total=non_eligible_total,
        pending_igst=pending_igst,
        pending_cgst=pending_cgst,
        pending_sgst=pending_sgst,
        pending_total=pending_total,
        eligible_invoices=int(eligible_mask.sum()),
        pending_invoices=int(pending_mask.sum()),
        non_eligible_invoices=int(non_eligible_mask.sum()),
        claim_rate_pct=claim_rate,
    )


def generate_ai_recommendations(dashboard: ITCDashboardSummary, enriched: pd.DataFrame) -> list[str]:
    tips: list[str] = [
        (
            f"GSTR-2B shows ₹{dashboard.available_total:,.2f} available ITC. "
            f"You can safely claim ₹{dashboard.eligible_total:,.2f} ({dashboard.claim_rate_pct}% of GSTR ITC) "
            f"in GSTR-3B Table 4 after final verification."
        )
    ]

    if dashboard.pending_total > 0:
        tips.append(
            f"₹{dashboard.pending_total:,.2f} ITC is in Pending Review across "
            f"{dashboard.pending_invoices} invoice(s). Resolve tax or date mismatches before filing."
        )

    if dashboard.non_eligible_total > 0:
        tips.append(
            f"₹{dashboard.non_eligible_total:,.2f} ITC is not eligible yet across "
            f"{dashboard.non_eligible_invoices} invoice(s). Follow up with suppliers or fix Vyapar entries."
        )

    if (enriched["Match Status"] == MATCH_MISSING_IN_PR).any():
        count = int((enriched["Match Status"] == MATCH_MISSING_IN_PR).sum())
        tips.append(
            f"{count} invoice(s) are in GSTR-2B but missing from your Purchase Register — book them in Vyapar first."
        )

    if (enriched["Match Status"] == MATCH_NOT).any():
        count = int((enriched["Match Status"] == MATCH_NOT).sum())
        tips.append(
            f"{count} Vyapar purchase(s) are not in GSTR-2B — ask suppliers to file GSTR-1 before claiming ITC."
        )

    if dashboard.claim_rate_pct >= 90:
        tips.append("Strong reconciliation. Review pending items, then proceed with GSTR-3B filing.")
    elif dashboard.claim_rate_pct >= 70:
        tips.append("Moderate match rate — clear pending and non-eligible invoices before claiming full ITC.")
    else:
        tips.append("Low match rate — complete supplier follow-ups before claiming ITC to avoid ineligible credit.")

    return tips


def dashboard_summary_rows(dashboard: ITCDashboardSummary) -> list[tuple[str, Any]]:
    return [
        ("Available ITC (GSTR-2B)", dashboard.available_total),
        ("Available IGST", dashboard.available_igst),
        ("Available CGST", dashboard.available_cgst),
        ("Available SGST", dashboard.available_sgst),
        ("", ""),
        ("Eligible ITC (Safe to Claim)", dashboard.eligible_total),
        ("Eligible IGST", dashboard.eligible_igst),
        ("Eligible CGST", dashboard.eligible_cgst),
        ("Eligible SGST", dashboard.eligible_sgst),
        ("", ""),
        ("Pending Review ITC", dashboard.pending_total),
        ("Pending IGST", dashboard.pending_igst),
        ("Pending CGST", dashboard.pending_cgst),
        ("Pending SGST", dashboard.pending_sgst),
        ("", ""),
        ("Non-Eligible ITC", dashboard.non_eligible_total),
        ("Non-Eligible IGST", dashboard.non_eligible_igst),
        ("Non-Eligible CGST", dashboard.non_eligible_cgst),
        ("Non-Eligible SGST", dashboard.non_eligible_sgst),
        ("", ""),
        ("Claim Rate (% of GSTR ITC)", dashboard.claim_rate_pct),
        ("Eligible Invoices", dashboard.eligible_invoices),
        ("Pending Invoices", dashboard.pending_invoices),
        ("Non-Eligible Invoices", dashboard.non_eligible_invoices),
    ]
