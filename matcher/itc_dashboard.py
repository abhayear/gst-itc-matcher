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
class ITCClaimPlan:
    claim_now_igst: float
    claim_now_cgst: float
    claim_now_sgst: float
    claim_now_total: float
    hold_itc: float
    blocked_itc: float
    recoverable_pending: float
    recoverable_missing_pr: float
    optimized_claim_total: float
    optimization_gain: float
    filing_readiness: str
    filing_advice: str


@dataclass
class ActionPlanItem:
    priority: str
    action: str
    impact_inr: float
    invoice_count: int
    timeline: str
    owner: str


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


def build_itc_claim_plan(dashboard: ITCDashboardSummary, enriched: pd.DataFrame) -> ITCClaimPlan:
    missing_pr = enriched[enriched["Match Status"] == MATCH_MISSING_IN_PR]
    recoverable_missing_pr = _round(missing_pr["Available ITC"].sum())

    recoverable_pending = dashboard.pending_total
    optimized = _round(
        dashboard.eligible_total + recoverable_pending + recoverable_missing_pr
    )
    optimization_gain = _round(max(optimized - dashboard.eligible_total, 0.0))

    if dashboard.claim_rate_pct >= 90 and dashboard.pending_total == 0:
        readiness = "Ready to File"
        advice = (
            f"Claim ₹{dashboard.eligible_total:,.2f} (IGST ₹{dashboard.eligible_igst:,.2f}, "
            f"CGST ₹{dashboard.eligible_cgst:,.2f}, SGST ₹{dashboard.eligible_sgst:,.2f}) "
            "in GSTR-3B Table 4. Reconciliation is strong — proceed with filing."
        )
    elif dashboard.claim_rate_pct >= 70:
        readiness = "File with Caution"
        advice = (
            f"Claim only the eligible ₹{dashboard.eligible_total:,.2f} now. "
            f"Hold ₹{dashboard.pending_total:,.2f} pending items until mismatches are cleared. "
            "Do not claim blocked ITC."
        )
    elif dashboard.eligible_total > 0:
        readiness = "Partial Claim Only"
        advice = (
            f"Claim ₹{dashboard.eligible_total:,.2f} only. "
            f"₹{dashboard.non_eligible_total:,.2f} is blocked and "
            f"₹{dashboard.pending_total:,.2f} needs review before any additional claim."
        )
    else:
        readiness = "Do Not Claim Yet"
        advice = (
            "No ITC is safe to claim this period. Complete supplier follow-ups and "
            "Vyapar booking corrections before filing GSTR-3B."
        )

    return ITCClaimPlan(
        claim_now_igst=dashboard.eligible_igst,
        claim_now_cgst=dashboard.eligible_cgst,
        claim_now_sgst=dashboard.eligible_sgst,
        claim_now_total=dashboard.eligible_total,
        hold_itc=dashboard.pending_total,
        blocked_itc=dashboard.non_eligible_total,
        recoverable_pending=recoverable_pending,
        recoverable_missing_pr=recoverable_missing_pr,
        optimized_claim_total=optimized,
        optimization_gain=optimization_gain,
        filing_readiness=readiness,
        filing_advice=advice,
    )


def generate_optimization_insights(
    dashboard: ITCDashboardSummary,
    plan: ITCClaimPlan,
    enriched: pd.DataFrame,
) -> list[str]:
    insights: list[str] = [
        (
            f"**Current claim:** ₹{plan.claim_now_total:,.2f} | "
            f"**Optimized potential:** ₹{plan.optimized_claim_total:,.2f} "
            f"(+₹{plan.optimization_gain:,.2f} if all actions completed)"
        ),
        (
            f"You are claiming {dashboard.claim_rate_pct}% of GSTR-2B available ITC. "
            f"₹{plan.blocked_itc:,.2f} remains blocked until mismatches are resolved."
        ),
    ]

    if plan.recoverable_pending > 0:
        insights.append(
            f"Fix {dashboard.pending_invoices} pending invoice(s) to unlock "
            f"₹{plan.recoverable_pending:,.2f} additional ITC."
        )

    if plan.recoverable_missing_pr > 0:
        count = int((enriched["Match Status"] == MATCH_MISSING_IN_PR).sum())
        insights.append(
            f"Book {count} missing purchase(s) in Vyapar to unlock "
            f"₹{plan.recoverable_missing_pr:,.2f} ITC already showing in GSTR-2B."
        )

    not_matched = enriched[enriched["Match Status"] == MATCH_NOT]
    if not not_matched.empty:
        supplier_count = not_matched["Supplier Name"].nunique()
        insights.append(
            f"Follow up with {supplier_count} supplier(s) on {len(not_matched)} invoice(s) "
            "not yet filed in GSTR-1 — no ITC until they appear in 2B."
        )

    if plan.optimization_gain > plan.claim_now_total * 0.1 and plan.optimization_gain > 1000:
        insights.append(
            "High optimization opportunity — resolving pending and missing items "
            "could significantly increase your eligible ITC before filing."
        )

    return insights


def generate_action_plan(dashboard: ITCDashboardSummary, enriched: pd.DataFrame) -> list[ActionPlanItem]:
    actions: list[ActionPlanItem] = []

    eligible = enriched[enriched["ITC Category"] == ITC_ELIGIBLE]
    if not eligible.empty:
        actions.append(
            ActionPlanItem(
                priority="P1 — Claim Now",
                action="Enter matched ITC in GSTR-3B Table 4 (IGST/CGST/SGST as per eligible breakup)",
                impact_inr=dashboard.eligible_total,
                invoice_count=len(eligible),
                timeline="Before GSTR-3B filing",
                owner="Accounts / Tax team",
            )
        )

    pending = enriched[enriched["ITC Category"] == ITC_PENDING]
    if not pending.empty:
        tax_pending = pending[pending["Match Status"] == MATCH_TAX]
        date_pending = pending[pending["Match Status"] == MATCH_INV_DATE]
        if not tax_pending.empty:
            actions.append(
                ActionPlanItem(
                    priority="P2 — Fix Tax Mismatch",
                    action="Reconcile tax amounts with suppliers or correct Vyapar purchase entries",
                    impact_inr=_round(tax_pending["Eligible ITC"].sum()),
                    invoice_count=len(tax_pending),
                    timeline="Within 3–5 days",
                    owner="Accounts + Supplier follow-up",
                )
            )
        if not date_pending.empty:
            actions.append(
                ActionPlanItem(
                    priority="P2 — Fix Date Mismatch",
                    action="Align invoice dates between Vyapar Purchase Register and supplier bills",
                    impact_inr=_round(date_pending["Eligible ITC"].sum()),
                    invoice_count=len(date_pending),
                    timeline="Within 3–5 days",
                    owner="Accounts team",
                )
            )

    missing_pr = enriched[enriched["Match Status"] == MATCH_MISSING_IN_PR]
    if not missing_pr.empty:
        actions.append(
            ActionPlanItem(
                priority="P2 — Book Missing Purchases",
                action="Record GSTR-2B invoices missing from Vyapar Purchase Register",
                impact_inr=_round(missing_pr["Available ITC"].sum()),
                invoice_count=len(missing_pr),
                timeline="Within 2–3 days",
                owner="Accounts / Vyapar entry",
            )
        )

    not_matched = enriched[enriched["Match Status"] == MATCH_NOT]
    if not not_matched.empty:
        actions.append(
            ActionPlanItem(
                priority="P3 — Supplier Follow-up",
                action="Ask suppliers to file or amend GSTR-1 so invoices appear in next GSTR-2B",
                impact_inr=_round(not_matched["Non-Eligible ITC"].sum()),
                invoice_count=len(not_matched),
                timeline="Before next return period",
                owner="Purchase / Vendor management",
            )
        )

    mismatch = enriched[enriched["Match Status"].isin([MATCH_GSTIN, MATCH_INV_NO])]
    if not mismatch.empty:
        actions.append(
            ActionPlanItem(
                priority="P3 — Correct Invoice Details",
                action="Fix GSTIN or invoice number mismatches in Vyapar to enable matching",
                impact_inr=_round(mismatch["Non-Eligible ITC"].sum()),
                invoice_count=len(mismatch),
                timeline="Within 1 week",
                owner="Accounts team",
            )
        )

    duplicates = enriched[enriched["Match Status"] == MATCH_DUPLICATE]
    if not duplicates.empty:
        actions.append(
            ActionPlanItem(
                priority="P1 — Remove Duplicates",
                action="Delete duplicate purchase entries in Vyapar to avoid excess ITC claim risk",
                impact_inr=0.0,
                invoice_count=len(duplicates),
                timeline="Immediately",
                owner="Accounts team",
            )
        )

    if dashboard.pending_total > 0:
        actions.append(
            ActionPlanItem(
                priority="Hold",
                action=f"Do not claim ₹{dashboard.pending_total:,.2f} pending ITC until mismatches are resolved",
                impact_inr=dashboard.pending_total,
                invoice_count=dashboard.pending_invoices,
                timeline="Until cleared",
                owner="Tax team",
            )
        )

    priority_order = {"P1 — Claim Now": 0, "P1 — Remove Duplicates": 1, "P2 — Fix Tax Mismatch": 2,
                      "P2 — Fix Date Mismatch": 3, "P2 — Book Missing Purchases": 4,
                      "P3 — Supplier Follow-up": 5, "P3 — Correct Invoice Details": 6, "Hold": 7}
    actions.sort(key=lambda item: priority_order.get(item.priority, 99))
    return actions


def claim_plan_rows(plan: ITCClaimPlan) -> list[tuple[str, Any]]:
    return [
        ("Filing Readiness", plan.filing_readiness),
        ("Filing Advice", plan.filing_advice),
        ("", ""),
        ("Claim Now - IGST (Table 4)", plan.claim_now_igst),
        ("Claim Now - CGST (Table 4)", plan.claim_now_cgst),
        ("Claim Now - SGST (Table 4)", plan.claim_now_sgst),
        ("Claim Now - Total", plan.claim_now_total),
        ("", ""),
        ("Hold ITC (Pending Review)", plan.hold_itc),
        ("Blocked ITC (Non-Eligible)", plan.blocked_itc),
        ("", ""),
        ("Recoverable if Pending Fixed", plan.recoverable_pending),
        ("Recoverable if Missing PR Booked", plan.recoverable_missing_pr),
        ("Optimized Claim Potential", plan.optimized_claim_total),
        ("Optimization Gain", plan.optimization_gain),
    ]


def action_plan_dataframe(actions: list[ActionPlanItem]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Priority": item.priority,
                "Action": item.action,
                "Impact (₹)": item.impact_inr,
                "Invoices": item.invoice_count,
                "Timeline": item.timeline,
                "Owner": item.owner,
            }
            for item in actions
        ]
    )


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
