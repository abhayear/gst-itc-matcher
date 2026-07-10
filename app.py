"""Streamlit app for GST ITC invoice matching — upload files, get results."""

from __future__ import annotations

import sys
from pathlib import Path

# Ensure project root is on path for Streamlit Cloud
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pandas as pd
import streamlit as st

st.set_page_config(page_title="GST ITC Matcher", page_icon=":bar_chart:", layout="wide")

try:
    from matcher.consolidate import (
        consolidated_gstr_to_display,
        consolidated_pr_to_display,
        export_consolidated_gstr,
        export_consolidated_purchase_register,
        gstr_summary_caption,
        label_from_filename,
    )
    from matcher.engine import (
        export_to_csv,
        export_to_excel,
        export_to_pdf,
        generate_ai_recommendations,
        load_and_match_with_consolidation,
    )
    from matcher.itc_dashboard import (
        ITC_ELIGIBLE,
        ITC_NON_ELIGIBLE,
        ITC_PENDING,
        action_plan_dataframe,
        build_itc_claim_plan,
        generate_action_plan,
        generate_optimization_insights,
    )
except Exception as import_error:
    st.error("App failed to start. Please refresh in a minute or contact support.")
    st.exception(import_error)
    st.stop()

st.title("GST ITC Matcher")
st.caption("Upload Purchase Register(s) and GSTR-2A/2B — consolidation and matching run automatically.")

col_pr, col_gstr = st.columns(2)
with col_pr:
    pr_mode = st.radio(
        "Purchase Register",
        options=["Single file", "Sales + Service (Consolidate)"],
        horizontal=True,
    )
with col_gstr:
    gstr_mode = st.radio(
        "GSTR-2A / 2B",
        options=["Single file", "Multiple periods (Consolidate)"],
        horizontal=True,
    )

sales_pr = service_pr = pr_file = None
gstr_file = None
gstr_period_files: list = []

st.subheader("Purchase Register")
if pr_mode == "Sales + Service (Consolidate)":
    c1, c2 = st.columns(2)
    with c1:
        sales_pr = st.file_uploader("Sales Purchase Register", type=["xlsx", "xls"], key="sales_pr")
    with c2:
        service_pr = st.file_uploader("Service Purchase Register", type=["xlsx", "xls"], key="service_pr")
else:
    st.caption(
        "Upload **Vyapar Purchase Report** Excel (must include **Purchase Items** sheet). "
        "Reports → Purchase Report → Excel export. "
        "The app converts Purchase Items into sample Purchase Register format automatically."
    )
    pr_file = st.file_uploader("Purchase Register (Excel)", type=["xlsx", "xls"], key="pr")

st.subheader("GSTR-2A / 2B")
if gstr_mode == "Multiple periods (Consolidate)":
    st.caption(
        "Upload one GSTR-2B Excel per period from the **GST portal** "
        "(e.g. 092025_GSTIN_GSTR2BQ_....xlsx). "
        "Do not upload Payment Register or Tally/Busy exports here."
    )
    gstr_period_files = st.file_uploader(
        "GSTR-2A/2B files (one per period)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="gstr_periods",
    ) or []
else:
    st.caption(
        "Upload the **GSTR-2B Excel downloaded from GST Portal** "
        "(e.g. `09GSTIN_062026_GSTR2BQ_08072026.xlsx` with B2B/CDNR sheets). "
        "Do **not** upload Payment Register, Tally payment exports, or files named only with a date."
    )
    gstr_file = st.file_uploader("GSTR-2A / 2B (Excel)", type=["xlsx", "xls"], key="gstr")

pr_ready = (sales_pr and service_pr) if pr_mode == "Sales + Service (Consolidate)" else bool(pr_file)
gstr_ready = len(gstr_period_files) >= 1 if gstr_mode == "Multiple periods (Consolidate)" else bool(gstr_file)
files_ready = pr_ready and gstr_ready


def _build_file_key() -> str:
    parts = [pr_mode, gstr_mode]
    if pr_mode == "Sales + Service (Consolidate)":
        parts.extend([sales_pr.name, str(sales_pr.size), service_pr.name, str(service_pr.size)])
    else:
        parts.extend([pr_file.name, str(pr_file.size)])
    if gstr_mode == "Multiple periods (Consolidate)":
        for uploaded in gstr_period_files:
            parts.extend([uploaded.name, str(uploaded.size)])
    else:
        parts.extend([gstr_file.name, str(gstr_file.size)])
    return "|".join(parts)


def _build_sources() -> tuple[list, list]:
    if pr_mode == "Sales + Service (Consolidate)":
        purchase_sources = [(sales_pr, "Sales"), (service_pr, "Service")]
    else:
        purchase_sources = [(pr_file, "Purchase Register")]

    if gstr_mode == "Multiple periods (Consolidate)":
        gstr_sources = [(f, label_from_filename(f.name)) for f in gstr_period_files]
    else:
        gstr_sources = [(gstr_file, "GSTR")]

    return purchase_sources, gstr_sources


if files_ready:
    file_key = _build_file_key()

    if st.session_state.get("last_file_key") != file_key:
        with st.spinner("Consolidating registers and matching invoices..."):
            try:
                purchase_sources, gstr_sources = _build_sources()
                for source, _ in purchase_sources + gstr_sources:
                    source.seek(0)

                consolidated_pr, consolidated_gstr, result, summary, dashboard = load_and_match_with_consolidation(
                    purchase_sources,
                    gstr_sources,
                )
                st.session_state["consolidated_pr"] = consolidated_pr
                st.session_state["consolidated_gstr"] = consolidated_gstr
                st.session_state["match_result"] = result
                st.session_state["match_summary"] = summary
                st.session_state["itc_dashboard"] = dashboard
                st.session_state["last_file_key"] = file_key
            except Exception as exc:
                for key in (
                    "match_result",
                    "match_summary",
                    "itc_dashboard",
                    "consolidated_pr",
                    "consolidated_gstr",
                    "last_file_key",
                ):
                    st.session_state.pop(key, None)
                st.error(f"Could not process files: {exc}")

    if "match_result" in st.session_state:
        result: pd.DataFrame = st.session_state["match_result"]
        summary = st.session_state["match_summary"]
        dashboard = st.session_state["itc_dashboard"]

        st.success("Done! Review the ITC dashboard below or download reports.")

        st.subheader("ITC Dashboard")
        d1, d2, d3, d4 = st.columns(4)
        d1.metric("Available ITC (GSTR-2B)", f"₹{dashboard.available_total:,.2f}")
        d2.metric(
            "Eligible ITC",
            f"₹{dashboard.eligible_total:,.2f}",
            help="Safe to claim in GSTR-3B after verification",
        )
        d3.metric("Pending Review", f"₹{dashboard.pending_total:,.2f}")
        d4.metric("Non-Eligible ITC", f"₹{dashboard.non_eligible_total:,.2f}")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**ITC balance overview**")
            balance_df = pd.DataFrame(
                {
                    "Amount": [
                        dashboard.eligible_total,
                        dashboard.pending_total,
                        dashboard.non_eligible_total,
                    ]
                },
                index=["Eligible", "Pending Review", "Non-Eligible"],
            )
            st.bar_chart(balance_df, height=280)
        with c2:
            st.markdown("**Available vs eligible by tax head**")
            tax_df = pd.DataFrame(
                {
                    "Available (GSTR-2B)": [
                        dashboard.available_igst,
                        dashboard.available_cgst,
                        dashboard.available_sgst,
                    ],
                    "Eligible ITC": [
                        dashboard.eligible_igst,
                        dashboard.eligible_cgst,
                        dashboard.eligible_sgst,
                    ],
                },
                index=["IGST", "CGST", "SGST"],
            )
            st.bar_chart(tax_df, height=280)

        inv1, inv2, inv3, inv4 = st.columns(4)
        inv1.metric("Eligible invoices", dashboard.eligible_invoices)
        inv2.metric("Pending invoices", dashboard.pending_invoices)
        inv3.metric("Non-eligible invoices", dashboard.non_eligible_invoices)
        inv4.metric("Claim rate", f"{dashboard.claim_rate_pct}%")

        claim_plan = build_itc_claim_plan(dashboard, result)
        optimization_insights = generate_optimization_insights(dashboard, claim_plan, result)
        action_plan = generate_action_plan(dashboard, result)

        st.subheader("ITC Claim & Planning")
        readiness = claim_plan.filing_readiness
        if readiness == "Ready to File":
            st.success(f"**{readiness}** — {claim_plan.filing_advice}")
        elif readiness in ("File with Caution", "Partial Claim Only"):
            st.warning(f"**{readiness}** — {claim_plan.filing_advice}")
        else:
            st.error(f"**{readiness}** — {claim_plan.filing_advice}")

        p1, p2, p3, p4 = st.columns(4)
        p1.metric("Claim Now (GSTR-3B)", f"₹{claim_plan.claim_now_total:,.2f}")
        p2.metric("Hold (Pending)", f"₹{claim_plan.hold_itc:,.2f}")
        p3.metric("Blocked", f"₹{claim_plan.blocked_itc:,.2f}")
        p4.metric(
            "Optimized Potential",
            f"₹{claim_plan.optimized_claim_total:,.2f}",
            delta=f"+₹{claim_plan.optimization_gain:,.2f}" if claim_plan.optimization_gain > 0 else None,
        )

        t1, t2 = st.columns(2)
        with t1:
            st.markdown("**GSTR-3B Table 4 — Claim amounts**")
            table4_df = pd.DataFrame(
                {
                    "Claim Now": [
                        f"₹{claim_plan.claim_now_igst:,.2f}",
                        f"₹{claim_plan.claim_now_cgst:,.2f}",
                        f"₹{claim_plan.claim_now_sgst:,.2f}",
                    ],
                    "Hold (Pending)": [
                        f"₹{dashboard.pending_igst:,.2f}",
                        f"₹{dashboard.pending_cgst:,.2f}",
                        f"₹{dashboard.pending_sgst:,.2f}",
                    ],
                },
                index=["IGST", "CGST", "SGST"],
            )
            st.dataframe(table4_df, use_container_width=True)
        with t2:
            st.markdown("**Current vs optimized claim**")
            claim_compare_df = pd.DataFrame(
                {
                    "Amount": [
                        claim_plan.claim_now_total,
                        claim_plan.optimized_claim_total,
                    ]
                },
                index=["Current Claim", "Optimized Potential"],
            )
            st.bar_chart(claim_compare_df, height=280)

        st.markdown("**Optimization insights**")
        for insight in optimization_insights:
            clean = insight.replace("**", "")
            st.info(clean)

        st.markdown("**Action plan**")
        st.dataframe(
            action_plan_dataframe(action_plan),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("AI Recommendations")
        for tip in generate_ai_recommendations(dashboard, result):
            st.info(tip)

        st.markdown("**Download ITC Results**")
        r1, r2, r3 = st.columns(3)
        with r1:
            st.download_button(
                label="Download Excel (.xlsx)",
                data=export_to_excel(result, summary, dashboard),
                file_name="itc_matching_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        with r2:
            st.download_button(
                label="Download CSV (.csv)",
                data=export_to_csv(result, summary, dashboard),
                file_name="itc_matching_report.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with r3:
            st.download_button(
                label="Download PDF (.pdf)",
                data=export_to_pdf(result, summary, dashboard),
                file_name="itc_matching_report.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        st.markdown("**Download Converted Registers**")
        d2, d3 = st.columns(2)
        with d2:
            if st.session_state.get("consolidated_pr") is not None:
                st.download_button(
                    label="Download Purchase Register (Sample Format)",
                    data=export_consolidated_purchase_register(st.session_state["consolidated_pr"]),
                    file_name="purchase_register_sample.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        with d3:
            if st.session_state.get("consolidated_gstr") is not None:
                st.download_button(
                    label="Download GSTR-2B (Sample Format)",
                    data=export_consolidated_gstr(st.session_state["consolidated_gstr"]),
                    file_name="consolidated_gstr2b.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        consolidated_pr = st.session_state.get("consolidated_pr")
        consolidated_gstr = st.session_state.get("consolidated_gstr")
        if consolidated_pr is not None or consolidated_gstr is not None:
            c1, c2 = st.columns(2)
            with c1:
                if consolidated_pr is not None:
                    st.subheader("Purchase Register (Sample Format)")
                    caption = f"{len(consolidated_pr)} invoices converted from your upload"
                    if (
                        "register_type" in consolidated_pr.columns
                        and consolidated_pr["register_type"].nunique() > 1
                    ):
                        sales_count = len(consolidated_pr[consolidated_pr["register_type"] == "Sales"])
                        service_count = len(consolidated_pr[consolidated_pr["register_type"] == "Service"])
                        caption = f"{sales_count} Sales + {service_count} Service invoices"
                    st.caption(caption)
                    st.dataframe(consolidated_pr_to_display(consolidated_pr), use_container_width=True, height=220)
            with c2:
                if consolidated_gstr is not None:
                    st.subheader("GSTR-2B (Sample Format)")
                    st.caption(
                        f"{len(consolidated_gstr)} invoices converted from GST portal format"
                        + (
                            f" — {gstr_summary_caption(consolidated_gstr)}"
                            if "gstr_source" in consolidated_gstr.columns
                            and consolidated_gstr["gstr_source"].nunique() > 1
                            else ""
                        )
                    )
                    st.dataframe(consolidated_gstr_to_display(consolidated_gstr), use_container_width=True, height=220)

        st.subheader("Matching Results")
        tab_all, tab_eligible, tab_pending, tab_blocked = st.tabs(
            ["All", "Eligible", "Pending Review", "Non-Eligible"]
        )
        with tab_all:
            st.dataframe(result, use_container_width=True, height=400)
        with tab_eligible:
            st.dataframe(
                result[result["ITC Category"] == ITC_ELIGIBLE],
                use_container_width=True,
                height=400,
            )
        with tab_pending:
            st.dataframe(
                result[result["ITC Category"] == ITC_PENDING],
                use_container_width=True,
                height=400,
            )
        with tab_blocked:
            st.dataframe(
                result[result["ITC Category"] == ITC_NON_ELIGIBLE],
                use_container_width=True,
                height=400,
            )

else:
    hints = []
    if not pr_ready:
        hints.append("Purchase Register file(s)")
    if not gstr_ready:
        if gstr_mode == "Multiple periods (Consolidate)":
            hints.append("at least one GSTR-2A/2B file per period")
        else:
            hints.append("GSTR-2A/2B file")
    st.info(f"Upload {' and '.join(hints)} to begin.")

st.divider()
st.caption(
    "Columns are auto-detected from Tally, Busy, SAP, and GST portal exports. "
    "ITC decisions are indicative — verify under GST rules before claiming."
)
