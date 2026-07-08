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
    from matcher.engine import export_to_excel, load_and_match_with_consolidation
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

                consolidated_pr, consolidated_gstr, result, summary = load_and_match_with_consolidation(
                    purchase_sources,
                    gstr_sources,
                )
                st.session_state["consolidated_pr"] = consolidated_pr
                st.session_state["consolidated_gstr"] = consolidated_gstr
                st.session_state["match_result"] = result
                st.session_state["match_summary"] = summary
                st.session_state["last_file_key"] = file_key
            except Exception as exc:
                for key in ("match_result", "match_summary", "consolidated_pr", "consolidated_gstr", "last_file_key"):
                    st.session_state.pop(key, None)
                st.error(f"Could not process files: {exc}")

    if "match_result" in st.session_state:
        result: pd.DataFrame = st.session_state["match_result"]
        summary = st.session_state["match_summary"]

        st.success("Done! Review results below or download the reports.")

        d1, d2, d3 = st.columns(3)
        with d1:
            st.download_button(
                label="Download ITC Taken Excel",
                data=export_to_excel(result, summary),
                file_name="itc_matching_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        with d2:
            if st.session_state.get("consolidated_pr") is not None:
                st.download_button(
                    label="Download Consolidated Purchase Register",
                    data=export_consolidated_purchase_register(st.session_state["consolidated_pr"]),
                    file_name="consolidated_purchase_register.xlsx",
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
                    st.subheader("Consolidated Purchase Register")
                    st.caption(
                        f"{len(consolidated_pr[consolidated_pr['register_type'] == 'Sales'])} Sales + "
                        f"{len(consolidated_pr[consolidated_pr['register_type'] == 'Service'])} Service invoices"
                    )
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

        st.subheader("Summary")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Fully Matched", summary.fully_matched)
        m2.metric("Tax Mismatch", summary.tax_mismatch)
        m3.metric("Not Matched", summary.not_matched)
        m4.metric("Duplicates", summary.duplicate)
        m5.metric("Total ITC Taken", f"₹{summary.total_itc:,.2f}")

        m6, m7, m8, m9 = st.columns(4)
        m6.metric("GSTIN Mismatch", summary.gstin_mismatch)
        m7.metric("Inv No Mismatch", summary.inv_no_mismatch)
        m8.metric("Inv Date Mismatch", summary.inv_date_mismatch)
        m9.metric("Missing in PR", summary.missing_in_pr)

        itc1, itc2, itc3 = st.columns(3)
        itc1.metric("Eligible IGST", f"₹{summary.total_itc_igst:,.2f}")
        itc2.metric("Eligible CGST", f"₹{summary.total_itc_cgst:,.2f}")
        itc3.metric("Eligible SGST", f"₹{summary.total_itc_sgst:,.2f}")

        st.subheader("Matching Results")
        st.dataframe(result, use_container_width=True, height=450)

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
