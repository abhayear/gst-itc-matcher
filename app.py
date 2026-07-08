"""Streamlit app for GST ITC invoice matching — upload files, get results."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from matcher.consolidate import (
    consolidated_to_display,
    export_consolidated_purchase_register,
)
from matcher.engine import export_to_excel, load_and_match, load_and_match_consolidated

st.set_page_config(page_title="GST ITC Matcher", page_icon="📊", layout="wide")

st.title("GST ITC Matcher")
st.caption("Upload Purchase Register(s) and GSTR-2A/2B — consolidation and matching run automatically.")

pr_mode = st.radio(
    "Purchase Register type",
    options=["Single file", "Sales + Service (Consolidate)"],
    horizontal=True,
)

sales_file = None
service_file = None
pr_file = None

if pr_mode == "Sales + Service (Consolidate)":
    st.info("Upload separate Sales and Service Purchase Registers. They will be merged automatically.")
    col_sales, col_service = st.columns(2)
    with col_sales:
        sales_file = st.file_uploader(
            "1a. Sales Purchase Register (Excel)",
            type=["xlsx", "xls"],
            key="sales_pr",
        )
    with col_service:
        service_file = st.file_uploader(
            "1b. Service Purchase Register (Excel)",
            type=["xlsx", "xls"],
            key="service_pr",
        )
else:
    pr_file = st.file_uploader(
        "1. Purchase Register (Excel)",
        type=["xlsx", "xls"],
        key="pr",
    )

gstr_file = st.file_uploader(
    "2. GSTR-2A / GSTR-2B (Excel)",
    type=["xlsx", "xls"],
    key="gstr",
)

pr_ready = (sales_file and service_file) if pr_mode == "Sales + Service (Consolidate)" else bool(pr_file)

if pr_ready and gstr_file:
    if pr_mode == "Sales + Service (Consolidate)":
        file_key = (
            f"{sales_file.name}|{sales_file.size}|"
            f"{service_file.name}|{service_file.size}|"
            f"{gstr_file.name}|{gstr_file.size}|{pr_mode}"
        )
    else:
        file_key = f"{pr_file.name}|{pr_file.size}|{gstr_file.name}|{gstr_file.size}|{pr_mode}"

    if st.session_state.get("last_file_key") != file_key:
        with st.spinner("Consolidating and matching invoices automatically..."):
            try:
                gstr_file.seek(0)
                if pr_mode == "Sales + Service (Consolidate)":
                    sales_file.seek(0)
                    service_file.seek(0)
                    consolidated, result, summary = load_and_match_consolidated(
                        [
                            (sales_file, "Sales"),
                            (service_file, "Service"),
                        ],
                        gstr_file,
                    )
                    st.session_state["consolidated_pr"] = consolidated
                else:
                    pr_file.seek(0)
                    result, summary = load_and_match(pr_file, gstr_file)
                    st.session_state.pop("consolidated_pr", None)

                st.session_state["match_result"] = result
                st.session_state["match_summary"] = summary
                st.session_state["last_file_key"] = file_key
            except Exception as exc:
                st.session_state.pop("match_result", None)
                st.session_state.pop("match_summary", None)
                st.session_state.pop("consolidated_pr", None)
                st.error(f"Could not process files: {exc}")

    if "match_result" in st.session_state:
        result: pd.DataFrame = st.session_state["match_result"]
        summary = st.session_state["match_summary"]

        st.success("Done! Review results below or download the reports.")

        download_col1, download_col2 = st.columns(2)
        with download_col1:
            st.download_button(
                label="Download ITC Taken Excel",
                data=export_to_excel(result, summary),
                file_name="itc_matching_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        with download_col2:
            if "consolidated_pr" in st.session_state:
                consolidated: pd.DataFrame = st.session_state["consolidated_pr"]
                st.download_button(
                    label="Download Consolidated Purchase Register",
                    data=export_consolidated_purchase_register(consolidated),
                    file_name="consolidated_purchase_register.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        if "consolidated_pr" in st.session_state:
            consolidated = st.session_state["consolidated_pr"]
            st.subheader("Consolidated Purchase Register")
            st.caption(
                f"Merged {len(consolidated[consolidated['register_type'] == 'Sales'])} Sales "
                f"+ {len(consolidated[consolidated['register_type'] == 'Service'])} Service invoices"
            )
            st.dataframe(
                consolidated_to_display(consolidated),
                use_container_width=True,
                height=250,
            )

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
    if pr_mode == "Sales + Service (Consolidate)":
        st.info("Upload Sales Purchase Register, Service Purchase Register, and GSTR-2A/2B Excel files.")
    else:
        st.info("Upload Purchase Register and GSTR-2A/2B Excel files. Matching will start automatically.")

st.divider()
st.caption(
    "Columns are auto-detected from Tally, Busy, SAP, and GST portal exports. "
    "ITC decisions are indicative — verify under GST rules before claiming."
)
