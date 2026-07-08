"""Streamlit app for GST ITC invoice matching — upload files, get results."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from matcher.consolidate import (
    consolidated_gstr_to_display,
    consolidated_pr_to_display,
    export_consolidated_gstr,
    export_consolidated_purchase_register,
)
from matcher.engine import (
    export_to_excel,
    load_and_match,
    load_and_match_fully_consolidated,
)

st.set_page_config(page_title="GST ITC Matcher", page_icon="📊", layout="wide")

st.title("GST ITC Matcher")
st.caption("Upload Purchase Register(s) and GSTR-2A/2B — consolidation and matching run automatically.")

pr_mode = st.radio(
    "Register type",
    options=["Single file", "Sales + Service (Consolidate)"],
    horizontal=True,
)

sales_pr = service_pr = pr_file = None
sales_gstr = service_gstr = gstr_file = None

if pr_mode == "Sales + Service (Consolidate)":
    st.info("Upload separate Sales and Service files for both Purchase Register and GSTR-2A/2B.")

    st.subheader("Purchase Register")
    col1, col2 = st.columns(2)
    with col1:
        sales_pr = st.file_uploader("Sales Purchase Register", type=["xlsx", "xls"], key="sales_pr")
    with col2:
        service_pr = st.file_uploader("Service Purchase Register", type=["xlsx", "xls"], key="service_pr")

    st.subheader("GSTR-2A / 2B")
    col3, col4 = st.columns(2)
    with col3:
        sales_gstr = st.file_uploader("Sales GSTR-2A/2B", type=["xlsx", "xls"], key="sales_gstr")
    with col4:
        service_gstr = st.file_uploader("Service GSTR-2A/2B", type=["xlsx", "xls"], key="service_gstr")
else:
    col1, col2 = st.columns(2)
    with col1:
        pr_file = st.file_uploader("1. Purchase Register (Excel)", type=["xlsx", "xls"], key="pr")
    with col2:
        gstr_file = st.file_uploader("2. GSTR-2A / GSTR-2B (Excel)", type=["xlsx", "xls"], key="gstr")

if pr_mode == "Sales + Service (Consolidate)":
    files_ready = sales_pr and service_pr and sales_gstr and service_gstr
else:
    files_ready = pr_file and gstr_file

if files_ready:
    if pr_mode == "Sales + Service (Consolidate)":
        file_key = (
            f"{sales_pr.name}|{sales_pr.size}|{service_pr.name}|{service_pr.size}|"
            f"{sales_gstr.name}|{sales_gstr.size}|{service_gstr.name}|{service_gstr.size}|{pr_mode}"
        )
    else:
        file_key = f"{pr_file.name}|{pr_file.size}|{gstr_file.name}|{gstr_file.size}|{pr_mode}"

    if st.session_state.get("last_file_key") != file_key:
        with st.spinner("Consolidating registers and matching invoices..."):
            try:
                if pr_mode == "Sales + Service (Consolidate)":
                    for f in (sales_pr, service_pr, sales_gstr, service_gstr):
                        f.seek(0)
                    consolidated_pr, consolidated_gstr, result, summary = load_and_match_fully_consolidated(
                        [(sales_pr, "Sales"), (service_pr, "Service")],
                        [(sales_gstr, "Sales"), (service_gstr, "Service")],
                    )
                    st.session_state["consolidated_pr"] = consolidated_pr
                    st.session_state["consolidated_gstr"] = consolidated_gstr
                else:
                    pr_file.seek(0)
                    gstr_file.seek(0)
                    result, summary = load_and_match(pr_file, gstr_file)
                    st.session_state.pop("consolidated_pr", None)
                    st.session_state.pop("consolidated_gstr", None)

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
            if "consolidated_pr" in st.session_state:
                st.download_button(
                    label="Download Consolidated Purchase Register",
                    data=export_consolidated_purchase_register(st.session_state["consolidated_pr"]),
                    file_name="consolidated_purchase_register.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        with d3:
            if "consolidated_gstr" in st.session_state:
                st.download_button(
                    label="Download Consolidated GSTR-2A/2B",
                    data=export_consolidated_gstr(st.session_state["consolidated_gstr"]),
                    file_name="consolidated_gstr2a_2b.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        if "consolidated_pr" in st.session_state:
            consolidated_pr = st.session_state["consolidated_pr"]
            consolidated_gstr = st.session_state["consolidated_gstr"]

            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Consolidated Purchase Register")
                st.caption(
                    f"{len(consolidated_pr[consolidated_pr['register_type'] == 'Sales'])} Sales + "
                    f"{len(consolidated_pr[consolidated_pr['register_type'] == 'Service'])} Service invoices"
                )
                st.dataframe(consolidated_pr_to_display(consolidated_pr), use_container_width=True, height=220)
            with c2:
                st.subheader("Consolidated GSTR-2A/2B")
                st.caption(
                    f"{len(consolidated_gstr[consolidated_gstr['gstr_type'] == 'Sales'])} Sales + "
                    f"{len(consolidated_gstr[consolidated_gstr['gstr_type'] == 'Service'])} Service invoices"
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
    if pr_mode == "Sales + Service (Consolidate)":
        st.info(
            "Upload Sales + Service Purchase Registers and Sales + Service GSTR-2A/2B files. "
            "Both will be consolidated automatically."
        )
    else:
        st.info("Upload Purchase Register and GSTR-2A/2B Excel files.")

st.divider()
st.caption(
    "Columns are auto-detected from Tally, Busy, SAP, and GST portal exports. "
    "ITC decisions are indicative — verify under GST rules before claiming."
)
